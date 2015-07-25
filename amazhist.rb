#!/usr/bin/env ruby
# -*- coding: utf-8 -*-
# Amazhist written by KIMATA Tetsuya <kimata@green-rabbit.net>

# Amazon の全購入履歴を JSON 形式で取得するスクリプトです．
#
# ■ できること
# 商品について，以下の情報が取得できます．
# - 商品名
# - 購入日
# - 数量
# - 購入価格
# - 売り手
# - 商品 URL
# - 商品画像 (実行フォルダに img フォルダを作成して中に保存)
#
# ■準備
#   このスクリプトでは次のライブラリを使っていますので，入っていない場合は
#   インストールしておいてください．
#   - Term::ANSIcolor
#   - Nokogiri
#   - Mechanize
#
# ■使い方
# 1. 次の環境変数に，Amazon の ID とパスワードを設定．
#    - amazon_id
#    - amazon_pass
#
# 2. スクリプトを実行
#    $ ./amazhist.rb -j amazhist.json -t img
#    引数の意味は以下
#    - j 履歴情報を保存する JSON ファイルのパス
#    - t サムネイル画像を保存するディレクトリのパス
#
# ■トラブルシュート
# 何度も実行すると，Amazon から迷惑がられて，画像認証をパスしないと
# アクセスできなくなります．
# 「画像認証を要求されたのでリトライします．」と表示された場合は，
# しばらく時間を空けてください．

require 'date'
require 'json'
require 'mechanize'
require 'nokogiri'
require 'optparse'
require 'pathname'
require 'term/ansicolor'
require 'uri'

# DEBUG = 1

class Color
  extend Term::ANSIColor
end

class Amazhist
  AMAZON_URL      = "http://www.amazon.co.jp/"
  HIST_URL_FORMAT = "https://www.amazon.co.jp/gp/css/order-history?" +
    "digitalOrders=1&unifiedOrders=1&orderFilter=year-%d&startIndex=%d"
  # NOTE: 下記アドレスの「?」以降を省略すると，時々ページの表示内容が変わり，
  # カテゴリを取得できなくなる
  ITEM_URL_FORMAT = "http://www.amazon.co.jp/gp/product/%s?*Version*=1&*entries*=0"
  CATEGORY_RETRY  = 5
  RETRY_WAIT_SEC  = 5

  def initialize(user_info, img_dir_path)
    @mech = Mechanize.new
    @mech.user_agent_alias = "Windows Chrome"
    @user_info = user_info
    @img_dir_path = Pathname.new(img_dir_path)
  end

  def self.error(message)
    STDERR.puts
    STDERR.puts "[%s] %s" % [ Color.bold(Color.red("ERROR")), message ]
    exit
  end

  def self.warn(message)
    STDERR.puts
    STDERR.puts "[%s] %s" % [ Color.bold(Color.yellow("WARN")), message ]
  end

  def hist_url(year, page)
    return HIST_URL_FORMAT % [ year, 10 * (page-1) ]
  end

  def login(page)
    page.form_with(name: "signIn") do |form|
      form.field_with(name: "email").value = @user_info[:id]
      form.field_with(name: "password").value = @user_info[:pass]
    end
    return page.form_with(name: "signIn").submit
  end

  def get_item_category(item_id)
    (0...CATEGORY_RETRY).each do
      begin
        page = @mech.get(ITEM_URL_FORMAT % [ item_id ])
        html = Nokogiri::HTML(page.body.toutf8, 'UTF-8')
        crumb = html.css("div.a-breadcrumb li")

        if (crumb.size == 0) then
          sleep(RETRY_WAIT_SEC)
          next
        end        

        return {
          category: crumb[0].text.strip,
          subcategory: crumb[2].text.strip,
        }
      rescue => e
        STDERR.puts(e.backtrace)
        sleep(RETRY_WAIT_SEC)
      end
    end

    warn("category is NOT determined: %s" % [ item_id ])    

    return {
      category: "",
      subcategory: "",
    }
  end

  def parse_item_page(html, item_list)
    # NOTE: for development
    if (defined? DEBUG) then
      f = File.open("debug.htm")
      html = Nokogiri::HTML(f)
      f.close
    end
    html.css("div.order").each do |order|
      begin
        date_text = order.css("div.order-info span.value")[0].text.strip
        date = Date.strptime(date_text, "%Y年%m月%d日")

        order.css("div.a-fixed-left-grid").each do |item|
          name = item.css("div.a-row")[0].text.strip
          url = URI.join(AMAZON_URL, item.css("div.a-row")[0].css("a")[0][:href]).to_s
          id = %r|/gp/product/([^/]+)/|.match(url)[1]
          count = 1
          if (%r|^商品名：(.+)、数量：(\d+)|.match(name)) then
            name = $1
            count = $2.to_i
          end
          price_str = item.css("div.a-row span.a-color-price").text.gsub(/¥|,/, "").strip
          price = %r|\d+|.match(price_str)[0].to_i

          seller = ""
          (1..2).each do |i| 
            seller_cand = item.css("div.a-row")[i].text
            next if (!%r|販売:|.match(seller_cand))

            if (item.css("div.a-row")[i].css("a")[0] != nil) then
              seller = item.css("div.a-row")[i].css("a")[0].text.strip
            else
              seller = item.css("div.a-row")[i].text.gsub("販売:", "").strip
            end
            break
          end

          img_url = item.css("div.item-view-left-col-inner img")[0][:src]
          img_file_name = "%s.%s" % [ id, %r|\.(\w+)$|.match(img_url)[1] ]
          @mech.get(img_url).save_as(@img_dir_path + img_file_name)

          category_info = get_item_category(id)

          item_list.push({
                           name: name,
                           id: id,
                           url: url,
                           count: count,
                           price: price,
                           category: category_info[:category],
                           subcategory: category_info[:subcategory],
                           seller: seller,
                           date: date
                         })
        end
        STDERR.print "."
        STDERR.flush
      rescue => e
        warn(e.message)    
      end
    end

    return html.css("div.pagination-full li.a-last").css("a").empty?
  end
  
  def get_item_list_by_page(year, page, item_list)
    # NOTE: for development
    if (defined? DEBUG) then
      parse_item_page(nil, item_list)
      p item_list
      exit
    end

    page = @mech.get(hist_url(year, page))
    2.times do |i|
      # html = Nokogiri::HTML(page.body.toutf8)
      if %r|サインイン|.match(page.title) then
        html = Nokogiri::HTML(page.body.toutf8, 'UTF-8')
        if !%r|画像に表示されている文字|.match(html.css("#ap_captcha_title").text) then
          error("ID もしくはパスワードが異なります．")
        end
        # 2回目以降は少し待つ
        if (i != 0) then
          sleep_time = 300
          warn("画像認証を要求されたので %d 分後にリトライします．" % [ sleep_time / 60 ])
          sleep(sleep_time) if (i != 0) 
        end
        page = login(page)
        next
      end
      return parse_item_page(page, item_list)
    end
    raise StandardError.new("ログインに失敗しました．")
  end

  def get_item_list(year)
    item_list = []
    @mech.get("http://www.amazon.co.jp/")
    page = 1
    loop do
      STDERR.print "%s Year %d page %d " % [ Color.bold(Color.green("Parsing")), 
                                           year, page ]
      STDERR.flush
      is_last = get_item_list_by_page(year, page, item_list)
      STDERR.puts
      break if is_last
      page += 1
      sleep 5
    end
    return item_list
  end
end

params = ARGV.getopts("j:t:")
if (params["j"] == nil) then
  Amazhist.error("履歴情報を保存するファイルのパスが指定されていません．" + 
                 "(-j で指定します)")
  exit
end
if (params["t"] == nil) then
  Amazhist.error("サムネイル画像を保存するディレクトリのパスが指定されていません．" + 
                 "(-t で指定します)")
  exit
end

json_file_path = params["j"]
img_dir_path = params["t"]

FileUtils.mkdir_p(img_dir_path)
amazhist = Amazhist.new({
                          id: ENV["amazon_id"], 	# Amazon の ID
                          pass: ENV["amazon_pass"],	# Amazon の パスワード
                        },
                        img_dir_path)

item_list = []
(2000..(Date.today.year)).each do |year|
  item_list.concat(amazhist.get_item_list(year))
end

File.open(json_file_path, "w") do |file|
  file.puts JSON.generate(item_list)
end

STDERR.puts Color.bold(Color.blue("Writing output file")) 

# Local Variables:
# coding: utf-8
# mode: ruby
# tab-width: 4
# indent-tabs-mode: nil
# End:

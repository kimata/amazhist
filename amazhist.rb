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

require "date"
require "json"
require "mechanize"
require "nokogiri"
require "optparse"
require "pathname"
require "term/ansicolor"
require "uri"
require "logger"
require "open-uri"
require "docopt"

DOCOPT = <<DOCOPT
Usage: amazhist.py [-h] [-j <json_path>] [-t <img_path>] [-o <excel_path>]

Amazon の購入履歴データを収集します．

Options:
  -j <json_path>    履歴情報存を記録する JSON ファイルのパス  (required) [default: amazhist.json]
  -t <img_path>     サムネイル画像が保存するディレクトリのパス (required) [default: img]
DOCOPT

# TRACE = 1 # 定義すると，取得した Web ページをデバッグ用に保存します
# DEBUG = 1 # 定義すると，デバッグ用に保存したファイルからページを読み込みます

if (ENV.has_key?("OCRA_EXECUTABLE"))
  ENV["SSL_CERT_FILE"] = File.join(File.dirname($0), "cert.pem")
end

class Color
  extend Term::ANSIColor
end

class Amazhist
  AMAZON_URL = "https://www.amazon.co.jp/"

  HIST_URL_FORMAT = "https://www.amazon.co.jp/gp/css/order-history?" +
                    "digitalOrders=1&unifiedOrders=1&orderFilter=year-%d&startIndex=%d"
  # NOTE: 下記アドレスの「?」以降を省略すると，時々ページの表示内容が変わり，
  # カテゴリを取得できなくなる
  ITEM_URL_FORMAT = "https://www.amazon.co.jp/gp/product/%s?*Version*=1&*entries*=0"
  RETRY_COUNT = 5
  RETRY_WAIT_SEC = 5
  YEAR_START = 2000
  COOKIE_DUMP = "cookie.txt"
  ITEM_LIST_DUMP = "item_list.cache"
  MECH_LOG_FILE = "mechanize.log"

  def initialize(user_info, img_dir_path)
    @mech = Mechanize.new
    @mech.user_agent_alias = "Windows Chrome"
    @mech.redirection_limit = 4
    @mech.cookie_jar.clear!
    if (defined? TRACE)
      @mech.log = Logger.new(MECH_LOG_FILE)
    end
    @user_info = user_info
    @img_dir_path = Pathname.new(img_dir_path)

    cookie_load()
    web_page = @mech.get(AMAZON_URL)
    sleep(RETRY_WAIT_SEC)
  end

  def self.error(message, is_nl = true)
    STDERR.puts if (is_nl)
    STDERR.puts "[%s] %s" % [Color.bold(Color.red("ERROR")), message]
    exit
  end

  def self.warn(message, is_nl = true)
    STDERR.puts if (is_nl)
    STDERR.puts "[%s] %s" % [Color.bold(Color.yellow("WARN")), message]
  end

  def cookie_save()
    @mech.cookie_jar.save_as(COOKIE_DUMP)
  end

  def cookie_load()
    if File.exist?(COOKIE_DUMP)
      @mech.cookie_jar.load(COOKIE_DUMP)
    end
  end

  def item_list_store(item_list, year, page)
    File.open(ITEM_LIST_DUMP, "w+b") do |dump|
      dump.write(Marshal.dump({ item_list: item_list, year: year, page: page }))
    end
  end

  def item_list_load()
    begin
      File.open(ITEM_LIST_DUMP, "r+b") do |dump|
        return Marshal.load(dump)
      end
    rescue
      return { item_list: [], year: YEAR_START, page: 0 }
    end
  end

  def hist_url(year, page)
    return HIST_URL_FORMAT % [year, 10 * (page - 1)]
  end

  def login(web_page)
    6.times do |i|
      if !%r|Amazonサインイン|.match(web_page.title)
        cookie_save()
        return web_page
      end

      if (defined? TRACE)
        File.open("debug_login_page_#{i}.htm", "w") do |file|
          file.puts(web_page.body.toutf8)
        end
      end

      sleep(1)

      html = Nokogiri::HTML(web_page.body.toutf8, "UTF-8")
      if (i == 0)
        web_page.form_with(name: "signIn") do |form|
          form.field_with(name: "email").value = @user_info[:id]
        end
        web_page = web_page.form_with(name: "signIn").submit
      elsif (i == 1)
        web_page.form_with(name: "signIn") do |form|
          form.field_with(name: "password").value = @user_info[:pass]
          form.checkbox_with(:name => "rememberMe").check
        end
        web_page = web_page.form_with(name: "signIn").submit
      elsif (i >= 2)
        begin
          captcha_url = web_page.search("#auth-captcha-image").attribute("src")
          File.open("captcha.png", "w+b") do |img|
            URI.open(captcha_url) do |data|
              img.write(data.read)
            end
          end
          captcha = request_input("\ncaptcha.png に書かれている文字列を入力してください．")
          web_page.form_with(name: "signIn") do |form|
            form.field_with(name: "password").value = @user_info[:pass]
            form.field_with(name: "guess").value = captcha
            form.checkbox_with(:name => "rememberMe").check
          end
          web_page = web_page.form_with(name: "signIn").submit
        rescue => e
          STDERR.puts(e.message)
          STDERR.puts(e.backtrace)
          error("不明なエラーです．")
        end
      else
        break
      end
    end
    error("ログインに失敗しました．")
  end

  def fetch_html(url, file_path)
    if (defined? DEBUG)
      File.open(file_path, "r") do |file|
        return Nokogiri::HTML(file)
      end
    end

    web_page = @mech.get(url)
    web_page = login(web_page)

    if (defined? TRACE)
      File.open(file_path, "w") do |file|
        file.puts(web_page.body.toutf8)
      end
    end

    return Nokogiri::HTML(web_page.body.toutf8, "UTF-8")
  end

  def get_item_category(item_id, name, offset = 0)
    default_category = {
      category: "",
      subcategory: "",
    }
    (0...RETRY_COUNT).each do
      url = ITEM_URL_FORMAT % [item_id]
      begin
        page = @mech.get(url)
        html = Nokogiri::HTML(page.body.toutf8, "UTF-8")
        crumb = html.css("div.a-breadcrumb li")

        if (crumb.size == 0)
          sleep(RETRY_WAIT_SEC)
          next
        end

        return {
                 category: crumb[0 + offset].text.strip,
                 subcategory: crumb[2 + offset].text.strip,
               }
      rescue Mechanize::ResponseCodeError => e
        case e.response_code
        when "404"
          self.class.warn("%s (ASIN: %s) のページが存在しないため，カテゴリが取得できませんでした．" %
                          [name, item_id])
          return default_category
        else
          STDERR.puts(e.message)
          STDERR.puts(e.backtrace.select { |item| %r|#{__FILE__}/|.match(item) }[0])
          self.class.warn("リトライします...(1) #{url}", false)
          sleep(RETRY_WAIT_SEC)
        end
      rescue => e
        STDERR.puts(e.message)
        STDERR.puts(e.backtrace.select { |item| %r|#{__FILE__}/|.match(item) }[0])
        self.class.warn("リトライします...(2) #{url}", false)
        sleep(RETRY_WAIT_SEC)
      end
    end

    self.class.warn("%s (ASIN: %s) のカテゴリを取得できませんでした．" %
                    [name, item_id])

    return default_category
  end

  def save_img(img_url, img_file_name, name, id)
    img_file_path = @img_dir_path + img_file_name
    5.times do |i|
      return if (File.size?(img_file_path) != nil)
      @mech.get(img_url).save!(img_file_path)
      @mech.back()
      sleep(RETRY_WAIT_SEC)
    end

    self.class.warn("%s (ASIN: %s) の画像を取得できませんでした．" %
                    [name, id])
  end

  def parse_order_normal(html, date)
    item_list = []

    html.xpath('//div[@class="a-box" or @class="a-box shipment" or @class="a-box shipment shipment-is-delivered"]' +
               '//div[@class="a-fixed-right-grid-col a-col-left"]/div/div').each do |item|
      if !%r/販売|コンディション/.match(item.text)
        next
      end

      name = item.css("div.a-row")[0].text.strip
      url = URI.join(AMAZON_URL, item.css("div.a-row")[0].css("a")[0][:href]).to_s
      id = %r|/gp/product/([^/]+)/|.match(url)[1]

      count = 1
      if (%r|^商品名：(.+)、数量：(\d+)|.match(name))
        name = $1
        count = $2.to_i
      end

      price_str = item.css("div.a-row span.a-color-price").text.gsub(/¥|,/, "").strip
      # NOTE: 「price」は数量分の値段とする

      price = %r|\d+|.match(price_str)[0].to_i * count

      seller = ""
      (1..2).each do |i|
        seller_cand = item.css("div.a-row")[i].text
        next if (!%r|販売:|.match(seller_cand))

        if (item.css("div.a-row")[i].css("a")[0] != nil)
          seller = item.css("div.a-row")[i].css("a")[0].text.strip
        else
          seller = item.css("div.a-row")[i].text.gsub("販売:", "").strip
        end
        break
      end

      img_url = nil
      begin
        img_url = item.css("div.item-view-left-col-inner img")[0][:'data-a-hires']
      rescue
        self.class.warn("%s (ASIN: %s) の画像を取得できませんでした．" %
                        [name, id])
      end

      if (img_url != nil)
        img_file_name = "%s.%s" % [id, %r|\.(\w+)$|.match(img_url)[1]]
        save_img(img_url, img_file_name, name, id)
      end

      category_info = get_item_category(id, name)

      item_list.push(
        {
          name: name,
          id: id,
          url: url,
          count: count,
          price: price,
          category: category_info[:category],
          subcategory: category_info[:subcategory],
          seller: seller,
          date: date,
        }
      )
    end

    return item_list
  end

  def parse_order_digital(html, date, img_url_map)
    item_list = []

    item = html.css("table.sample")

    seller_str = item.css("table table tr:nth-child(2) td:nth-child(1)")[0].text.strip
    price_str = item.css("table table tr:nth-child(2) td:nth-child(2)")[0].text.gsub(/¥|￥|,/, "").strip

    name = item.css("table table tr:nth-child(2) b")[0].text.strip
    url = ""
    id = ""
    seller = "Amazon Japan G.K."
    count = 1
    price = %r|\d+|.match(price_str)[0].to_i

    if m = %r|販売: (.+)$|.match(seller_str)
      seller = m[1]
    end
    category_info = {
      category: "",
      subcategory: "",
    }

    begin
      url = item.css("table table tr:nth-child(2) a")[0][:href]
      id = %r|/dp/([^/]+)/|.match(url)[1]
      category_info = get_item_category(id, name, 2)
    rescue
      # NOTE: 商品のページが消失
      self.class.warn("%s の URL, ID, カテゴリ を取得できませんでした．" % [name])
    end

    if (!img_url_map.empty?)
      # NOTE: 一律 img_url_map.values.first でもいいはずだけど自信ないので
      img_url = img_url_map.has_key?(id) ? img_url_map[id] : img_url_map.values.first
      img_file_name = "%s.%s" % [id, %r|\.(\w+)$|.match(img_url)[1]]
      save_img(img_url, img_file_name, name, id)
    else
      self.class.warn("%s (ASIN: %s) の画像を取得できませんでした．" %
                      [name, id])
    end

    item_list.push(
      {
        name: name,
        id: id,
        url: url,
        count: count,
        price: price,
        category: category_info[:category],
        subcategory: category_info[:subcategory],
        seller: seller,
        date: date,
      }
    )

    return item_list
  end

  def parse_order_page(url, date, img_url_map)
    (0...RETRY_COUNT).each do
      begin
        html = fetch_html(url, "debug_order_page.htm")

        error_message = html.xpath('//div[@class="a-box a-alert a-alert-warning a-spacing-large"]' +
                                   '//h4[@class="a-alert-heading"]')
        if (!error_message.empty?)
          raise StandardError.new(error_message.text.strip)
        end

        if (!html.xpath('//b[contains(text(), "デジタル注文")]').empty?)
          return parse_order_digital(html, date, img_url_map)
        else
          return parse_order_normal(html, date)
        end
      rescue => e
        STDERR.puts(e.message)
        STDERR.puts(e.backtrace.select { |item| %r|#{__FILE__}/|.match(item) }[0])
        self.class.warn("リトライします...(3) #{url}", false)
        sleep(RETRY_WAIT_SEC)
      end
    end

    return []
  end

  def parse_order_list_page(html)
    item_list = []

    html.css("div.order").each do |order|
      begin
        date_text = order.css("div.order-info span.value")[0].text.strip
        date = Date.strptime(date_text, "%Y年%m月%d日")

        # Kindle とかの場合はここで画像を取得しておく
        img_url_map = {}
        order.css("div.a-fixed-left-grid").each do |item|
          url = URI.join(AMAZON_URL, item.css("div.a-row")[0].css("a")[0][:href]).to_s
          id = %r|/gp/product/([^/]+)/|.match(url)[1]

          begin
            img_url = item.css("div.item-view-left-col-inner img")[0][:'data-a-hires']
            img_url_map[id] = img_url if (img_url != nil)
          rescue
            # do nothing
          end
        end

        detail_url = order.css("a").select { |e| e.text =~ /注文内容を表示/ }[0][:href]
        detail_url = (@mech.page.uri.merge(detail_url)).to_s

        order_item = parse_order_page(detail_url, date, img_url_map)

        if (order_item.empty?)
          self.class.warn("注文詳細を読み取れませんでした．")
          self.class.warn("URL: %s" % [detail_url], false)
        end
        item_list.concat(order_item)
        STDERR.print "."
        STDERR.flush
      rescue Mechanize::Error => e
        self.class.warn("URL: %s" % [e.page.uri.to_s])
        STDERR.puts(e.message)
        STDERR.puts(e.backtrace.select { |item| %r|#{__FILE__}|.match(item) }[0])
      rescue => e
        STDERR.puts(e.message)
        STDERR.puts(e.backtrace.select { |item| %r|#{__FILE__}|.match(item) }[0])
      ensure
        sleep(2)
      end
    end

    return {
             item_list: item_list,
             is_last: html.css("div.pagination-full li.a-last").css("a").empty?,
           }
  end

  def get_item_list_by_page(year, page)
    html = fetch_html(hist_url(year, page), "debug_order_list_page.htm")
    return parse_order_list_page(html)
  end

  def get_item_list(year, page = 1)
    item_list = []

    loop do
      STDERR.print "%s Year %d page %d " % [Color.bold(Color.green("Parsing")),
                                            year, page]
      STDERR.flush
      page_info = get_item_list_by_page(year, page)
      item_list.concat(page_info[:item_list])
      STDERR.puts
      break if page_info[:is_last]
      page += 1
      sleep(2)
    end
    return item_list
  end
end

def info(message)
  STDERR.puts "[%s] %s" % [Color.bold(Color.yellow("INFO")), message]
end

def error(message)
  STDERR.puts "[%s] %s" % [Color.bold(Color.red("ERROR")), message]
  exit
end

def show_usage()
  puts <<"EOS"
■使い方
#{File.basename(__FILE__)} -j amazhist.json -t img -o amazhist.xlsx

引数の意味は以下なります．省略した場合，上記と同じ内容で実行します．

  -j 履歴情報を保存する JSON ファイルのパス (出力)

  -t サムネイル画像が保存されているディレクトリのパス (出力)

EOS
end

def check_arg(args)
  return if (defined?(Ocra))

  pass = args[:amazon_pass][0] + ("*" * (args[:amazon_pass].length - 2)) + args[:amazon_pass][-1]

  info(<<"EOS")
次の設定で実行します．
- ログイン ID               : #{args[:amazon_id]}
- ログイン PASS             : #{pass} (伏字処理済)

- 履歴情報ファイル          : #{args["-j"]}
- サムネイルディレクトリ    : #{args["-t"]}

続けますか？ [Y/n]
EOS

  answer = gets().strip

  if ((answer != "") && (answer.downcase != "y"))
    error("中断しました")
    exit
  end

  info(<<"EOS")
開始します．

【注意事項】
- 画像認証に対応する必要がある場合があります．
  メッセージが表示されましたら，画像ファイルに書かれている文字を入力
  お願いします．購入履歴が多い場合，何度か必要になります．

- Amazon のロボット対策を回避する為，時間がかかります．
  他のことをしてお待ち願います．

- スクリプトが途中で終了した場合，読み取りが完了した年から再開します．
  最初からやり直す場合，item_list.cache を削除してください．

EOS
end

def request_input(label, echo = true)
  print "#{label}: "

  loop do
    if (echo)
      input = gets().strip
    else
      input = (STDIN.noecho &:gets).strip
    end
    return input if (input != "")
  end
end

def login_info(arg)
  return if (defined?(Ocra))

  if (ENV["amazon_id"] == nil)
    arg[:amazon_id] = request_input("Amazon ログイン ID  ")
  else
    arg[:amazon_id] = ENV["amazon_id"]
  end

  if (ENV["amazon_pass"] == nil)
    arg[:amazon_pass] = request_input("Amazon ログイン PASS", false)
  else
    arg[:amazon_pass] = ENV["amazon_pass"]
  end
end

begin
  args = Docopt::docopt(DOCOPT)

  login_info(args)
  check_arg(args)

  FileUtils.mkdir_p(args["-t"])
  amazhist = Amazhist.new(
    {
      id: args[:amazon_id],     # Amazon の ID
      pass: args[:amazon_pass], # Amazon の パスワード
    },
    args["-t"]
  )

  exit if (defined?(Ocra))

  cache = amazhist.item_list_load()
  item_list = cache[:item_list]

  if ((cache[:year] > Amazhist::YEAR_START) || (cache[:page] > 1))
    info(<<"EOS")
#{cache[:year]-1}年 までのデータはキャッシュファイル(#{Amazhist::ITEM_LIST_DUMP})の内容を利用し，
データの取得を再開します．

EOS
  end

  year_start = cache[:year]
  page = cache[:page] + 1
  (year_start..(Date.today.year)).each do |year|
    year_item_list = amazhist.get_item_list(year, page)

    # NOTE: 前年までの内容をキャッシュファイルに書き出す
    amazhist.item_list_store(item_list, year, 0)

    item_list.concat(year_item_list)
    page = 1
  end

  File.open(args["-j"], "w") do |file|
    file.puts JSON.pretty_generate(item_list)
  end

  STDERR.puts Color.bold(Color.blue("Writing output file"))
rescue Docopt::Exit => e
  puts e.message
end

# Local Variables:
# coding: utf-8
# mode: ruby
# tab-width: 4
# indent-tabs-mode: nil
# End:

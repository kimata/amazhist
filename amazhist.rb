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

# TRACE = 1 # 定義すると，取得した Web ページをデバッグ用に保存します
# DEBUG = 1 # 定義すると，デバッグ用に保存したファイルからページを読み込みます

class Color
  extend Term::ANSIColor
end

class Amazhist
  AMAZON_URL      = 'http://www.amazon.co.jp/'

  HIST_URL_FORMAT = 'https://www.amazon.co.jp/gp/css/order-history?' +
                    'digitalOrders=1&unifiedOrders=1&orderFilter=year-%d&startIndex=%d'
  # NOTE: 下記アドレスの「?」以降を省略すると，時々ページの表示内容が変わり，
  # カテゴリを取得できなくなる
  ITEM_URL_FORMAT = 'http://www.amazon.co.jp/gp/product/%s?*Version*=1&*entries*=0'
  RETRY_COUNT     = 5
  RETRY_WAIT_SEC  = 5
  COOKIE_DUMP     = 'cookie.txt'

  def initialize(user_info, img_dir_path)
    @mech = Mechanize.new
    @mech.user_agent_alias = 'Windows Chrome'
    @mech.cookie_jar.clear!
    @user_info = user_info
    @img_dir_path = Pathname.new(img_dir_path)

    cookie_load()
    web_page = @mech.get(AMAZON_URL)
    sleep(RETRY_WAIT_SEC)
  end

  def self.error(message, is_nl=true)
    STDERR.puts if (is_nl)
    STDERR.puts '[%s] %s' % [ Color.bold(Color.red('ERROR')), message ]
    exit
  end

  def self.warn(message, is_nl=true)
    STDERR.puts if (is_nl)
    STDERR.puts '[%s] %s' % [ Color.bold(Color.yellow('WARN')), message ]
  end

  def cookie_save()
    @mech.cookie_jar.save_as(COOKIE_DUMP)
  end

  def cookie_load()
    if File.exist?(COOKIE_DUMP) then
      @mech.cookie_jar.load(COOKIE_DUMP)
    end
  end

  def hist_url(year, page)
    return HIST_URL_FORMAT % [ year, 10 * (page-1) ]
  end

  def login(web_page)
    2.times do |i|
      if !%r|Amazonログイン|.match(web_page.title) then
        cookie_save()
        return web_page
      end

      sleep(1)

      html = Nokogiri::HTML(web_page.body.toutf8, 'UTF-8')
      if (i == 0) then
        web_page.form_with(name: 'signIn') do |form|
          form.field_with(name: 'email').value = @user_info[:id]
          form.field_with(name: 'password').value = @user_info[:pass]
        end
        web_page = web_page.form_with(name: 'signIn').submit
      else
        if (defined? TRACE) then
          File.open('debug_login_page.htm', 'w') do |file|
            file.puts(web_page.body.toutf8)
          end
        end
        break
      end
    end
    raise StandardError.new('ログインに失敗しました．')
  end

  def fetch_html(url, file_path)
    if (defined? DEBUG) then
      File.open(file_path, 'r') do |file|
        return Nokogiri::HTML(file)
      end
    end

    web_page = @mech.get(url)
    web_page = login(web_page)

    if (defined? TRACE) then
      File.open(file_path, 'w') do |file|
        file.puts(web_page.body.toutf8)
      end
    end

    return Nokogiri::HTML(web_page.body.toutf8, 'UTF-8')
  end

  def get_item_category(item_id, name, offset = 0)
    default_category = {
      category: '',
      subcategory: '',
    }
    (0...RETRY_COUNT).each do
      url = ITEM_URL_FORMAT % [ item_id ]
      begin
        page = @mech.get(url)
        html = Nokogiri::HTML(page.body.toutf8, 'UTF-8')
        crumb = html.css('div.a-breadcrumb li')

        if (crumb.size == 0) then
          sleep(RETRY_WAIT_SEC)
          next
        end

        return {
          category: crumb[0 + offset].text.strip,
          subcategory: crumb[2 + offset].text.strip,
        }
      rescue Mechanize::ResponseCodeError => e
        case e.response_code
        when '404'
          self.class.warn('%s (ASIN: %s) のページが存在しないため，カテゴリが取得できませんでした．' %
                          [ name , item_id])
          return default_category
        else
          STDERR.puts(e.message)
          STDERR.puts(e.backtrace.select{|item| %r|#{__FILE__}/|.match(item) }[0])
          self.class.warn("リトライします... #{url}", false)
          sleep(RETRY_WAIT_SEC)
        end
      rescue => e
        STDERR.puts(e.message)
        STDERR.puts(e.backtrace.select{|item| %r|#{__FILE__}/|.match(item) }[0])
        self.class.warn("リトライします... #{url}", false)
        sleep(RETRY_WAIT_SEC)
      end
    end

    self.class.warn('%s (ASIN: %s) のカテゴリを取得できませんでした．' %
                    [ name , item_id])

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

    self.class.warn('%s (ASIN: %s) の画像を取得できませんでした．' %
                    [ name , id])
  end

  def parse_order_normal(html, date)
    item_list = []

    html.xpath('//div[@class="a-box" or @class="a-box shipment" or @class="a-box shipment shipment-is-delivered"]' +
               '//div[@class="a-fixed-right-grid-col a-col-left"]/div/div').each do |item|
      if !%r/販売|コンディション/.match(item.text) then
        next
      end

      name = item.css('div.a-row')[0].text.strip
      url = URI.join(AMAZON_URL, item.css('div.a-row')[0].css('a')[0][:href]).to_s
      id = %r|/gp/product/([^/]+)/|.match(url)[1]

      count = 1
      if (%r|^商品名：(.+)、数量：(\d+)|.match(name)) then
        name = $1
        count = $2.to_i
      end

      price_str = item.css('div.a-row span.a-color-price').text.gsub(/¥|,/, '').strip
      # NOTE: 「price」は数量分の値段とする

      price = %r|\d+|.match(price_str)[0].to_i * count

      seller = ''
      (1..2).each do |i|
        seller_cand = item.css('div.a-row')[i].text
        next if (!%r|販売:|.match(seller_cand))

        if (item.css('div.a-row')[i].css('a')[0] != nil) then
          seller = item.css('div.a-row')[i].css('a')[0].text.strip
        else
          seller = item.css('div.a-row')[i].text.gsub('販売:', '').strip
        end
        break
      end

      img_url = nil
      begin
        img_url = item.css('div.item-view-left-col-inner img')[0][:'data-a-hires']
      rescue
        self.class.warn('%s (ASIN: %s) の画像を取得できませんでした．' %
                        [ name , id])
      end

      if (img_url != nil) then
        img_file_name = '%s.%s' % [ id, %r|\.(\w+)$|.match(img_url)[1] ]
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
          date: date
        }
      )
    end

    return item_list
  end

  def parse_order_digital(html, date, img_url_map)
    item_list = []

    item = html.css('table.sample')

    seller_str = item.css('table table tr:nth-child(2) td:nth-child(1)')[0].text.strip
    price_str = item.css('table table tr:nth-child(2) td:nth-child(2)')[0].text.gsub(/¥|￥|,/, '').strip

    name = item.css('table table tr:nth-child(2) b')[0].text.strip
    url = ''
    id = ''
    seller = 'Amazon Japan G.K.'
    count = 1
    price = %r|\d+|.match(price_str)[0].to_i

    if m = %r|販売: (.+)$|.match(seller_str) then
      seller = m[1]
    end
    category_info = {
      category: '',
      subcategory: '',
    }

    begin
      url = item.css('table table tr:nth-child(2) a')[0][:href]
      id = %r|/dp/([^/]+)/|.match(url)[1]
      category_info = get_item_category(id, name, 2)
    rescue
      # NOTE: 商品のページが消失
      self.class.warn('%s の URL, ID, カテゴリ を取得できませんでした．' % [ name ])
    end

    if (!img_url_map.empty?) then
      # NOTE: 一律 img_url_map.values.first でもいいはずだけど自信ないので
      img_url = img_url_map.has_key?(id) ? img_url_map[id] : img_url_map.values.first
      img_file_name = '%s.%s' % [ id, %r|\.(\w+)$|.match(img_url)[1] ]
      save_img(img_url, img_file_name, name, id)
    else
      self.class.warn('%s (ASIN: %s) の画像を取得できませんでした．' %
                      [ name , id])
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
        date: date
      }
    )

    return item_list
  end

  def parse_order_page(url, date, img_url_map)
    (0...RETRY_COUNT).each do
      begin
        html = fetch_html(url, 'debug_order_page.htm')

        error_message = html.xpath('//div[@class="a-box a-alert a-alert-warning a-spacing-large"]' +
                                   '//h4[@class="a-alert-heading"]')
        if (!error_message.empty?) then
          raise StandardError.new(error_message.text.strip)
        end

        if (!html.xpath('//b[contains(text(), "デジタル注文")]').empty?) then
          return parse_order_digital(html, date, img_url_map)
        else
          return parse_order_normal(html, date)
        end
      rescue => e
        STDERR.puts(e.message)
        STDERR.puts(e.backtrace.select{|item| %r|#{__FILE__}/|.match(item) }[0])
        self.class.warn("リトライします... #{url}", false)
        sleep(RETRY_WAIT_SEC)
      end
    end

    return []
  end

  def parse_order_list_page(html, item_list)
    html.css('div.order').each do |order|
      begin
        date_text = order.css('div.order-info span.value')[0].text.strip
        date = Date.strptime(date_text, '%Y年%m月%d日')

        # Kindle とかの場合はここで画像を取得しておく
        img_url_map = {}
        order.css('div.a-fixed-left-grid').each do |item|
          url = URI.join(AMAZON_URL, item.css('div.a-row')[0].css('a')[0][:href]).to_s
          id = %r|/gp/product/([^/]+)/|.match(url)[1]

          begin
            img_url = item.css('div.item-view-left-col-inner img')[0][:'data-a-hires']
            img_url_map[id] = img_url if (img_url != nil)
          rescue
            # do nothing
          end
        end

        detail_url = order.css('a').select{|e| e.text =~ /注文の詳細/}[0][:href]
        order_item = parse_order_page(detail_url, date, img_url_map)

        if (order_item.empty?) then
          self.class.warn('注文詳細を読み取れませんでした．')
          self.class.warn('URL: %s' % [ detail_url], false)
        end

        item_list.concat(order_item)
        STDERR.print '.'
        STDERR.flush
      rescue Mechanize::Error => e
        self.class.warn('URL: %s' % [ e.page.uri.to_s ])
        STDERR.puts(e.message)
        STDERR.puts(e.backtrace.select{|item| %r|#{__FILE__}|.match(item) }[0])
      rescue => e
        STDERR.puts(e.message)
        STDERR.puts(e.backtrace.select{|item| %r|#{__FILE__}|.match(item) }[0])
      ensure
        sleep(1)
      end
    end

    return html.css('div.pagination-full li.a-last').css('a').empty?
  end

  def get_item_list_by_page(year, page, item_list)
    html = fetch_html(hist_url(year, page), 'debug_order_list_page.htm')
    return parse_order_list_page(html, item_list)
  end

  def get_item_list(year)
    item_list = []

    page = 1
    loop do
      STDERR.print '%s Year %d page %d ' % [ Color.bold(Color.green('Parsing')), 
                                           year, page ]
      STDERR.flush
      is_last = get_item_list_by_page(year, page, item_list)
      STDERR.puts
      break if is_last
      page += 1
      sleep 30
    end
    return item_list
  end
end

params = ARGV.getopts('j:t:')
if (params['j'] == nil) then
  Amazhist.error('履歴情報を保存するファイルのパスが指定されていません．' +
                 '(-j で指定します)')
  exit
end
if (params['t'] == nil) then
  Amazhist.error('サムネイル画像を保存するディレクトリのパスが指定されていません．' + 
                 '(-t で指定します)')
  exit
end

json_file_path = params['j']
img_dir_path = params['t']

if ((ENV['amazon_id'] == nil) || (ENV['amazon_pass'] == nil)) then
  STDERR.puts '[%s] %s' % [ Color.bold(Color.red('ERROR')),
                            '環境変数 amazon_id と amazon_pass を設定してください．' ]
  exit(-1)
end

FileUtils.mkdir_p(img_dir_path)
amazhist = Amazhist.new({
                          id: ENV['amazon_id'],     # Amazon の ID
                          pass: ENV['amazon_pass'], # Amazon の パスワード
                        },
                        img_dir_path)

item_list = []
(2000..(Date.today.year)).each do |year|
  item_list.concat(amazhist.get_item_list(year))
end

File.open(json_file_path, 'w') do |file|
  file.puts JSON.generate(item_list)
end

STDERR.puts Color.bold(Color.blue('Writing output file'))

# Local Variables:
# coding: utf-8
# mode: ruby
# tab-width: 4
# indent-tabs-mode: nil
# End:

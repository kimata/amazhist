#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

require 'term/ansicolor'
require 'pathname'
require 'nokogiri'
require 'mechanize'
require 'date'
require 'uri'
require 'json'

class Color
  extend Term::ANSIColor
end

class Amazhist
  AMAZON_URL      = "http://www.amazon.co.jp/"
  HIST_URL_FORMAT = "https://www.amazon.co.jp/gp/css/order-history?" +
    "digitalOrders=1&unifiedOrders=1&orderFilter=year-%d&startIndex=%d"

  def initialize(user_info, img_dir_path)
    @mech = Mechanize.new
    @mech.user_agent_alias = "Windows Chrome"
    @user_info = user_info
    @img_dir_path = Pathname.new(img_dir_path)
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

  def parse_item_page(html, item_list)
    # NOTE: for development
    # f = File.open("a.html")
    # html = Nokogiri::HTML(f)
    # f.close
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
          price = item.css("div.a-row span.a-color-price").text.gsub(/￥|,/, "").strip
          seller = ""
          if (item.css("div.a-row")[1].css("a")[0] != nil) then
            seller = item.css("div.a-row")[1].css("a")[0].text.strip
          else
            seller = item.css("div.a-row")[1].text.gsub("販売:", "").strip
          end

          img_url = item.css("div.item-view-left-col-inner img")[0][:src]
          img_file_name = "%s.%s" % [ id, %r|\.(\w+)$|.match(img_url)[1] ]
          @mech.get(img_url).save_as(@img_dir_path + img_file_name)

          item_list.push({
                           name: name,
                           id: id,
                           url: url,
                           count: count,
                           price: price,
                           seller: seller,
                           date: date
                         })
        end
        STDERR.print "."
        STDERR.flush
        # NOTE: for development
        # exit
      rescue => e
        STDERR.puts(e.message)
      end
    end
    return !html.css("div.pagination-full li.a-last").css("a").empty?
  end
  
  def get_item_page(year, page, item_list)
    # NOTE: for development
    # parse_item_page(nil, item_list)
    @mech.get(hist_url(year, page)) do |page|
      2.times do |i|
        html = Nokogiri::HTML(page.body)
        if %r|サインイン|.match(html.title) then
          # 2回目以降は少し待つ
          if (i != 0) then
            warn "画像認証を要求されたのでリトライします．"
            sleep(60) if (i != 0) 
          end
          page = login(page)
          next
        end
        return parse_item_page(html, item_list)
      end
      raise StandardError.new("ログインに失敗しました．")
    end
  end

  def get_item_list(year)
    item_list = []
    @mech.get("http://www.amazon.co.jp/")
    page = 1
    loop do
      STDERR.print "%s Year %d page %d " % [ Color.bold(Color.green("Parsing")), 
                                           year, page ]
      STDERR.flush
      break if !get_item_page(year, page, item_list)
      STDERR.puts
      page += 1
      sleep 5
    end
    return item_list
  end
end

IMG_PATH = './img'

FileUtils.mkdir_p(IMG_PATH)
amazhist = Amazhist.new({
                          id: ENV["amazon_id"], 	# Amazon の ID
                          pass: ENV["amazon_pass"],	# Amazon の パスワード
                        },
                        IMG_PATH)

item_list = []
(2000..(Date.today.year)).each do |year|
  item_list.merge(amazhist.get_item_list(year))
end

puts JSON.generate(item_list)

# Local Variables:
# coding: utf-8
# mode: ruby
# tab-width: 4
# indent-tabs-mode: nil
# End:

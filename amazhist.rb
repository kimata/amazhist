#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

require 'nokogiri'
require 'mechanize'

class Amazhist
  HIST_URL_FORMAT = "https://www.amazon.co.jp/gp/css/order-history?" +
    "digitalOrders=1&unifiedOrders=1&orderFilter=year-%d&startIndex=%d"

  def initialize()
    @mech = Mechanize.new
    @mech.user_agent_alias = "Windows Chrome"
  end

  def hist_url(year, page)
    return HIST_URL_FORMAT % [ year, (page-1) ]
  end

  def get_item_page(year, page)
    @mech.get(hist_url(year, page)) do |page|
      html = Nokogiri::HTML(page.body)
      p html
    end
  end
  def get_item(year)
    page = 1
    get_item_page(year, page)
  end
end

amazhist = Amazhist.new
amazhist.get_item(2014)


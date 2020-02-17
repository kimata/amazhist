#!/usr/bin/env ruby
# -*- coding: utf-8 -*-
# Amazexcel written by KIMATA Tetsuya <kimata@green-rabbit.net>

# Amazon の全購入履歴を Excel ファイルに見やすく出力するスクリプトです．
# amazhist.rb と組み合わせて使用します．

require 'date'
require 'json'
require 'optparse'
require 'pathname'
require 'set'
require 'term/ansicolor'
require 'axlsx'
require 'pathname'
require 'rmagick'

# DEBUG = 1

class Color
  extend Term::ANSIColor
end

class AmazExcel
  SHEET_NAME = {
    hist_data: "購入履歴一覧",
    category_stat: "【集計】カテゴリ別",
    yearly_stat: "【集計】購入年別",
    monthly_stat: "【集計】購入月別",
    wday_stat: "【集計】曜日別",
  }
  HIST_CONFIG = {
    header: {
      row: {
        pos: 1,
        height: 18,
        style: {
          font_name: 'メイリオ',
          bg_color: '333333',
          fg_color: 'FFFFFF',
          alignment: {
            indent: 1,
          },
        },
      },
      col: {
        date:           { label: '日付',            pos: 1,     width: 23, },
        name:           { label: '商品名',          pos: 2,     width: 70, },
        image:          { label: '画像',            pos: 3,     width: 8,  },
        count:          { label: '数量',            pos: 4,     width: 8,  },
        price:          { label: '価格',            pos: 5,     width: 16, },
        category:       { label: 'カテゴリ',        pos: 6,     width: 21, },
        subcategory:    { label: 'サブカテゴリ',    pos: 7,     width: 30, },
        seller:         { label: '売り手',          pos: 8,     width: 29, },
        id:             { label: '商品ID',          pos: 9,     width: 17, },
        url:            { label: '商品URL',         pos: 10,    width: 11, },
      },
    },
    data: {
      row: {
        height: 50,
        style: {
          font_name: 'メイリオ',
          fg_color: '333333',
          alignment: {
            vertical: :center,
            indent: 1,
          },
          border: {
            edges: [:top, :bottom],
            style: :thin,
            color: '333333',
          },
        },
      },
      col: {
        date: {
          type: :date,
          style: {
            format_code: 'yyyy年mm月dd日(aaa)',
          },
        },
        name: {
          style: {
            alignment: {
              wrap_text: true,
            },
          },
        },
        count: {
          type: :integer,
          style: {
            format_code: %|0_ |,
          },
        },
        price: {
          type: :integer,
          style: {
            format_code: %|_ ¥* #,##0_ ;_ ¥* -#,##0_ ;_ ¥* -_ ;_ @_ |, # NOTE: 末尾の空白要
          },
        },
        seller: {
          style: {
            alignment: {
              wrap_text: true,
            },
          },
        },
        url: {
          style: {
            alignment: {
              horizontal: :center,
            },
          },
        },
      },
    },
  }
  STAT_CONFIG = {
    header: {
      row: {
        pos: 1,
        height: 18,
        style: {
          font_name: 'メイリオ',
          bg_color: '333333',
          fg_color: 'FFFFFF',
          alignment: {
            indent: 1,
          },
        },
      },
      col: {
        target:         {                           pos: 1,     width: 23, },
        count:          { label: '合計件数',        pos: 2,     width: 12,  },
        price:          { label: '合計価格',        pos: 3,     width: 18, },
      },
    },
    data: {
      row: {
        height: 18,
        style: {
          font_name: 'メイリオ',
          fg_color: '333333',
          alignment: {
            vertical: :center,
            indent: 1,
          },
          border: {
            edges: [:top, :bottom],
            style: :thin,
            color: '333333',
          },
        },
      },
      col: {
        target: {
          style: {
            format_code: %|@|,
          },
        },
        count: {
          databar_color: '63C384',
          style: {
            format_code: %|0_ |,
          },
        },
        price: {
          databar_color: 'FF555A',
          style: {
            format_code: %|_ ¥* #,##0_ ;_ ¥* -#,##0_ ;_ ¥* -_ ;_ @_ |, # NOTE: 末尾の空白要
          },
        },
      },
    },
  }

  def initialize
    @package = Axlsx::Package.new
    @tmp_dir_path = Dir.mktmpdir('amazexcel-')
  end

  def create_style(sheet, table_config)
    style = {}

    table_config.each do |type, config|
      style[type] = {
        col: {},
      }

      if (config[:row].has_key?(:style)) then
        style[type][:row] =  sheet.styles.add_style(config[:row][:style])
      end
      config[:col].each do |key, value|
        next if (!value.has_key?(:style))

        col_style = config[:row][:style].dup
        value[:style].each do |k, v|
          if (col_style[k].is_a?(Hash)) then
            col_style[k] = col_style[k].merge(v)
          else
            col_style[k] = v
          end
        end
        style[type][:col][key] =  sheet.styles.add_style(col_style)
      end
    end

    sheet.sheet_view.show_grid_lines = false

    return style
  end

  def set_style(cell, name, style_map)
    if (style_map.has_key?(:row)) then
      cell.style = style_map[:row]
    end

    if (style_map[:col].has_key?(name)) then
      cell.style = style_map[:col][name]
    end
  end

  def insert_header(sheet, style, table_config)
    STDERR.print Color.cyan("    - テーブルヘッダを挿入します ")
    STDERR.flush

    col_max = table_config[:header][:col].values.map {|col_config| col_config[:pos] }.max

    row = table_config[:header][:row][:pos]

    (row+1).times do
      sheet.add_row(Array.new(col_max + 1, ''))
    end

    width_list = Array.new(col_max, 8)

    table_config[:header][:col].each do |name, col_config|
      col = col_config[:pos]

      sheet.rows[row].cells[col].value = col_config[:label]
      set_style(sheet.rows[row].cells[col], name, style[:header])

      if col_config.has_key?(:width) then
        width_list[col] = col_config[:width]
      end
    end
    sheet.column_widths(*width_list)

    STDERR.puts
  end

  def insert_item(sheet, table_config, style, row, col_max, item)
    sheet.add_row(Array.new(col_max + 1, ''))
    if (table_config[:data][:row].has_key?(:height)) then
      sheet.rows[row].height = table_config[:data][:row][:height]
    end

    table_config[:header][:col].each_key do |name|
      col = table_config[:header][:col][name][:pos]

      if (name == :url) then
        sheet.rows[row].cells[col].value = 'URL'
        sheet.add_hyperlink(
          :location => item[name.to_s],
          :ref => sheet.rows[row].cells[col]
        )
      else
        if (table_config[:data][:col].has_key?(name) &&
            table_config[:data][:col][name].has_key?(:type)) then
          sheet.rows[row].cells[col].type = table_config[:data][:col][name][:type]
        end
        if (item.has_key?(name.to_s)) then
          sheet.rows[row].cells[col].value = item[name.to_s]
        end
      end
      set_style(sheet.rows[row].cells[col], name, style[:data])
    end
  end

  def insert_hist_data(sheet, style, table_config, hist_data)
    STDERR.print Color.cyan('    - 履歴データを挿入します ')
    STDERR.flush

    col_max = table_config[:header][:col].values.map {|col_config| col_config[:pos] }.max

    hist_data.each_with_index do |item, i|
      row = table_config[:header][:row][:pos] + 1 + i
      insert_item(sheet, table_config, style, row, col_max, item)

      STDERR.print '.'
      STDERR.flush

      # NOTE: for development
      if (defined? DEBUG) then
        if (i > 10) then
          break
        end
      end
    end

    row = sheet.rows.size
    col = table_config[:header][:col][:price][:pos]
    sheet.add_row(Array.new(col_max + 1, ''))

    sheet.rows[row].cells[col].value =
      %|=SUM(%s%d:%s%d)| %
      [
        num2alpha(table_config[:header][:col][:price][:pos]),
        table_config[:header][:row][:pos] + 1,
        num2alpha(table_config[:header][:col][:price][:pos]),
        row
      ]
    set_style(sheet.rows[row].cells[col], :price_sum, style[:data])

    STDERR.puts
  end

  def crete_padding_image(img_path)
    # NOTE: Axlsx で画像を挿入する場合，画像の左上はセルの境界に必ずなってしまい，
    # 下記の観点で見た目が良くない．
    # - 罫線が隠れる
    # - 横長の画像が上に張り付いたように配置される
    # そのため，貼り付ける前に上下に透明の帯を着けて，セルにぴったり貼ったときにも
    # 見た目が良くなるようにする．
    #
    # 正確には，Axlsx の Pic.add_iamge には下記のコメントがあり，Axlsx
    # だけで何とかなるのかもしれない．
    #
    # This is a short cut method to set the start anchor position If
    # you need finer granularity in positioning use
    # graphic_frame.anchor.from.colOff / rowOff.

    image = Magick::Image.read(img_path).first

    width = image.columns
    height = image.rows

    if (width >= (height*1.1)) then
      height = width
    else
      height = (height*1.1).to_i
      width  = height
    end

    pad_image = Magick::Image.new(width, height).matte_floodfill(1, 1)
    pad_image.composite!(image, Magick::CenterGravity, Magick::OverCompositeOp)

    pad_image_path = (Pathname(@tmp_dir_path) + Pathname(img_path).basename).sub_ext('.png')
    pad_image.write(pad_image_path)

    return {
      path: pad_image_path,
      width: width,
      height: height,
    }
  end

  def insert_hist_image(sheet, table_config, hist_data, img_dir)
    STDERR.print Color.cyan('    - サムネイルを挿入します ')
    STDERR.flush

    hist_data.each_with_index do |item, i|
      img_path = img_dir + ('%s.jpg' % [ item['id'] ])
      if (!img_path.exist?) then
        STDERR.print item['id']
        next
      end

      pad_image = crete_padding_image(img_path)

      row = table_config[:header][:row][:pos] + 1 + i
      col = table_config[:header][:col][:pos]

      # NOTE: この変の計算式は出力をみて合わせ込み
      cell_width = table_config[:header][:col][:image][:width] * 10 * 0.8
      cell_height = sheet.rows[row].height * 83 / 50 * 0.8

      width = cell_width
      height = cell_height

      if ((pad_image[:width] / cell_width) > (pad_image[:height] / cell_height)) then
        height = pad_image[:height] * cell_width / pad_image[:width]
      else
        width = pad_image[:width] * cell_height / pad_image[:height]
      end

      sheet.add_image(image_src: pad_image[:path].to_s,
                      noSelect: true, noMove: false) do |image|
        image.width = width.to_i
        image.height = height.to_i
        image.start_at(table_config[:header][:col][:image][:pos],
                       table_config[:header][:row][:pos] + 1 + i)
      end

      STDERR.print '.'
      STDERR.flush

      # NOTE: for development
      if (defined? DEBUG) then
        if (i > 10) then
          break
        end
      end
    end
    STDERR.puts
  end

  def num2alpha(number)
    alpha = 'A'
    number.times { alpha.succ! }
    return alpha
  end

  def config_view(sheet, table_config, pin_name)
    STDERR.print Color.cyan('    - ビューを設定します ')
    STDERR.flush

    col_min = table_config[:header][:col].values.map {|col_config| col_config[:pos] }.min
    col_max = table_config[:header][:col].values.map {|col_config| col_config[:pos] }.max

    sheet.auto_filter = '%s%d:%s%s' % [
      num2alpha(col_min), table_config[:header][:row][:pos]+1,
      num2alpha(col_max), sheet.rows.size
    ]

    sheet.sheet_view.pane do |pane|
      spilit_col = table_config[:header][:col][pin_name][:pos]+1
      spilit_row = table_config[:header][:row][:pos]+1

      pane.state = :frozen_split
      pane.top_left_cell = '%s%d' % [ num2alpha(spilit_col), spilit_row+1 ]
      pane.x_split = spilit_col
      pane.y_split = spilit_row
    end

    STDERR.puts
  end

  def create_hist_sheet(book, hist_data, img_dir)
    STDERR.puts Color.bold(Color.green("「#{SHEET_NAME[:hist_data]}」シートを作成します:"))
    STDERR.flush

    sheet = book.add_worksheet(name: SHEET_NAME[:hist_data])

    style = create_style(sheet, HIST_CONFIG)
    insert_header(sheet, style, HIST_CONFIG)
    insert_hist_data(sheet, style, HIST_CONFIG, hist_data)
    insert_hist_image(sheet, HIST_CONFIG, hist_data, img_dir)
    config_view(sheet, HIST_CONFIG, :image)

    STDERR.puts

    return {
      start: {
        row: HIST_CONFIG[:header][:row][:pos] + 1,
        col: HIST_CONFIG[:header][:col].values.map {|col_config| col_config[:pos] }.min,
      },
      last: {
        row: sheet.rows.size - 2, # NOTE: 末尾に合計金額の行があるので -1 ではなく -2
        col: HIST_CONFIG[:header][:col].values.map {|col_config| col_config[:pos] }.max,
      },
    }
  end

  def build_stat_formula_category(table_config, hist_sheet_info, row)
    # NOTE: row に +1 しているのは Excel のマクロが，one-based なため
    count = %|=COUNTIF(%s!%s%d:%s%d,%s%d)| %
            [
              hist_sheet_info[:name],

              num2alpha(hist_sheet_info[:config][:header][:col][:category][:pos]),
              hist_sheet_info[:range][:start][:row] + 1,
              num2alpha(hist_sheet_info[:config][:header][:col][:category][:pos]),
              hist_sheet_info[:range][:last][:row] + 1,

              num2alpha(table_config[:header][:col][:target][:pos]),
              row + 1
            ]
    price = %|=SUMIF(%s!%s%d:%s%d,%s%d,%s!%s%d:%s%d)| %
            [
              hist_sheet_info[:name],

              num2alpha(hist_sheet_info[:config][:header][:col][:category][:pos]),
              hist_sheet_info[:range][:start][:row] + 1,
              num2alpha(hist_sheet_info[:config][:header][:col][:category][:pos]),
              hist_sheet_info[:range][:last][:row] + 1,

              num2alpha(table_config[:header][:col][:target][:pos]),
              row + 1,

              hist_sheet_info[:name],

              num2alpha(hist_sheet_info[:config][:header][:col][:price][:pos]),
              hist_sheet_info[:range][:start][:row] + 1,
              num2alpha(hist_sheet_info[:config][:header][:col][:price][:pos]),
              hist_sheet_info[:range][:last][:row] + 1
            ]

    return {
      count: count,
      price: price,
    }
  end

  def build_stat_formula_yearly(table_config, hist_sheet_info, row)
    # NOTE: row に +1 しているのは Excel のマクロが，one-based なため
    count = %|=SUMPRODUCT(--(YEAR(%s!%s%d:%s%d)=VALUE(%s%d)))| %
            [
              hist_sheet_info[:name],

              num2alpha(hist_sheet_info[:config][:header][:col][:date][:pos]),
              hist_sheet_info[:range][:start][:row] + 1,
              num2alpha(hist_sheet_info[:config][:header][:col][:date][:pos]),
              hist_sheet_info[:range][:last][:row] + 1,

              num2alpha(table_config[:header][:col][:target][:pos]),
              row + 1,
            ]
    price = %|=SUMPRODUCT((YEAR(%s!%s%d:%s%d)=VALUE(%s%d))*%s!%s%d:%s%d)| %
            [
              hist_sheet_info[:name],

              num2alpha(hist_sheet_info[:config][:header][:col][:date][:pos]),
              hist_sheet_info[:range][:start][:row] + 1,
              num2alpha(hist_sheet_info[:config][:header][:col][:date][:pos]),
              hist_sheet_info[:range][:last][:row] + 1,

              num2alpha(table_config[:header][:col][:target][:pos]),
              row + 1,

              hist_sheet_info[:name],

              num2alpha(hist_sheet_info[:config][:header][:col][:price][:pos]),
              hist_sheet_info[:range][:start][:row] + 1,
              num2alpha(hist_sheet_info[:config][:header][:col][:price][:pos]),
              hist_sheet_info[:range][:last][:row] + 1,
            ]

    return {
      count: count,
      price: price,
    }
  end

  def build_stat_formula_monthly(table_config, hist_sheet_info, row)
    # NOTE: row に +1 しているのは Excel のマクロが，one-based なため
    count = %|=SUMPRODUCT(--(TEXT(%s!%s%d:%s%d,"yyyy年mm月")=%s%d))| %
            [
              hist_sheet_info[:name],

              num2alpha(hist_sheet_info[:config][:header][:col][:date][:pos]),
              hist_sheet_info[:range][:start][:row] + 1,
              num2alpha(hist_sheet_info[:config][:header][:col][:date][:pos]),
              hist_sheet_info[:range][:last][:row] + 1,

              num2alpha(table_config[:header][:col][:target][:pos]),
              row + 1,
            ]

    price = %|=SUMPRODUCT((TEXT(%s!%s%d:%s%d,"yyyy年mm月")=%s%d)*%s!%s%d:%s%d)| %
            [
              hist_sheet_info[:name],

              num2alpha(hist_sheet_info[:config][:header][:col][:date][:pos]),
              hist_sheet_info[:range][:start][:row] + 1,
              num2alpha(hist_sheet_info[:config][:header][:col][:date][:pos]),
              hist_sheet_info[:range][:last][:row] + 1,

              num2alpha(table_config[:header][:col][:target][:pos]),
              row + 1,

              hist_sheet_info[:name],

              num2alpha(hist_sheet_info[:config][:header][:col][:price][:pos]),
              hist_sheet_info[:range][:start][:row] + 1,
              num2alpha(hist_sheet_info[:config][:header][:col][:price][:pos]),
              hist_sheet_info[:range][:last][:row] + 1,
            ]

    return {
      count: count,
      price: price,
    }
  end

  def build_stat_formula_wday(table_config, hist_sheet_info, row)
    # NOTE: row に +1 しているのは Excel のマクロが，one-based なため
    count = %|=SUMPRODUCT(--(WEEKDAY(%s!%s%d:%s%d,3)=%d))| %
            [
              hist_sheet_info[:name],

              num2alpha(hist_sheet_info[:config][:header][:col][:date][:pos]),
              hist_sheet_info[:range][:start][:row] + 1,
              num2alpha(hist_sheet_info[:config][:header][:col][:date][:pos]),
              hist_sheet_info[:range][:last][:row] + 1,

              row - (table_config[:header][:row][:pos] + 1)
            ]

    price = %|=SUMPRODUCT((WEEKDAY(%s!%s%d:%s%d,3)=%d)*%s!%s%d:%s%d)| %
            [
              hist_sheet_info[:name],

              num2alpha(hist_sheet_info[:config][:header][:col][:date][:pos]),
              hist_sheet_info[:range][:start][:row] + 1,
              num2alpha(hist_sheet_info[:config][:header][:col][:date][:pos]),
              hist_sheet_info[:range][:last][:row] + 1,

              row - (table_config[:header][:row][:pos] + 1),

              hist_sheet_info[:name],

              num2alpha(hist_sheet_info[:config][:header][:col][:price][:pos]),
              hist_sheet_info[:range][:start][:row] + 1,
              num2alpha(hist_sheet_info[:config][:header][:col][:price][:pos]),
              hist_sheet_info[:range][:last][:row] + 1,
            ]

    return {
      count: count,
      price: price,
    }
  end

  def create_stat_sheet_impl(book, sheet_config, hist_sheet_info)
    STDERR.puts Color.bold(Color.green("「#{sheet_config[:name]}」シートを作成します:"))
    STDERR.flush

    table_config = STAT_CONFIG.dup
    table_config[:header][:col][:target][:label] = sheet_config[:label]

    sheet = book.add_worksheet(name: sheet_config[:name])
    style = create_style(sheet, table_config)
    insert_header(sheet, style, table_config)

    col_max = table_config[:header][:col].values.map {|col_config| col_config[:pos] }.max

    STDERR.print Color.cyan('    - データを挿入します ')
    STDERR.flush

    sheet_config[:target_list].each_with_index do |target, i|
      row = table_config[:header][:row][:pos] + 1 + i

      sheet.add_row(Array.new(col_max + 1, ''))
      if (table_config[:data][:row].has_key?(:height)) then
        sheet.rows[row].height = table_config[:data][:row][:height]
      end

      formula = sheet_config[:formula_func].call(table_config, hist_sheet_info, row)

      item = {
        'target' => target,
        'count' => formula[:count],
        'price' => formula[:price],
      }
      insert_item(sheet, table_config, style, row, col_max, item)

      STDERR.print '.'
      STDERR.flush
    end

    table_config[:data][:col].each do |name, col_config|
      next if !col_config.has_key?(:databar_color)

      sheet.add_conditional_formatting(
        %|%s%d:%s%d| %
            [
              num2alpha(table_config[:header][:col][name][:pos]),
              table_config[:header][:row][:pos] + 2,
              num2alpha(table_config[:header][:col][name][:pos]),
              sheet.rows.size
            ],
        {
          type: :dataBar,
          priority: 1,
          data_bar: Axlsx::DataBar.new(color: col_config[:databar_color]),
        }
      )
    end
    STDERR.puts

    config_view(sheet, table_config, :target)

    STDERR.puts
  end

  def create_stat_sheet(book, stat_type, hist_data, hist_data_range)
    hist_sheet_info = {
      name: SHEET_NAME[:hist_data],
      config: HIST_CONFIG,
      data: hist_data,
      range: hist_data_range,
    }

    sheet_config = {
      name: SHEET_NAME[stat_type]
    }

    case stat_type
    when :category_stat
      category_set = Set.new
      hist_sheet_info[:data].each do |item|
        category_set.add(item['category'])
      end
      category_set.delete('')
      sheet_config.merge!({
        label: hist_sheet_info[:config][:header][:col][:category][:label],
        target_list: category_set.sort.to_a,
        formula_func: method(:build_stat_formula_category)
      })
    when :yearly_stat
      year_start = hist_data[0]['date'].year
      year_end = hist_data[-1]['date'].year

      sheet_config.merge!({
        label: '年',
        target_list: (year_start..year_end),
        formula_func: method(:build_stat_formula_yearly)
      })
    when :monthly_stat
      year_start = hist_data[0]['date'].year
      year_end = hist_data[-1]['date'].year
      month_start = hist_data[0]['date'].month
      month_end = hist_data[-1]['date'].month

      target = []
      (year_start..year_end).each do |year|
        (1..12).each do |month|
          next if ((year == year_start) && (month < month_start))
          next if ((year == year_end) && (month > month_end))

          target.push('%02d年%02d月' % [ year, month ])
        end
      end
      sheet_config.merge!({
        label: '年月',
        target_list: target,
        formula_func: method(:build_stat_formula_monthly)
      })
    when :wday_stat
      sheet_config.merge!({
        label: '曜日',
        target_list: %w(月 火 水 木 金 土 日),
        formula_func: method(:build_stat_formula_wday)
      })
    end

    create_stat_sheet_impl(book, sheet_config, hist_sheet_info)
  end

  def convert(json_path, img_dir_path, excel_path)
    begin
      img_dir = Pathname.new(img_dir_path)
      hist_data = open(json_path) {|io| JSON.load(io) }
      hist_data = hist_data.map do |item|
        item['date'] = Date.strptime(item['date'], "%Y-%m-%d")
        item
      end
      hist_data.sort_by! {|item| item['date'] }

      # MEMO: サンプルデータ作成用
      tmp_hist_data = hist_data
      hist_data = []
      (2013..2015).each do |year|
        (1..12).each do |month|
          data = tmp_hist_data.select {|item|
            (item['date'].year == year) && (item['date'].month == month)
          }
          if (data.size < 2) then
            hist_data.concat(data)
          else
            hist_data.concat(data[(0..rand(1..3))])
          end
        end
      end

      book = @package.workbook

      hist_data_range = create_hist_sheet(book, hist_data, img_dir)

      [
        :category_stat, :yearly_stat, :monthly_stat, :wday_stat
      ].each_with_index do |stat_type, i|
        create_stat_sheet(book, stat_type, hist_data, hist_data_range)
      end
      @package.serialize(excel_path)

      STDERR.puts
      STDERR.puts '完了しました．'
      STDERR.puts
    end
  end
end

def error(message)
  STDERR.puts '[%s] %s' % [ Color.bold(Color.red('ERROR')), message ]
  exit
end

def show_usage()
  puts <<"EOS"
■使い方
#{File.basename(__FILE__)} -j amazhist.json -t img -o amazhist.xlsx

引数の意味は以下なります．省略した場合，上記と同じ内容で実行します．

  -j 履歴情報を保存する JSON ファイルのパス (amazhist.rb にて生成したもの)

  -t サムネイル画像が保存されているディレクトリのパス (amazhist.rb にて生成したもの)

  -o 生成する Excel ファイルのパス

EOS
end

def check_arg(arg)
  puts <<"EOS"
次の設定で実行します．
- 履歴情報ファイル          : #{arg[:json_file_path]}
- サムネイルディレクトリ    : #{arg[:img_dir_path]}
- エクセルファイル(出力先)  : #{arg[:excel_file_path]}

続けますか？ [Y/n]
EOS
  answer = gets().strip

  if ((answer != '') && (answer.downcase != 'y')) then
    error('中断しました')
    exit
  end
end

ARG_DEFAULT = {
  json_file_path:   'amazhist.json',
  img_dir_path:     'img',
  excel_file_path:  'amazhist.xlsx'
}

params = ARGV.getopts('j:t:o:h')

if (params['h']) then
  show_usage()
  exit
end

arg = ARG_DEFAULT.dup
arg[:json_file_path]    = params['j'] if params['j']
arg[:img_dir_path]      = params['t'] if params['t']
arg[:excel_file_path]   = params['o'] if params['o']

check_arg(arg)

amazexcel = AmazExcel.new
amazexcel.convert(arg[:json_file_path], arg[:img_dir_path], arg[:excel_file_path])

# Local Variables:
# coding: utf-8
# mode: ruby
# tab-width: 4
# indent-tabs-mode: nil
# End:

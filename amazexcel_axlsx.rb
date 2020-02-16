#!/usr/bin/env ruby
# -*- coding: utf-8 -*-
# Amazexcel written by KIMATA Tetsuya <kimata@green-rabbit.net>

# Amazon の全購入履歴を Excel ファイルに見やすく出力するスクリプトです．
# amazhist.rb と組み合わせて使用します．
# WIN32OLE を使用するため，基本的に Windows で実行することを想定しています．
#
# ■準備
#   このスクリプトでは次のライブラリを使っていますので，入っていない場合は
#   インストールしておいてください．
#   - Term::ANSIcolor
#
# ■使い方
# 1. スクリプトを実行
#    $ ruby amazexcel.rb -j amazhist.json -t img -o amazhist.xlsx
#    引数の意味は以下
#    - j 履歴情報を保存する JSON ファイルのパス (amazhist.rb にて生成したもの)
#    - t サムネイル画像が保存されているディレクトリのパス (amazhist.rb にて生成したもの)
#    - o 生成する Excel ファイルのパス


require 'date'
require 'json'
require 'optparse'
require 'pathname'
require 'set'
require 'term/ansicolor'
require 'axlsx'
require 'pathname'
require 'rmagick'

DEBUG = 1

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
            format_code: %|_ ¥* #,##0_ ;_ ¥* -#,##0_ ;_ ¥* '-'_ ;_ @_ |, # NOTE: 末尾の空白要
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

  # TARGET_LABEL = {
  #   category_stat: HIST_HEADER[:col][:category][:label],
  #   yearly_stat: "年",
  #   monthly_stat: "年月",
  #   wday_stat: "曜日",
  # }
  GRAPH_CONFIG = {
    category_stat: {
      width: 700,
      height: 420,
    },
    yearly_stat: {
      width: 500,
      height: 300,
    },
    monthly_stat: {
      width: 1200,
      height: 360,
    },
    wday_stat: {
      width: 500,
      height: 300,
    },
  }
  IMG_SPACEING = 2

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
    STDERR.print Color.green("    - テーブルヘッダを挿入します ")
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
    STDERR.print Color.green('    - 履歴データを挿入します ')
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
    STDERR.print Color.green('    - サムネイルを挿入します ')
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
        image.start_at(2, 2+i)
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

  def config_view(sheet, table_config)
    STDERR.print Color.green('    - ビューを設定します ')
    STDERR.flush

    col_min = table_config[:header][:col].values.map {|col_config| col_config[:pos] }.min
    col_max = table_config[:header][:col].values.map {|col_config| col_config[:pos] }.max

    sheet.auto_filter = '%s%d:%s%s' % [
      num2alpha(col_min), table_config[:header][:row][:pos]+1,
      num2alpha(col_max), sheet.rows.size
    ]

    sheet.sheet_view.pane do |pane|
      spilit_col = table_config[:header][:col][:image][:pos]+1
      spilit_row = table_config[:header][:row][:pos]+1

      pane.state = :frozen_split
      pane.top_left_cell = '%s%d' % [ num2alpha(spilit_col), spilit_row+1 ]
      pane.x_split = spilit_col
      pane.y_split = spilit_row
    end

    STDERR.puts
  end

  def create_hist_sheet(book, hist_data, img_dir)
    STDERR.puts Color.bold(Color.green('購入履歴シートを作成します:'))
    STDERR.flush

    sheet = book.add_worksheet(name: '購入履歴')

    style = create_style(sheet, HIST_CONFIG)
    insert_header(sheet, style, HIST_CONFIG)
    insert_hist_data(sheet, style, HIST_CONFIG, hist_data)
    insert_hist_image(sheet, HIST_CONFIG, hist_data, img_dir)
    config_view(sheet, HIST_CONFIG)

    STDERR.puts

    return {
      start: {
        row: HIST_CONFIG[:header][:row][:pos] + 1,
        col: HIST_CONFIG[:header][:col].values.map {|col_config| col_config[:pos] }.min,
      },
      last: {
        row: sheet.rows.size - 1,
        col: HIST_CONFIG[:header][:col].values.map {|col_config| col_config[:pos] }.max,
      },
    }
  end

  def create_category_stat_sheet(book, hist_data, hist_data_range )
    category_set = Set.new

    hist_data.each do |item|
      category_set.add(item["category"])
    end
    category_set.delete('')

    category_set.sort.each_with_index do |category, i|
      # count_formula = %|=COUNTIF(%s!R%dC%d:R%dC%d,RC[-1])| %
      #   [
      #     SHEET_NAME[:hist_data],
      #     hist_data_range[:start_row], HIST_HEADER[:col][:category][:pos],
      #     hist_data_range[:last_row], HIST_HEADER[:col][:category][:pos],
      #   ]
    #   price_formula = %|=SUMIF(%s!R%dC%d:R%dC%d,RC[-2],%s!R%dC%d:R%dC%d)| %
    #     [
    #      SHEET_NAME[:hist_data],
    #      hist_data_range_info[:start_row], HIST_HEADER[:col][:category][:pos],
    #      hist_data_range_info[:last_row], HIST_HEADER[:col][:category][:pos],
    #      SHEET_NAME[:hist_data],
    #      hist_data_range_info[:start_row], HIST_HEADER[:col][:price][:pos],
    #      hist_data_range_info[:last_row], HIST_HEADER[:col][:price][:pos],
    #     ]

    #   insert_stat_data(sheet,
    #                    {
    #                      target: category,
    #                      row: STAT_HEADER[:row][:pos] + 1 + i,
    #                      count_formula: count_formula,
    #                      price_formula: price_formula,
    #                    },
    #                    hist_data_range_info)
    end
  end

  def create_stat_sheet(book, stat_type, hist_data, hist_data_range)
    STDERR.puts Color.bold(Color.green('Create Category Statistics:'))
    STDERR.flush

    case stat_type
      when :category_stat; create_category_stat_sheet(book, hist_data, hist_data_range)
      # when :yearly_stat; insert_yearly_stat_data(sheet, hist_data, hist_data_range)
      # when :monthly_stat; insert_monthly_stat_data(sheet, hist_data, hist_data_range)
      # when :wday_stat; insert_wday_stat_data(sheet, hist_data, hist_data_range)
    end

    # format_data(sheet, SHEET_NAME[stat_type], STAT_HEADER)
    # insert_header(sheet, STAT_HEADER, { nil => TARGET_LABEL[stat_type] })
    # set_border(sheet, STAT_HEADER)
    # insert_graph(sheet, stat_type)
  end

  # def insert_stat_data(sheet, param, hist_data_range_info)
  #   target_range = sheet.Cells[param[:row], STAT_HEADER[:col][:target][:pos]]
  #   # NOTE: 値を入力した場合に，数値に変換されるのを防止
  #   target_range.NumberFormatLocal = "@"
  #   target_range.Value = param[:target]

  #   count_range = sheet.Cells[param[:row], STAT_HEADER[:col][:count][:pos]]
  #   count_range.FormulaR1C1 = param[:count_formula]

  #   price_range = sheet.Cells[param[:row], STAT_HEADER[:col][:price][:pos]]
  #   price_range.FormulaR1C1 = param[:price_formula]
  # end


  # def insert_yearly_stat_data(sheet, hist_data, hist_data_range_info)
  #   year_start = Date.strptime(hist_data[0]["date"], "%Y-%m-%d").year
  #   year_end = Date.strptime(hist_data[-1]["date"], "%Y-%m-%d").year

  #   (year_start..year_end).each_with_index do |year, i|
  #     count_formula = %|=SUMPRODUCT(--(YEAR(%s!R%dC%d:R%dC%d)=RC[-1]))| %
  #       [
  #        SHEET_NAME[:hist_data],
  #        hist_data_range_info[:start_row], HIST_HEADER[:col][:date][:pos],
  #        hist_data_range_info[:last_row], HIST_HEADER[:col][:date][:pos],
  #       ]

  #     price_formula = %|=SUMPRODUCT((YEAR(%s!R%dC%d:R%dC%d)=RC[-2])*%s!R%dC%d:R%dC%d)| %
  #       [
  #        SHEET_NAME[:hist_data],
  #        hist_data_range_info[:start_row], HIST_HEADER[:col][:date][:pos],
  #        hist_data_range_info[:last_row], HIST_HEADER[:col][:date][:pos],
  #        SHEET_NAME[:hist_data],
  #        hist_data_range_info[:start_row], HIST_HEADER[:col][:price][:pos],
  #        hist_data_range_info[:last_row], HIST_HEADER[:col][:price][:pos],
  #       ]

  #     insert_stat_data(sheet,
  #                      {
  #                        target: year,
  #                        row: STAT_HEADER[:row][:pos] + 1 + i,
  #                        count_formula: count_formula,
  #                        price_formula: price_formula,
  #                      },
  #                      hist_data_range_info)
  #   end
  # end

  # def insert_monthly_stat_data(sheet, hist_data, hist_data_range_info)
  #   date_start = Date.strptime(hist_data[0]["date"], "%Y-%m-%d")
  #   date_end = Date.strptime(hist_data[-1]["date"], "%Y-%m-%d")

  #   year_start = date_start.year
  #   year_end = date_end.year
  #   month_start = date_start.month
  #   month_end = date_end.month

  #   i = 0
  #   (year_start..year_end).each do |year|
  #     (1..12).each do |month|
  #       next if ((year == year_start) && (month < month_start))
  #       next if ((year == year_end) && (month > month_end))

  #       year_month = "%02d年%02d月" % [ year, month ]

  #       count_formula = %|=SUMPRODUCT(--(TEXT(%s!R%dC%d:R%dC%d,"yyyy年mm月")="%s"))| %
  #         [
  #          SHEET_NAME[:hist_data],
  #          hist_data_range_info[:start_row], HIST_HEADER[:col][:date][:pos],
  #          hist_data_range_info[:last_row], HIST_HEADER[:col][:date][:pos],
  #          year_month,
  #         ]
  #       price_formula = %|=SUMPRODUCT((TEXT(%s!R%dC%d:R%dC%d,"yyyy年mm月")="%s")*%s!R%dC%d:R%dC%d)| %
  #         [
  #          SHEET_NAME[:hist_data],
  #          hist_data_range_info[:start_row], HIST_HEADER[:col][:date][:pos],
  #          hist_data_range_info[:last_row], HIST_HEADER[:col][:date][:pos],
  #          year_month,
  #          SHEET_NAME[:hist_data],
  #          hist_data_range_info[:start_row], HIST_HEADER[:col][:price][:pos],
  #          hist_data_range_info[:last_row], HIST_HEADER[:col][:price][:pos],
  #         ]

  #       insert_stat_data(sheet,
  #                        {
  #                          target: year_month,
  #                          row: STAT_HEADER[:row][:pos] + 1 + i,
  #                          count_formula: count_formula,
  #                          price_formula: price_formula,
  #                        },
  #                        hist_data_range_info)
  #       i += 1
  #     end
  #   end
  # end

  # def insert_wday_stat_data(sheet, hist_data, hist_data_range_info)
  #   year_start = Date.strptime(hist_data[0]["date"], "%Y-%m-%d").year
  #   year_end = Date.strptime(hist_data[-1]["date"], "%Y-%m-%d").year

  #   %w(月 火 水 木 金 土 日).each_with_index do |wday, i|
  #     count_formula = %|=SUMPRODUCT(--(WEEKDAY(%s!R%dC%d:R%dC%d,2)=%d))| %
  #       [
  #        SHEET_NAME[:hist_data],
  #        hist_data_range_info[:start_row], HIST_HEADER[:col][:date][:pos],
  #        hist_data_range_info[:last_row], HIST_HEADER[:col][:date][:pos],
  #        i + 1,
  #       ]

  #     price_formula = %|=SUMPRODUCT((WEEKDAY(%s!R%dC%d:R%dC%d,2)=%d)*%s!R%dC%d:R%dC%d)| %
  #       [
  #        SHEET_NAME[:hist_data],
  #        hist_data_range_info[:start_row], HIST_HEADER[:col][:date][:pos],
  #        hist_data_range_info[:last_row], HIST_HEADER[:col][:date][:pos],
  #        i + 1,
  #        SHEET_NAME[:hist_data],
  #        hist_data_range_info[:start_row], HIST_HEADER[:col][:price][:pos],
  #        hist_data_range_info[:last_row], HIST_HEADER[:col][:price][:pos],
  #       ]

  #     insert_stat_data(sheet,
  #                      {
  #                        target: wday,
  #                        row: STAT_HEADER[:row][:pos] + 1 + i,
  #                        count_formula: count_formula,
  #                        price_formula: price_formula,
  #                      },
  #                      hist_data_range_info)
  #   end
  # end

  # def insert_graph(sheet, stat_type)
  #   STDERR.print Color.green("    - Insert Graph ")
  #   STDERR.flush

  #   chart_width = GRAPH_CONFIG[stat_type][:width]
  #   chart_height = GRAPH_CONFIG[stat_type][:height]

  #   graph_col = STAT_HEADER[:col][get_last_col_id(STAT_HEADER)][:pos] + 2
  #   graph_range = sheet.Cells[STAT_HEADER[:row][:pos], graph_col]

  #   chart = sheet.ChartObjects.Add(graph_range.Left, graph_range.Top, chart_width, chart_height).Chart
  #   chart.SeriesCollection.NewSeries()
  #   chart.SeriesCollection.NewSeries()
  #   chart.HasTitle = true
  #   chart.ChartTitle.Text = sheet.Name
  #   chart.HasLegend = true
  #   chart.Legend.Position = ExcelConst::XlLegendPositionBottom

  #   series = chart.SeriesCollection(1)
  #   series.ChartType = ExcelConst::XlColumnClustered
  #   series.XValues = get_data_col_range(sheet, STAT_HEADER, :target)
  #   series.Values = get_data_col_range(sheet, STAT_HEADER, :count)
  #   series.Name = STAT_HEADER[:col][:count][:label]

  #   series = chart.SeriesCollection(2)
  #   series.ChartType = ExcelConst::XlLine
  #   series.XValues = get_data_col_range(sheet, STAT_HEADER, :target)
  #   series.Values = get_data_col_range(sheet, STAT_HEADER, :price)
  #   series.Name = STAT_HEADER[:col][:price][:label]
  #   series.AxisGroup = ExcelConst::XlSecondary

  #   yaxis_0 = chart.Axes(ExcelConst::XlValue, ExcelConst::XlPrimary)
  #   yaxis_1 = chart.Axes(ExcelConst::XlValue, ExcelConst::XlSecondary)
  #   xaxis = chart.Axes(ExcelConst::XlCategory, ExcelConst::XlPrimary)

  #   yaxis_0.HasTitle = true
  #   yaxis_0.AxisTitle.Text = "数量"
  #   set_font(yaxis_0.AxisTitle, 11)
  #   yaxis_1.HasTitle = true
  #   yaxis_1.AxisTitle.Text = "金額"
  #   set_font(yaxis_1.AxisTitle, 11)

  #   set_font(chart.ChartTitle, 14)
  #   set_font(yaxis_0.TickLabels, 11)
  #   set_font(yaxis_1.TickLabels, 11)
  #   set_font(xaxis.TickLabels, 11)
  #   set_font(chart.Legend, 11)

  #   STDERR.puts
  # end

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
        break
      end
      @package.serialize(excel_path)
    end
  end
end

def error(message)
  STDERR.puts '[%s] %s' % [ Color.bold(Color.red('ERROR')), message ]
  exit
end


params = ARGV.getopts("j:t:o:")
if (params["j"] == nil) then
  error("履歴情報が保存されたファイルのパスが指定されていません．" +
        "(-j で指定します)")
end
if (params["t"] == nil) then
  error("サムネイル画像が保存されたディレクトリのパスが指定されていません．" +
        "(-t で指定します)")
end
if (params["o"] == nil) then
  error("生成する Excel ファイルのパスが指定されていません．" +
        "(-o で指定します)")
end

json_file_path = params["j"]
img_dir_path = params["t"]
excel_file_path = params["o"]

amazexcel = AmazExcel.new
amazexcel.convert(json_file_path, img_dir_path, excel_file_path)

# Local Variables:
# coding: utf-8
# mode: ruby
# tab-width: 4
# indent-tabs-mode: nil
# End:

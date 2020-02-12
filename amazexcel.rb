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
require 'win32ole'

# DEBUG = 1

class Color
  extend Term::ANSIColor
end

class ExcelConst; end

class ExcelApp
  def initialize
    @excel = WIN32OLE.new("Excel.Application")
    WIN32OLE.const_load(@excel, ExcelConst)
    @excel.DisplayAlerts = false
  end

  def create_book(sheet_num = 1)
    @excel.SheetsInNewWorkbook = sheet_num
    return @excel.Workbooks.Add
  end

  def save(book, path)
    fso = WIN32OLE.new('Scripting.FileSystemObject')    
    full_path = fso.GetAbsolutePathName(path)

    book.SaveAs({
                  Filename: full_path,
                  ReadOnlyRecommended: true,
                })
  end

  def freeze_pane
    @excel.ActiveWindow.FreezePanes = true
  end
  
  def quit
    if @excel != nil then
      @excel.Quit
    end
    @excel = nil
  end
end

class AmazExcel
  SHEET_NAME = {
    hist_data: "購入履歴一覧",
    category_stat: "【集計】カテゴリ別",
    yearly_stat: "【集計】購入年別",
    monthly_stat: "【集計】購入月別",
    wday_stat: "【集計】曜日別",
  }
  HIST_HEADER = {
    row: {
      pos: 2,
      height: 50,
    },
    col: {
      date:	{
        label: "日付",			pos: 2,
        format: %|yyyy"年"mm"月"dd"日"|,
      },
      image: {
        label: "画像",			pos: 3,
        width: 8,
      },
      name: {
        label: "商品名",    	pos: 4,
        width: 70,
        wrap: true,
      },
      count: {
        label: "数量",      	pos: 5,
        format: %|0_ |,
        width: 8,
      },
      price: {
        label: "価格",      	pos: 6,
        format: %|_ ¥* #,##0_ ;_ ¥* -#,##0_ ;_ ¥* "-"_ ;_ @_ |, # NOTE: 末尾の空白要
      },
      category: {
        label: "カテゴリ",    	pos: 7,
        width: 15,
      },
      subcategory: {
        label: "サブカテゴリ", 	pos: 8,
        width: 22,
      },
      seller: {
        label: "売り手",    	pos: 9,
        width: 29,
        wrap: true,
      },
      id: {
        label: "商品ID",    	pos: 10,
        width: 17,
        format: %|@|
      },
      url: {
        label: "商品URL",   	pos: 11, 
        width: 11,
      },
    }
  }
  STAT_HEADER = {
    row: {
      pos: 2,      height: 20,
    },
    col: {
      target:	{
        label: nil,   			pos: 2, 
        format: %|@|
      },
      count: {
        label: "合計数量", 		pos: 3, 
        width: 12,
        format: %|0_ |,
      },
      price: {
        label: "合計価格", 		pos: 4, 
        format: %|_ ¥* #,##0_ ;_ ¥* -#,##0_ ;_ ¥* "-"_ ;_ @_ |, # NOTE: 末尾の空白要
        width: 17,
      },
    },
  }
  TARGET_LABEL = {
    category_stat: HIST_HEADER[:col][:category][:label],
    yearly_stat: "年",
    monthly_stat: "年月",
    wday_stat: "曜日",
  }
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
    @excel_app = ExcelApp.new
  end

  def RGB(r, g, b)
    return r | (g << 8) | (b << 16)
  end

  def set_font(target, size)
    target.Font.Name = "メイリオ"
    target.Font.Size = size
  end

  def get_first_col_id(header)
    first_col_id = nil
    col = 100
    header[:col].each do |col_id, cell_config|
      if (cell_config[:pos] < col) then
        col =  cell_config[:pos]
        first_col_id = col_id
      end
    end
    return first_col_id
  end

  def get_last_col_id(header)
    last_col_id = nil
    col = 0
    header[:col].each do |col_id, cell_config|
      if (cell_config[:pos] > col) then
        col =  cell_config[:pos]
        last_col_id = col_id
      end
    end
    return last_col_id
  end

  def get_table_range(sheet, header)
    if (header == HIST_HEADER) then
      return sheet.Cells[HIST_HEADER[:row][:pos],
                         HIST_HEADER[:col][:name][:pos]].CurrentRegion
    elsif (header == STAT_HEADER) then
      return sheet.Cells[STAT_HEADER[:row][:pos],
                         STAT_HEADER[:col][:target][:pos]].CurrentRegion
    else
      raise StandardError.new("BUG: 未知のヘッダです．")
    end
  end

  def get_data_col_range(sheet, header, col_id=nil)
    table_range = get_table_range(sheet, header)
    col_id = get_first_col_id(header) if col_id == nil
    last_row = table_range.Rows(table_range.Rows.Count).Row

    return sheet.Range(sheet.Cells[header[:row][:pos]+1,
                                   header[:col][col_id][:pos]],
                       sheet.Cells[last_row,
                                   header[:col][col_id][:pos]])
  end

  def get_data_range(sheet, header)
    table_range = get_table_range(sheet, header)
    last_row = table_range.Rows(table_range.Rows.Count).Row

    return sheet.Range(sheet.Cells[header[:row][:pos]+1,
                                   table_range.Columns(1).Column],
                       sheet.Cells[last_row,
                                   table_range.Columns(table_range.Columns.Count).Column])
    end

  def insert_picture(sheet, row, col, img_path)
    fso = WIN32OLE.new('Scripting.FileSystemObject')    
    img_full_path = fso.GetAbsolutePathName(img_path)

    cell_range = sheet.Cells[row, col]
    shape = sheet.Shapes.AddPicture({
                              FIleName: img_full_path,
                              LinkToFile: false,
                              SaveWithDocument: true,
                              Left: cell_range.Left + IMG_SPACEING,
                              Top: cell_range.Top + IMG_SPACEING,
                              Width: 0,
                              Height: 0
                            })
    shape.ScaleHeight(1, true)
    shape.ScaleWidth(1, true)
    
    scale = 1
    cell_width = sheet.Cells[1, col].Width - (IMG_SPACEING*2)
    cell_height = sheet.Cells[row, 1].Height - (IMG_SPACEING*2)
    if ((cell_width / shape.Width) < (cell_height / shape.Height)) then
      scale = cell_width / shape.Width
    else
      scale = cell_height / shape.Height
    end
    shape.Height *= scale
    shape.Width *= scale
    shape.Left = cell_range.Left + ((sheet.Cells[1, col].Width - shape.Width) / 2)
    shape.Placement = ExcelConst::XlMoveAndSize
  end

  def insert_header(sheet, header, label_map={})
    STDERR.print Color.green("    - Insert Header ")
    STDERR.flush

    header[:col].each_value do |cell_config|
      label = cell_config[:label]
      label = label_map[label] if (label_map.has_key?(label))
      header_range = sheet.Cells[header[:row][:pos], cell_config[:pos]]
      header_range.Value = label
      header_range.Font.Color = RGB(255, 255, 255)
      header_range.Interior.Color = RGB(38, 38, 38)
    end

    STDERR.puts
  end

  def insert_hist_data(sheet, hist_data)
    STDERR.print Color.green("    - Insert Data ")
    STDERR.flush

    hist_data.each_with_index do |item, i|
      item.each_key do |key|
        cell_range = sheet.Cells[HIST_HEADER[:row][:pos] + 1 + i,
                                 HIST_HEADER[:col][key.to_sym][:pos]]
        if (key.to_sym == :url) then
          sheet.Hyperlinks.Add({
                                 Anchor: cell_range,
                                 Address: item[key],
                                 TextToDisplay: "URL",
                               })
        else
          cell_range.Value = item[key]
        end
      end
      #NOTE: 画像用のセルにダミーデータを挿入
      img_cell_range = sheet.Cells[HIST_HEADER[:row][:pos] + 1 + i,
                                   HIST_HEADER[:col][:image][:pos]]
      img_cell_range.Value = "@"
      img_cell_range.Font.ColorIndex = 2

      STDERR.print "."
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

  def format_data(sheet, sheet_name, header)
    STDERR.print Color.green("    - Format Data ")
    STDERR.flush

    sheet.Name = sheet_name
    sheet.Cells.Interior.ColorIndex = 2
    sheet.Cells.Font.Name = "メイリオ"

    get_data_col_range(sheet, header).RowHeight = header[:row][:height]

    header[:col].each do |col_id, cell_config|
      col_range = sheet.Columns(header[:col][col_id][:pos])
      if (cell_config.has_key?(:format)) then
        col_range.NumberFormatLocal = cell_config[:format]
      end
      if (cell_config.has_key?(:wrap)) then
        col_range.WrapText = cell_config[:wrap]
      end
      if (col_id == :count) then
        col_range.HorizontalAlignment = ExcelConst::XlHAlignRight
      else
        col_range.HorizontalAlignment = ExcelConst::XlHAlignLeft
        col_range.IndentLevel = 1
        col_range.AddIndent = true
      end
    end

    sheet.Columns.AutoFit

    header[:col].each do |col_id, cell_config|
      col_range = sheet.Columns(header[:col][col_id][:pos])
      if (cell_config.has_key?(:width)) then
        col_range.ColumnWidth = cell_config[:width]
      end
    end
    STDERR.puts
  end

  def set_border(sheet, header)
    STDERR.print Color.green("    - Set Border ")
    STDERR.flush
    data_range = get_data_range(sheet, header)
    data_range.Borders(ExcelConst::XlInsideHorizontal).LineStyle = ExcelConst::XlContinuous
    data_range.Borders(ExcelConst::XlEdgeBottom).LineStyle = ExcelConst::XlContinuous
    STDERR.puts
  end

  def insert_hist_image(sheet, hist_data, img_dir)
    STDERR.print Color.green("    - Insert Image ")
    STDERR.flush

    hist_data.each_with_index do |item, i|
      cell_range = sheet.Cells[HIST_HEADER[:row][:pos] + 1 + i,
                               HIST_HEADER[:col][:image][:pos]]
      img_path = img_dir + ("%s.jpg" % [ item["id"] ])
      if (img_path.exist?) then
        insert_picture(sheet,
                       HIST_HEADER[:row][:pos] + 1 + i,
                       HIST_HEADER[:col][:image][:pos],
                       img_path.to_s)
      else
        STDERR.print item["id"]
      end
      STDERR.print "."
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

  def config_view(sheet, header)
    get_table_range(sheet, header).AutoFilter
    sheet.Cells[header[:row][:pos]+1,
                header[:col][get_first_col_id(header)][:pos]].Select
    @excel_app.freeze_pane
  end

  def create_hist_sheet(sheet, hist_data, img_dir)
    STDERR.puts Color.bold(Color.green("Create History Sheet:"))
    STDERR.flush

    insert_hist_data(sheet, hist_data)
    format_data(sheet, SHEET_NAME[:hist_data], HIST_HEADER)
    set_border(sheet, HIST_HEADER)
    insert_header(sheet, HIST_HEADER)
    insert_hist_image(sheet, hist_data, img_dir)
    config_view(sheet, HIST_HEADER)
    
    data_range = get_data_range(sheet, HIST_HEADER)
    
    return {
      start_row: data_range.Rows(1).Row,
      last_row: data_range.Rows(data_range.Rows.Count).Row,
      start_col: data_range.Columns(1).Column,
      last_col: data_range.Columns(data_range.Columns.Count).Column,
    }
  end

  def insert_stat_data(sheet, param, hist_data_range_info)
    target_range = sheet.Cells[param[:row], STAT_HEADER[:col][:target][:pos]]
    # NOTE: 値を入力した場合に，数値に変換されるのを防止
    target_range.NumberFormatLocal = "@"
    target_range.Value = param[:target]

    count_range = sheet.Cells[param[:row], STAT_HEADER[:col][:count][:pos]]
    count_range.FormulaR1C1 = param[:count_formula]

    price_range = sheet.Cells[param[:row], STAT_HEADER[:col][:price][:pos]]
    price_range.FormulaR1C1 = param[:price_formula]
  end

  def insert_category_stat_data(sheet, hist_data, hist_data_range_info)
    category_set = Set.new
    
    hist_data.each do |item|
      category_set.add(item["category"])
    end
    category_set.delete("")

    category_set.sort.each_with_index do |category, i|
      count_formula = %|=COUNTIF(%s!R%dC%d:R%dC%d,RC[-1])| %
        [
         SHEET_NAME[:hist_data],
         hist_data_range_info[:start_row], HIST_HEADER[:col][:category][:pos],
         hist_data_range_info[:last_row], HIST_HEADER[:col][:category][:pos],
        ]
      price_formula = %|=SUMIF(%s!R%dC%d:R%dC%d,RC[-2],%s!R%dC%d:R%dC%d)| %
        [
         SHEET_NAME[:hist_data],
         hist_data_range_info[:start_row], HIST_HEADER[:col][:category][:pos],
         hist_data_range_info[:last_row], HIST_HEADER[:col][:category][:pos],
         SHEET_NAME[:hist_data],
         hist_data_range_info[:start_row], HIST_HEADER[:col][:price][:pos],
         hist_data_range_info[:last_row], HIST_HEADER[:col][:price][:pos],
        ]

      insert_stat_data(sheet,
                       {
                         target: category,
                         row: STAT_HEADER[:row][:pos] + 1 + i,
                         count_formula: count_formula,
                         price_formula: price_formula,
                       },
                       hist_data_range_info)
    end
  end

  def insert_yearly_stat_data(sheet, hist_data, hist_data_range_info)
    year_start = Date.strptime(hist_data[0]["date"], "%Y-%m-%d").year
    year_end = Date.strptime(hist_data[-1]["date"], "%Y-%m-%d").year

    (year_start..year_end).each_with_index do |year, i|
      count_formula = %|=SUMPRODUCT(--(YEAR(%s!R%dC%d:R%dC%d)=RC[-1]))| %
        [
         SHEET_NAME[:hist_data],
         hist_data_range_info[:start_row], HIST_HEADER[:col][:date][:pos],
         hist_data_range_info[:last_row], HIST_HEADER[:col][:date][:pos],
        ]

      price_formula = %|=SUMPRODUCT((YEAR(%s!R%dC%d:R%dC%d)=RC[-2])*%s!R%dC%d:R%dC%d)| %
        [
         SHEET_NAME[:hist_data],
         hist_data_range_info[:start_row], HIST_HEADER[:col][:date][:pos],
         hist_data_range_info[:last_row], HIST_HEADER[:col][:date][:pos],
         SHEET_NAME[:hist_data],
         hist_data_range_info[:start_row], HIST_HEADER[:col][:price][:pos],
         hist_data_range_info[:last_row], HIST_HEADER[:col][:price][:pos],
        ]

      insert_stat_data(sheet,
                       {
                         target: year,
                         row: STAT_HEADER[:row][:pos] + 1 + i,
                         count_formula: count_formula,
                         price_formula: price_formula,
                       },
                       hist_data_range_info)
    end
  end

  def insert_monthly_stat_data(sheet, hist_data, hist_data_range_info)
    date_start = Date.strptime(hist_data[0]["date"], "%Y-%m-%d")
    date_end = Date.strptime(hist_data[-1]["date"], "%Y-%m-%d")

    year_start = date_start.year
    year_end = date_end.year
    month_start = date_start.month
    month_end = date_end.month

    i = 0
    (year_start..year_end).each do |year|
      (1..12).each do |month|
        next if ((year == year_start) && (month < month_start))
        next if ((year == year_end) && (month > month_end))

        year_month = "%02d年%02d月" % [ year, month ]

        count_formula = %|=SUMPRODUCT(--(TEXT(%s!R%dC%d:R%dC%d,"yyyy年mm月")="%s"))| %
          [
           SHEET_NAME[:hist_data],
           hist_data_range_info[:start_row], HIST_HEADER[:col][:date][:pos],
           hist_data_range_info[:last_row], HIST_HEADER[:col][:date][:pos],
           year_month,
          ]
        price_formula = %|=SUMPRODUCT((TEXT(%s!R%dC%d:R%dC%d,"yyyy年mm月")="%s")*%s!R%dC%d:R%dC%d)| %
          [
           SHEET_NAME[:hist_data],
           hist_data_range_info[:start_row], HIST_HEADER[:col][:date][:pos],
           hist_data_range_info[:last_row], HIST_HEADER[:col][:date][:pos],
           year_month,
           SHEET_NAME[:hist_data],
           hist_data_range_info[:start_row], HIST_HEADER[:col][:price][:pos],
           hist_data_range_info[:last_row], HIST_HEADER[:col][:price][:pos],
          ]

        insert_stat_data(sheet,
                         {
                           target: year_month,
                           row: STAT_HEADER[:row][:pos] + 1 + i,
                           count_formula: count_formula,
                           price_formula: price_formula,
                         },
                         hist_data_range_info)
        i += 1
      end
    end
  end

  def insert_wday_stat_data(sheet, hist_data, hist_data_range_info)
    year_start = Date.strptime(hist_data[0]["date"], "%Y-%m-%d").year
    year_end = Date.strptime(hist_data[-1]["date"], "%Y-%m-%d").year

    %w(月 火 水 木 金 土 日).each_with_index do |wday, i|
      count_formula = %|=SUMPRODUCT(--(WEEKDAY(%s!R%dC%d:R%dC%d,2)=%d))| %
        [
         SHEET_NAME[:hist_data],
         hist_data_range_info[:start_row], HIST_HEADER[:col][:date][:pos],
         hist_data_range_info[:last_row], HIST_HEADER[:col][:date][:pos],
         i + 1,
        ]

      price_formula = %|=SUMPRODUCT((WEEKDAY(%s!R%dC%d:R%dC%d,2)=%d)*%s!R%dC%d:R%dC%d)| %
        [
         SHEET_NAME[:hist_data],
         hist_data_range_info[:start_row], HIST_HEADER[:col][:date][:pos],
         hist_data_range_info[:last_row], HIST_HEADER[:col][:date][:pos],
         i + 1,
         SHEET_NAME[:hist_data],
         hist_data_range_info[:start_row], HIST_HEADER[:col][:price][:pos],
         hist_data_range_info[:last_row], HIST_HEADER[:col][:price][:pos],
        ]

      insert_stat_data(sheet,
                       {
                         target: wday,
                         row: STAT_HEADER[:row][:pos] + 1 + i,
                         count_formula: count_formula,
                         price_formula: price_formula,
                       },
                       hist_data_range_info)
    end
  end

  def insert_graph(sheet, stat_type)
    STDERR.print Color.green("    - Insert Graph ")
    STDERR.flush

    chart_width = GRAPH_CONFIG[stat_type][:width]
    chart_height = GRAPH_CONFIG[stat_type][:height]

    graph_col = STAT_HEADER[:col][get_last_col_id(STAT_HEADER)][:pos] + 2
    graph_range = sheet.Cells[STAT_HEADER[:row][:pos], graph_col]

    chart = sheet.ChartObjects.Add(graph_range.Left, graph_range.Top, chart_width, chart_height).Chart
    chart.SeriesCollection.NewSeries()
    chart.SeriesCollection.NewSeries()
    chart.HasTitle = true
    chart.ChartTitle.Text = sheet.Name
    chart.HasLegend = true
    chart.Legend.Position = ExcelConst::XlLegendPositionBottom

    series = chart.SeriesCollection(1)
    series.ChartType = ExcelConst::XlColumnClustered
    series.XValues = get_data_col_range(sheet, STAT_HEADER, :target)
    series.Values = get_data_col_range(sheet, STAT_HEADER, :count)
    series.Name = STAT_HEADER[:col][:count][:label]

    series = chart.SeriesCollection(2)
    series.ChartType = ExcelConst::XlLine
    series.XValues = get_data_col_range(sheet, STAT_HEADER, :target)
    series.Values = get_data_col_range(sheet, STAT_HEADER, :price)
    series.Name = STAT_HEADER[:col][:price][:label]
    series.AxisGroup = ExcelConst::XlSecondary

    yaxis_0 = chart.Axes(ExcelConst::XlValue, ExcelConst::XlPrimary)
    yaxis_1 = chart.Axes(ExcelConst::XlValue, ExcelConst::XlSecondary)
    xaxis = chart.Axes(ExcelConst::XlCategory, ExcelConst::XlPrimary)

    yaxis_0.HasTitle = true
    yaxis_0.AxisTitle.Text = "数量"
    set_font(yaxis_0.AxisTitle, 11)
    yaxis_1.HasTitle = true
    yaxis_1.AxisTitle.Text = "金額"
    set_font(yaxis_1.AxisTitle, 11)

    set_font(chart.ChartTitle, 14)
    set_font(yaxis_0.TickLabels, 11)
    set_font(yaxis_1.TickLabels, 11)
    set_font(xaxis.TickLabels, 11)
    set_font(chart.Legend, 11)

    STDERR.puts
  end

  def create_stat_sheet(sheet, stat_type, hist_data, hist_data_range)
    STDERR.puts Color.bold(Color.green("Create Category Statistics:"))
    STDERR.flush

    case stat_type
      when :category_stat; insert_category_stat_data(sheet, hist_data, hist_data_range)
      when :yearly_stat; insert_yearly_stat_data(sheet, hist_data, hist_data_range)
      when :monthly_stat; insert_monthly_stat_data(sheet, hist_data, hist_data_range)
      when :wday_stat; insert_wday_stat_data(sheet, hist_data, hist_data_range)
    end

    format_data(sheet, SHEET_NAME[stat_type], STAT_HEADER)
    insert_header(sheet, STAT_HEADER, { nil => TARGET_LABEL[stat_type] }) 
    set_border(sheet, STAT_HEADER)   
    insert_graph(sheet, stat_type)
  end

  def conver(json_path, img_dir_path, excel_path)
    begin
      img_dir = Pathname.new(img_dir_path)
      hist_data = open(json_path) {|io| JSON.load(io) }      
      hist_data.sort_by! {|item| Date.strptime(item["date"], "%Y-%m-%d") }

      # MEMO: サンプルデータ作成用
      # tmp_hist_data = hist_data
      # hist_data = []
      # (2013..2015).each do |year|
      #   (1..12).each do |month|        
      #     data = tmp_hist_data.select {|item|
      #       date = Date.strptime(item["date"], "%Y-%m-%d")
      #       (date.year == year) && (date.month == month)
      #     }
      #     if (data.size < 2) then
      #       hist_data.concat(data)
      #     else
      #       hist_data.concat(data[(0..rand(1..3))])
      #     end
      #   end
      # end

      book = @excel_app.create_book(5)

      hist_data_range_info = create_hist_sheet(book.Sheets[1], hist_data, img_dir)

      [:category_stat, :yearly_stat, :monthly_stat, :wday_stat].each_with_index do |stat_type, i|
        create_stat_sheet(book.Sheets[2 + i], stat_type, hist_data, hist_data_range_info)
      end

      @excel_app.save(book, excel_path)
    ensure
      @excel_app.quit
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
amazexcel.conver(json_file_path, img_dir_path, excel_file_path)

# Local Variables:
# coding: utf-8
# mode: ruby
# tab-width: 4
# indent-tabs-mode: nil
# End:

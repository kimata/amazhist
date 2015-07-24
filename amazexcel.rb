#!/usr/bin/env ruby
# -*- coding: utf-8 -*-
# Amazexcel written by KIMATA Tetsuya <kimata@green-rabbit.net>

# Amazhist が生成した JSON から Excel ファイルを生成するスクリプトです．
# WIN32OLE の機能を使いますので，Windows でのみ実行できます．
#
# ■準備
#   このスクリプトでは Term::ANSIcolor を使っていますので，入っていない場合は
#   インストールしておいてください．
#   > gem install term-ansicolor
#
# ■使い方
#   $ ruby amazexcel.rb JSON EXCEL
#   第一引数に指定された JSON ファイルを読み込み，
#   第二引数に指定されたファイル名の Excel ファイルを生成します．
#

require 'term/ansicolor'
require 'json'
require 'set'
require 'pathname'

require 'win32ole'

DEBUG = 1

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
    category_stat: "カテゴリ別集計",
  }
  HIST_HEADER = {
    row: {
      pos: 2,
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
        width: 8,
      },
      price: {
        label: "価格",      	pos: 6,
        format: %|_ ¥* #,##0_ ;_ ¥* -#,##0_ ;_ ¥* "-"_ ;_ @_ |, # NOTE: 末尾の空白要
      },
      category: {
        label: "カテゴリ",    	pos: 7,
      },
      subcategory: {
        label: "サブカテゴリ", 	pos: 8,
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
      pos: 2,
    },
    col: {
      target:	{
        label: nil,   			pos: 2, 
      },
      price: {
        label: "合計価格", 		pos: 3, 
      },
      count: {
        label: "合計数量", 		pos: 4, 
      },
    },
  }

  IMG_SPACEING = 2

  def initialize
    @excel_app = ExcelApp.new
  end

  def RGB(r, g, b)
    return r | (g << 8) | (b << 16)
  end

  def get_first_col(header)
    col = 100
    HIST_HEADER[:col].each do |col_id, cell_config|
      if (cell_config[:pos] < col) then
       col =  cell_config[:pos]
      end
    end
    return col
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

  def get_data_col_range(sheet, header, col_id)
    if (header == HIST_HEADER) then
      table_range = get_table_range(sheet, header)
      last_row = table_range.Rows(table_range.Rows.Count).Row

      return sheet.Range(sheet.Cells[HIST_HEADER[:row][:pos]+1,
                                     HIST_HEADER[:col][col_id][:pos]],
                         sheet.Cells[last_row,
                                     HIST_HEADER[:col][col_id][:pos]])
    else
      
    end
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

  def insert_header(sheet, header, label_map)
    STDERR.print Color.green("    Insert Header ")
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
    STDERR.print Color.green("    Insert Data ")
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

  def format_hist_data(sheet)
    STDERR.print Color.green("    Format Data ")
    STDERR.flush

    sheet.Name = SHEET_NAME[:hist_data]
    sheet.Cells.Interior.ColorIndex = 2
    sheet.Cells.Font.Name = "メイリオ"
    sheet.Columns.AutoFit

    get_data_col_range(sheet, HIST_HEADER, :name).RowHeight = 50

    HIST_HEADER[:col].each do |col_id, cell_config|
      col_range = sheet.Columns(HIST_HEADER[:col][col_id][:pos])
      if (cell_config.has_key?(:width)) then
        col_range.ColumnWidth = cell_config[:width]
      end
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
    STDERR.puts
  end

  def set_border(sheet, header)
    STDERR.print Color.green("    Set Border ")
    STDERR.flush
    data_range = get_data_range(sheet, header)
    data_range.Borders(ExcelConst::XlInsideHorizontal).LineStyle = ExcelConst::XlContinuous
    data_range.Borders(ExcelConst::XlEdgeBottom).LineStyle = ExcelConst::XlContinuous
    STDERR.puts
  end

  def insert_hist_image(sheet, hist_data, img_dir)
    STDERR.print Color.green("    Insert Image ")
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
        puts item["id"]
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

  def config_hist_view(sheet)
    get_table_range(sheet, HIST_HEADER).AutoFilter
    sheet.Cells[HIST_HEADER[:row][:pos]+1, get_first_col()].Select
    @excel_app.freeze_pane
  end

  def create_hist_sheet(sheet, hist_data, img_dir)
    STDERR.puts Color.bold(Color.green("Create History Sheet:"))
    STDERR.flush

    insert_hist_data(sheet, hist_data)
    format_hist_data(sheet)
    set_border(sheet, HIST_HEADER)
    insert__header(sheet, HIST_HEADER)
    insert_hist_image(sheet, hist_data, img_dir)
    config_hist_view(sheet)
  end

  def insert_category_data(sheet, hist_data)
    category_set = Set.new
    
    hist_data.each do |item|
      category_set.add(item["category"])
    end

    category_set.sort.each_with_index do |category, i|
      cell_range = sheet.Cells[STAT_HEADER[:row][:pos] + 1 + i,
                               STAT_HEADER[:col][:target][:pos]]
      cell_range.Value = category
    end
    
  end

  def create_category_stat_sheet(sheet, hist_data)
    insert_category_data(sheet, hist_data)
    insert_header(sheet, STAT_HEADER, { nil: HIST_HEADER[:col][:category][:label] }) 
    set_border(sheet, STAT_HEADER)   

  end

  def create_yearly_stat_sheet(sheet)

  end

  def create_monthly_stat_sheet(sheet)

  end

  def create_wday_stat_sheet(sheet)

  end


  def conver(img_dir_path, json_path, excel_path)
    begin
      img_dir = Pathname.new(img_dir_path)
      hist_data = open(json_path) { |io| JSON.load(io) }      
      book = @excel_app.create_book(4)

      create_category_stat_sheet(book.Sheets[2], hist_data)

      create_hist_sheet(book.Sheets[1], hist_data, img_dir)

      create_yearly_stat_sheet(book.Sheets[2])
      create_monthly_stat_sheet(book.Sheets[3])
      create_wday_stat_sheet(book.Sheets[4])

      @excel_app.save(book, excel_path)
    ensure
      @excel_app.quit
    end
  end
end

amazexcel = AmazExcel.new
amazexcel.conver("./img", "amaz.json", "amazecel.xlsx")

# Local Variables:
# coding: utf-8
# mode: ruby
# tab-width: 4
# indent-tabs-mode: nil
# End:

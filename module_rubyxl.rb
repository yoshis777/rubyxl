require 'rubyXL'

module ModuleRubyXL
  def begin
    # ワークブックの読み込み
    workbook = RubyXL::Parser.parse('excels/for_rubytest.xlsx')

    # コピー元シート
        org_sheet = workbook['コピーしたいシート']

    # 同じワークブックにシートをコピーして作成
        copied_sheet = workbook.add_worksheet #シート新規作成
    #copied_sheet = workbook.add_worksheet('コピーされたシート')でもよい
        copied_sheet.sheet_data = org_sheet.sheet_data
        copied_sheet.sheet_name = 'コピーされたシート'

    # セルの扱いについて
    # A1 = (0, 0)
    #add_cell(row = 0, column = 0, data = '', formula = nil, overwrite = true)
    # = add_cell(y, x, '', nil, false)

    # 文字列追記
        cell = copied_sheet.add_cell(1, 0, 'A2セル', nil, false)
        copied_sheet.add_cell(1, 1, '', 'A2') #「=A2」を追加

    # 文字列参照
        value = copied_sheet[0][0].value
        puts value

        file_name = "outputs/output_test_#{Time.now.strftime('%Y%m%d%H%M%s')}.xlsx"
        workbook.write(file_name)
    file_name
  end
end
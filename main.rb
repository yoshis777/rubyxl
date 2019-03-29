require './module_rubyxl'
require 'libreconv'
include ModuleRubyXL

# RubyXLを用いてexcelを操作
output_file_name = ModuleRubyXL.begin

# excelからpdfに変換
Libreconv.convert(output_file_name, output_file_name.gsub(/\.(.+)$/, ".pdf"))
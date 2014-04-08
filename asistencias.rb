require 'spreadsheet'
require 'pp'

Spreadsheet.client_encoding = 'UTF-8'
book = Spreadsheet.open './asistencias2.xls'

sheet1 = book.worksheet 0
asistencias = 0
sheet1.each do |row|
  
  a = row.grep /cp/
  
  if a.size > 0
    puts a
    if row.last.include?('none') || row.last.include?('orga') || row.last.include?('ref')
      puts row.first
      asistencias += row.first
    end
  end
  
end

puts asistencias
#   Prosledjujemo path do excel fajla

#   Mora biti implementiran enumarable model

#   Vodi racuna o mergovanju

#   Obogacena sintaksa

#   Biblioteka prepoznaje ukoliko postoji ključna reč "total" ili "subtotal" unutar excel fajla
#   i ignoriše taj red

#   Ignorise prazne redove  
#__________________________________________________________________________________________________#

#   Ovo radi za xlsx
require 'roo'
require 'matrix'

#   xlsx
#workbook = Roo::Spreadsheet.open 'C:\Users\3eka\Documents\Faks\Nepolozeni\III godina\Skript jezici\Domaci3-Ruby\Domaci_test\Skript-Ruby\Test_fajlovi\test1_xlsx.xlsx'
workbook = Roo::Spreadsheet.open 'C:\Users\3eka\Documents\Faks\Nepolozeni\III godina\Skript jezici\Domaci3-Ruby\Domaci_test\Skript-Ruby\Test_fajlovi\test3_xlsx.xlsx'
#   xls
#workbook = Roo::Spreadsheet.open 'C:\Users\3eka\Documents\Faks\Nepolozeni\III godina\Skript jezici\Domaci3-Ruby\Domaci_test\Skript-Ruby\Test_fajlovi\test2_xls.xls'

#matrix = Matrix[[0, 1][2, 3]]
#puts "Matrica #{matrix}"

worksheets = workbook.sheets

puts "Found #{worksheets.count}"

#   ws nam je iterable
worksheets.each do |ws|
    puts "Reading #{ws}"
    num_rows = 0
    
    workbook.sheet(ws).each_row_streaming do |row|
        
        if row.include? 'total' or row.include? 'subtotal' 
            #puts row_cells.join ' '
            puts 'Preskacemo'
        elsif row.empty?
            p row
            puts '------------------'
        end
        #row_cells = row.map { |cell| cell.value }
        #num_rows += 1
        
        #   Stampamo vrednosti celija
        #unless !row.include? 'total' or !row.include? 'subtotal'
        row_cells = row.map { |cell| cell.value }
        num_rows += 1

        puts row_cells.join ' '
    end
    
    puts "Reading #{num_rows} rows"
end

puts "Done"

#puts workbook.info

#xlsx = Roo::Spreadsheet.open('C:\Users\3eka\Documents\Faks\Nepolozeni\III godina\Skript jezici\Domaci3-Ruby\Domaci_test\Skript-Ruby\Test_fajlovi\test3_xlsx.xlsx.xlsx')
#p xlsx
#xlsx.info
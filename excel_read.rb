require 'roo'
require 'roo-xls'
require 'matrix'
#require 'daru'


#matrix = Matrix[[0, 1][2, 3]]


#   Pretvaramo xls u xlsx
'''
def convert(path)
    workbook = Roo::Spreadsheet.open path
    worksheets = workbook.sheets

end
'''

def read_file(path)
    
    workbook = Roo::Spreadsheet.open path
    workbook = Roo::Excelx.new(path, {:expand_merged_ranges => true})       #   Resenje za mergovane celije
    worksheets = workbook.sheets
    puts "Found #{worksheets.count}"

    mat = []

    #   ws nam je iterable
    worksheets.each do |ws|
        puts "Reading #{ws}..."
        num_rows = 0
        
        workbook.sheet(ws).each_row_streaming do |row|
            

            #puts row                                           #DEBUG
            #puts a.include? 'Roo::Excelx::Cell::Empty'         #DEBUG
            
            #   Ako je celija prazna, onda sadrzi 'Roo::Excelx::Cell::Empty'
            #   Pa filtriramo preko toga
            a = row
            #   Proverava da li je praznina
            unless a.to_s.include? 'Roo::Excelx::Cell::Empty'
                if a.to_s.include? 'total' or row.include? 'subtotal' 
                    #   Ako red sadrzi jednu od ove dve reci biva preskocen
                    next
                end
                
                #   Celije koje vracamo
                row_cells = row.map { |cell| cell.value }
                num_rows += 1
                
                #   Dodajemo elemente u niz
                mat.append(row_cells)

                #   Za ispis
                #puts row_cells.join '      '            #DEBUG
                
                #puts mat.join '\n'                      #DEBUG
                #puts '-------------'                    #DEBUG

            end
        end
        puts "Reading #{num_rows} rows..."
        
        
    end
    puts "Done"
    
    return mat
    
end

#puts workbook.info

#p xlsx
#xlsx.info

def print_mat(mat)
    mat.each do |el|
        if el[1][3]
            puts el.join '  '
            #next                   #DEBUG
        else
            puts el.join '                      ' 
        end
    end
end

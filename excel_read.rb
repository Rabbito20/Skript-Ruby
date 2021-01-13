require 'roo'
require 'roo-xls'
require 'matrix'


#matrix = Matrix[[0, 1][2, 3]]
#puts "Matrica #{matrix}"
#puts "-----------------------"

#   Pretvaramo xls u xlsx
def convert(path)
    workbook = Roo::Spreadsheet.open path
    worksheets = workbook.sheets


end

def read_file(path, save=false)
    
    workbook = Roo::Spreadsheet.open path
    workbook = Roo::Excelx.new(path, {:expand_merged_ranges => true})       #   Resenje za mergovane celije
    worksheets = workbook.sheets
    puts "Found #{worksheets.count}"

    #   ws nam je iterable
    worksheets.each do |ws|
        puts "Reading #{ws}..."
        num_rows = 0
        
        workbook.sheet(ws).each_row_streaming do |row|

            #puts row     #hmmmm
            #puts '-----------'
            #puts a.include? 'Roo::Excelx::Cell::Empty'
            
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
                
                #   Za ispis
                puts row_cells.join ' '
            end
        end
        puts "Reading #{num_rows} rows"
        

        #   Cuvanje fajla
        if save != false
            #TODO
            
            #ws.to_matrix
        end
        
    end
    puts "Done"
end

#puts workbook.info

#p xlsx
#xlsx.info
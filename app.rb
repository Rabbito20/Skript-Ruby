#   Prosledjujemo path do excel fajla

#   Mora biti implementiran enumarable model

#   Vodi racuna o mergovanju

#   Obogacena sintaksa

#   Biblioteka prepoznaje ukoliko postoji ključna reč "total" ili "subtotal" unutar excel fajla
#   i ignoriše taj red

#   Ignorise prazne redove  
#__________________________________________________________________________________________________#

require './excel_read'

#   xlsx
#path = '.\Test_fajlovi\test1_xlsx.xlsx'
path = '.\Test_fajlovi\test3_xlsx.xlsx'
#   xls
#path = '.\Test_fajlovi\test2_xls.xls'


#convert(path)

#   Citamo fajl
a = read_file(path)

#   Stampanje elemenata tabele
print_table(a)

#   Pristup redu tabele
#b = row(a, 4)
#print "b is #{b}"





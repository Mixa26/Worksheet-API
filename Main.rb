require "./GoogleSheets.rb"

table = Table.new($currentSheet)
#Biblioteka vraca dvodimenzioni niz
puts "Ucitana tabela:"
table.load_table
table.print_contents
print "\n"
#pristupanje redu preko table.row(index)
puts "Drugi red tabele:"
puts table.row(1)
print "\n"
#pristupanje elementima tabele preko each funkcije
puts "Pristupanje elementima preko each funkcije:"
table.each do |element|
    puts element
end
print "\n"
#pristupanje kolonama na razne nacine
puts "Pristupanje koloni upitom table[\"imeKolone\"]"
puts table["indeks"].to_s
print "\n"
puts "Pristupanje elementima unutar kolone koristeci sintaksu table[\"imeKolone\"][index]"
puts table["indeks"][1]
print "\n"
puts "Izmena vrednosti unutar tabele upotrebom sintakse table[\"imeKolone\"][index]"
puts "Stanje pre izmene"
table.print_contents
table["indeks"][2] = "6"
puts "Stanje posle izmene"
table.print_contents
print "\n"
#direktan pristup kolonama preko istoimenih metoda
puts "Pristup koloni upotrebom table.imeKolone"
puts table.header2.to_s
print "\n"
puts "Sum preko table.imeKolone.sum sintakse"
puts table.header2.sum
print "\n"
puts "Avg preko table.imeKolone.avg sintakse"
puts table.header2.avg
print "\n"
puts "Izvlacenje pojedinacnog reda tabele upotrebom table.imeKolone.sadrzajCelije sintakse"
puts "Stanje tabele:"
open_sheet(1)
table1 = Table.new($currentSheet)
table1.load_table
table1.print_contents
puts "Pristupanje preko indeksa rn5520"
print table1.indeks.rn5520
print "\n"
puts "Upotreba funkcija map, select, reduce posle izvlacenja reda sa prethodnom sintaksom"
puts "Stanje tabele:"
table.print_contents
puts "Map HEADER kolone upotrebom funkcije cell+=1"
puts table.HEADER.map { |cell| cell+=1 }
print "\n"
puts "Select HEADER kolone parnih vrednosti"
puts table.HEADER.select { |cell| cell.even? }
print "\n"
puts "Reduce HEADER kolone sabiranjem"
puts table.HEADER.reduce { |sum, cell| sum+cell }
print "\n"
puts "Demonstracija ignorisanja praznih redova i redova sa total/subtotal kljucnom reci"
puts "Red ispod headera je ignorisan jer sadzri subtotal kljucnu rec"
puts "Treci red (ukljucujuci header u brojanje) je prazan i ignorisan"
table.print_contents
print "\n"
puts "Sabiranje dve tabele:"
puts "Stanje prve tabele:"
table.print_contents
puts "Stanje druge tabele:"
table1.print_contents
puts "Rezultat sabiranja:"
table2 = table + table1
table2.print_contents
print "\n"
puts "Oduzimanje dve tabele:"
puts "Stanje prve tabele:"
table.print_contents
puts "Stanje druge tabele:"
table1.print_contents
puts "Rezultat oduzimanja:"
table2 = table-table1
table2.print_contents
print "\n"
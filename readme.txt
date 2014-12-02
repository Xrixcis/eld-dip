(23:40:38) Xrixcis: Tik_Tki_Tii.xsl je to cos mi poslal
(23:40:50) Xrixcis: obce.xlsx je ten seznam obci
(23:41:04) Xrixcis: matice.csv je prvni list te matice v csv
(23:41:16) Xrixcis: result.csv je to co snad chces
(23:41:30) Xrixcis: Gemfile te moc nezajima
(23:41:53) Xrixcis: db.db je SQLite databaze do ktere jsou nalite ty obce, matice a tiky
(23:42:20) Xrixcis: export.rb je ruby script ktery udela dotaz nad tou sqlite databazi a vybleje vysledek do result.csv
(23:43:15) Xrixcis: a create_db.rb je ruby script ktery precte ty obce, matici a tiky (pojmenovane tak jak tam jsou) a nableje je do databaze db.db
(23:44:28) Xrixcis: jinak kdyz si otevres ty scripty
(23:44:34) Xrixcis: export.rb
(23:44:57) Xrixcis: spousti se to po 
def run
(23:45:02) Xrixcis: = funkce run
(23:45:30) Xrixcis: db = SQLite3::Database.new 'db.db'
otevre tu sqlite databazi
(23:45:59) Xrixcis: File.open('result.csv', 'w+') do |f|	
otevre soubor result.csv k zapisu a soupne ho do promenne f
(23:46:11) Xrixcis: f.puts 'Region,KODOB_VYJ,NAZOB_VYJ,SumOfVYJ_DENNE,SumOfDOJ_DENNE,KODOB_DOJ,NAZOB_DOJ,Tik,Tki,Tii,Tjk,Tkj,SumOfTij,SumOfTii'
zapise radek do souboru
(23:46:47) Xrixcis: db.execute 'select .... ' do |row|
spusti ten select a kazdy radek postupne ulozi do promenne row
(23:47:15) Xrixcis: f.puts "#{row[0]},#{row[3]},#{row[1]},#{row[13]},#{row[12]},#{row[4]},#{row[2]},#{row[5]},#{row[6]},#{row[7]},#{row[8]},#{row[9]},#{row[10]},#{row[11]}"
vybleje radek do csv, to funguje nasledovne
(23:47:26) Xrixcis: " a " ohranicuji retezec
(23:47:48) Xrixcis: mezi #{ a } v retezci jde napsat nejakou promennou ktera se na to misto v retezci vlozi
(23:48:15) Xrixcis: a row je radek jako pole indexovane od 0, stejne poradi sloupcu jako v tom dotazu
(23:48:45) Xrixcis: do pole se pristupuje pres [] to je asi jasne
23:50:13) Xrixcis: jak vidis tak create_db je celkem trivialni, otevre databazi, udela tri tabulky, otevre ty xls a csv a z nich si vyzobe co potrebuje a vlozi to do db
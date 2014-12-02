require 'dbf'
require 'sqlite3'
require 'roo'

class Parser

  def create_db
    # otevreme databazi v pameti
    db = SQLite3::Database.new ':memory:'

    #db.execute('PRAGMA foreign_keys = ON')

    # vytvorime tabulky
    db.execute('create table obce (kod integer primary key, nazev character varying, region character varying, jadro boolean default false, skupina integer default null)')
    db.execute('create table matice (kod_from integer references obce(kod), kod_to integer references obce(kod), suma integer, primary key (kod_from, kod_to))')
    db.execute('create table tiky (kod integer references obce(kod) primary key, tii integer, tik integer, tki integer)')
    db
  end

  def write_db(db, file)
    # smazeme starou databazi
    File.delete file if File.exists? file
    # vytvorime databazi na disku
    disk_db = SQLite3::Database.new file

    # vyblejeme databazi z pameti na disk
    backup = SQLite3::Backup.new(disk_db, 'main', db, 'main')
    backup.step -1
    backup.finish
  end

  def read_rows(file, first_header = true)
    case File.extname(file)
    when '.xlsx'
      spreadsheet = Roo::Excelx.new file
      first = true
      spreadsheet.each do |row|
        yield row unless first && first_header
        first = false
      end
    when '.dbf'
      table = DBF::Table.new file, nil, 'cp1250'
      table.each {|row| yield row }
    when '.xls'
      spreadsheet = Roo::Excel.new file
      first = true
      spreadsheet.each do |row|
        yield row unless first && first_header
        first = false
      end
    when '.csv'
      spreadsheet = Roo::CSV.new file
      first = true
      spreadsheet.each do |row|
        yield row unless first && first_header
        first = false
      end
    else
      fail 'Unknown file type: ' + file
    end
  end

  def handle_obce(row, stmt)
    stmt.execute(row['ICZUJ'].to_i, row['NAZEV'], row['NAZEV_OBC2'], row['BYLO_JADRE'])
  end

  def handle_matice(row, stmt)
    stmt.execute(row[3].to_i, row[6].to_i, row[5].to_i)
  end

  def handle_tik(row, stmt)
    stmt.execute(row[0].to_i, row[4].to_i, row[2].to_i, row[3].to_i)
  end

  def handle_regiony(row, stmt)
    @i ||= 1
    row.each do |obec|
      stmt.execute(@i, obec) unless obec.nil? or obec.empty?
    end
    @i = @i + 1
  end

  def run

    db = create_db

    files = [
        {
            name: 'obce.dbf',
            handler: Proc.new {|row, stmt| handle_obce row, stmt },
            headers: true,
            query: 'insert into obce (kod, nazev, region, jadro) values (?,?,?,?)'
        },
        {
            name:'matice.csv',
            handler: Proc.new {|row, stmt| handle_matice row, stmt },
            headers: true,
            query: 'insert into matice (kod_from, kod_to, suma) values (?,?,?)'
        },
        {
            name: 'Tik_Tki_Tii.xls',
            handler: Proc.new {|row, stmt| handle_tik row, stmt },
            headers: true,
            query: 'insert into tiky (kod, tii, tik, tki) values (?,?,?,?)'
        },
        {
            name: 'RGR_regiony_rozdeleni.xlsx',
            handler: Proc.new {|row, stmt| handle_regiony row, stmt },
            headers: false,
            query: 'update obce set skupina=? where region=?'
        }
    ]

    db.transaction

    files.each do |file|
      puts "file #{file[:name]}"
      read_rows "data/#{file[:name]}", file[:headers] do |row|
        db.prepare file[:query] do |stmt|
          file[:handler].call row, stmt
        end
      end
    end

    db.commit

    write_db db, 'database.sqlite'
  end
end

Parser.new.run
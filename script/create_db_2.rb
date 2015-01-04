require 'dbf'
require 'sqlite3'
require 'roo'

class Parser

  def create_db
    # otevreme databazi v pameti
    db = SQLite3::Database.new ':memory:'

    #db.execute('PRAGMA foreign_keys = ON')

    # vytvorime tabulky
    db.execute('create table obce (kod integer primary key, nazev character varying, region character varying, jadro boolean default false, pos_x real, pos_y real, area real default null, perimeter real default null, obyvatel integer default null)')
    db.execute('create table tiky (kod integer references obce(kod) primary key, tii integer, tik integer, tki integer, class integer)')
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
    stmt.execute(row['ICZUJ'].to_i, row['NAZEV'], row['NAZEV_OBC2'], row['BYLO_JADRE'], row['POINT_X'], row['POINT_Y'])
  end

  def handle_tik(row, stmt)
    stmt.execute(row[0].to_i, row[4].to_i, row[2].to_i, row[3].to_i, row[5])
  end

  def handle_obce_xls(row, stmt)
    stmt.execute(row[0].to_f, row[1].to_i, row[4].to_i, row[2].to_i)
  end

  def run

    db = create_db

    files = [
        {
            name: 'obce.dbf',
            handler: Proc.new {|row, stmt| handle_obce row, stmt },
            headers: true,
            query: 'insert into obce (kod, nazev, region, jadro, pos_x, pos_y) values (?,?,?,?,?,?)'
        },
        {
            name: 'obce.xlsx',
            handler: Proc.new {|row, stmt| handle_obce_xls row, stmt },
            headers: true,
            query: 'update obce set area=?, perimeter=?, obyvatel=? where kod=?'
        },
        {
            name: 'Tik_Tki_Tii.xls',
            handler: Proc.new {|row, stmt| handle_tik row, stmt },
            headers: true,
            query: 'insert into tiky (kod, tii, tik, tki, class) values (?,?,?,?,?)'
        }
    ]

    db.transaction

    files.each do |file|
      puts "file #{file[:name]}"
      read_rows "data2/#{file[:name]}", file[:headers] do |row|
        db.prepare file[:query] do |stmt|
          file[:handler].call row, stmt
          affected = db.changes
          raise StandardError, "Unexpected count of changes: #{affected}" unless affected == 1
        end
      end
    end

    db.commit

    write_db db, 'database2.sqlite'
  end
end

Parser.new.run
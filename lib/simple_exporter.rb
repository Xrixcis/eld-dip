require_relative 'spreadsheet_writer'
require 'sqlite3'

class SimpleExporter

  def initialize(sum_result_file, headers, formulas, sheet2 = nil)
    @sum_file = sum_result_file

    @headers = headers
    @formulas = formulas
    @sheet2 = sheet2

    @sum_writer = SpreadsheetWriter.new @headers, @formulas, @sheet2
  end

  def add_row(row)
    @sum_writer.add row
  end

  def write_sum
    @sum_writer.save @sum_file
  end

  def finish
    write_sum
  end

  def read_db(filename)
    # otevri db
    db_file = SQLite3::Database.new filename

    # vytvorime db v pameti
    db = SQLite3::Database.new ':memory:'

    # a nacteme do ni db z disku
    backup = SQLite3::Backup.new(db, 'main', db_file, 'main')
    backup.step -1
    backup.finish

    db
  end

  def run(db_file, data_query)

    db = read_db db_file
      db.execute data_query do |row|
        add_row row
      end
    finish
  end

end
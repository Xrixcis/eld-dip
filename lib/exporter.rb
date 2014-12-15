require_relative 'spreadsheet_writer'
require 'sqlite3'

class Exporter

  def initialize(sum_result_file, unit_results_dir, headers, formulas)
    @sum_file = sum_result_file
    @unit_dir = unit_results_dir

    @headers = headers
    @formulas = formulas

    @sum_writer = SpreadsheetWriter.new @headers, @formulas
  end

  def add_row(row)
    r1 = @sum_writer.add row
    r2 = @unit_writer.add row
    [r1, r2]
  end

  def new_unit(name)
    write_unit unless @unit_writer.nil?
    @unit_writer = SpreadsheetWriter.new @headers, @formulas
    @unit_name = name
  end

  def write_unit
    @unit_writer.save "#{@unit_dir}#{File::SEPARATOR}#{@unit_name}.xlsx"
  end

  def write_sum
    @sum_writer.save @sum_file
  end

  def finish
    write_unit unless @unit_writer.nil?
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

  def run(db_file, unit_query, data_query, &block)

    db = read_db db_file

    Dir.mkdir @unit_dir unless Dir.exists? @unit_dir

    # vyber vsechny regiony
    db.execute unit_query do |unit_row|
      unit_name = unit_row[0]
      new_unit unit_name

      index = 1
      # spust dotaz a pro kazdy region
      db.execute data_query, unit_name do |row|
        rows = add_row row
        block.call(rows[0], index) unless block.nil?
        block.call(rows[1], index) unless block.nil?

        index += 1
      end
    end

    finish
  end

end
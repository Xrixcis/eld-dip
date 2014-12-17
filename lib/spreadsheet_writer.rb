require 'xlsx_writer'

class XlsxWriter
  class Cell
    def to_xml
      if empty?
        %{<c r="#{Cell.column_letter(x)}#{y}" s="0" t="s" />}
      elsif (value =~ /=/) == 0
        %{<c r="#{Cell.column_letter(x)}#{y}" s="#{Cell.style_number(type, faded?)}" t="#{Cell.type_name(type)}"><f>#{value}</f></c>}
      else
        %{<c r="#{Cell.column_letter(x)}#{y}" s="#{Cell.style_number(type, faded?)}" t="#{Cell.type_name(type)}"><v>#{escaped_value}</v></c>}
      end
    end
  end
end

class SpreadsheetWriter

  def initialize(headers, formulas, sheet2)
    @buff = []
    @formulas = formulas
    @headers = headers
    @sheet2 = sheet2
  end

  def add(row)
    rownum = @buff.length + 2
    @buff << (row + @formulas.map {|x| x.gsub(/\$i/, rownum.to_s) })
    @buff.last
  end

  def save(file)
    doc = XlsxWriter.new
    sheet = doc.add_sheet 'sheet 1'
    sheet.add_row @headers
    @buff.each { |row| sheet.add_row row }
    unless @sheet2.nil?
      sheet = doc.add_sheet @sheet2[:name]
      @sheet2[:data].each {|row| sheet.add_row row}
    end
    FileUtils.mv doc.path, file
    doc.cleanup
  end
end

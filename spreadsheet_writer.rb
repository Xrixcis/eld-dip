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

  def initialize(headers, formulas)
    @buff = []
    @formulas = formulas
    @headers = headers
  end

  def add(row)
    rownum = @buff.length + 2
    @buff << (row + @formulas.map {|x| x.gsub(/\$i/, rownum.to_s) })
  end

  def save(file)
    doc = XlsxWriter.new
    sheet = doc.add_sheet 'sheet 1'
    sheet.add_row @headers
    @buff.each { |row| sheet.add_row row }
    FileUtils.mv doc.path, file
    doc.cleanup
  end
end

require "google_drive"

module SheetHelper

  def row(row_index)
    self.rows[row_index - 1]
  end

  def each
    self.rows.each do |row|
      row.each do |cell|
        yield cell
      end
    end
  end

  def rows
    original_rows = super
    filtered_rows = original_rows.reject do |row|
      row.any? do |cell|
        cell.to_s.match?(/total|subtotal/i)
      end
    end
    filtered_rows
  end

  def any_empty_rows(ws)
    empty_rows = []
    ws.rows.each_with_index do |row, index|
      is_row_empty = row.all? { |cell| cell.empty? }
      empty_rows << index if is_row_empty
    end
    empty_rows
  end

  def headers
    self.rows.first
  end

  def method_missing(name, *args, &block)
    if name.to_s.match?(/^[a-z]+Kolona$/)
      column_name = format_name_to_header(name)
      column(column_name)
    elsif args.empty? && block.nil?
      find_row_by_value(name.to_s)
    else
      super
    end
  end

  def method_missing(name, *args, &block)
    if name.to_s.match?(/^[a-z]+Kolona$/)
      column_name = format_name_to_header(name)
      column(column_name)
    elsif args.empty? && block.nil? && (row = rows.find { |r| r.include?(name.to_s) })
      row
    else
      super
    end
  end

  private

  def format_name_to_header(name)
    name_string_split = name.to_s.split(/(?=[A-Z])/)
    caps_name = name_string_split.map {|part| part.capitalize}
    formatted_name = caps_name.join(" ")
  end

  def column(header_name)
    index = headers.index(header_name)
    rows.map { |row| row[index] }.extend(ColumnHelper)
  end

end

module ColumnHelper

  def sum
    converted_to_floats = map { |cell| cell.to_s.strip.to_f }
    converted_to_floats.sum
  end

  def avg
    return 0.0 if empty?
    filtered_data = select do |cell|
      cell_as_string = cell.to_s.strip
      !cell_as_string.empty? && cell.to_f != 0.0
    end
    filtered_data = filtered_data.map { |cell| cell.to_f }
    return 0.0 if filtered_data.empty?
    sum = filtered_data.sum
    sum / filtered_data.length
  end

end

# Creates a session. This will prompt the credential via command line for the
# first time and save it to config.json file for later usages.
# See this document to learn how to create config.json:
# https://github.com/gimite/google-drive-ruby/blob/master/doc/authorization.md
##Kristina Grbic RN 104/22
session = GoogleDrive::Session.from_config("config.json")

# First worksheet of
# https://docs.google.com/spreadsheet/ccc?key=pz7XtlQC-PYx-jrVMJErTcg
# Or https://docs.google.com/a/someone.com/spreadsheets/d/pz7XtlQC-PYx-jrVMJErTcg/edit?usp=drive_web
ws = session.spreadsheet_by_key("1vCS-0saNwpYfGO-oMfmeTx1mC6aqhSy1ukl5rDwYjl0").worksheets[0]

ws.extend(SheetHelper)
ws.extend(ColumnHelper)

##1. zadatak
# array = []
# ws.rows.each do |row|
#   array << row
# end
# p array

##2. zadatak
# t = ws
# p t.row(1) ##indeksiranje pocinje od 1, a ne od 0

# #3. zadatak
# ws.each do |cell|
#   puts cell
# end

##6. zadatak
# t = ws
# prva_kolona = t.prvaKolona
# p prva_kolona
# sum = t.prvaKolona.sum
# p sum
# avg = t.prvaKolona.avg
# p avg

##7. zadatak
# ws.rows.each do |row|
#   p row
# end

##10. zadatak
# empty_rows = ws.any_empty_rows(ws)
# ws.rows.each_with_index do |row, row_index|
#   p "Row #{row_index + 1} is empty" if empty_rows.include?(row_index)
#   p row unless empty_rows.include?(row_index)
# end

ws.reload
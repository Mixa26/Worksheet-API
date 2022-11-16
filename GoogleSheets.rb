
require "google_drive"

# Creates a session. This will prompt the credential via command line for the
# first time and save it to config.json file for later usages.
# See this document to learn how to create config.json:
# https://github.com/gimite/google-drive-ruby/blob/master/doc/authorization.md
session = GoogleDrive::Session.from_config("config.json")

# First worksheet of
# https://docs.google.com/spreadsheet/ccc?key=pz7XtlQC-PYx-jrVMJErTcg
# Or https://docs.google.com/a/someone.com/spreadsheets/d/pz7XtlQC-PYx-jrVMJErTcg/edit?usp=drive_web
$ws = session.spreadsheet_by_key("1JSNrQrqQZkiVOd5QT10Lm5OKUh5jD5U6BnmCfjramjI").worksheets[0]
  
$currentSheet = 0

#Opens a sheet specified by the argument
def open_sheet(sheetNumber)
  $ws = session.spreadsheet_by_key("1JSNrQrqQZkiVOd5QT10Lm5OKUh5jD5U6BnmCfjramjI").worksheets[sheetNumber]
  $currentSheet = sheetNumber
end

class Table
  include Enumerable

  def initialize(worksheet)
    @worksheet = worksheet
    @table = []
  end

  def each
    (0..@table.size-1).each do |row|
      (0..@table[0].size-1).each do |col|
        yield @table[row][col]
      end
    end
  end

  def find_table
    (1..$ws.num_rows).each do |row|
      (1..$ws.num_cols).each do |col|
        return row, col unless $ws[row, col].empty?
      end
    end
  end

  def row_empty?(row)
    (1..$ws.num_cols).each do |col|
      #print("$ws[row, col].empty? is #{$ws[row, col]}\n")
      return false unless $ws[row, col].empty? || $ws[row, col].nil?
    end
    true
  end

  def row_contains_total?(row)
    (1..$ws.num_cols).each do |col|
      cell_contents = $ws[row, col].downcase
      return true if cell_contents.include?("total") || cell_contents.include?("subtotal")
    end
    false
  end

  def load_table
    table_row, table_col = find_table
    @table = []

    (table_row..$ws.num_rows).each do |row|
      if row_empty?(row) || row_contains_total?(row)
        table_row += 1
        next
      end

      dummy_row = []
      (table_col..$ws.num_cols).each do |col|
        dummy_row[col-table_col] = $ws[row, col]
      end
      @table[row-table_row] = dummy_row
    end
  end

  def row(index)
    @table[index]
  end

  def print_contents
    @table.each do |row|
      print "#{row}\n"
    end
  end
end

table = Table.new($currentSheet)
table.load_table
print table.row(1)

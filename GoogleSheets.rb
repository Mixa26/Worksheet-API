
require "google_drive"

# Creates a session. This will prompt the credential via command line for the
# first time and save it to config.json file for later usages.
# See this document to learn how to create config.json:
# https://github.com/gimite/google-drive-ruby/blob/master/doc/authorization.md
$session = GoogleDrive::Session.from_config("config.json")

# First worksheet of
# https://docs.google.com/spreadsheet/ccc?key=pz7XtlQC-PYx-jrVMJErTcg
# Or https://docs.google.com/a/someone.com/spreadsheets/d/pz7XtlQC-PYx-jrVMJErTcg/edit?usp=drive_web
$ws = $session.spreadsheet_by_key("1JSNrQrqQZkiVOd5QT10Lm5OKUh5jD5U6BnmCfjramjI").worksheets[0]
  
$currentSheet = 0

#Opens a sheet specified by the argument
def open_sheet(sheetNumber)
  $ws = $session.spreadsheet_by_key("1JSNrQrqQZkiVOd5QT10Lm5OKUh5jD5U6BnmCfjramjI").worksheets[sheetNumber]
  $currentSheet = sheetNumber
end

def is_number?(str)
  true if Float(str) rescue false
end

class Column
  include Enumerable

  def initialize(name, index, column, table)
    @name = name
    @index = index
    @column = column
    @table = table
  end

  def method_missing(method, *args)
    index = -1
    (0..@column.size-1).each do |row|
      if method.to_s == @column[row]
        index = row 
        break
      end
    end
    @table.table[index+1] unless index == -1
  end

  def each
    (0..@column.size-1).each do |row|
      yield @column[row].to_i
    end
  end

  def [](row)
    @column[row]
  end

  def []=(row,val)
    @table.table[row][@index] = val
    #value can be updated on the server here
    #1.first save the current worksheet
    #2.add a attr_reader :worksheet to the class
    #Table
    #3.load up the worksheet from table by using
    #open_sheet(@table.worksheet)
    #4.use $ws.save here
    #5.and now load up the old worksheet you saved
    #in step 1
  end

  def sum
    sum = 0
    (0..@column.size-1).each do |row| 
      sum += @column[row].to_i if is_number?(@column[row])
    end
    sum
  end

  def avg
    self.sum.to_f / @column.size.to_f
  end

  def to_s
    print @column
  end
end

class Table
  include Enumerable

  attr_reader :table

  def initialize(worksheet)
    @worksheet = worksheet
    @table = []
  end

  def method_missing(method, *args)
    self[method.to_s]
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

  def  get_row(table_col_pos, row)
    dummy_row = []
    (table_col_pos..$ws.num_cols).each do |col|
      dummy_row[col-table_col_pos] = $ws[row, col]
    end
    dummy_row
  end

  def load_table
    table_row, table_col = find_table
    @table = []

    (table_row..$ws.num_rows).each do |row|
      if row_empty?(row) || row_contains_total?(row)
        table_row += 1
        next
      end

      @table[row-table_row] = get_row(table_col, row)
    end
  end

  def row(index)
    @table[index]
  end

  def row_exists?(table, row_to_compare)
    (0..table.size-1).each do |row|
      same = true
      (0..table[0].size-1).each do |col|
        same = false unless table[row][col] == row_to_compare[col]
      end
      return true if same
    end
    false
  end

  def headers_equal?(table1, table2)
    return false unless table1.size == table2.size
    (0..table1.size).each do |col|
      return false unless table1[0][col] == table2[0][col]
    end
    true
  end

  def +(table)
    return Table.new(-1) unless headers_equal?(@table, table.table)

    table_result = []
    result_row = 0
    
    (0..@table.size-2).each do |row|
      dummy_row = []
      (0..@table[0].size-1).each do |col|
        dummy_row[col] = @table[row+1][col]
      end
      table_result[result_row] = dummy_row
      result_row += 1
    end

    (0..table.table.size-2).each do |row|
      dummy_row = []
      (0..table.table[0].size-1).each do |col|
        dummy_row[col] = table.table[row+1][col]
      end
      next if row_exists?(table_result, dummy_row)
      table_result[result_row] = dummy_row
      result_row += 1
    end    

    table_return = Table.new(-1)
    table_return.table = table_result
    table_return
  end

  def -(table)
    return Table.new(-1) unless headers_equal?(@table, table.table)

    table_result = []
    result_row = 0
    
    (0..@table.size-1).each do |row|
      unless row_exists?(table.table, @table[row])
        table_result[result_row] = @table[row]
        result_row += 1
      end
    end

    table_return = Table.new(-1)
    table_return.table = table_result
    table_return
  end

  def get_col_index(col_name)
    (0..@table[0].size-1).each do |col|
      return col if @table[0][col] == col_name
    end
    -1
  end

  def [](col_name)
    index = get_col_index(col_name)
    return -1 if index == -1
    
    dummy_col = []
    (1..@table.size-1).each do |row|
      dummy_col[row-1] = @table[row][index]
    end
    Column.new(col_name, index, dummy_col, self)
  end

  def print_contents
    @table.each do |row|
      print "#{row}\n"
    end
  end
end

table = Table.new($currentSheet)
table.load_table
print table.header2.reduce(0) { |sum, num| sum + num }
require 'roo'
require 'forwardable'
require 'pry' # debug

# Has fields for a JIT produce item.
JitProduceItem =
  Struct.new(:vendor_num, :description, :size, :price, :id)

# Initializes a Roo::Spreadsheet with the default sheet equal to the first
# non-empty sheet.
class RooWrapper
  attr_reader :book

  def initialize(path)
    @book = Roo::Spreadsheet.open(path)
    @book.default_sheet = non_empty_sheets.first
  end

  private

  def non_empty_sheets
    @book.sheets.reject do |name|
      @book.sheet(name).empty?(2, 1) # scientifically determined
    end
  end
end

# Holds fields for a row in the output spreadsheet.
class ReportRow
  extend Forwardable

  attr_reader :prev, :cur
  def_delegators :@cur, :vendor_num, :id, :price, :size

  def initialize(cur, prev)
    binding.pry if prev && (cur.id != prev.id)

    @prev = prev
    @cur = cur
  end

  def last_price
    prev&.price
  end

  def diff
    return nil unless last_price
    price - last_price
  end

  def description
    prefix = prev.nil? ? "*** NEW *** " : ""
    prefix + cur.description
  end

  def to_h
    cur.to_h.merge({
      diff: diff,
      description: description
    })
  end

  def values
    [
      quoted(id),
      quoted(vendor_num),
      description,
      size,
      currency(price),
      currency(last_price),
      currency(diff)
    ]
  end

  def quoted(val)
    val ? "'#{val}" : nil
  end

  def currency(val)
    "%.2f" % val.to_f
  end
end

# Everything to run the script. Outputs a CSV with differences.
class Script
  DATA_DIR = 'data'
  OUTPUT_DIR = 'output'
  OUTPUT_HEADERS = %w[ID Vendor_Num Description Size Price Last_Price Difference]

  attr_reader :code_map

  def run
    @code_map = make_code_map('codes.csv')
    prev_items = xlsx_to_items('xdock-list-3.1-3.5.xlsx')
    cur_items = xlsx_to_items('xdock-list-3.8-3.13.xlsx')

    prev_items_by_id = prev_items.each_with_object({}) do |item, hsh|
      hsh[item.id] = item
    end

    report_rows = cur_items.map do |cur|
      prev = prev_items_by_id[cur.id]
      ReportRow.new(cur, prev)
    end.sort_by do |row|
      row.id
    end.sort_by do |row|
      next -1e10 if row.diff.nil?
      -1 * (row.diff || 0).abs
    end

    csv_path = File.join(OUTPUT_DIR, 'jit.csv')
    CSV.open(csv_path, 'wb', write_headers: true, headers: OUTPUT_HEADERS) do |csv|
      report_rows.each do |row|
        next if row.diff&.zero?
        csv << row.values
      end
    end
  end

  private

  def xlsx_to_items(file_name)
    path = File.join(DATA_DIR, file_name)
    raise 'no file' unless File.file?(path)

    book = RooWrapper.new(path).book

    book.parse.map do |row|
      row[1] = row[1].gsub(/\s+/, " ")
      row[4] = code_map[row[4].to_s] || row[4].to_s
      JitProduceItem.new(*row)
    end
  end

  def make_code_map(file_name)
    codes_csv = CSV.readlines(File.join(DATA_DIR, file_name), headers: true)

    codes_csv.each_with_object({}) do |row, hsh|
      alpha = cleaned(row['Alpha'])
      retalix = cleaned(row['Retalix'])
      hsh[alpha] = retalix || alpha
    end
  end

  def cleaned(str)
    result = str&.strip&.sub!("'", "")
    result == "" ? nil : result
  end
end

Script.new.run

require_relative '../lib/fast_excel'

def filename
  "example_table.xlsx"
end

workbook = FastExcel.open("example_table.xlsx", constant_memory: false)

worksheet = workbook.add_worksheet

worksheet.add_row(["Number", "Series", "Example", "Test", "Testing"])

for i in 0..5
  for n in 0..3
    worksheet.write_number(i, n, (i + 1) * (n + 1), nil)
  end
end

table = worksheet.add_table(0, 5, 0, 3)

workbook.close
puts "Saved to file #{filename}"

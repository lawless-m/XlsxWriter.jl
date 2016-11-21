include("XlsxWriter.jl")

using XlsxWriter


# Create an new Excel file and add a worksheet.
workbook = Workbook("demo.xlsx")
worksheet = add_worksheet!(workbook)

# Widen the first column to make the text clearer.
set_column!(worksheet, "A:A", 20)
set_column!(worksheet, "D:D", 21)
set_column!(worksheet, "E:E", 22)

# Add a bold format to use to highlight cells.
bold = add_format!(workbook, Dict("bold"=>true))
date_format = add_format!(workbook, Dict("num_format"=>"d mmmm yyyy"))

# Write some simple text.
write!(worksheet, "A1", "Hello")

# Text with formatting.
write!(worksheet, 1,1, "World", bold)

# Write some numbers, with row/column notation.
write!(worksheet, 2, 2, 123)
write!(worksheet, 3, 1, 123.456)
write!(worksheet, 3, 2, true)
write!(worksheet, 3, 3, now(), date_format)
write!(worksheet, 3, 4, Url("http://localhost"))
write!(worksheet, 3, 5, "=3 + 4")


close_workbook(workbook)

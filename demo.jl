include("XlsxWriter.jl")

using XlsxWriter


# Create an new Excel file and add a worksheet.
wb = workbook("demo.xlsx")
ws = add_worksheet!(wb)

# Widen the first column to make the text clearer.
set_column!(ws, "A:A", 20)
set_column!(ws, "D:D", 21)
set_column!(ws, "E:E", 22)

# Add a bold format to use to highlight cells.
bold = add_format!(wb, Dict("bold"=>true))
set_font_name!(bold, "Courier New")
set_font_size!(bold, 16)

date_format = add_format!(wb, Dict("num_format"=>"d mmmm yyyy"))
set_font_color!(date_format, "red")
# Write some simple text.
write!(ws, "A1", "Hello")

# Text with formatting.
write!(ws, 1,1, "World", bold)

# Write some numbers, with row/column notation.
write!(ws, 2, 2, 123)
write!(ws, 3, 1, 123.456)
write!(ws, 3, 2, true)
write!(ws, 3, 3, now(), date_format)
write!(ws, 3, 4, Url("http://localhost"))
write!(ws, 3, 5, "=3 + 4")
write!(ws, 3, 6, [["6" 7 8.8]; ["4x6" 47 4.7]])

freeze_panes!(ws, 1, 1)

merge_range!(ws, 4, 9, 4, 12, "Merged")

close_workbook(wb)

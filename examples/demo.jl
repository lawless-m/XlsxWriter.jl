include("../XlsxWriter.jl")

using XlsxWriter

# Create an new Excel file and add a worksheet.
wb = Workbook("demo.xlsx")
set_calc_mode!(wb, "manual")
ws = add_worksheet!(wb)

# Widen the first column to make the text clearer.
set_column!(ws, "A:A", 20)
set_column!(ws, 3, 3, 21)
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

write_formula!(ws, 3, 6, "=13 + 14", bold, result=9)

define_name!(wb, "duck", "=Sheet1!\$C\$3")
write_formula!(ws, 3, 7, "=duck*2", result=8)

write_row!(ws, 3, 7, ["6", 7, 8.8])
write_column!(ws, 4, 7, ["46", 47, 48.8])
write_matrix!(ws, 7, 1, [["105" "106" "107"]; ["201" 202 203]], bold)

write_row!(ws, 10, 1, [-2, 2, 3, -1])
add_sparkline!(ws, 10, 5, Dict("range"=>"Sheet1!B11:E11"))





#freeze_panes!(ws, 1, 1)

merge_range!(ws, 5, 9, 5, 12, "Merged")

close(wb)

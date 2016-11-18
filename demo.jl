include("XlsxWriter.jl")

using XlsxWriter


# Create an new Excel file and add a worksheet.
workbook = Workbook("demo.xlsx")
worksheet = add_worksheet(workbook)

# Widen the first column to make the text clearer.
set_column!(worksheet, "A:A", 20)

# Add a bold format to use to highlight cells.
#bold = add_format!(workbook, :bold=>true)

# Write some simple text.
write!(worksheet, "A1", "Hello")

# Text with formatting.
#write!(worksheet, "A2", "World", bold)

# Write some numbers, with row/column notation.
write!(worksheet, 2, 2, 123)
write!(worksheet, 3, 1, 123.456)
write!(worksheet, 3, 2, true)
write!(worksheet, 3, 3, now())
write!(worksheet, 3, 4, Url("http://localhost"))
write!(worksheet, 3, 5, "=3 + 4")


# Insert an image.
#insert_image!(worksheet, 'B5', 'logo.png')

close_workbook(workbook)

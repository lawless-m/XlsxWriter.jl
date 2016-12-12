###############################################################################
#
# An example of writing cell comments to a worksheet using XlsxWriter.
#
# Each of the worksheets demonstrates different features of cell comments.
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter


include("../XlsxWriter.jl")

using XlsxWriter

wb = Workbook("comments2.xlsx")

ws1 = add_worksheet!(wb)
ws2 = add_worksheet!(wb)
ws3 = add_worksheet!(wb)
ws4 = add_worksheet!(wb)
ws5 = add_worksheet!(wb)
ws6 = add_worksheet!(wb)
ws7 = add_worksheet!(wb)
ws8 = add_worksheet!(wb)

text_wrap = add_format!(wb, Dict("text_wrap"=> 1, "valign"=> "top"))


###############################################################################
#
# Example 1. Demonstrates a simple cell comments without formatting.
#            comments.
#

# Set up some formatting.
set_column!(ws1, "C:C", 25)
set_row!(ws1, 2, 50)
set_row!(ws1, 5, 50)

# Simple ASCII string.
cell_text = "Hold the mouse over this cell to see the comment."

comment = "This is a comment."

write!(ws1, "C3", cell_text, text_wrap)
write_comment!(ws1, "C3", comment)


###############################################################################
#
# Example 2. Demonstrates visible and hidden comments.
#

# Set up some formatting.
set_column!(ws2, "C:C", 25)
set_row!(ws2, 2, 50)
set_row!(ws2, 5, 50)

cell_text = "This cell comment is visible."
comment = "Hello."

write!(ws2, "C3", cell_text, text_wrap)
write_comment!(ws2, "C3", comment, Dict("visible"=> true))

cell_text = "This cell comment isn\"t visible (the default)."

write!(ws2, "C6", cell_text, text_wrap)
write_comment!(ws2, "C6", comment)


###############################################################################
#
# Example 3. Demonstrates visible and hidden comments set at the worksheet
#            level.
#

# Set up some formatting.
set_column!(ws3, "C:C", 25)
set_row!(ws3, 2, 50)
set_row!(ws3, 5, 50)
set_row!(ws3, 8, 50)

# Make all comments on the worksheet visible.
show_comments!(ws3)

cell_text = "This cell comment is visible, explicitly."
comment = "Hello."

write!(ws3, "C3", cell_text, text_wrap)
write_comment!(ws3, "C3", comment, Dict("visible"=> 1))

cell_text = "This cell comment is also visible because of show_comments()."

write!(ws3, "C6", cell_text, text_wrap)
write_comment!(ws3, "C6", comment)

cell_text = "However, we can still override it locally."

write!(ws3, "C9", cell_text, text_wrap)
write_comment!(ws3, "C9", comment, Dict("visible"=> false))


###############################################################################
#
# Example 4. Demonstrates changes to the comment box dimensions.
#

# Set up some formatting.
set_column!(ws4, "C:C", 25)
set_row!(ws4, 2, 50)
set_row!(ws4, 5, 50)
set_row!(ws4, 8, 50)
set_row!(ws4, 15, 50)

show_comments!(ws4)

cell_text = "This cell comment is default size."
comment = "Hello."

write!(ws4, "C3", cell_text, text_wrap)
write_comment!(ws4, "C3", comment)

cell_text = "This cell comment is twice as wide."

write!(ws4, "C6", cell_text, text_wrap)
write_comment!(ws4, "C6", comment, Dict("x_scale"=> 2))

cell_text = "This cell comment is twice as high."

write!(ws4, "C9", cell_text, text_wrap)
write_comment!(ws4, "C9", comment, Dict("y_scale"=> 2))

cell_text = "This cell comment is scaled in both directions."

write!(ws4, "C16", cell_text, text_wrap)
write_comment!(ws4, "C16", comment, Dict("x_scale"=> 1.2, "y_scale"=> 0.8))

cell_text = "This cell comment has width and height specified in pixels."

write!(ws4, "C19", cell_text, text_wrap)
write_comment!(ws4, "C19", comment, Dict("width"=> 200, "height"=> 20))


###############################################################################
#
# Example 5. Demonstrates changes to the cell comment position.
#
set_column!(ws5, "C:C", 25)
set_row!(ws5, 2, 50)
set_row!(ws5, 5, 50)
set_row!(ws5, 8, 50)
set_row!(ws5, 11, 50)

show_comments!(ws5, )

cell_text = "This cell comment is in the default position."
comment = "Hello."

write!(ws5, "C3", cell_text, text_wrap)
write_comment!(ws5, "C3", comment)

cell_text = "This cell comment has been moved to another cell."

write!(ws5, "C6", cell_text, text_wrap)
write_comment!(ws5, "C6", comment, Dict("start_cell"=> "E4"))

cell_text = "This cell comment has been moved to another cell."

write!(ws5, "C9", cell_text, text_wrap)
write_comment!(ws5, "C9", comment, Dict("start_row"=> 8, "start_col"=> 4))

cell_text = "This cell comment has been shifted within its default cell."

write!(ws5, "C12", cell_text, text_wrap)
write_comment!(ws5, "C12", comment, Dict("x_offset"=> 30, "y_offset"=> 12))


###############################################################################
#
# Example 6. Demonstrates changes to the comment background color.
#
set_column!(ws6, "C:C", 25)
set_row!(ws6, 2, 50)
set_row!(ws6, 5, 50)
set_row!(ws6, 8, 50)

show_comments!(ws6, )

cell_text = "This cell comment has a different color."
comment = "Hello."

write!(ws6 ,"C3", cell_text, text_wrap)
write_comment!(ws6, "C3", comment, Dict("color"=> "green"))

cell_text = "This cell comment has the default color."

write!(ws6, "C6", cell_text, text_wrap)
write_comment!(ws6, "C6", comment)

cell_text = "This cell comment has a different color."

write!(ws6, "C9", cell_text, text_wrap)
write_comment!(ws6, "C9", comment, Dict("color"=> "#CCFFCC"))


###############################################################################
#
# Example 7. Demonstrates how to set the cell comment author.
#
set_column!(ws7, "C:C", 30)
set_row!(ws7, 2, 50)
set_row!(ws7, 5, 50)
set_row!(ws7, 8, 50)

author = ""
cell = "C3"

cell_text = "Move the mouse over this cell and you will see \"Cell commented by (blank)\" in the status bar at the bottom"

comment = "Hello."

write!(ws7, cell, cell_text, text_wrap)
write_comment!(ws7, cell, comment)

author = "Julia"
cell = "C6"
cell_text = "Move the mouse over this cell and you will see \"Cell commented by Julia\" in the status bar at the bottom"

write!(ws7, cell, cell_text, text_wrap)
write_comment!(ws7, cell, comment, Dict("author"=> author))


###############################################################################
#
# Example 8. Demonstrates the need to explicitly set the row height.
#
# Set up some formatting.
set_column!(ws8, "C:C", 25)
set_row!(ws8, 2, 80)

show_comments!(ws8)

cell_text = "The height of this row has been adjusted explicitly using set_row(). The size of the comment box is adjusted accordingly by XlsxWriter."

comment = "Hello."

write!(ws8, "C3", cell_text, text_wrap)
write_comment!(ws8, "C3", comment)

cell_text = "The height of this row has been adjusted by Excel due to the text wrap property being set. Unfortunately this means that the height of the row is unknown to XlsxWriter at run time and thus the comment box is stretched as well.\n\nUse set_row() to specify the row height explicitly to avoid this problem."

write!(ws8, "C6", cell_text, text_wrap)
write_comment!(ws8, "C6", comment)

close(wb)

###############################################################################
#
# An example of writing cell comments to a worksheet using XlsxWriter.
#
# For more advanced comment options see comments2.jl
#
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter

include("../XlsxWriter.jl")

using XlsxWriter

wb = Workbook("comments1.xlsx")
ws = add_worksheet!(wb)

write!(ws, "A1", "Hello")
write_comment!(ws, "A1", "This is a comment")

close(wb)

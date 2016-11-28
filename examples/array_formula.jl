#######################################################################
#
# Example of how to use Python and the XlsxWriter module to write
# simple array formulas.
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter

include("../XlsxWriter.jl")

using XlsxWriter

wb = workbook("array_formula.xlsx")

ws = add_worksheet!(wb)

# Write some test data.
write!(ws, "B1", 500)
write!(ws, "B2", 10)
write!(ws, "B5", 1)
write!(ws, "B6", 2)
write!(ws, "B7", 3)
write!(ws, "C1", 300)
write!(ws, "C2", 15)
write!(ws, "C5", 20234)
write!(ws, "C6", 21003)
write!(ws, "C7", 10000)

# Write an array formula that returns a single value
write_formula!(ws, "A1", "{=SUM(B1:C1*B2:C2)}")
# 9500 in A1

# Same as above but more verbose.
write_array_formula!(ws, "A2", "{=SUM(B1:C1*B2:C2)}")
# 9500 in A2

# Write an array formula that returns a range of values
write_array_formula!(ws, "A5:A7", "{=TREND(C5:C7,B5:B7)}")

close(wb)

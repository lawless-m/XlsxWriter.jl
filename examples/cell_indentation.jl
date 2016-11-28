#######################################################################
#
# Example of how to use Python and the XlsxWriter module to write
# simple array formulas.
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter

include("../XlsxWriter.jl")

using XlsxWriter

wb = workbook("cell_indentation.xlsx")

ws = add_worksheet!(wb)

indent1 = add_format!(wb, Dict("indent"=>1))
indent2 = add_format!(wb, Dict("indent"=>2))

set_column!(ws, "A:A", 40)

write!(ws, "A1", "This text is indented 1 level", indent1)
write!(ws, "A2", "This text is indented 2 levels", indent2)

close(wb)

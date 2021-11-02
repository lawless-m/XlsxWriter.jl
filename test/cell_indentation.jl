#######################################################################
# A simple formatting example using XlsxWriter.
# This program demonstrates the indentation cell format.
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter


using XlsxWriter

function test()

	wb = Workbook("cell_indentation.xlsx")

	ws = add_worksheet!(wb)

	indent1 = add_format!(wb, Dict("indent"=>1))
	indent2 = add_format!(wb, Dict("indent"=>2))

	set_column!(ws, "A:A", 40)

	write!(ws, "A1", "This text is indented 1 level", fmt=indent1)
	write!(ws, "A2", "This text is indented 2 levels", fmt=indent2)

	close(wb)
	isfile("cell_indentation.xlsx")
end

test()

##############################################################################
#
# Simple Python program to benchmark several Python Excel writing modules.
#
# python bench_excel_writers.jl [num_rows] [num_cols]
#
# Original code
# Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
#

using XlsxWriter

row_max = 10000
col_max = 50

function time_xlsxwriter(const_mem=false)
	rm = Int64(floor(row_max)/2)

    wb = const_mem ? Workbook("xlsxwriter_const_mem.xlsx", Dict("constant_memory"=>true)) : Workbook("xlsxwriter.xlsx")
    ws = add_worksheet!(wb)

    for row in 1:rm
        for col in 1:col_max
            write_string!(ws, row * 2, col, "Row: $row Col: $col")
		end
        for col in 1:col_max
            write_number!(ws, row * 2 + 1, col, row + col)
		end
	end
    close(wb)
end	

time_xlsxwriter(false)

println("False")
@time time_xlsxwriter(false)
println("True")
@time time_xlsxwriter(true)

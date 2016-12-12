#######################################################################
#
# An example showing all 48 default chart styles available in Excel 2007
# using Python and XlsxWriter. Note, these styles are not the same as
# the styles available in Excel 2013.

# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter

include("../XlsxWriter.jl")
using XlsxWriter

wb = Workbook("chart_styles.xlsx")

# sythesis Python's title string function 
title(words) = join([ucfirst(t) for t in split(words, ' ')], " ")

# Show the styles for all of these chart types.
for chart_type in ["column" "area" "line" "pie"]
    # Add a worksheet for each chart type.
	ws = add_worksheet!(wb, title(chart_type))
    set_zoom!(ws, 30)
    style_number = 1

    # Create 48 charts, each with a different style.
    for row_num in 0:16:90, col_num in 0:9:64 
		chart = add_chart!(wb, Dict("type"=> chart_type))
		add_series!(chart, Dict("values"=> "=Data!\$A\$1:\$A\$6"))
		set_title!(chart, Dict("name"=> "Style $style_number"))
		set_legend!(chart, Dict("none"=> true))
		set_style!(chart, style_number)
		insert_chart!(ws, row_num, col_num , chart)
		style_number += 1
	end
end	
# Create a worksheet with data for the charts.
ws = add_worksheet!(wb, "Data")
data = [10, 40, 50, 20, 10, 50]
write_column!(ws, "A1", data)
hide!(ws)

close(wb)

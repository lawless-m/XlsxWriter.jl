#######################################################################
#
# An example of creating an Excel charts with a date axis using
# XlsxWriter.
#
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter

using Dates
using XlsxWriter

function test()

	wb = Workbook("chart_date_axis.xlsx")
	ws = add_worksheet!(wb)

	# Add a format for the headings.

	chart = add_chart!(wb, Dict("type"=> "line"))
	date_format = add_format!(wb, Dict("num_format"=> "dd/mm/yyyy"))

	# Widen the first column to display the dates.
	set_column!(ws, "A:A", 12)

	# Some data to be plotted in the worksheet.
	dates = [Date(2013, 1, 1),
			 Date(2013, 1, 2),
			 Date(2013, 1, 3),
			 Date(2013, 1, 4),
			 Date(2013, 1, 5),
			 Date(2013, 1, 6),
			 Date(2013, 1, 7),
			 Date(2013, 1, 8),
			 Date(2013, 1, 9),
			 Date(2013, 1, 10)]

	values = [10, 30, 20, 40, 20, 60, 50, 40, 30, 30]

	# Write the date to the worksheet.
	write_column!(ws, "A1", dates, fmt=date_format)
	write_column!(ws, "B1", values)

	# Add a series to the chart.
	add_series!(chart, Dict(
		"categories"=> "=Sheet1!\$A\$1:\$A\$10",
		"values"=> "=Sheet1!\$B\$1:\$B\$10",
	))

	# Configure the X axis as a Date axis and set the max and min limits.
	set_x_axis!(chart, Dict(
		"date_axis"=> true,
		"min"=> Date(2013, 1, 2),
		"max"=> Date(2013, 1, 9),
	))

	# Turn off the legend.
	set_legend!(chart, Dict("none"=> true))

	# Insert the chart into the worksheet.
	insert_chart!(ws, "D2", chart)

	close(wb)
	isfile("chart_date_axis.xlsx")
end

test()

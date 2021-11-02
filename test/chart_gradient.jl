#######################################################################
# An example of creating an Excel charts with gradient fills using
#  XlsxWriter.
#
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.rg
# https://github.com/jmcnamara/XlsxWriter



using XlsxWriter

function test()

	wb = Workbook("chart_gradient.xlsx")
	ws = add_worksheet!(wb)
	bold = add_format!(wb, Dict("bold"=> 1))


	# Add the worksheet data that the charts will refer to.
	headings = ["Number" "Batch 1" "Batch 2"]
	data = [
		[2 3 4 5 6 7];
		[10 40 50 20 10 50];
		[30 60 70 50 40 30];
	]

	write_row!(ws, "A1", headings, fmt=bold)
	write_column!(ws, "A2", data[1, :])
	write_column!(ws, "B2", data[2, :])
	write_column!(ws, "C2", data[3, :])


	# Create a new column chart.
	chart = add_chart!(wb, Dict("type"=> "column"))

	# Configure the first series, including a gradient.
	add_series!(chart, Dict(
		"name"=>       "=Sheet1!\$B\$1",
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$B\$2:\$B\$7",
		"gradient"=>   Dict("colors"=> ["#963735", "#F1DCDB"])
	))

	# Configure the second series, including a gradient.
	add_series!(chart, Dict(
		"name"=>       "=Sheet1!\$C\$1",
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$C\$2:\$C\$7",
		"gradient"=>   Dict("colors"=> ["#E36C0A", "#FCEADA"])
	))

	# Set a gradient for the plotarea.
	set_plotarea!(chart, Dict(
		"gradient"=> Dict("colors"=> ["#FFEFD1", "#F0EBD5", "#B69F66"])
	))


	# Add some axis labels.
	set_x_axis!(chart, Dict("name"=> "Test number"))
	set_y_axis!(chart, Dict("name"=> "Sample length (mm)"))

	# Turn off the chart legend.
	set_legend!(chart, Dict("none"=> true))

	# Insert the chart into the worksheet.
	insert_chart!(ws, "E2", chart)

	close(wb)
	isfile("chart_gradient.xlsx")
end

test()

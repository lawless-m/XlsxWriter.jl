#######################################################################
#
# An example of creating Excel Radar charts with XlsxWriter.
#
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter



using XlsxWriter

function test()


	wb = Workbook("chart_radar.xlsx")
	ws = add_worksheet!(wb)
	# Formats used in the workbook.
	bold = add_format!(wb, Dict("bold"=> 1))

	# Add the worksheet data that the charts will refer to.
	headings = ["Number", "Batch 1", "Batch 2"]
	data = [
		[2 3 4 5 6 7];
		[30 60 70 50 40 30];
		[25 40 50 30 50 40];
	]

	write_row!(ws, "A1", headings, fmt=bold)
	write_column!(ws, "A2", data[1,:])
	write_column!(ws, "B2", data[2,:])
	write_column!(ws, "C2", data[3,:])

	#######################################################################
	#
	# Create a new radar chart.
	#
	chart1 = add_chart!(wb, Dict("type"=> "radar"))

	# Configure the first series.
	add_series!(chart1, Dict(
		"name"=>       "=Sheet1!\$B\$1",
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$B\$2:\$B\$7",
	))

	# Configure second series. Note use of alternative syntax to define ranges.
	add_series!(chart1, Dict(
		"name"=>       ["Sheet1", 0, 2],
		"categories"=> ["Sheet1", 1, 0, 6, 0],
		"values"=>     ["Sheet1", 1, 2, 6, 2],
	))

	# Add a chart title and some axis labels.
	set_title!(chart1, Dict("name"=> "Results of sample analysis"))
	set_x_axis!(chart1, Dict("name"=> "Test number"))
	set_y_axis!(chart1, Dict("name"=> "Sample length (mm)"))

	# Set an Excel chart style.
	set_style!(chart1, 11)

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "D2", chart1, Dict("x_offset"=> 25, "y_offset"=> 10))

	#######################################################################
	#
	# Create a radar chart with markers chart sub-type.
	#
	chart2 = add_chart!(wb, Dict("type"=> "radar", "subtype"=> "with_markers"))

	# Configure the first series.
	add_series!(chart2, Dict(
		"name"=>       "=Sheet1!\$B\$1",
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$B\$2:\$B\$7",
	))

	# Configure second series.
	add_series!(chart2, Dict(
		"name"=>       "=Sheet1!\$C\$1",
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$C\$2:\$C\$7",
	))

	# Add a chart title and some axis labels.
	set_title!(chart2, Dict("name"=> "Radar Chart With Markers"))
	set_x_axis!(chart2, Dict("name"=> "Test number"))
	set_y_axis!(chart2, Dict("name"=> "Sample length (mm)"))

	# Set an Excel chart style.
	set_style!(chart2, 12)

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "D18", chart2, Dict("x_offset"=> 25, "y_offset"=> 10))

	#######################################################################
	#
	# Create a filled radar chart sub-type.
	#
	chart3 = add_chart!(wb, Dict("type"=> "radar", "subtype"=> "filled"))

	# Configure the first series.
	add_series!(chart3, Dict(
		"name"=>       "=Sheet1!\$B\$1",
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$B\$2:\$B\$7",
	))

	# Configure second series.
	add_series!(chart3, Dict(
		"name"=>       "=Sheet1!\$C\$1",
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$C\$2:\$C\$7",
	))

	# Add a chart title and some axis labels.
	set_title!(chart3, Dict("name"=> "Filled Radar Chart"))
	set_x_axis!(chart3, Dict("name"=> "Test number"))
	set_y_axis!(chart3, Dict("name"=> "Sample length (mm)"))

	# Set an Excel chart style.
	set_style!(chart3, 13)

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "D34", chart3, Dict("x_offset"=> 25, "y_offset"=> 10))

	close(wb)
	isfile("chart_radar.xlsx")
end

test()

#######################################################################
#
# An example of creating Excel Scatter charts with  XlsxWriter.
#
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter



using XlsxWriter

function test()

	wb = Workbook("chart_scatter.xlsx")
	ws = add_worksheet!(wb)
	bold = add_format!(wb, Dict("bold"=> 1))


	# Add the worksheet data that the charts will refer to.
	headings = ["Number", "Batch 1", "Batch 2"]
	data = [
		[2 3 4 5 6 7];
		[10 40 50 20 10 50];
		[30 60 70 50 40 30];
	]

	write_row!(ws, "A1", headings, fmt=bold)
	write_column!(ws, "A2", data[1,:])
	write_column!(ws, "B2", data[2,:])
	write_column!(ws, "C2", data[3,:])


	#######################################################################
	#
	# Create a new scatter chart.
	#
	chart1 = add_chart!(wb, Dict("type"=> "scatter"))

	# Configure the first series.
	add_series!(chart1, Dict(
		"name"=> "=Sheet1!\$B\$1",
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=> "=Sheet1!\$B\$2:\$B\$7",
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
	# Create a scatter chart sub-type with straight lines and markers.
	#
	chart2 = add_chart!(wb, Dict("type"=> "scatter",
								 "subtype"=> "straight_with_markers"))

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
	set_title!(chart2, Dict("name"=> "Straight line with markers"))
	set_x_axis!(chart2, Dict("name"=> "Test number"))
	set_y_axis!(chart2, Dict("name"=> "Sample length (mm)"))

	# Set an Excel chart style.
	set_style!(chart2, 12)

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "D18", chart2, Dict("x_offset"=> 25, "y_offset"=> 10))

	#######################################################################
	#
	# Create a scatter chart sub-type with straight lines and no markers.
	#
	chart3 = add_chart!(wb, Dict("type"=> "scatter",
								 "subtype"=> "straight"))

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
	set_title!(chart3, Dict("name"=> "Straight line"))
	set_x_axis!(chart3, Dict("name"=> "Test number"))
	set_y_axis!(chart3, Dict("name"=> "Sample length (mm)"))

	# Set an Excel chart style.
	set_style!(chart3, 13)

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "D34", chart3, Dict("x_offset"=> 25, "y_offset"=> 10))

	#######################################################################
	#
	# Create a scatter chart sub-type with smooth lines and markers.
	#
	chart4 = add_chart!(wb, Dict("type"=> "scatter",
								 "subtype"=> "smooth_with_markers"))

	# Configure the first series.
	add_series!(chart4, Dict(
		"name"=>       "=Sheet1!\$B\$1",
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$B\$2:\$B\$7",
	))

	# Configure second series.
	add_series!(chart4, Dict(
		"name"=>       "=Sheet1!\$C\$1",
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$C\$2:\$C\$7",
	))

	# Add a chart title and some axis labels.
	set_title!(chart4, Dict("name"=> "Smooth line with markers"))
	set_x_axis!(chart4, Dict("name"=> "Test number"))
	set_y_axis!(chart4, Dict("name"=> "Sample length (mm)"))

	# Set an Excel chart style.
	set_style!(chart4, 14)

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "D50", chart4, Dict("x_offset"=> 25, "y_offset"=> 10))

	#######################################################################
	#
	# Create a scatter chart sub-type with smooth lines and no markers.
	#
	chart5 = add_chart!(wb, Dict("type"=> "scatter",
								 "subtype"=> "smooth"))

	# Configure the first series.
	add_series!(chart5, Dict(
		"name"=>       "=Sheet1!\$B\$1",
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$B\$2:\$B\$7",
	))

	# Configure second series.
	add_series!(chart5, Dict(
		"name"=>       "=Sheet1!\$C\$1",
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$C\$2:\$C\$7",
	))

	# Add a chart title and some axis labels.
	set_title!(chart5, Dict("name"=> "Smooth line"))
	set_x_axis!(chart5, Dict("name"=> "Test number"))
	set_y_axis!(chart5, Dict("name"=> "Sample length (mm)"))

	# Set an Excel chart style.
	set_style!(chart5, 15)

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "D66", chart5, Dict("x_offset"=> 25, "y_offset"=> 10))

	close(wb)
	isfile("chart_scatter.xlsx")
end

test()

#######################################################################
#
# A demo of an various Excel chart data tools that are available via
# an XlsxWriter chart.
#
# These include, Trendlines, Data Labels, Error Bars, Drop Lines,
# High-Low Lines and Up-Down Bars.
#
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter
#


using XlsxWriter

function test()

	wb = Workbook("chart_data_table.xlsx")
	ws = add_worksheet!(wb)

	# Add a format for the headings.
	bold = add_format!(wb, Dict("bold"  => 1))

	# Add the worksheet data that the charts will refer to.
	headings = ["Number" "Data 1" "Data 2"]
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
	# Trendline example.
	#
	# Create a Line chart.
	chart1 = add_chart!(wb, Dict("type"=> "line"))

	# Configure the first series with a polynomial trendline.
	add_series!(chart1, Dict(
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$B\$2:\$B\$7",
		"trendline"=> Dict(
			"type"=> "polynomial",
			"order"=> 3,
		),
	))

	# Configure the second series with a moving average trendline.
	add_series!(chart1, Dict(
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$C\$2:\$C\$7",
		"trendline"=> Dict("type"=> "linear"),
	))

	# Add a chart title. and some axis labels.
	set_title!(chart1, Dict("name"=> "Chart with Trendlines"))

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "D2", chart1, Dict("x_offset"=> 25, "y_offset"=> 10))

	#######################################################################
	#
	# Data Labels and Markers example.
	#
	# Create a Line chart.
	chart2 = add_chart!(wb, Dict("type"=> "line"))

	# Configure the first series.
	add_series!(chart2, Dict(
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$B\$2:\$B\$7",
		"data_labels"=> Dict("value"=> 1),
		"marker"=> Dict("type"=> "automatic"),
	))

	# Configure the second series.
	add_series!(chart2, Dict(
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$C\$2:\$C\$7",
	))

	# Add a chart title. and some axis labels.
	set_title!(chart2, Dict("name"=> "Chart with Data Labels and Markers"))

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "D18", chart2, Dict("x_offset"=> 25, "y_offset"=> 10))

	#######################################################################
	#
	# Error Bars example.
	#
	# Create a Line chart.
	chart3 = add_chart!(wb, Dict("type"=> "line"))

	# Configure the first series.
	add_series!(chart3, Dict(
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$B\$2:\$B\$7",
		"y_error_bars"=> Dict("type"=> "standard_error"),
	))

	# Configure the second series.
	add_series!(chart3, Dict(
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=> "=Sheet1!\$C\$2:\$C\$7",
	))

	# Add a chart title. and some axis labels.
	set_title!(chart3, Dict("name"=> "Chart with Error Bars"))

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "D34", chart3, Dict("x_offset"=> 25, "y_offset"=> 10))

	#######################################################################
	#
	# Up-Down Bars example.
	#
	# Create a Line chart.
	chart4 = add_chart!(wb, Dict("type"=> "line"))

	# Add the Up-Down Bars.
	set_up_down_bars!(chart4)

	# Configure the first series.
	add_series!(chart4, Dict(
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$B\$2:\$B\$7",
	))

	# Configure the second series.
	add_series!(chart4, Dict(
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$C\$2:\$C\$7",
	))

	# Add a chart title. and some axis labels.
	set_title!(chart4, Dict("name"=> "Chart with Up-Down Bars"))

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "D50", chart4, Dict("x_offset"=> 25, "y_offset"=> 10))

	#######################################################################
	#
	# High-Low Lines example.
	#
	# Create a Line chart.
	chart5 = add_chart!(wb, Dict("type"=> "line"))

	# Add the High-Low lines.
	set_high_low_lines!(chart5)

	# Configure the first series.
	add_series!(chart5, Dict(
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$B\$2:\$B\$7",
	))

	# Configure the second series.
	add_series!(chart5, Dict(
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$C\$2:\$C\$7",
	))

	# Add a chart title. and some axis labels.
	set_title!(chart5, Dict("name"=> "Chart with High-Low Lines"))

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "D66", chart5, Dict("x_offset"=> 25, "y_offset"=> 10))

	#######################################################################
	#
	# Drop Lines example.
	#
	# Create a Line chart.
	chart6 = add_chart!(wb, Dict("type"=> "line"))

	# Add Drop Lines.
	set_drop_lines!(chart6)

	# Configure the first series.
	add_series!(chart6, Dict(
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$B\$2:\$B\$7",
	))

	# Configure the second series.
	add_series!(chart6, Dict(
		"categories"=> "=Sheet1!\$A\$2:\$A\$7",
		"values"=>     "=Sheet1!\$C\$2:\$C\$7",
	))

	# Add a chart title. and some axis labels.
	set_title!(chart6, Dict("name"=> "Chart with Drop Lines"))

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "D82", chart6, Dict("x_offset"=> 25, "y_offset"=> 10))

	close(wb)
	isfile("chart_data_table.xlsx")
end

test()

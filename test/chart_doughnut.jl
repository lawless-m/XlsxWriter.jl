#######################################################################
#
# An example of creating Excel Doughnut charts with Python and XlsxWriter.
#
# The demo also shows how to set segment colors. It is possible to
# define chart colors for most types of XlsxWriter charts
# via the add_series() method. However, Pie/Doughnut charts are a special
# case since each segment is represented as a point so it is necessary to
# assign formatting to each point in the series.

# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter


using XlsxWriter

function test()

	wb = Workbook("chart_doughnut.xlsx")
	ws = add_worksheet!(wb)
	bold = add_format!(wb, Dict("bold"=> 1))

	# Add the worksheet data that the charts will refer to.
	headings = ["Category" "Values"]
	data = [
		["Glazed" "Chocolate" "Cream"];
		[50 35 15];
	]

	write_row!(ws, "A1", headings, fmt=bold)
	write_column!(ws, "A2", data[1,:])
	write_column!(ws, "B2", data[2,:])

	#######################################################################
	#
	# Create a new chart object.
	#
	chart1 = add_chart!(wb, Dict("type"=> "doughnut"))

	# Configure the series. Note the use of the list syntax to define ranges:
	add_series!(chart1, Dict(
		"name"=>       "Doughnut sales data",
		"categories"=> ["Sheet1", 1, 0, 3, 0],
		"values"=>     ["Sheet1", 1, 1, 3, 1]
	))

	# Add a title.
	set_title!(chart1, Dict("name"=> "Popular Doughnut Types"))

	# Set an Excel chart style. Colors with white outline and shadow.
	set_style!(chart1, 10)

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "C2", chart1, Dict("x_offset"=> 25, "y_offset"=> 10))

	#######################################################################
	#
	# Create a Doughnut chart with user defined segment colors.
	#

	# Create an example Doughnut chart like above.
	chart2 = add_chart!(wb, Dict("type"=> "doughnut"))

	# Configure the series and add user defined segment colors.
	add_series!(chart2, Dict(
		"name"=> "Doughnut sales data",
		"categories"=> "=Sheet1!\$A\$2:\$A\$4",
		"values"=>     "=Sheet1!\$B\$2:\$B\$4",
		"points"=> [
			Dict("fill"=> Dict("color"=> "#FA58D0")),
			Dict("fill"=> Dict("color"=> "#61210B")),
			Dict("fill"=> Dict("color"=> "#F5F6CE")),
		];
	))

	# Add a title.
	set_title!(chart2, Dict("name"=> "Doughnut Chart with user defined colors"))

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "C18", chart2, Dict("x_offset"=> 25, "y_offset"=> 10))

	#######################################################################
	#
	# Create a Doughnut chart with rotation of the segments.
	#

	# Create an example Doughnut chart like above.
	chart3 = add_chart!(wb, Dict("type"=> "doughnut"))

	# Configure the series.
	add_series!(chart3, Dict(
		"name"=> "Doughnut sales data",
		"categories"=> "=Sheet1!\$A\$2:\$A\$4",
		"values"=>     "=Sheet1!\$B\$2:\$B\$4",
	))

	# Add a title.
	set_title!(chart3, Dict("name"=> "Doughnut Chart with segment rotation"))

	# Change the angle/rotation of the first segment.
	set_rotation!(chart3, 90)

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "C34", chart3, Dict("x_offset"=> 25, "y_offset"=> 10))


	#######################################################################
	#
	# Create a Doughnut chart with user defined hole size and other options.
	#

	# Create an example Doughnut chart like above.
	chart4 = add_chart!(wb, Dict("type"=> "doughnut"))

	# Configure the series.
	add_series!(chart4, Dict(
		"name"=> "Doughnut sales data",
		"categories"=> "=Sheet1!\$A\$2:\$A\$4",
		"values"=>     "=Sheet1!\$B\$2:\$B\$4",
		"points"=> [
			Dict("fill"=> Dict("color"=> "#FA58D0")),
			Dict("fill"=> Dict("color"=> "#61210B")),
			Dict("fill"=> Dict("color"=> "#F5F6CE")),
		];
	))

	# Set a 3D style.
	set_style!(chart4, 26)

	# Add a title.
	set_title!(chart4, Dict("name"=> "Doughnut Chart with options applied"))

	# Change the angle/rotation of the first segment.
	set_rotation!(chart4, 28)

	# Change the hole size.
	set_hole_size!(chart4, 33)

	# Insert the chart into the worksheet (with an offset).
	insert_chart!(ws, "C50", chart4, Dict("x_offset"=> 25, "y_offset"=> 10))


	close(wb)
	isfile("chart_doughnut.xlsx")
end

test()

#######################################################################
#
# An example of an Excel chart with patterns using XlsxWriter.
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter



using XlsxWriter

function test()

	wb = Workbook("chart_pattern.xlsx")
	ws = add_worksheet!(wb)
	# Formats used in the workbook.
	bold = add_format!(wb, Dict("bold"=> 1))

	# Add the worksheet data that the charts will refer to.
	headings = ["Shingle", "Brick"]
	data = [
		[105 150 130 90 ];
		[50  120 100 110];
	]

	write_row!(ws, "A1", headings, fmt=bold)
	write_column!(ws, "A2", data[1,:])
	write_column!(ws, "B2", data[2,:])

	# Create a new Chart object.
	chart = add_chart!(wb, Dict("type"=> "column"))

	# Configure the charts. Add two series with patterns. The gap is used to make
	# the patterns more visible.
	add_series!(chart, Dict(
		"name"=>   "=Sheet1!\$A\$1",
		"values"=> "=Sheet1!\$A\$2:\$A\$5",
		"pattern"=> Dict(
			"pattern"=>  "shingle",
			"fg_color"=> "#804000",
			"bg_color"=> "#c68c53"
		),
		"border"=>  Dict("color"=> "#804000"),
		"gap"=>     70,
	))

	add_series!(chart, Dict(
		"name"=>   "=Sheet1!\$B\$1",
		"values"=> "=Sheet1!\$B\$2:\$B\$5",
		"pattern"=> Dict(
			"pattern"=>  "horizontal_brick",
			"fg_color"=> "#b30000",
			"bg_color"=> "#ff6666"
		),
		"border"=>  Dict("color"=> "#b30000"),
	))

	# Add a chart title and some axis labels.
	set_title!(chart, Dict("name"=> "Cladding types"))
	set_x_axis!(chart, Dict("name"=> "Region"))
	set_y_axis!(chart, Dict("name"=> "Number of houses"))

	# Insert the chart into the worksheet.
	insert_chart!(ws, "D2", chart)

	close(wb)
	isfile("chart_pattern.xlsx")
end

test()

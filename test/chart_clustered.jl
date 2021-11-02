#######################################################################
# A demo of a clustered category chart in XlsxWriter.
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter


using XlsxWriter

function test()

	wb = Workbook("chart_clustered.xlsx")
	ws = add_worksheet!(wb)
	bold = add_format!(wb, Dict("bold" => 1))

	# Add the worksheet data that the charts will refer to.
	headings = ["Types" "Sub Type" "Value 1" "Value 2" "Value 3"]

	data = [
		["Type 1" "Sub Type A" 5000      8000      6000];
		[""       "Sub Type B" 2000      3000      4000];
		[""       "Sub Type C" 250       1000      2000];
		["Type 2" "Sub Type D" 6000      6000      6500];
		[""       "Sub Type E" 500       300        200];
	]

	write_row!(ws, "A1", headings, fmt=bold)

	for row_num in 1:size(data,1)
		write_row!(ws, row_num, 0, data[row_num,:])
	end
	# Create a new chart object. In this case an embedded chart.
	chart = add_chart!(wb, Dict("type" => "column"))

	# Configure the series. Note, that the categories are 2D ranges (from column A
	# to column B). This creates the clusters. The series are shown as formula
	# strings for clarity but you can also use the list syntax. See the docs.
	add_series!(chart, Dict(
		"name" => "=Sheet1!\$C\$1",
		"categories" => "=Sheet1!\$A\$2:\$B\$6",
		"values" => "=Sheet1!\$C\$2:\$C\$6",
	))

	add_series!(chart, Dict(
		"name" => "=Sheet1!\$D\$1",
		"categories" => "=Sheet1!\$A\$2:\$B\$6",
		"values" => "=Sheet1!\$D\$2:\$D\$6",
	))

	add_series!(chart, Dict(
		"name" => "=Sheet1!\$E\$1",
		"categories" => "=Sheet1!\$A\$2:\$B\$6",
		"values" => "=Sheet1!\$E\$2:\$E\$6",
	))

	# Set the Excel chart style.
	set_style!(chart, 37)

	# Turn off the legend.
	set_legend!(chart, Dict("position" => "none"))

	# Insert the chart into the worksheet.
	insert_chart!(ws, "G3", chart)

	close(wb)

	isfile("chart_clustered.xlsx")
end

test()


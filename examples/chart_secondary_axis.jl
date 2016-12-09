#######################################################################
#
# An example of creating an Excel Line chart with a secondary axis
# using XlsxWriter.
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter

include("../XlsxWriter.jl")

using XlsxWriter

wb = Workbook("chart_secondary_axis.xlsx")
ws = add_worksheet!(wb)
bold = add_format!(wb, Dict("bold"=> 1))

# Add the worksheet data that the charts will refer to.
headings = ["Aliens", "Humans"]
data = [
    [2 3 4 5 6 7];
    [10 40 50 20 10 50];
]

write_row!(ws, "A1", headings, bold)
write_column!(ws, "A2", data[1,:])
write_column!(ws, "B2", data[2,:])


# Create a new chart object. In this case an embedded chart.
chart = add_chart!(wb, Dict("type"=> "line"))

# Configure a series with a secondary axis
add_series!(chart, Dict(
    "name"=>   "=Sheet1!\$A\$1",
    "values"=> "=Sheet1!\$A\$2:\$A\$7",
    "y2_axis"=> 1,
))

add_series!(chart, Dict(
    "name"=>   "=Sheet1!\$B\$1",
    "values"=> "=Sheet1!\$B\$2:\$B\$7",
))

set_legend!(chart, Dict("position"=> "right"))

# Add a chart title and some axis labels.
set_title!(chart, Dict("name"=> "Survey results"))
set_x_axis!(chart, Dict("name"=> "Days", ))
set_y_axis!(chart, Dict("name"=> "Population", "major_gridlines"=> Dict("visible"=> 0)))
set_y2_axis!(chart, Dict("name"=> "Laser wounds"))

# Insert the chart into the worksheet (with an offset).
insert_chart!(ws, "D2", chart, Dict("x_offset"=> 25, "y_offset"=> 10))

close(wb)

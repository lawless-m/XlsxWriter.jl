#######################################################################
#
# An example of creating Excel Line charts with XlsxWriter.
#
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@ org
# https://github.com/jmcnamara/XlsxWriter

include("../XlsxWriter.jl")

using XlsxWriter

wb = Workbook("chart_line.xlsx")
ws = add_worksheet!(wb)
bold = add_format!(wb, Dict("bold"=> 1))


# Add the worksheet data that the charts will refer to.
headings = ["Number", "Batch 1", "Batch 2"]
data = [
    [2 3 4 5 6 7];
    [10 40 50 20 10 50];
    [30 60 70 50 40 30];
]

write_row!(ws, "A1", headings, bold)
write_column!(ws, "A2", data[1,:])
write_column!(ws, "B2", data[2,:])
write_column!(ws, "C2", data[3,:])

# Create a new chart object. In this case an embedded chart.
chart1 = add_chart!(wb, Dict("type"=> "line"))

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

# Set an Excel chart style. Colors with white outline and shadow.
set_style!(chart1, 10)

# Insert the chart into the worksheet (with an offset).
insert_chart!(ws, "D2", chart1, Dict("x_offset"=> 25, "y_offset"=> 10))


close(wb)

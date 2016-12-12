#######################################################################
#
# An example of creating an Excel chart in a chartsheet with XlsxWriter.
#
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https:// com/jmcnamara/XlsxWriter

include("../XlsxWriter.jl")

using XlsxWriter

wb = Workbook("chartsheet.xlsx")
ws = add_worksheet!(wb)
bold = add_format!(wb, Dict("bold"=>1))

# Add a chartsheet. A worksheet that only holds a chart.
cs = add_chartsheet!(wb)

# Add the worksheet data that the charts will refer to.
headings = ["Number" "Batch 1" "Batch 2"]
data = [
    [2 3 4 5 6 7];
    [10 40 50 20 10 50];
    [30 60 70 50 40 30];
]

write_row!(ws, "A1", headings, bold)
write_column!(ws, "A2", data[1,:])
write_column!(ws, "B2", data[2,:])
write_column!(ws, "C2", data[3,:])


# Create a new bar chart.
chart1 = add_chart!(wb, Dict("type"=> "bar"))

# Configure the first series.
add_series!(chart1, Dict(
    "name"=>       "=Sheet1!\$B\$1",
    "categories"=> "=Sheet1!\$A\$2:\$A\$7",
    "values"=>     "=Sheet1!\$B\$2:\$B\$7",
))

# Configure a second series. Note use of alternative syntax to define ranges.
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

# Add the chart to the chartsheet.
set_chart!(cs, chart1)

close(wb)

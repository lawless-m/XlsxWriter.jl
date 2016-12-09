######################################################################
# An example of a Combined chart in XlsxWriter.
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@ org
# https://github.com/jmcnamara/XlsxWriter

include("../XlsxWriter.jl")

using XlsxWriter

wb = Workbook("chart_combined.xlsx")
ws = add_worksheet!(wb)

# Add a format for the headings.
bold = add_format!(wb, Dict("bold"  => 1))


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

#
# In the first example we will create a combined column and line chart.
# They will share the same X and Y axes.
#

# Create a new column  This will use this as the primary chart.
column_chart1 = add_chart!(wb, Dict("type" => "column"))

# Configure the data series for the primary chart.
add_series!(column_chart1, Dict(
    "name" => "=Sheet1!\$B\$1",
    "categories" => "=Sheet1!\$A\$2:\$A\$7",
    "values" => "=Sheet1!\$B\$2:\$B\$7",
))

# Create a new column  This will use this as the secondary chart.
line_chart1 = add_chart!(wb, Dict("type" => "line"))

# Configure the data series for the secondary chart.
add_series!(line_chart1, Dict(
    "name" => "=Sheet1!\$C\$1",
    "categories" => "=Sheet1!\$A\$2:\$A\$7",
    "values" => "=Sheet1!\$C\$2:\$C\$7",
))

# Combine the 
combine!(column_chart1, line_chart1)

# Add a chart title and some axis  Note, this is done via the
# primary chart.
set_title!(column_chart1, Dict( "name" => "Combined chart - same Y axis"))
set_x_axis!(column_chart1, Dict("name" => "Test number"))
set_y_axis!(column_chart1, Dict("name" => "Sample length (mm)"))

# Insert the chart into the worksheet
insert_chart!(ws, "E2", column_chart1)

#
# In the second example we will create a similar combined column and line
# chart except that the secondary chart will have a secondary Y axis.
#

# Create a new column  This will use this as the primary chart.
column_chart2 = add_chart!(wb, Dict("type" => "column"))

# Configure the data series for the primary chart.
add_series!(column_chart2, Dict(
    "name" => "=Sheet1!\$B\$1",
    "categories" => "=Sheet1!\$A\$2:\$A\$7",
    "values" => "=Sheet1!\$B\$2:\$B\$7",
))

# Create a new column  This will use this as the secondary chart.
line_chart2 = add_chart!(wb, Dict("type" => "line"))

# Configure the data series for the secondary  We also set a
# secondary Y axis via (y2_axis). This is the only difference between
# this and the first example, apart from the axis label below.
add_series!(line_chart2, Dict(
    "name" => "=Sheet1!\$C\$1",
    "categories" => "=Sheet1!\$A\$2:\$A\$7",
    "values" => "=Sheet1!\$C\$2:\$C\$7",
    "y2_axis" => true,
))

# Combine the charts.
combine!(column_chart2, line_chart2)

# Add a chart title and some axis labels.
set_title!(column_chart2, Dict(  "name" => "Combine chart - secondary Y axis"))
set_x_axis!(column_chart2, Dict( "name" => "Test number"))
set_y_axis!(column_chart2, Dict( "name" => "Sample length (mm)"))

# Note: the y2 properties are on the secondary chart.
set_y2_axis!(line_chart2, Dict("name" => "Target length (mm)"))

# Insert the chart into the worksheet
insert_chart!(ws, "E18", column_chart2)

close(wb)

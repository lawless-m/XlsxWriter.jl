#######################################################################
#
# An example of creating of a Pareto chart with  XlsxWriter.
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@ org
# https://github.com/jmcnamara/XlsxWriter

include("../XlsxWriter.jl")

using XlsxWriter

wb = Workbook("chart_pareto.xlsx")
ws = add_worksheet!(wb)
# Formats used in the workbook.
bold = add_format!(wb, Dict("bold"=> 1))

percent_format = add_format!(wb, Dict("num_format"=> "0.0%"))

# Widen the columns for visibility.
set_column!(ws, "A:A", 15)
set_column!(ws, "B:C", 10)

# Add the worksheet data that the charts will refer to.
headings = ["Reason", "Number", "Percentage"]

reasons = [
    "Traffic", "Child care", "Public Transport", "Weather",
    "Overslept", "Emergency",
]

numbers  = [60,   40,    20,  15,  10,    5]
percents = [0.44, 0.667, 0.8, 0.9, 0.967, 1]

write_row!(ws, "A1", headings, bold)
write_column!(ws, "A2", reasons)
write_column!(ws, "B2", numbers)
write_column!(ws, "C2", percents, percent_format)


# Create a new column chart. This will be the primary chart.
column_chart = add_chart!(wb, Dict("type"=> "column"))

# Add a series.
add_series!(column_chart, Dict(
    "categories"=> "=Sheet1!\$A\$2:\$A\$7",
    "values"=>     "=Sheet1!\$B\$2:\$B\$7",
))

# Add a chart title.
set_title!(column_chart, Dict("name"=> "Reasons for lateness"))

# Turn off the chart legend.
set_legend!(column_chart, Dict("position"=> "none"))

# Set the title and scale of the Y axes. Note, the secondary axis is set from
# the primary chart.
set_y_axis!(column_chart, Dict(
    "name"=> "Respondents (number)",
    "min"=> 0,
    "max"=> 120
))
set_y2_axis!(column_chart, Dict("max"=> 1))

# Create a new line chart. This will be the secondary chart.
line_chart = add_chart!(wb, Dict("type"=> "line"))

# Add a series, on the secondary axis.
add_series!(line_chart, Dict(
    "categories"=> "=Sheet1!\$A\$2:\$A\$7",
    "values"=>     "=Sheet1!\$C\$2:\$C\$7",
    "marker"=>     Dict("type"=> "automatic"),
    "y2_axis"=>    1,
))

# Combine the charts.
combine!(column_chart, line_chart)

# Insert the chart into the worksheet.
insert_chart!(ws, "F2", column_chart)

close(wb)

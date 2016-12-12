#######################################################################
#
# An example of creating Excel Stock charts with XlsxWriter.
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter

include("../XlsxWriter.jl")

using XlsxWriter

wb = Workbook("chart_stock.xlsx")
ws = add_worksheet!(wb)
bold = add_format!(wb, Dict("bold"=> 1))


date_format = add_format!(wb, Dict("num_format"=> "dd/mm/yyyy"))

chart = add_chart!(wb, Dict("type"=> "stock"))

# Add the worksheet data that the charts will refer to.
headings = ["Date", "High", "Low", "Close"]
data = [
    [Date(2007, 1, 1) Date(2007, 1, 2) Date(2007, 1, 3) Date(2007, 1, 4) Date(2007, 1, 5)];
    [27.2 25.03 19.05 20.34 18.5];
    [23.49 19.55 15.12 17.84 16.34];
    [25.45 23.05 17.32 20.45 17.34];
]

write_row!(ws, "A1", headings, bold)

for row in 1:5
    write!(ws, row, 0, data[1, row], date_format)
    write!(ws, row, 1, data[2, row])
    write!(ws, row, 2, data[3, row])
    write!(ws, row, 3, data[4, row])
end

set_column!(ws, "A:D", 11)

# Add a series for each of the High-Low-Close columns.
add_series!(chart, Dict(
    "categories"=> "=Sheet1!\$A\$2:\$A\$6",
    "values"=> "=Sheet1!\$B\$2:\$B\$6",
))

add_series!(chart, Dict(
    "categories"=> "=Sheet1!\$A\$2:\$A\$6",
    "values"=>     "=Sheet1!\$C\$2:\$C\$6",
))

add_series!(chart, Dict(
    "categories"=> "=Sheet1!\$A\$2:\$A\$6",
    "values"=> "=Sheet1!\$D\$2:\$D\$6",
))

# Add a chart title and some axis labels.
set_title!(chart, Dict("name"=> "High-Low-Close"))
set_x_axis!(chart, Dict("name"=> "Date"))
set_y_axis!(chart, Dict("name"=> "Share price"))

insert_chart!(ws, "E9", chart)

close(wb)

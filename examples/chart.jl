#######################################################################
# An example of a simple Excel chart with XlsxWriter.
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter

include("../XlsxWriter.jl")

using XlsxWriter

wb = workbook("chart.xlsx")

ws = add_worksheet!(wb)

chart = add_chart!(wb, Dict("type"=>"column"))

data = [
[1 2 3 4 5];
[2 4 6 8 10];
[3 6 9 12 15];
]


write_column!(ws, "A1", data[1,:])
write_column!(ws, "B1", data[2,:])
write_column!(ws, "C1", data[3,:])

add_series!(chart, Dict("values"=>"=Sheet1!\$A\$1:\$A\$5"))
add_series!(chart, Dict("values"=>"=Sheet1!\$B\$1:\$B\$5"))
add_series!(chart, Dict("values"=>"=Sheet1!\$C\$1:\$C\$5"))

insert_chart!(ws, "A7", chart)

close(wb)

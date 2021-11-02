# XLSXWriter - it's working again!

This is a wrapper around 

https://github.com/jmcnamara/XlsxWriter

which is the *best* Excel writing module for Python, hands down.

Today I finally got it to work as a proper Julia package so you can just Pkg.add("https://github.com/lawless-m/XlsxWriter.jl") and start using it, after a little more setup.

To make it work, it is up to you, dear user, to arrange for the Python code to be available

I've used an environment variable for the task : ENV["XLSXWRITER_PATH"]

and this should be set to the directory inside which the directory xlsxwriter can be found

So, for example, if you Clone the above code to /opt/XlsxWriter.py then set ENV["XLSXWRITER_PATH"] = "/opt/XlsxWriter.py"

otherwise Julia will throw an error to tell you to do that.

Excitingly, here in November 2021, Release 3.0.2 is coming up, with new stuff in it, so it looks like I have some work to do!

Here's the example from the XlxsWriter homepage but in Julia

	wb = Workbook("demo.xlsx")
	set_calc_mode!(wb, manual)
	ws = add_worksheet!(wb)

	# Widen the first column to make the text clearer.
	set_column!(ws, "A:A", 20)
	set_column!(ws, 3, 3, 21)
	set_column!(ws, "E:E", 22)

	# Add a bold format to use to highlight cells.
	bold = add_format!(wb, Dict("bold"=>true))
	set_font_name!(bold, "Courier New")
	set_font_size!(bold, 16)

	date_format = add_format!(wb, Dict("num_format"=>"d mmmm yyyy"))
	set_font_color!(date_format, "red")
	# Write some simple text.
	write!(ws, "A1", "Hello")

	# Text with formatting.
	write!(ws, 1,1, "World", fmt=bold)

	# Write some numbers, with row/column notation.
	write!(ws, 2, 2, 123)
	write!(ws, 3, 1, 123.456)
	write!(ws, 3, 2, true)
	write!(ws, 3, 3, now(), fmt=date_format)
	write!(ws, 3, 4, Url("http://localhost"))
	write!(ws, 3, 5, "=3 + 4")

	write_formula!(ws, 3, 6, "=13 + 14", fmt=bold, result=9)

	define_name!(wb, "duck", "=Sheet1!\$C\$3")
	write_formula!(ws, 3, 7, "=duck*2", result=8)

	write_row!(ws, 3, 7, ["6", 7, 8.8])
	write_column!(ws, 4, 7, ["46", 47, 48.8])
	write_matrix!(ws, 7, 1, [["105" "106" "107"]; ["201" 202 203]], fmt=bold)

	write_row!(ws, 10, 1, [-2, 2, 3, -1])
	add_sparkline!(ws, 10, 5, Dict("range"=>"Sheet1!B11:E11"))

	#freeze_panes!(ws, 1, 1)

	merge_range!(ws, 5, 9, 5, 12, "Merged")

	close(wb)

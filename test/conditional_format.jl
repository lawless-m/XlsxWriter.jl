###############################################################################
#
# Example of how to add conditional formatting to an XlsxWriter file.
#
# Conditional formatting allows you to apply a format to a cell or a
# range of cells based on certain criteria.

# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter


using XlsxWriter

function test()
	wb = Workbook("conditional_format.xlsx")

	ws1 = add_worksheet!(wb)
	ws2 = add_worksheet!(wb)
	ws3 = add_worksheet!(wb)
	ws4 = add_worksheet!(wb)
	ws5 = add_worksheet!(wb)
	ws6 = add_worksheet!(wb)
	ws7 = add_worksheet!(wb)
	ws8 = add_worksheet!(wb)

	# Add a format. Light red fill with dark red text.
	format1 = add_format!(wb, Dict("bg_color"=> "#FFC7CE", "font_color"=> "#9C0006"))

	# Add a format. Green fill with dark green text.
	format2 = add_format!(wb, Dict("bg_color"=> "#C6EFCE", "font_color"=> "#006100"))

	# Some sample data to run the conditional formatting against.
	data = [
		[34 72 38 30 75 48 75 66 84 86];
		[6 24 1 84 54 62 60 3 26 59];
		[28 79 97 13 85 93 93 22 5 14];
		[27 71 40 17 18 79 90 93 29 47];
		[88 25 33 23 67 1 59 79 47 36];
		[24 100 20 88 29 33 38 54 54 88];
		[6 57 88 28 10 26 37 7 41 48];
		[52 78 1 96 26 45 47 33 96 36];
		[60 54 81 66 81 90 80 93 12 55];
		[70 5 46 14 71 19 66 36 41 21];
	]
	###############################################################################
	#
	# Example 1.
	#
	caption = "Cells with values >= 50 are in light red. Values < 50 are in light green."

	# Write the data.
	write!(ws1, "A1", caption)

	for row in size(data, 2)
		write_row!(ws1, row + 2, 1, data[row,:])
	end

	# Write a conditional format over a range.
	conditional_format!(ws1, "B3:K12", Dict("type"=> "cell",
											 "criteria"=> ">=",
											 "value"=> 50,
											 "format"=> format1))

	# Write another conditional format over the same range.
	conditional_format!(ws1, "B3:K12", Dict("type"=> "cell",
											 "criteria"=> "<",
											 "value"=> 50,
											 "format"=> format2))
	###############################################################################
	#
	# Example 2.
	#
	caption = "Values between 30 and 70 are in light red. Values outside that range are in light green."

	write!(ws2, "A1", caption)

	for row in size(data, 2)
		write_row!(ws2, row + 2, 1, data[row,:])
	end

	conditional_format!(ws2, "B3:K12", Dict("type"=> "cell",
											 "criteria"=> "between",
											 "minimum"=> 30,
											 "maximum"=> 70,
											 "format"=> format1))

	conditional_format!(ws2, "B3:K12", Dict("type"=> "cell",
											 "criteria"=> "not between",
											 "minimum"=> 30,
											 "maximum"=> 70,
											 "format"=> format2))
	###############################################################################
	#
	# Example 3.
	#
	caption = "Duplicate values are in light red. Unique values are in light green."

	write!(ws3, "A1", caption)

	for row in size(data, 2)
		write_row!(ws3, row + 2, 1, data[row,:])
	end
	conditional_format!(ws3, "B3:K12", Dict("type"=> "duplicate",
											 "format"=> format1))

	conditional_format!(ws3, "B3:K12", Dict("type"=> "unique",
											 "format"=> format2))
	###############################################################################
	#
	# Example 4.
	#
	caption = "Above average values are in light red. Below average values are in light green."

	write!(ws4, "A1", caption)

	for row in size(data, 2)
		write_row!(ws4, row + 2, 1, data[row,:])
	end
	conditional_format!(ws4, "B3:K12", Dict("type"=> "average",
											 "criteria"=> "above",
											 "format"=> format1))

	conditional_format!(ws4, "B3:K12", Dict("type"=> "average",
											 "criteria"=> "below",
											 "format"=> format2))

	###############################################################################
	#
	# Example 5.
	#
	caption = "Top 10 values are in light red. Bottom 10 values are in light green."

	write!(ws5, "A1", caption)

	for row in size(data, 2)
		write_row!(ws5, row + 2, 1, data[row,:])
	end
	conditional_format!(ws5, "B3:K12", Dict("type"=> "top",
											 "value"=> "10",
											 "format"=> format1))

	conditional_format!(ws5, "B3:K12", Dict("type"=> "bottom",
											 "value"=> "10",
											 "format"=> format2))
	###############################################################################
	#
	# Example 6.
	#
	caption = "Cells with values >= 50 are in light red. Values < 50 are in light green. Non-contiguous ranges."

	# Write the data.
	write!(ws6, "A1", caption)

	for row in size(data, 2)
		write_row!(ws6, row + 2, 1, data[row,:])
	end
	# Write a conditional format over a range.
	conditional_format!(ws6, "B3:K6", Dict("type"=> "cell",
											"criteria"=> ">=",
											"value"=> 50,
											"format"=> format1,
											"multi_range"=> "B3:K6 B9:K12"))

	# Write another conditional format over the same range.
	conditional_format!(ws6, "B3:K6", Dict("type"=> "cell",
											"criteria"=> "<",
											"value"=> 50,
											"format"=> format2,
											"multi_range"=> "B3:K6 B9:K12"))


	###############################################################################
	#
	# Example 7.
	#
	caption = "Examples of color scales and data bars. Default colors."

	data = collect(1:13)

	write!(ws7, "A1", caption)

	write!(ws7, "B2", "2 Color Scale")
	write!(ws7, "D2", "3 Color Scale")
	write!(ws7, "F2", "Data Bars")

	for row in 1:length(data)
		write!(ws7, row + 2, 1, data[row])
		write!(ws7, row + 2, 3, data[row])
		write!(ws7, row + 2, 5, data[row])
	end

	conditional_format!(ws7, "B3:B14", Dict("type"=> "2_color_scale"))
	conditional_format!(ws7, "D3:D14", Dict("type"=> "3_color_scale"))
	conditional_format!(ws7, "F3:F14", Dict("type"=> "data_bar"))


	###############################################################################
	#
	# Example 8.
	#
	caption = "Examples of color scales and data bars. Modified colors."

	data = collect(1:13)

	write!(ws8, "A1", caption)

	write!(ws8, "B2", "2 Color Scale")
	write!(ws8, "D2", "3 Color Scale")
	write!(ws8, "F2", "Data Bars")

	for row in 1:length(data)
		write!(ws8, row + 2, 1, data[row])
		write!(ws8, row + 2, 3, data[row])
		write!(ws8, row + 2, 5, data[row])
	end
	conditional_format!(ws8, "B3:B14", Dict("type"=> "2_color_scale",
											 "min_color"=> "#FF0000",
											 "max_color"=> "#00FF00"))

	conditional_format!(ws8, "D3:D14", Dict("type"=> "3_color_scale",
											 "min_color"=> "#C5D9F1",
											 "mid_color"=> "#8DB4E3",
											 "max_color"=> "#538ED5"))

	conditional_format!(ws8, "F3:F14", Dict("type"=> "data_bar",
											 "bar_color"=> "#63C384"))

	close(wb)
	isfile("conditional_format.xlsx")

end

test()


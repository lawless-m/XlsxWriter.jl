###############################################################################
#
# An example of how to create autofilters with XlsxWriter.
#
# An autofilter is a way of adding drop down lists to the headers of a 2D
# range of worksheet data. This allows users to filter the data based on
# simple criteria so that some data is shown and some is hidden.
#
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# https://github.com/jmcnamara/XlsxWriter
#
# 

include("../XlsxWriter.jl")

using XlsxWriter

wb = workbook("autofilter.xlsx")

# Add a worksheet for each autofilter example.
sheet = Worksheet[]
for i in 1:6
	push!(sheet, add_worksheet!(wb))
end

# Add a bold format for the headers.
bold = add_format!(wb, Dict("bold"=>true))

# Open a text file with autofilter example data.
(data, headers) = readdlm("autofilter_data.txt", header=true)

# Set up several sheets with the same data.
for ws in sheet
    # Make the columns wider.
    set_column!(ws, "A:D", 12)
    # Make the header row larger and bold.
    set_row!(ws, 0, 20, bold)
	# Add header text
    write_row!(ws, 0, 0, headers)
end

###############################################################################
# Example 1. Autofilter without conditions.

# Set the autofilter.
autofilter!(sheet[1], "A1:D51")
write_matrix!(sheet[1], 1, 0, data)

###############################################################################
# Example 2. Autofilter with a filter condition in the first column.

# Autofilter range using Row-Column notation.
autofilter!(sheet[2], 0, 0, 50, 3)

# Add filter criteria. The placeholder "Region" in the filter is
# ignored and can be any string that adds clarity to the expression.
filter_column!(sheet[2], 0, "Region == East")

# Hide the rows that don't match the filter criteria.

for r in 1:size(data, 1)
    # Check for rows that match the filter.
    if data[r,1] != "East"
        set_row!(sheet[2], r, Dict("hidden"=>"1"))
	end
	write_row!(sheet[2], r, 0, data[r,:])
end

###############################################################################
# Example 3. Autofilter with a dual filter condition in one of the columns.

# Set the autofilter.
autofilter!(sheet[3], "A1:D51")

# Add filter criteria.
filter_column!(sheet[3], "A", "x == East or x == South")

# Hide the rows that don't match the filter criteria.
for r in 1:size(data, 1)
    # Check for rows that match the filter.
    if data[r,1] != "East" && data[r,1] != "South"
        set_row!(sheet[3], r, Dict("hidden"=>true))
	end
	write_row!(sheet[3], r, 0, data[r,:])
end


###############################################################################
# Example 4. Autofilter with filter conditions in two columns.

# Set the autofilter.
autofilter!(sheet[4], "A1:D51")

# Add filter criteria.
filter_column!(sheet[4], "A", "x == East")
filter_column!(sheet[4], "C", "x > 3000 and x < 8000")

# Hide the rows that don't match the filter criteria.
for r in 1:size(data, 1)
    # Hide rows that don't match the filter.
    if !(data[r,1] == "East" && data[r,3] > 3000 && data[r,3] < 8000)
		set_row!(sheet[4], r, Dict("hidden"=>true))
	end
    write_row!(sheet[4], r, 0, data[r,:])
end


###############################################################################
# Example 5. Autofilter with filter for blanks.
# Create a blank cell in our test data.

# Set the autofilter.
autofilter!(sheet[5], "A1:D51")

# Add filter criteria.
filter_column!(sheet[5], "A", "x == Blanks")

# Simulate a blank cell in the data.
data[5, 1] = ""

for r in 1:size(data, 1)
    # Check for rows that match the filter.
    if data[r,1] != ""
        # Row matches the filter, no further action required.
        set_row!(sheet[5], r, Dict("hidden"=>true))
	end
    write_row!(sheet[5], r, 0, data[r,:])
end

###############################################################################
# Example 6. Autofilter with filter for non-blanks.
# Set the autofilter.
autofilter!(sheet[6], "A1:D51")

# Add filter criteria.
filter_column!(sheet[6], "A", "x == NonBlanks")

for r in 1:size(data, 1)
    # Check for rows that match the filter.
    if data[r,1] == ""
        # Row matches the filter, no further action required.
        set_row!(sheet[6], r, Dict("hidden"=>true))
	end
    write_row!(sheet[6], r, 0, data[r,:])
end

close(wb)



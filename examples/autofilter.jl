###############################################################################
#
# An example of how to create autofilters with XlsxWriter.
#
# An autofilter is a way of adding drop down lists to the headers of a 2D
# range of worksheet data. This allows users to filter the data based on
# simple criteria so that some data is shown and some is hidden.
#
# Original Python Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
# http://xlsxwriter.readthedocs.io/example_autofilter.html

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
    if data[r,:][1] != "East"
        set_row!(sheet[2], r, Dict("hidden"=>"1"))
	end
	write_row!(sheet[2], r, 0, data[r,:])
end

###############################################################################
#
#
# Example 3. Autofilter with a dual filter condition in one of the columns.
#

# Set the autofilter.
autofilter!(sheet[3], "A1:D51")

# Add filter criteria.
filter_column!(sheet[3], "A", "x == East or x == South")

# Hide the rows that don't match the filter criteria.
row = 1
for r in 1:size(data, 1)
    # Check for rows that match the filter.
    if data[r,:][1] != "East" && data[r,:][1] != "South"
        set_row!(sheet[3], r, Dict("hidden"=>true))
	end
	write_row!(sheet[3], r, 0, data[r,:])
end

#=
###############################################################################
#
#
# Example 4. Autofilter with filter conditions in two columns.
#

# Set the autofilter.
worksheet4.autofilter('A1:D51')

# Add filter criteria.
worksheet4.filter_column('A', 'x == East')
worksheet4.filter_column('C', 'x > 3000 and x < 8000')

# Hide the rows that don't match the filter criteria.
row = 1
for row_data in (data):
    region = row_data[0]
    volume = int(row_data[2])

    # Check for rows that match the filter.
    if region == 'East' and volume > 3000 and volume < 8000:
        # Row matches the filter, no further action required.
        pass
    else:
        # We need to hide rows that don't match the filter.
        worksheet4.set_row(row, options={'hidden': True})

    worksheet4.write_row(row, 0, row_data)

    # Move on to the next worksheet row.
    row += 1


###############################################################################
#
#
# Example 5. Autofilter with filter for blanks.
#
# Create a blank cell in our test data.

# Set the autofilter.
worksheet5.autofilter('A1:D51')

# Add filter criteria.
worksheet5.filter_column('A', 'x == Blanks')

# Simulate a blank cell in the data.
data[5][0] = ''

# Hide the rows that don't match the filter criteria.
row = 1
for row_data in (data):
    region = row_data[0]

    # Check for rows that match the filter.
    if region == '':
        # Row matches the filter, no further action required.
        pass
    else:
        # We need to hide rows that don't match the filter.
        worksheet5.set_row(row, options={'hidden': True})

    worksheet5.write_row(row, 0, row_data)

    # Move on to the next worksheet row.
    row += 1


###############################################################################
#
#
# Example 6. Autofilter with filter for non-blanks.
#

# Set the autofilter.
worksheet6.autofilter('A1:D51')

# Add filter criteria.
worksheet6.filter_column('A', 'x == NonBlanks')

# Hide the rows that don't match the filter criteria.
row = 1
for row_data in (data):
    region = row_data[0]

    # Check for rows that match the filter.
    if region != '':
        # Row matches the filter, no further action required.
        pass
    else:
        # We need to hide rows that don't match the filter.
        worksheet6.set_row(row, options={'hidden': True})

    worksheet6.write_row(row, 0, row_data)

    # Move on to the next worksheet row.
    row += 1
=#

close(wb)



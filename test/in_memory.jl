#######################################################################
#
# An example of creating an in-memory XLS file with XlsxWriter.

using XlsxWriter

function test()

	buff = IOBuffer()
	wb = Workbook(buff)
	ws = add_worksheet!(wb)
	write!(ws, 0, 0, "In memory")
	close(wb)

	# dump it to a file for testing but really for throwing across the wire via HTTP or whatever

	fid = open("in_memory.xlsx", "w+")
	write(fid, take!(buff))
	close(fid)
	
	isfile("in_memory.xlsx")
end

test()

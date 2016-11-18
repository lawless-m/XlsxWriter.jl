module XlsxWriter

export Workbook, add_worksheet, set_column!, write!, Url, close_workbook

typealias Format Dict{AbstractString, Bool}
type Url
	url::AbstractString
end


using PyCall
@pyimport xlsxwriter

Workbook(fn::AbstractString) = xlsxwriter.Workbook(fn)

add_worksheet(wb::PyObject) = wb[:add_worksheet]()


# add_format!(ws::PyObject, args...) = ws[:add_format](args...)

set_column!(ws::PyObject, args...) = ws[:set_column](args...)

# write_string / write_formula

function write!(ws::PyObject, row::Int64, col::Int64, data::AbstractString, fmt::Format=Format())
	if length(data) > 0
		if data[1] == '='
			fn = :write_formula
		else
			fn = :write_string
		end
	else
		fn = :write_blank
	end
	ws[fn](row, col, data, fmt)
end
function write!(ws::PyObject, cell::AbstractString, data::AbstractString, fmt::Format=Format())
	if length(data) > 0
		if data[1] == '='
			fn = :write_formula
		else
			fn = :write_string
		end
	else
		fn = :write_blank
	end
	ws[fn](cell, data, fmt)
end

# write_number
write!(ws::PyObject, row::Int64, col::Int64, num::Real, fmt::Format=Format()) = ws[:write_number](row, col, num, fmt)
write!(ws::PyObject, cell::AbstractString, num::Real, fmt::Format=Format()) = ws[:write_number](cell, num, fmt)

#write_blank
write!(ws::PyObject, row::Int64, col::Int64, fmt::Format=Format()) = ws[:write_blank](row, col, fmt)
write!(ws::PyObject, cell::AbstractString, fmt::Format=Format()) = ws[:write_blank](cell, fmt)

# write_datetime
write!(ws::PyObject, row::Int64, col::Int64, dt::DateTime, fmt::Format=Format()) = ws[:write_datetime](row, col, dt, fmt)
write!(ws::PyObject, cell::AbstractString, dt::DateTime, fmt::Format=Format()) = ws[:write_datetime](cell, dt, fmt)

# write_boolean
write!(ws::PyObject, row::Int64, col::Int64, bool::Bool, fmt::Format=Format()) = ws[:write_boolean](row, col, bool, fmt)
write!(ws::PyObject, cell::AbstractString, bool::Bool, fmt::Format=Format()) = ws[:write_boolean](cell, bool, fmt)

# write_url
write!(ws::PyObject, row::Int64, col::Int64, u::Url, fmt::Format=Format()) = ws[:write_url](row, col, u.url, fmt)
write!(ws::PyObject, cell::AbstractString, u::Url, fmt::Format=Format()) = ws[:write_url](cell, u.url, fmt)


#insert_image!(ws::PyObject, args...) = ws[:insert_image](args...)

close_workbook(wb::PyObject) = wb[:close]()

end

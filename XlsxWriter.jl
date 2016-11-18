module XlsxWriter

export Workbook, add_worksheet!, add_format!, set_column!, write!, Url, close_workbook



type Url
	url::AbstractString
end


using PyCall
@pyimport xlsxwriter


type Format
	fmt::Union{PyObject, Void}
	Format() = new(nothing)
	Format(p::PyObject) = new(p)
end

Workbook(fn::AbstractString) = xlsxwriter.Workbook(fn)

add_worksheet!(wb::PyObject) = wb[:add_worksheet]()


add_format!(ws::PyObject, f::Dict) = Format(ws[:add_format](f))

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
	fmt.fmt == nothing ? ws[fn](row, col, data) : ws[fn](row, col, data, fmt.fmt)
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
	
	fmt.fmt == nothing ? ws[fn](cell, data) : ws[fn](cell, data, fmt.fmt)
end

# write_number
write!(ws::PyObject, row::Int64, col::Int64, num::Real, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:write_number](row, col, num) : ws[:write_number](row, col, num, fmt.fmt)
write!(ws::PyObject, cell::AbstractString, num::Real, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:write_number](cell, num) : ws[:write_number](cell, num, fmt.fmt)

#write_blank
write!(ws::PyObject, row::Int64, col::Int64, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:write_blank](row, col) : ws[:write_blank](row, col, fmt.fmt)
write!(ws::PyObject, cell::AbstractString, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:write_blank](cell) : ws[:write_blank](cell, fmt.fmt)

# write_datetime
write!(ws::PyObject, row::Int64, col::Int64, dt::DateTime, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:write_datetime](row, col, dt) : ws[:write_datetime](row, col, dt, fmt.fmt)
write!(ws::PyObject, cell::AbstractString, dt::DateTime, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:write_datetime](cell, dt) : ws[:write_datetime](cell, dt, fmt.fmt)

# write_boolean
write!(ws::PyObject, row::Int64, col::Int64, bool::Bool, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:write_boolean](row, col, bool) : ws[:write_boolean](row, col, bool, fmt.fmt)
write!(ws::PyObject, cell::AbstractString, bool::Bool, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:write_boolean](cell, bool) : ws[:write_boolean](cell, bool, fmt.fmt)

# write_url
write!(ws::PyObject, row::Int64, col::Int64, u::Url, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:write_url](row, col, u.url) : ws[:write_url](row, col, u.url, fmt.fmt)
write!(ws::PyObject, cell::AbstractString, u::Url, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:write_url](cell, u.url) : ws[:write_url](cell, u.url, fmt.fmt)


#insert_image!(ws::PyObject, args...) = ws[:insert_image](args...)

close_workbook(wb::PyObject) = wb[:close]()

end

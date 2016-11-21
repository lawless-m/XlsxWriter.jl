module XlsxWriter

export Workbook, add_worksheet!, add_format!, set_column!, write!, Url, close_workbook, define_name!, worksheets, get_worksheet_by_name, set_first_sheet!, merge_range!, freeze_panes!, split_panes!, Xls, write_matrix!

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

function rc2cell(row::Int, col::Int64)
	cell = string(Char(mod(col, 26) + 65)) * "$(row+1)"
	col = div(col, 26)
	while col > 0
		cell = string(Char(mod(col, 26) + 64)) * cell
		col = div(col, 26)
	end
	cell
end

Workbook(fn::AbstractString) = xlsxwriter.Workbook(fn)

add_worksheet!(wb::PyObject) = wb[:add_worksheet]()
add_worksheet!(wb::PyObject, name::AbstractString) = wb[:add_worksheet](name)



add_format!(ws::PyObject, f::Dict) = Format(ws[:add_format](f))

set_column!(ws::PyObject, args...) = ws[:set_column](args...)

# write_string / write_formula

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

# convert r,c into cell format
write!(ws::PyObject, row::Int64, col::Int64, data, fmt::Format=Format()) = write!(ws, rc2cell(row, col), data, fmt)

# write_number
write!(ws::PyObject, cell::AbstractString, num::Real, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:write_number](cell, num) : ws[:write_number](cell, num, fmt.fmt)

#write_blank
write!(ws::PyObject, cell::AbstractString, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:write_blank](cell) : ws[:write_blank](cell, fmt.fmt)

# write_datetime
write!(ws::PyObject, cell::AbstractString, dt::DateTime, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:write_datetime](cell, dt) : ws[:write_datetime](cell, dt, fmt.fmt)

# write_boolean
write!(ws::PyObject, cell::AbstractString, bool::Bool, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:write_boolean](cell, bool) : ws[:write_boolean](cell, bool, fmt.fmt)

# write_url
write!(ws::PyObject, cell::AbstractString, u::Url, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:write_url](cell, u.url) : ws[:write_url](cell, u.url, fmt.fmt)

#write matrix

function write_matrix!(ws::PyObject, row::Int64, col::Int64, data::Array, fmt::Format=Format())
	re = size(data, 1)-1
	ce = size(data, 2)-1
	for r in 0:re, c in 0:ce
		write!(ws, rc2cell(row+r, col+c), data[r+1, c+1], fmt)
	end
end

define_name!(wb::PyObject, name::AbstractString, target::AbstractString) = wb[:define_name](name, target)

worksheets(wb::PyObject) = wb[:worksheets]()

get_worksheet_by_name(wb::PyObject) = wb[:get_worksheet_by_name]()

set_first_sheet!(ws::PyObject) = ws[:set_first_sheet]()

merge_range!(ws::PyObject, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64, contents, fmt::Format=Format()) = fmt.fmt==nothing ? ws[:merge_range](first_row, first_col, last_row, last_col, contents) : ws[:merge_range](first_row, first_col, last_row, last_col, contents, fmt.fmt) 


freeze_panes!(ws::PyObject, row::Int64, col::Int64) = ws[:freeze_panes](row, col)
freeze_panes!(ws::PyObject, cell::AbstractString) = ws[:freeze_panes](cell)
freeze_panes!(ws::PyObject, row::Int64, col::Int64, top_row::Int64) = ws[:freeze_panes](row, col, top_row)
freeze_panes!(ws::PyObject, row::Int64, col::Int64, top_row::Int64, left_col::Int64) = ws[:freeze_panes](row, col, top_row, left_col)

split_panes!(ws::PyObject,x::Int64, y::Int64) = ws[:split_panes](x, y)
split_panes!(ws::PyObject,x::Int64, y::Int64, top_row::Int64) = ws[:split_panes](x, y, top_row)
split_panes!(ws::PyObject,x::Int64, y::Int64, top_row::Int64, left_col::Int64) = ws[:split_panes](x, y, top_row, left_col)

#insert_image!(ws::PyObject, args...) = ws[:insert_image](args...)

close_workbook(wb::PyObject) = wb[:close]()

end

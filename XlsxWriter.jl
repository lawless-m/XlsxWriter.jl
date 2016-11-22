module XlsxWriter

export workbook, add_worksheet!, add_format!, set_column!, write!, Url, close_workbook, define_name!, worksheets, get_worksheet_by_name, set_first_sheet!, merge_range!, freeze_panes!, split_panes!, Xls, write_matrix!, set_font_name!, set_font_size!, set_font_color!, set_bold!, set_italic!, set_underline!, set_font_strikeout!, set_font_script!, set_num_format!, set_locked!, set_hidden!, set_align!, set_center_across!, set_text_wrap!, set_rotation!, set_indent!, set_shrink!, set_text_justlast!, set_pattern!, set_bg_color!, set_fg_color!, set_border!, set_bottom!, set_top!, set_left!, set_right!, set_border_color!, set_bottom_color!, set_top_color!, set_left_color!, set_right_color!, set_diag_border!, set_diag_type!, set_diag_color!

type Url
	url::AbstractString
end

using PyCall
@pyimport xlsxwriter


type Workbook
	py::PyObject
end

type Worksheet
	py::PyObject
end

type Format
	py::PyObject
end

typealias MaybeFormat Union{Format, Void}

function rc2cell(row::Int, col::Int64)
	cell = string(Char(mod(col, 26) + 65)) * "$(row+1)"
	col = div(col, 26)
	while col > 0
		cell = string(Char(mod(col, 26) + 64)) * cell
		col = div(col, 26)
	end
	cell
end

workbook(fn::AbstractString) = Workbook(xlsxwriter.Workbook(fn))

add_worksheet!(wb::Workbook) = Worksheet(wb.py[:add_worksheet]())
add_worksheet!(wb::Workbook, name::AbstractString) = Worksheet(wb.py[:add_worksheet](name))

add_format!(wb::Workbook, f::Dict) = Format(wb.py[:add_format](f))

set_column!(ws::Worksheet, args...) = ws.py[:set_column](args...)

# write_string / write_formula

function write!(ws::Worksheet, cell::AbstractString, data::AbstractString, fmt::MaybeFormat=nothing)
	if length(data) > 0
		if data[1] == '='
			fn = :write_formula
		else
			fn = :write_string
		end
	else
		fn = :write_blank
	end
	write!(ws, cell, fn, data, fmt)
end

function write!(ws::Worksheet, cell::AbstractString, fn::Symbol, data, fmt::Void)
	ws.py[fn](cell, data)
end

function write!(ws::Worksheet, cell::AbstractString, fn::Symbol, data, fmt::Format)
	ws.py[fn](cell, data, fmt.py)
end

# convert r,c into cell format
write!(ws::Worksheet, row::Int64, col::Int64, data, fmt::MaybeFormat=nothing) = write!(ws, rc2cell(row, col), data, fmt)
write!(ws::Worksheet, cell::AbstractString, num::Real, fmt::MaybeFormat=nothing) = write!(ws, cell, :write_number, num, fmt)
write!(ws::Worksheet, cell::AbstractString, fmt::MaybeFormat=nothing) = write!(ws, cell, :write_blank, fmt)
write!(ws::Worksheet, cell::AbstractString, dt::DateTime, fmt::MaybeFormat=nothing) = write!(ws, cell, :write_datetime, dt, fmt)
write!(ws::Worksheet, cell::AbstractString, bool::Bool, fmt::MaybeFormat=nothing) = write!(ws, cell, :write_boolean, bool, fmt)
write!(ws::Worksheet, cell::AbstractString, u::Url, fmt::MaybeFormat=nothing) = write!(ws, cell, :write_url, u.url, fmt)

function write!(ws::Worksheet, row::Int64, col::Int64, data::Array, fmt::MaybeFormat=nothing)
	re = size(data, 1)-1
	ce = size(data, 2)-1
	for r in 0:re, c in 0:ce
		write!(ws, row+r, col+c, data[r+1, c+1], fmt)
	end
end

define_name!(wb::Workbook, name::AbstractString, target::AbstractString) = wb.py[:define_name](name, target)

worksheets(wb::Workbook) = wb.py[:worksheets]()

get_worksheet_by_name(wb::Workbook) = wb.py[:get_worksheet_by_name]()

set_first_sheet!(ws::Worksheet) = ws.py[:set_first_sheet]()

merge_range!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64, contents, fmt::MaybeFormat=nothing) = merge_range!(first_row, first_col, last_row, last_col, contents, fmt)

merge_range!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64, contents, fmt::Format) = ws.py[:merge_range](first_row, first_col, last_row, last_col, contents, fmt.py)

merge_range!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64, contents, fmt::Void) = ws.py[:merge_range](first_row, first_col, last_row, last_col, contents)


freeze_panes!(ws::Worksheet, row::Int64, col::Int64) = ws.py[:freeze_panes](row, col)
freeze_panes!(ws::Worksheet, cell::AbstractString) = ws.py[:freeze_panes](cell)
freeze_panes!(ws::Worksheet, row::Int64, col::Int64, top_row::Int64) = ws.py[:freeze_panes](row, col, top_row)
freeze_panes!(ws::Worksheet, row::Int64, col::Int64, top_row::Int64, left_col::Int64) = ws.py[:freeze_panes](row, col, top_row, left_col)

split_panes!(ws::Worksheet,x::Int64, y::Int64) = ws.py[:split_panes](x, y)
split_panes!(ws::Worksheet,x::Int64, y::Int64, top_row::Int64) = ws.py[:split_panes](x, y, top_row)
split_panes!(ws::Worksheet,x::Int64, y::Int64, top_row::Int64, left_col::Int64) = ws.py[:split_panes](x, y, top_row, left_col)

set_font_name!(fmt::Format, name::AbstractString) = fmt.py[:set_font_name](name)
set_font_size!(fmt::Format, sz::Int64) = fmt.py[:set_font_size](sz)
set_font_color!(fmt::Format, color::AbstractString) = fmt.py[:set_font_color](color)
set_bold!(fmt::Format, state::Bool=true) = fmt.py[:set_bold](state)
set_italic!(fmt::Format, state::Bool=true) = fmt.py[:set_italic](state)
set_underline!(fmt::Format, state::Bool=true) = fmt.py[:set_underline](state)
set_font_strikeout!(fmt::Format, state::Bool=true) = fmt.py[:set_font_strikeout](state)
set_font_script!(fmt::Format, state::Int64) = fmt.py[:set_font_script](state)
set_num_format!(fmt::Format, format::AbstractString) = fmt.py[:set_font_num_format](format)
set_locked!(fmt::Format, state::Bool=true) = fmt.py[:set_locked](state)
set_hidden!(fmt::Format, state::Bool=true) = fmt.py[:set_hidden](state)
set_align!(fmt::Format, alignment::AbstractString) = fmt.py[:set_align](alignment)
set_center_across!(fmt::Format, across::AbstractString) = fmt.py[:set_center_across](across)
set_text_wrap!(fmt::Format, state::Bool=true) = fmt.py[:set_text_wrap](state)
set_rotation!(fmt::Format, angle::Int64) = fmt.py[:set_rotation](angle)
set_indent!(fmt::Format, level::Int64) = fmt.py[:set_indent](level)
set_shrink!(fmt::Format, state::Bool=true) = fmt.py[:set_shrink](state)
set_text_justlast!(fmt::Format, state::Bool=true) = fmt.py[:set_text_justlast](state)
set_pattern!(fmt::Format, index::Int64) = fmt.py[:set_pattern](index)
set_bg_color!(fmt::Format, color::AbstractString) = fmt.py[:set_bg_color](color)
set_fg_color!(fmt::Format, color::AbstractString) = fmt.py[:set_fg_color](color)
set_border!(fmt::Format, style::Int64) = fmt.py[:set_border](style)
set_bottom!(fmt::Format, style::Int64) = fmt.py[:set_bottom](style)
set_top!(fmt::Format, style::Int64) = fmt.py[:set_top](style)
set_left!(fmt::Format, style::Int64) = fmt.py[:set_left](style)
set_right!(fmt::Format, style::Int64) = fmt.py[:set_right](style)
set_border_color!(fmt::Format, color::AbstractString) = fmt.py[:set_border_color](color)
set_bottom_color!(fmt::Format, color::AbstractString) = fmt.py[:set_bottom_color](color)
set_top_color!(fmt::Format, color::AbstractString) = fmt.py[:set_top_color](color)
set_left_color!(fmt::Format, color::AbstractString) = fmt.py[:set_left_color](color)
set_right_color!(fmt::Format, color::AbstractString) = fmt.py[:set_right_color](color)
set_diag_border!(fmt::Format, style::Int64) = fmt.py[:set_diag_border](style)
set_diag_type!(fmt::Format, style::Int64) = fmt.py[:set_diag_type](style)
set_diag_color!(fmt::Format, color::AbstractString) = fmt.py[:set_diag_color](style)




#insert_image!(ws::PyObject, args...) = ws[:insert_image](args...)

close_workbook(wb::Workbook) = wb.py[:close]()

end

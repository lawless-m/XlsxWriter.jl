module XlsxWriter

export workbook, add_worksheet!, add_format!, set_properties!, set_custom_property!, set_calc_mode!, set_column!, set_row!, write!, write_string!, write_blank!, write_formula!, write_datetime!, write_bool!, write_url!, write_number!, write_array_formula!, Url, write_row!, write_column!, write_matrix!, close_workbook, define_name!, worksheets, get_worksheet_by_name, set_first_sheet!, merge_range!, freeze_panes!, split_panes!, Xls, write_matrix!, set_font_name!, set_font_size!, set_font_color!, set_bold!, set_italic!, set_underline!, set_font_strikeout!, set_font_script!, set_num_format!, set_locked!, set_hidden!, set_align!, set_center_across!, set_text_wrap!, set_rotation!, set_indent!, set_shrink!, set_text_justlast!, set_pattern!, set_bg_color!, set_fg_color!, set_border!, set_bottom!, set_top!, set_left!, set_right!, set_border_color!, set_bottom_color!, set_top_color!, set_left_color!, set_right_color!, set_diag_border!, set_diag_type!, set_diag_color!, data_validation!, conditional_format!, add_table!, add_sparkline!

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

macro Fmt()
	:(fmt == nothing ? fmt : fmt.py)
end

typealias Data Union{Real, AbstractString, DateTime, Bool, Url}
typealias MaybeFormat Union{Format, Void}
typealias MaybeData Union{Data, Void}

function rc2cell(row::Int64, col::Int64)
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
define_name!(wb::Workbook, name::AbstractString, target::AbstractString) = wb.py[:define_name](name, target)
worksheets(wb::Workbook) = wb.py[:worksheets]()
close_workbook(wb::Workbook) = wb.py[:close]()
get_worksheet_by_name(wb::Workbook) = wb.py[:get_worksheet_by_name]()
set_properties!(wb::Workbook, p::Dict{AbstractString, AbstractString}) = wb.py[:set_properties](p)
add_format!(wb::Workbook, f::Dict) = Format(wb.py[:add_format](f))
set_custom_property!(wb::Workbook, name::AbstractString, value::Data) = wb.py[:set_custom_property](name, value)
set_calc_mode!(wb::Workbook, mode::AbstractString) = wb.py[:set_calc_mode](mode)

set_column!(ws::Worksheet, first_col::Int64, last_col::Int64, width::Real, fmt::MaybeFormat=nothing, options::Dict=Dict()) = ws.py[:set_column](first_col, last_col, width, @Fmt, options)
set_column!(ws::Worksheet, cols::AbstractString, width::Real, fmt::MaybeFormat=nothing, options::Dict=Dict()) = ws.py[:set_column](cols, width, @Fmt, options)
set_column!(ws::Worksheet, first_col::Int64, last_col::Int64, width::Real, options::Dict=Dict()) = ws.py[:set_column](first_col, last_col, width, options)
set_column!(ws::Worksheet, cols::AbstractString, width::Real, options::Dict=Dict()) = ws.py[:set_column](cols, width, options)

set_row!(ws::Worksheet, row::Int64, height::Real, fmt::MaybeFormat=nothing, options::Dict=Dict()) = ws.py[:set_row](row, height, @Fmt, options)
set_row!(ws::Worksheet, row::Int64, height::Real, options::Dict=Dict()) = ws.py[:set_row](row, height, options)



# write_string / write_formula


function write!(ws::Worksheet, row::Int64, col::Int64, data::AbstractString, fmt::MaybeFormat=nothing)
	write!(ws, rc2cell(row, col), data, fmt)
end

function write!(ws::Worksheet, cell::AbstractString, data::AbstractString, fmt::MaybeFormat=nothing)
	if length(data) > 0
		if data[1] == '=' || (data[1:2] == "{=" && data[end] == '}')
			write!(ws, cell, :write_formula, data, fmt)
		else
			write!(ws, cell, :write_string, data, fmt)
		end
	else
		write!(ws, cell, :write_blank, data, fmt)
	end
end
#write_formula! = write!

function write!(ws::Worksheet, cell::AbstractString, fn::Symbol, data::Data, fmt::MaybeFormat=nothing)
	ws.py[fn](cell, data, @Fmt)
end

# convert r,c into cell format
write!(ws::Worksheet, row::Int64, col::Int64, data, fmt::MaybeFormat=nothing) = write!(ws, rc2cell(row, col), data, fmt)
write_string! = write!

write!(ws::Worksheet, cell::AbstractString, num::Real, fmt::MaybeFormat=nothing) = write!(ws, cell, :write_number, num, fmt)
write_number! = write!

write!(ws::Worksheet, cell::AbstractString, fmt::MaybeFormat=nothing) = write!(ws, cell, :write_blank, fmt)
write_blank! = write!

write!(ws::Worksheet, cell::AbstractString, dt::DateTime, fmt::MaybeFormat=nothing) = write!(ws, cell, :write_datetime, dt, fmt)
write_datetime! = write!

write!(ws::Worksheet, cell::AbstractString, bool::Bool, fmt::MaybeFormat=nothing) = write!(ws, cell, :write_boolean, bool, fmt)
write_bool! = write!

write!(ws::Worksheet, cell::AbstractString, u::Url, fmt::MaybeFormat=nothing) = write!(ws, cell, :write_url, u.url, fmt)
write_url! = write!

function write_matrix!(ws::Worksheet, row::Int64, col::Int64, data::Matrix, fmt::MaybeFormat=nothing)
	re = size(data, 1)-1
	ce = size(data, 2)-1
	if re > ce
		for c in 0:ce
			write_column!(ws, row, col+c, squeeze(data[:, c+1], 1), fmt)
		end
	else
		for r in 0:re
			write_row!(ws, row+r, col, squeeze(data[r+1, :], 1), fmt)
		end
	end
end

function write_formula!(ws::Worksheet, row::Int64, col::Int64, formula::AbstractString, fmt::MaybeFormat=nothing; result::MaybeData=nothing)
	ws.py[:write_formula](row, col, formula, @Fmt, result)
end

function write_formula!(ws::Worksheet, cell::AbstractString, formula::AbstractString, fmt::MaybeFormat=nothing; result::MaybeData=nothing)
	ws.py[:write_formula](cell, formula, @Fmt, result)
end

function write_array_formula!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64, formula::AbstractString, fmt::MaybeFormat=nothing)
	ws.py[:write_array_formula](first_row, first_col, last_row, last_col, formula, @Fmt)
end

function write_array_formula!(ws::Worksheet, first_cell::AbstractString, last_cell::Int64, last_col::Int64, formula::AbstractString, fmt::MaybeFormat=nothing)
	ws.py[:write_array_formula](first_cell, last_cell, formula, @Fmt)
end

function write_row!(ws::Worksheet, row::Int64, col::Int64, data::Vector, fmt::MaybeFormat=nothing)
	ws.py[:write_row](row, col, data, @Fmt)
end

function write_row!(ws::Worksheet, cell::AbstractString, data::Vector, fmt::MaybeFormat=nothing)
	ws.py[:write_row](cell, data)
end

function write_column!(ws::Worksheet, row::Int64, col::Int64, data::Vector, fmt::MaybeFormat=nothing)
	ws.py[:write_column](row, col, data, @Fmt)
end

function write_row!(ws::Worksheet, cell::AbstractString, data::Vector, fmt::MaybeFormat=nothing)
	ws.py[:write_column](cell, data, @Fmt)
end



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

data_validation!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64, options::Dict) = ws.py[:data_validation](first_row, first_col, last_row, last_col, options)
data_validation!(ws::Worksheet, first_cell::AbstractString, last_cell::AbstractString, options::Dict) = ws.py[:data_validation](first_cell, last_cell, options)

conditional_format!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64, options::Dict) = ws.py[:data_validation](first_row, first_col, last_row, last_col, options)
conditional_format!(ws::Worksheet, first_cell::AbstractString, last_cell::AbstractString, options::Dict) = ws.py[:data_validation](first_cell, last_cell, options)

add_table!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64, options::Dict) = ws.py[:data_validation](first_row, first_col, last_row, last_col, options)
add_table!(ws::Worksheet, first_cell::AbstractString, last_cell::AbstractString, options::Dict) = ws.py[:data_validation](first_cell, last_cell, options)

add_sparkline!(ws::Worksheet, row::Int64, col::Int64, options::Dict) = ws.py[:add_sparkline](row, col, options)
add_sparkline!(ws::Worksheet, cell::AbstractString, options::Dict) = ws.py[:add_sparkline](cell, options)

#insert_image!(ws::PyObject, args...) = ws[:insert_image](args...)
end

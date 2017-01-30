module XlsxWriter

#=
http://xlsxwriter.readthedocs.io/
=#

import Base.close

using PyCall
@pyimport xlsxwriter

export Workbook, add_worksheet!, add_chartsheet!, add_format!, set_properties!, set_custom_property!, set_calc_mode!, set_column!, set_row!, add_chart!, close, rc2cell, colNtocolA

export Chart, add_series!, set_x_axis!, set_y_axis!, set_x2_axis!, set_y2_axis!, combine!, set_size!, set_title!, set_legend!, set_chartarea!, set_plotarea!, set_style!, set_table!, set_up_down_bars!, set_drop_lines!, set_high_low_lines!, set_blanks_as!, show_hidden_data!, set_rotation!, set_hole_size!

export Format, set_font_name!, set_font_size!, set_font_color!, set_bold!, set_italic!, set_underline!, set_font_strikeout!, set_font_script!, set_num_format!, set_locked!, set_hidden!, set_align!, set_center_across!, set_text_wrap!, set_rotation!, set_indent!, set_shrink!, set_text_justlast!, set_pattern!, set_bg_color!, set_fg_color!, set_border!, set_bottom!, set_top!, set_left!, set_right!, set_border_color!, set_bottom_color!, set_top_color!, set_left_color!, set_right_color!, set_diag_border!, set_diag_type!, set_diag_color!

export Url 

export Worksheet, Chartsheet, set_chart!, write!, write_string!, write_blank!, write_formula!, write_datetime!, write_bool!, write_url!, write_number!, write_array_formula!, write_row!, write_column!, write_matrix!, define_name!, worksheets, get_worksheet_by_name, set_first_sheet!, merge_range!, freeze_panes!, split_panes!, Xls, write_matrix!, data_validation!, conditional_format!, add_table!, add_sparkline!, activate!, select!, hide!, set_first_sheet!, t!, protect!, set_zoom!, set_tab_color!, set_landscape!, set_portrait!, set_paper!, set_margins!, set_header!, set_header!, set_footer!, set_footer!, get_name, write_rich_string!, insert_image!, insert_chart!, insert_textbox!, insert_button!, write_comment!, show_comments!, set_comments_author!, autofilter!, filter_column!, filter_column_list!, set_selection!, set_default_row!, outline_settings!, set_vba_name!, add_vba_project!

type Url
	url::AbstractString
end

type Workbook
	py::PyObject
	io::Union{AbstractString, IOBuffer}
	Workbook() = Workbook(IOBuffer())
	Workbook(fn::AbstractString, opts::Dict=Dict()) = fn == "" ? Workbook(IOBuffer(), opts) : new(xlsxwriter.Workbook(fn, opts), fn)
	function Workbook(io::IOBuffer, opts::Dict=Dict())
		if !haskey(opts, "in_memory")
			opts["in_memory"] = true
		end
		new(xlsxwriter.Workbook(io, opts), io)
	end
end

type Worksheet
	py::PyObject
end

type Format
	py::PyObject
end

type Chart
	py::PyObject
end

type Chartsheet
	py::PyObject
end


typealias Data Union{Real, AbstractString, DateTime, Date, Bool, Url}
typealias MaybeFormat Union{Format, Void}
typealias MaybeData Union{Data, Void}

function colNtocolA(n::Int64)
	a = string(Char(mod(n, 26) + 65))
	n = div(n, 26)
	while n > 0
		a = string(Char(mod(n, 26) + 64)) * a
		n = div(n, 26)
	end
	a
end

function rc2cell(row::Int64, col::Int64)
	colNtocolA(col) * "$(row+1)"
end

function cell2rc(cell::AbstractString)
	col = 0
	while length(cell) > 0 && cell[1]>='A'
		col *= 26
		col += 1 + Int(cell[1]-'A') 
		cell = cell[2:end]
	end
	parse(Int, cell)-1, col-1
end

# Workbook
macro WbFn(fn)
	:(wb.py[$fn]())
end
macro WbOpts(fn)
	:(wb.py[$fn](options))
end

add_worksheet!(wb::Workbook) = Worksheet(wb.py[:add_worksheet]())
add_worksheet!(wb::Workbook, name::AbstractString) = Worksheet(wb.py[:add_worksheet](name))
define_name!(wb::Workbook, name::AbstractString, target::AbstractString) = wb.py[:define_name](name, target)
worksheets(wb::Workbook) = @WbFn :worksheets

close(wb::Workbook) = @WbFn :close

get_worksheet_by_name(wb::Workbook) = @WbFn :get_worksheet_by_name
set_properties!(wb::Workbook, options::Dict{AbstractString, AbstractString}) = @WbOpts set_properties
add_format!(wb::Workbook, options::Dict=Dict()) = Format(@WbOpts :add_format)
set_custom_property!(wb::Workbook, name::AbstractString, value::Data) = wb.py[:set_custom_property](name, value)
set_calc_mode!(wb::Workbook, mode::AbstractString) = wb.py[:set_calc_mode](mode)
add_chart!(wb::Workbook, options::Dict) = Chart(@WbOpts :add_chart)
add_chartsheet!(wb::Workbook) = Chartsheet(@WbFn :add_chartsheet)
set_size!(wb::Workbook, width::Int64, height::Int64) = wb.py[:set_size](width, height)
add_vba_project!(wb::Workbook, filename::AbstractString, is_stream::Bool=false) = wb.py[:add_vba_project](filename, is_stream)
set_vba_name!(wb::Workbook, name::AbstractString) = wb.py[:set_vba_name](name)
use_zip64!(wb::Workbook) = @WbFn :use_zip64



# Worksheet
macro Fmt()
	:(fmt == nothing ? fmt : fmt.py)
end



set_column!(ws::Worksheet, first_col::Int64, last_col::Int64, width::Real, fmt::MaybeFormat=nothing, options::Dict=Dict()) = ws.py[:set_column](first_col, last_col, width, @Fmt, options)
set_column!(ws::Worksheet, cols::AbstractString, width::Real, fmt::MaybeFormat=nothing, options::Dict=Dict()) = ws.py[:set_column](cols, width, @Fmt, options)
set_column!(ws::Worksheet, first_col::Int64, last_col::Int64, width::Real, options::Dict=Dict()) = ws.py[:set_column](first_col, last_col, width, options)
set_column!(ws::Worksheet, cols::AbstractString, width::Real, options::Dict=Dict()) = ws.py[:set_column](cols, width, options)

set_row!(ws::Worksheet, row::Int64, height::Real, fmt::Format, options::Dict) = ws.py[:set_row](row, height, @Fmt, options)
set_row!(ws::Worksheet, row::Int64, height::Real, fmt::Format) = ws.py[:set_row](row, height, @Fmt)
set_row!(ws::Worksheet, row::Int64, height::Real) = ws.py[:set_row](row, height)
set_row!(ws::Worksheet, row::Int64, fmt::Format, options::Dict=Dict()) = ws.py[:set_row](row, nothing, @Fmt, options)
set_row!(ws::Worksheet, row::Int64, options::Dict) = ws.py[:set_row](row, nothing, nothing, options)


# write_string / write_formula


function write!(ws::Worksheet, cell::AbstractString, data, fmt::MaybeFormat=nothing)
	write!(ws, cell2rc(cell)..., data, fmt)
end

function write!(ws::Worksheet, row::Int64, col::Int64, data::AbstractString, fmt::MaybeFormat=nothing)
	if length(data) > 0
		if data[1] == '=' || (length(data) > 1 && (data[1:2] == "{=" && data[end] == '}'))
			write!(ws, row, col, :write_formula, data, fmt)
		else
			write!(ws, row, col, :write_string, data, fmt)
		end
	else
		write!(ws, row, col, :write_blank, data, fmt)
	end
end
#write_formula! = write!

function write!(ws::Worksheet, row::Int64, col::Int64, fn::Symbol, data::Data, fmt::MaybeFormat=nothing)
	ws.py[fn](row, col, data, @Fmt)
	1
end

# convert r,c into cell format
write!(ws::Worksheet, cell::AbstractString, data, fmt::MaybeFormat=nothing) = write!(ws, cell2rc(cell)..., data, fmt)
write_string! = write!

write!(ws::Worksheet, row::Int64, col::Int64, num::Real, fmt::MaybeFormat=nothing) = write!(ws, row, col, :write_number, num, fmt)
write_number! = write!

write!(ws::Worksheet, row::Int64, col::Int64, fmt::MaybeFormat=nothing) = write!(ws, row, col, :write_blank, fmt)
write_blank! = write!

write!(ws::Worksheet, row::Int64, col::Int64, dt::DateTime, fmt::MaybeFormat=nothing) = write!(ws, row, col, :write_datetime, dt, fmt)
write!(ws::Worksheet, row::Int64, col::Int64, dt::Date, fmt::MaybeFormat=nothing) = write!(ws, row, col, :write_datetime, dt, fmt)
write_datetime! = write!

write!(ws::Worksheet, row::Int64, col::Int64, bool::Bool, fmt::MaybeFormat=nothing) = write!(ws, row, col, :write_boolean, bool, fmt)
write_bool! = write!

write!(ws::Worksheet, row::Int64, col::Int64, u::Url, fmt::MaybeFormat=nothing) = write!(ws, row, col, :write_url, u.url, fmt)
write_url! = write!

function write_matrix!(ws::Worksheet, row::Int64, col::Int64, data::Matrix, fmt::MaybeFormat=nothing)
	re = size(data, 1)-1
	ce = size(data, 2)-1
	if re > ce
		for c in 0:ce
			write_column!(ws, row, col+c, vec(data[:, c+1]), fmt)
		end
	else
		for r in 0:re
			write_row!(ws, row+r, col, vec(data[r+1, :]), fmt)
		end
	end
end

function write_formula!(ws::Worksheet, row::Int64, col::Int64, formula::AbstractString, fmt::MaybeFormat=nothing; result::Data=0)
	ws.py[:write_formula](row, col, formula, @Fmt, result)
end

function write_formula!(ws::Worksheet, cell::AbstractString, formula::AbstractString, fmt::MaybeFormat=nothing; result::Data=0)
	ws.py[:write_formula](cell2rc(cell)..., formula, @Fmt, result)
end

function write_formula!(ws::Worksheet, row::Int64, col::Int64, formula::AbstractString, result::Data=0)
	ws.py[:write_formula](row, col, formula, nothing, result)
end

function write_formula!(ws::Worksheet, cell::AbstractString, formula::AbstractString, result::Data=0)
	ws.py[:write_formula](cell2rc(cell)..., formula, nothing, result)
end




function write_array_formula!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64, formula::AbstractString, fmt::MaybeFormat=nothing)
	ws.py[:write_array_formula](first_row, first_col, last_row, last_col, formula, @Fmt)
end

function write_array_formula!(ws::Worksheet, first_row::Int64, first_col::Int64, formula::AbstractString, fmt::MaybeFormat=nothing)
	ws.py[:write_array_formula](first_row, first_col, first_row, first_col, formula, @Fmt)
end

function write_array_formula!(ws::Worksheet, first_cell::AbstractString, last_cell::AbstractString, formula::AbstractString, fmt::MaybeFormat=nothing)
	ws.py[:write_array_formula](cell2rc(first_cell)..., cell2rc(last_cell)..., formula, @Fmt)
end

function write_array_formula!(ws::Worksheet, cell::AbstractString, formula::AbstractString, fmt::MaybeFormat=nothing)
	if search(cell, ':') > 0
		first, last = split(cell, ':')
	else
		first = last = cell
	end
	
	ws.py[:write_array_formula](cell2rc(first)..., cell2rc(last)..., formula, @Fmt)
end


function write_row!(ws::Worksheet, row::Int64, col::Int64, data::Array, fmt::MaybeFormat=nothing)
	ws.py[:write_row](row, col, vec(data), @Fmt)
	length(vec(data))
end

function write_row!(ws::Worksheet, cell::AbstractString, data::Array, fmt::MaybeFormat=nothing)
	ws.py[:write_row](cell2rc(cell)..., vec(data), @Fmt)
	length(vec(data))
end

function write_column!(ws::Worksheet, row::Int64, col::Int64, data::Array, fmt::MaybeFormat=nothing)
	ws.py[:write_column](row, col, vec(data), @Fmt)
	length(vec(data))
end

function write_column!(ws::Worksheet, cell::AbstractString, data::Array, fmt::MaybeFormat=nothing)
	ws.py[:write_column](cell2rc(cell)..., vec(data), @Fmt)
	length(vec(data))
end

function write_rich_string!(ws::Worksheet, row::Int64, col::Int64, parts...)
	unwrap(f::Format) = f.py
	unwrap(a) = a
	ws.py[:write_rich_string](row, col, [unwrap(p) for p in parts]...)
end
function write_rich_string!(ws::Worksheet, cell::AbstractString, parts...)
	unwrap(f::Format) = f.py
	unwrap(a) = a
	ws.py[:write_rich_string](cell2rc(cell)..., [unwrap(p) for p in parts]...)
end

set_first_sheet!(ws::Worksheet) = ws.py[:set_first_sheet]()

merge_range!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64, contents, fmt::MaybeFormat=nothing) = merge_range!(first_row, first_col, last_row, last_col, contents, fmt)

merge_range!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64, contents, fmt::Format) = ws.py[:merge_range](first_row, first_col, last_row, last_col, contents, fmt.py)

merge_range!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64, contents, fmt::Void) = ws.py[:merge_range](first_row, first_col, last_row, last_col, contents)

freeze_panes!(ws::Worksheet, row::Int64, col::Int64) = ws.py[:freeze_panes](row, col)
freeze_panes!(ws::Worksheet, cell::AbstractString) = ws.py[:freeze_panes](cell2rc(cell)...)
freeze_panes!(ws::Worksheet, row::Int64, col::Int64, top_row::Int64) = ws.py[:freeze_panes](row, col, top_row)
freeze_panes!(ws::Worksheet, row::Int64, col::Int64, top_row::Int64, left_col::Int64) = ws.py[:freeze_panes](row, col, top_row, left_col)

split_panes!(ws::Worksheet,x::Int64, y::Int64) = ws.py[:split_panes](x, y)
split_panes!(ws::Worksheet,x::Int64, y::Int64, top_row::Int64) = ws.py[:split_panes](x, y, top_row)
split_panes!(ws::Worksheet,x::Int64, y::Int64, top_row::Int64, left_col::Int64) = ws.py[:split_panes](x, y, top_row, left_col)

data_validation!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64, options::Dict) = ws.py[:data_validation](first_row, first_col, last_row, last_col, options)
data_validation!(ws::Worksheet, first_cell::AbstractString, last_cell::AbstractString, options::Dict) = ws.py[:data_validation](cell2rc(first_cell)..., cell2rc(last_cell)..., options)

function conditional_format!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64, options::Dict)
	for k in collect(keys(options))
		@printf STDERR "k %s type:%s\n" k typeof(options[k])
		if typeof(options[k]) == "XlsxWriter.Format"
			options[k] = options[k].py
		end
	end
	ws.py[:conditional_format](first_row, first_col, last_row, last_col, options)
end

conditional_format!(ws::Worksheet, first_cell::AbstractString, last_cell::AbstractString, options::Dict) = conditional_format!(ws, cell2rc(first_cell)..., cell2rc(last_cell)..., options)
function conditional_format!(ws::Worksheet, cell::AbstractString, options::Dict)
	if search(cell, ':') > 0
		f, s = split(cell, ':')
		conditional_format!(ws, cell2rc(f)..., cell2rc(s)..., options)
	else
		conditional_format!(ws, cell2rc(cell)..., cell2rc(cell)..., options)
	end
end


add_table!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64, options::Dict) = ws.py[:add_table](first_row, first_col, last_row, last_col, options)
add_table!(ws::Worksheet, first_cell::AbstractString, last_cell::AbstractString, options::Dict) = ws.py[:add_table](cell2rc(first_cell), cell2rc(last_cell), options)

add_sparkline!(ws::Worksheet, row::Int64, col::Int64, options::Dict) = ws.py[:add_sparkline](row, col, options)
add_sparkline!(ws::Worksheet, cell::AbstractString, options::Dict) = ws.py[:add_sparkline](cell2rc(cell)..., options)


insert_image!(ws::Worksheet, row::Int64, col::Int64, image::AbstractString, options::Dict=Dict()) = ws.py[:insert_image](row, col, image, options)
insert_image!(ws::Worksheet, cell::AbstractString, image::AbstractString, options::Dict=Dict()) = ws.py[:insert_image](cell2rc(cell)..., image, options)

insert_chart!(ws::Worksheet, row::Int64, col::Int64, ch::Chart, options::Dict=Dict()) = ws.py[:insert_chart](row, col, ch.py, options)
insert_chart!(ws::Worksheet, cell::AbstractString, ch::Chart, options::Dict=Dict()) = ws.py[:insert_chart](cell2rc(cell)..., ch.py, options)

insert_textbox!(ws::Worksheet, row::Int64, col::Int64, text::AbstractString, options::Dict=Dict()) = ws.py[:insert_textbox](row, col, text, options)
insert_textbox!(ws::Worksheet, cell::AbstractString, text::AbstractString, options::Dict=Dict()) = ws.py[:insert_textbox](cell2rc(cell)..., text, options)

insert_button!(ws::Worksheet, row::Int64, col::Int64, options::Dict) = wb.py[:insert_button](row, col, options)
insert_button!(ws::Worksheet, cell::AbstractString, options::Dict) = wb.py[:insert_button](cell2rc(cell)..., options)

write_comment!(ws::Worksheet, row::Int64, col::Int64, comment::AbstractString, options::Dict=Dict()) = ws.py[:write_comment](row, col, comment, options)
write_comment!(ws::Worksheet, cell::AbstractString, comment::AbstractString, options::Dict=Dict()) = ws.py[:write_comment](cell2rc(cell)..., comment, options)

show_comments!(ws::Worksheet) = ws.py[:show_comments]()
set_comments_author!(ws::Worksheet, author::AbstractString) = ws.py[:set_comments_author](author)

autofilter!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64) = ws.py[:autofilter](first_row, first_col, last_row, last_col)
autofilter!(ws::Worksheet, cells::AbstractString) = ws.py[:autofilter](cells)

filter_column!(ws::Worksheet, col::Int64, criteria::AbstractString) = ws.py[:filter_column](col, criteria)
filter_column!(ws::Worksheet, col::AbstractString, criteria::AbstractString) = ws.py[:filter_column](col, criteria)

filter_column_list!(ws::Worksheet, col::Int64, criteria::Array{AbstractString}) = ws.py[:filter_column](col, vec(criteria))
filter_column!(ws::Worksheet, col::AbstractString, criteria::Array{AbstractString}) = ws.py[:filter_column](col, vec(criteria))

set_selection!(ws::Worksheet, first_row::Int64, first_col::Int64, last_row::Int64, last_col::Int64) = ws.py[:set_selection](first_row, first_col, last_row, last_col)
set_selection!(ws::Worksheet, cells::AbstractString) = ws.py[:set_selection](cells)

set_default_row!(ws::Worksheet, height::Float64=15; hide_unused_rows::Bool=false) = ws.py[:set_default_row](height, hide_unused_rows)

outline_settings!(ws::Worksheet, visible::Bool=true, symbols_below::Bool=true, symbols_right::Bool=true, auto_style::Bool=false) = ws.py[:outline_settings](visible, symbols_below, symbols_right, auto_style)

macro PyFn(fn)
	:(sh.py[$fn]())
end

# worksheet
right_to_left!(sh::Worksheet) = @PyFn :right_to_left
hide_zero!(sh::Worksheet) = @PyFn :hide_zero


# worksheet or chartsheet
typealias Sheet Union{Worksheet, Chartsheet}
activate!(sh::Sheet) = @PyFn :activate
select!(sh::Sheet) = @PyFn :select
hide!(sh::Sheet) = @PyFn :hide
set_first_sheet!(sh::Sheet) = @PyFn :set_first_sheet
t!(sh::Sheet) = @PyFn :right_to_left
protect!(sh::Sheet, password::AbstractString, options::Dict) = sh.py[:protect](password, options)
set_zoom!(sh::Sheet, zoom::Int64) = sh.py[:set_zoom](zoom)
set_tab_color!(sh::Sheet, color::AbstractString) = sh.py[:set_tab_color](color)
set_landscape!(sh::Sheet) = @PyFn :set_landscape
set_portrait!(sh::Sheet) = @PyFn :set_portrait
set_paper!(sh::Sheet, index::Int64) = sh.py[:set_paper](index)
set_margins!(sh::Sheet, left::Float64=0.7, right::Float64=0.7, top::Float64=0.75, bottom::Float64=0.75) = sh.py[:set_margins](left, right, top, bottom)
set_header!(sh::Sheet, header::AbstractString, options::Dict=Dict()) = sh.py[:set_header](header, options)
set_header!(sh::Sheet, options::Dict) = sh.py[:set_header]("", options)
set_footer!(sh::Sheet, footer::AbstractString, options::Dict=Dict()) = sh.py[:set_footer](footer, options)
set_footer!(sh::Sheet, options::Dict) = sh.py[:set_footer]("", options)
get_name(sh::Sheet) = sh.py[:get_name]()




# Format
macro SFmt(sym)
	:(fmt.py[$sym](opt))
end 

set_font_name!(fmt::Format, opt::AbstractString) = @SFmt :set_font_name
set_font_size!(fmt::Format, opt::Int64) = @SFmt :set_font_size 
set_font_color!(fmt::Format, opt::AbstractString) = @SFmt :set_font_color
set_bold!(fmt::Format, opt::Bool=true) = @SFmt :set_bold
set_italic!(fmt::Format, opt::Bool=true) = @SFmt :set_italic
set_underline!(fmt::Format, opt::Bool=true) = @SFmt :set_underline
set_font_strikeout!(fmt::Format, opt::Bool=true) = @SFmt :set_font_strikeout
set_font_script!(fmt::Format, opt::Int64) = @SFmt :set_font_script
set_num_format!(fmt::Format, opt::AbstractString) = @SFmt :set_font_num_format
set_locked!(fmt::Format, opt::Bool=true) = @SFmt :set_locked
set_hidden!(fmt::Format, opt::Bool=true) = @SFmt :set_hidden
set_align!(fmt::Format, opt::AbstractString) = @SFmt :set_align
set_center_across!(fmt::Format, opt::AbstractString) = @SFmt :set_center_across
set_text_wrap!(fmt::Format, opt::Bool=true) = @SFmt :set_text_wrap
set_rotation!(fmt::Format, opt::Int64) = @SFmt :set_rotation
set_indent!(fmt::Format, opt::Int64) = @SFmt :set_indent
set_shrink!(fmt::Format, opt::Bool=true) = @SFmt :set_shrink
set_text_justlast!(fmt::Format, opt::Bool=true) = @SFmt :set_text_justlast
set_pattern!(fmt::Format, opt::Int64) = @SFmt :set_pattern
set_bg_color!(fmt::Format, opt::AbstractString) = @SFmt :set_bg_color
set_fg_color!(fmt::Format, opt::AbstractString) = @SFmt :set_fg_color
set_border!(fmt::Format, opt::Int64) = @SFmt :set_border
set_bottom!(fmt::Format, opt::Int64) = @SFmt :set_bottom
set_top!(fmt::Format, opt::Int64) = @SFmt :set_top
set_left!(fmt::Format, opt::Int64) = @SFmt :set_left
set_right!(fmt::Format, opt::Int64) = @SFmt :set_right
set_border_color!(fmt::Format, opt::AbstractString) = @SFmt :set_border_color
set_bottom_color!(fmt::Format, opt::AbstractString) = @SFmt :set_bottom_color
set_top_color!(fmt::Format, opt::AbstractString) = @SFmt :set_top_color
set_left_color!(fmt::Format, opt::AbstractString) = @SFmt :set_left_color
set_right_color!(fmt::Format, opt::AbstractString) = @SFmt :set_right_color
set_diag_border!(fmt::Format, opt::Int64) = @SFmt :set_diag_border
set_diag_type!(fmt::Format, opt::Int64) = @SFmt :set_diag_type
set_diag_color!(fmt::Format, opt::AbstractString) = @SFmt :set_diag_color

# Chart

macro ChOpts(fn)
	:(set_options!(ch.py, $fn, options))
end
set_options!(py::PyObject, fn::Symbol, options::Dict) = py[fn](options)

add_series!(ch::Chart, options::Dict=Dict()) = @ChOpts :add_series
set_x_axis!(ch::Chart, options::Dict=Dict()) = @ChOpts :set_x_axis
set_y_axis!(ch::Chart, options::Dict=Dict()) = @ChOpts :set_y_axis
set_x2_axis!(ch::Chart, options::Dict=Dict()) = @ChOpts :set_x2_axis
set_y2_axis!(ch::Chart, options::Dict=Dict()) = @ChOpts :set_y2_axis
combine!(ch1::Chart, ch2::Chart) = ch1.py[:combine](ch2.py)
set_size!(ch::Chart, options::Dict=Dict()) = @ChOpts :set_size
set_title!(ch::Chart, options::Dict=Dict()) = @ChOpts :set_title
set_legend!(ch::Chart, options::Dict=Dict()) = @ChOpts :set_legend
set_chartarea!(ch::Chart, options::Dict=Dict()) = @ChOpts :set_chartarea
set_plotarea!(ch::Chart, options::Dict=Dict()) = @ChOpts :set_plotarea
set_style!(ch::Chart, style_id::Int64) = ch.py[:set_style](style_id)
set_table!(ch::Chart, options::Dict=Dict()) = @ChOpts :set_table
set_up_down_bars!(ch::Chart, options::Dict=Dict()) = @ChOpts :set_up_down_bars
set_drop_lines!(ch::Chart, options::Dict=Dict()) = @ChOpts :set_drop_lines
set_high_low_lines!(ch::Chart, options::Dict=Dict()) = @ChOpts :set_high_low_lines
show_blanks_as!(ch::Chart, option::AbstractString) = ch.py[:show_blanks_as](option)
show_hidden_data!(ch::Chart) = ch.py[:show_hidden_data]()
set_rotation!(ch::Chart, angle::Int64) = ch.py[:set_rotation](angle)
set_hole_size!(ch::Chart, hole::Int64) = ch.py[:set_hole_size](hole)

# Chartsheet

set_chart!(cs::Chartsheet, ch::Chart) = cs.py[:set_chart](ch.py)


end

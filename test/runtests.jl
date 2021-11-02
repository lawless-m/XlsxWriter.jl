
ENV["XLSXWRITER_PATH"] = "C:\\Users\\matthew.heath\\repos\\XlsxWriter.py"

using SHA
using XlsxWriter
using Test

broadcast(rm, filter(f->endswith(f, ".xlsx"), readdir()));

filehash(fn) = open(fn) do f  bytes2hex(sha2_256(f))  end

@testset "XlsxWriter.jl" begin
	@test include("array_formula.jl")
	@test include("autofilter.jl")
	@test include("cell_indentation.jl")
	@test include("chart.jl")
	@test include("chart_area.jl")
	@test include("chart_bar.jl")
	@test include("chart_clustered.jl")
	@test include("chart_column.jl")
	@test include("chart_combined.jl")
	@test include("chart_data_table.jl")
	@test include("chart_data_tools.jl")
	@test include("chart_date_axis.jl")
	@test include("chart_doughnut.jl")
	@test include("chart_gradient.jl")
	@test include("chart_line.jl")
	@test include("chart_pareto.jl")
	@test include("chart_pattern.jl")
	@test include("chart_pie.jl")
	@test include("chart_radar.jl")
	@test include("chart_scatter.jl")
	@test include("chart_secondary_axis.jl")
	@test include("chart_stock.jl")
	@test include("chart_styles.jl")
	@test include("chartsheet.jl")
	@test include("comments1.jl")
	@test include("comments2.jl")
	@test include("conditional_format.jl")
	@test include("demo.jl") 
	@test include("in_memory.jl")
end

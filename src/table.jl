# __rowtable

struct XLSheetRowIter
    sheet::AbstractXLSheet
    current_row::Int
    total_rows::Int
    columns::Vector{Symbol}
end

function XLSheetRowIter(x::AbstractXLSheet; header::Bool = false, alt_keys = nothing)
    columns = header ? copy(x[1, :]) : index_to_column_letter.(1:xl_ncol(x))

    if isa(alt_keys, AbstractVector)
        length(alt_keys) == xl_ncol(x) || error("Column header mismatch: expected $(xl_ncol(x)) headers, but received $(length(alt_keys)).")
        columns = alt_keys
    elseif isa(alt_keys, AbstractDict)
        columns = replace(columns, alt_keys...)
    end

    return XLSheetRowIter(x, header ? 2 : 1, xl_nrow(x), Symbol.(columns))
end

function Base.iterate(itr::XLSheetRowIter, state::Int = itr.current_row)
    state > itr.total_rows && return nothing
    rd = [itr.sheet[state, j] for j = 1:length(itr.columns)]
    return (NamedTuple{Tuple(itr.columns)}(rd), state + 1)
end

function Base.length(iter::XLSheetRowIter)
    return iter.total_rows - iter.current_row + 1
end

function Base.eachrow(x::AbstractXLSheet; kw...)
    return XLSheetRowIter(x; kw...)
end

"""
    xl_rowtable(sheet::AbstractXLSheet; kw...) -> Vector{NamedTuple}

Converts sheet rows to a `Vector` of `NamedTuples`.

## Keyword arguments
- `alt_keys::Dict{String,String}`: Alternative custom column headers.
- `header::Bool = false`: Use first row elements as column headers.

## Examples

```julia-repl
julia> xlsx = xl_parse(xl_sample_stock_xlsx())
1-element XLWorkbook:
 41x6 XLSheet("Stock")

julia> xl_rowtable(xlsx["Stock"]["A1:C30"], header = true)
29-element Vector{NamedTuple{(:name, :price, :h24)}}:
 (name = "MSFT", price = "430.16", h24 = "0.0007")
 (name = "AAPL", price = "189.98", h24 = "-0.0005")
 (name = "NVDA", price = "1064.69", h24 = "0.0045")
 (name = "GOOG", price = "176.33", h24 = "-0.0006")
 ⋮
 (name = "LLY", price = 807.43, h24 = -0.0024)
 (name = "AVGO", price = 1407.84, h24 = 0.0036)
```
"""
function xl_rowtable(x::AbstractXLSheet; kw...)
    return map(x -> x, XLSheetRowIter(x; kw...))
end

"""
    xl_columntable(sheet::AbstractXLSheet; kw...) -> Vector{NamedTuple}

Converts sheet columns to a `Vector` of `NamedTuples`.

## Keyword arguments
- `alt_keys::Dict{String,String}`: Alternative custom column headers.
- `header::Bool = false`: Use first row elements as column headers.

## Examples

```julia-repl
julia> xlsx = xl_parse(xl_sample_stock_xlsx())
1-element XLWorkbook:
 41x6 XLSheet("Stock")

julia> alt_keys = Dict("A" => "Name", "B" => "Price", "C" => "H24");

julia> xl_columntable(xlsx["Stock"][2:end, 1:3]; alt_keys)
(
    Name = Any["MSFT", "AAPL"  …  "JNJ", "ORCL"]
    Price = Any[430.16, 189.98  …  146.97, 122.91],
    H24 = Any[0.0007, -0.0005  …  0.0007, -0.0003],
)
```
"""
function xl_columntable(x::AbstractXLSheet; kw...)
    iter = XLSheetRowIter(x; kw...)
    return NamedTuple{Tuple(iter.columns)}(eachcol(x[iter.current_row:iter.total_rows, :]))
end

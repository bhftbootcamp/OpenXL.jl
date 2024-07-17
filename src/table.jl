# xl_table

struct XLSheetRowIter{names} <: AbstractVector{NamedTuple{names}}
    sheet::AbstractXLSheet
    current_row::Int
    total_rows::Int
    columns::Vector{Symbol}
end

function XLSheetRowIter(x::AbstractXLSheet; header::Bool = false, alt_keys = nothing)
    headers = header ? copy(x[1, :]) : index_to_column_letter.(1:xl_ncol(x))

    columns = if isa(alt_keys, AbstractVector)
        length(alt_keys) == xl_ncol(x) || error("Column header mismatch: expected $(xl_ncol(x)) headers, but received $(length(alt_keys)).")
        Symbol.(alt_keys)
    elseif isa(alt_keys, AbstractDict)
        Symbol.(replace(headers, alt_keys...))
    else
        Symbol.(headers)
    end

    return XLSheetRowIter{Tuple(columns)}(x, header ? 2 : 1, xl_nrow(x), columns)
end

function Base.show(io::IO, x::XLSheetRowIter)
    print(io, x.total_rows, "-rows XLSheetRowIter(\"", xl_sheetname(x.sheet), "\")")
end

function Base.show(io::IO, ::MIME"text/plain", x::XLSheetRowIter)
    show(io, x)
    print(io, "\n")
    show(io, x.sheet.table)
end

Base.length(x::XLSheetRowIter) = x.total_rows - x.current_row + 1
Base.size(x::XLSheetRowIter) = (length(x),)

function Base.eltype(::XLSheetRowIter{names}) where {names}
    return NamedTuple{names}
end

function Base.iterate(iter::XLSheetRowIter{names}, state::Int = iter.current_row) where {names}
    state > iter.total_rows && return nothing
    rd = [iter.sheet[state, j] for j = 1:length(iter.columns)]
    return (NamedTuple{names}(rd), state + 1)
end

Base.getindex(x::XLSheetRowIter, i::Int) = first(iterate(x, i))

"""
    eachrow(x::AbstractXLSheet; kw...)

Creates a table row iterator. Row slices are returned as `NamedTuple`.

## Keyword arguments
- `alt_keys`: Alternative custom column headers (`Dict{String,String}` or `Vector{String}`).
- `header::Bool = false`: Use first row elements as column headers.

## Examples
```julia-repl
julia> xlsx = xl_parse(xl_sample_stock_xlsx())

julia for row in eachrow(xlsx["Stock"]; header = true)
          println(row)
      end
(name = "MSFT", price = "430.16", h24 = "0.0007", volume = "11855456", ...)
(name = "AAPL", price = "189.98", h24 = "-0.0005", volume = "36327000", ...)
(name = "NVDA", price = "1064.69", h24 = "0.0045", volume = "42948000", ...)
(name = "GOOG", price = "176.33", h24 = "-0.0006", volume = "11404000", ...)
 ⋮
(name = "JNJ", price = 146.97, h24 = 0.0007, volume = 7.173e6, ...)
(name = "ORCL", price = 122.91, h24 = -0.0003, volume = 5.984e6, ...)
```
"""
function Base.eachrow(x::AbstractXLSheet; kw...)
    return XLSheetRowIter(x; kw...)
end

"""
    xl_rowtable(sheet::AbstractXLSheet; kw...) -> Vector{NamedTuple}

Converts sheet rows to a `Vector` of `NamedTuples`.

## Keyword arguments
- `alt_keys`: Alternative custom column headers (`Dict{String,String}` or `Vector{String}`).
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
    return collect(XLSheetRowIter(x; kw...))
end

"""
    xl_columntable(sheet::AbstractXLSheet; kw...) -> Vector{NamedTuple}

Converts sheet columns to a `Vector` of `NamedTuples`.

## Keyword arguments
- `alt_keys`: Alternative custom column headers (`Dict{String,String}` or `Vector{String}`).
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
    return eltype(iter)(eachcol(x[iter.current_row:iter.total_rows, :]))
end

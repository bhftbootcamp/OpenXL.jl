# interface

"""
    AbstractXLSheet <: AbstractArray{Any,2}

Abstract supertype for [`XLSheet`](@ref) and [`XLTable`](@ref).
"""
abstract type AbstractXLSheet <: AbstractArray{Any,2} end

#__XLTable

"""
    XLTable <: AbstractXLSheet

Table representation of the [`XLSheet`](@ref) object.
Supports the same table operations as `XLSheet`.

## Fields
- `data::Matrix`: Table data.

## Accessors
- `xl_nrow(x::XLTable)`: Number of rows.
- `xl_ncol(x::XLTable)`: Number of columns.

See also: [`xl_rowtable`](@ref), [`xl_columntable`](@ref), [`xl_print_sheet`](@ref).

## Examples
```julia-repl
julia> xlsx = xl_parse(xl_sample_stock_xlsx())
1-element XLWorkbook:
 41x6 XLSheet("Stock")

julia> xlsx["Stock"]["A1:D25"]
25x4 XLTable
 Sheet │ A      B         C        D
───────┼───────────────────────────────────────
     1 │ name   price     h24      volume
     2 │ MSFT   430.16    0.0007   11855456
     3 │ AAPL   189.98    -0.0005  36327000
     4 │ NVDA   1064.69   0.0045   42948000
     ⋮ │ ⋮      ⋮         ⋮        ⋮
    24 │ NVDA   1064.69   0.0045   4.2948e7
    25 │ GOOG   176.33    -0.0006  1.1404e7
```
"""
struct XLTable <: AbstractXLSheet
    data::Matrix
end

Base.size(x::XLTable) = size(x.data)
Base.size(x::XLTable, dim) = size(x.data, dim)

xl_nrow(x::AbstractXLSheet) = size(x, 1)
xl_ncol(x::AbstractXLSheet) = size(x, 2)

Base.:(==)(x::XLTable, y::XLTable) = x.data == y.data
Base.isequal(x::XLTable, y::XLTable) = isequal(x.data, y.data)

Base.getindex(x::XLTable, i::Int) = x.data[i]

function Base.getindex(x::XLTable, inds::Vararg{Any,2})
    data = x.data[inds...]
    return data isa Matrix ? XLTable(data) : data
end

function Base.getindex(x::XLTable, addr::AbstractString)
    row_range, column_range = parse_cell_range(addr, size(x))
    return getindex(x, row_range, column_range)
end

function Base.setindex!(x::XLTable, value::Any, addr::AbstractString)
    row_range, column_range = parse_cell_range(addr, size(x))
    return setindex!(x, value, row_range, column_range)
end

function Base.setindex!(x::XLTable, value::Any, inds...)
    return setindex!(x.data, value, inds...)
end

Base.convert(::Type{XLTable}, x::AbstractMatrix) = XLTable(x)
Base.convert(::Type{Matrix}, x::XLTable) = x.data

function Base.show(io::IO, x::XLTable; kw...)
    xl_print_sheet(io, x; kw...)
end

function Base.show(io::IO, ::MIME"text/plain", x::XLTable; kw...)
    num_rows, num_cols = size(x)
    print(io, num_rows, "x", num_cols, " XLTable\n")
    xl_print_sheet(io, x; kw...)
end

function Base.print(io::IO, x::XLTable; kw...)
    xl_print_sheet(io, x; compact = false, kw...)
end

#__ XLSheet

"""
    XLSheet <: AbstractXLSheet

Sheet of the [`XLWorkbook`](@ref).
Supports indexing like a regular `Matrix`, as well as address indexing (e.g. `A`, `A1`, `AB3` or range `D:E`, `B1:C10`, etc.).

The sheet slice will be converted into a separate [`XLTable`](@ref).

## Fields
- `name::String`: Sheet name.
- `table::XLTable`: Table representation.

## Accessors
- `xl_name(x::XLSheet)`: Sheet name.
- `xl_nrow(x::XLTable)`: Number of rows.
- `xl_ncol(x::XLTable)`: Number of columns.

See also: [`xl_rowtable`](@ref), [`xl_columntable`](@ref), [`xl_print_sheet`](@ref).

## Examples

```julia-repl
julia> xlsx = xl_parse(xl_sample_stock_xlsx())
1-element XLWorkbook:
 41x6 XLSheet("Stock")

julia> sheet = xlsx["Stock"]
41x6 XLSheet("Stock")
 Sheet │ A      B         C        D         E              F
───────┼─────────────────────────────────────────────────────────────────────
     1 │ name   price     h24      volume    mkt            sector
     2 │ MSFT   430.16    0.0007   11855456  3197000000000  Technology Serv…
     3 │ AAPL   189.98    -0.0005  36327000  2913000000000  Electronic Tech…
     4 │ NVDA   1064.69   0.0045   42948000  2662000000000  Electronic Tech…
     ⋮ │ ⋮      ⋮         ⋮        ⋮         ⋮              ⋮
    40 │ JNJ    146.97    0.0007   7.173e6   3.5371e11      Health Technolo…
    41 │ ORCL   122.91    -0.0003  5.984e6   3.3782e11      Technology Serv…
```
"""
struct XLSheet <: AbstractXLSheet
    name::String
    id::Int64
    table::XLTable
end

xl_name(x::XLSheet) = x.name

Base.size(x::XLSheet) = size(x.table)
Base.size(x::XLSheet, dim) = size(x.table, dim)

Base.:(==)(x::XLSheet, y::XLSheet) = x.table == y.table
Base.isequal(x::XLSheet, y::XLSheet) = isequal(x.table, y.table)

Base.getindex(x::XLSheet, inds::Any...) = x.table[inds...]
Base.setindex!(x::XLSheet, value, inds...) = setindex!(x.table, value, inds...)

function Base.show(io::IO, x::XLSheet)
    num_rows, num_cols = size(x)
    print(io, num_rows, "x", num_cols, " XLSheet(\"", x.name, "\")")
end

function Base.show(io::IO, ::MIME"text/plain", x::XLSheet)
    show(io, x)
    print(io, "\n")
    show(io, x.table)
end

function Base.print(io::IO, x::XLSheet; kw...)
    show(io, x)
    print(io, "\n")
    print(io, x.table; kw...)
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
function xl_rowtable(
    x::AbstractXLSheet;
    alt_keys::Dict{String,String} = Dict{String,String}(),
    header::Bool = false,
)
    column_keys, t_data = if header
        headers = x[1, :]
        replace!(headers, alt_keys...)
        Symbol.(headers), view(x, 2:xl_nrow(x), :)
    else
        gen_column_keys(xl_ncol(x), alt_keys), x
    end

    return map(eachrow(t_data)) do row
        return NamedTuple{Tuple(column_keys)}(row)
    end
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
function xl_columntable(
    x::AbstractXLSheet;
    alt_keys::Dict{String,String} = Dict{String,String}(),
    header::Bool = false,
)
    column_keys, t_data = if header
        headers = x[1, :]
        replace!(headers, alt_keys...)
        Symbol.(headers), view(x, 2:xl_nrow(x), :)
    else
        gen_column_keys(xl_ncol(x), alt_keys), x
    end

    replace!(column_keys, alt_keys...)

    return NamedTuple{Tuple(column_keys)}(eachcol(t_data))
end

#__ XLWorkbook

"""
    XLWorkbook <: AbstractVector{XLSheet}

Represents an Excel workbook containing [`XLSheet`](@ref).

## Fields
- `sheets::Vector{XLSheet}`: Workbook sheets.

## Accessors
- `xl_names(x::XLWorkbook)`: Workbook sheet names.

See also: [`xl_parse`](@ref).
"""
struct XLWorkbook <: AbstractVector{XLSheet}
    sheets::Vector{XLSheet}

    function XLWorkbook(xl::XL)
        r = map(xl.workbook_xml.sheets.sheet) do ws
            sheet = xl[ws]
            result = Matrix{Any}(nothing, nrow(sheet), ncol(sheet))
            for row in sheet.sheetData.row
                for cell in row.c
                    column_addr, _ = parse_cell_addr(cell.r)
                    result[row.r, column_addr] = xl[cell]
                end
            end
            XLSheet(ws.name, ws.sheetId, result)
        end
        return new(r)
    end
end

xl_names(x::XLWorkbook) = map(xl_name, x.sheets)

Base.size(x::XLWorkbook) = size(x.sheets)
Base.getindex(x::XLWorkbook) = x.sheets
Base.getindex(x::XLWorkbook, i::Int) = x.sheets[i]

function Base.getindex(x::XLWorkbook, k::AbstractString)
    sheets = x[]
    i = findfirst(sheet -> xl_name(sheet) == k, sheets)
    i === nothing && throw(KeyError(k))
    return sheets[i]
end

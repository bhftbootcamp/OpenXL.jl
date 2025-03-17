# xl_interface

"""
    AbstractXLCell

Abstract supertype for table cells.
"""
abstract type AbstractXLCell end

"""
    XLCell{T} <: AbstractXLCell

A table cell containing a value and information about its formatting.

!!! note
    To extract the value of a cell, the empty index operator (`[]`) can be used.

## Fields
- `value::T`: Cell value.
- `format::Int`: Number format id (see [Number Format](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-3.0.1)).
"""
struct XLCell{T} <: AbstractXLCell
    value::T
    format::Int

    function XLCell(val::T, id::Int) where {T<:Union{Nothing,String,Float64,Bool,DateTime}}
        return new{T}(val, id)
    end
end

Base.eltype(::XLCell{T}) where {T} = T
Base.convert(::Type{XLCell}, x::Nothing) = XLCell(nothing, 0)

Base.:(==)(x::XLCell, y::XLCell) = cell_value(x) == cell_value(y)
Base.isequal(x::XLCell, y::XLCell) = isequal(cell_value(x), cell_value(y))

Base.string(x::XLCell{Nothing}) = "*"
Base.string(x::XLCell{String}) = repr(cell_value(x))
Base.string(x::XLCell{Float64}) = format_number(cell_value(x), cell_format(x))
Base.string(x::XLCell{DateTime}) = format_datetime(cell_value(x), cell_format(x))

cell_value(x::XLCell) = x.value
cell_format(x::XLCell) = x.format
cell_string(x::XLCell) = string(x)

Base.show(io::IO, x::XLCell{T}) where {T} = print(io, "XLCell{", T, "}(", cell_string(x), ")")

function Base.getindex(x::XLCell)
    return cell_value(x)
end

"""
    AbstractXLSheet <: AbstractArray{Any,2}

Abstract supertype for [`XLSheet`](@ref) and [`SubXLSheet`](@ref).
"""
abstract type AbstractXLSheet <: AbstractArray{Any,2} end

"""
    XLSheet <: AbstractXLSheet

Sheet of the [`XLWorkbook`](@ref).
Supports indexing like a regular `Matrix`, as well as address indexing (e.g. `A`, `A1`, `AB3` or range `D:E`, `B1:C10`, etc.).

The sheet slice will be converted into a [`SubXLSheet`](@ref).

## Fields
- `name::String`: Sheet name.
- `table::Matrix{AbstractXLCell}`: Table representation.

## Accessors
- `xl_sheetname(x::XLSheet)`: Sheet name.
- `xl_nrow(x::XLSheet)`: Number of rows.
- `xl_ncol(x::XLSheet)`: Number of columns.

See also: [`xl_rowtable`](@ref), [`xl_columntable`](@ref), [`xl_print`](@ref).

## Examples

```julia-repl
julia> xlsx = xl_parse(read("assets/stock_sample_data.xlsx"))
1-element XLWorkbook:
 41x6 XLSheet("Stock")

julia> sheet = xlsx["Stock"]
41x6 XLSheet("Stock")
 Sheet │ A     B        C        D            E          F                 
───────┼──────────────────────────────────────────────────────────────────
     1 │ name  price    h24      volume       mkt        sector            
     2 │ MSFT  430.16   0.0007   1.1855456e7  3.197e12   Technology Serv…  
     3 │ AAPL  189.98   -0.0005  3.6327e7     2.913e12   Electronic Tech…  
     4 │ NVDA  1064.69  0.0045   4.2948e7     2.662e12   Electronic Tech…  
     ⋮ │ ⋮     ⋮        ⋮        ⋮            ⋮          ⋮                 
    40 │ JNJ   146.97   0.0007   7.173e6      3.5371e11  Health Technolo…  
    41 │ ORCL  122.91   -0.0003  5.984e6      3.3782e11  Technology Serv… 
```
"""
struct XLSheet <: AbstractXLSheet
    name::String
    id::Int64
    table::Matrix{XLCell}
end

Base.size(x::AbstractXLSheet) = (size(x.table, 1), size(x.table, 2))
Base.size(x::AbstractXLSheet, dim) = size(x.table, dim)

xl_sheetname(x::XLSheet) = x.name
xl_table(x::XLSheet) = collect(x)

xl_nrow(x::AbstractXLSheet) = size(x, 1)
xl_ncol(x::AbstractXLSheet) = size(x, 2)

Base.:(==)(x::AbstractXLSheet, y::AbstractXLSheet) = x.table == y.table
Base.isequal(x::AbstractXLSheet, y::AbstractXLSheet) = isequal(x.table, y.table)

Base.convert(::Type{Matrix}, x::AbstractXLSheet) = xl_table(x)

function Base.show(io::IO, x::T; kw...) where {T<:AbstractXLSheet}
    num_rows, num_cols = size(x)
    print(io, num_rows, "x", num_cols, " $T(\"", xl_sheetname(x), "\")")
end

function Base.show(io::IO, ::MIME"text/plain", x::AbstractXLSheet; kw...)
    show(io, x)
    print(io, "\n")
    xl_print(io, x; kw...)
end

function Base.print(io::IO, x::AbstractXLSheet; kw...)
    xl_print(io, x; compact = false, kw...)
end

"""
    SubXLSheet <: AbstractXLSheet

Slice view of the [`XLSheet`](@ref) object.
Supports the same operations as `XLSheet`.

## Fields
- `data::Matrix`: Table data.

## Accessors
- `parent(x::SubXLSheet)`: Parent sheet.
- `xl_nrow(x::SubXLSheet)`: Number of rows.
- `xl_ncol(x::SubXLSheet)`: Number of columns.

See also: [`xl_rowtable`](@ref), [`xl_columntable`](@ref), [`xl_print`](@ref).

## Examples
```julia-repl
julia> xlsx = xl_parse(read("assets/stock_sample_data.xlsx"))
1-element XLWorkbook:
 41x6 XLSheet("Stock")

julia> xlsx["Stock"]["A1:D25"]
25x4 SubXLSheet("Stock")
 Sheet │ A     B        C        D            
───────┼─────────────────────────────────────
     1 │ name  price    h24      volume       
     2 │ MSFT  430.16   0.0007   1.1855456e7  
     3 │ AAPL  189.98   -0.0005  3.6327e7     
     4 │ NVDA  1064.69  0.0045   4.2948e7     
     ⋮ │ ⋮     ⋮        ⋮        ⋮              
    24 │ NVDA  1064.69  0.0045   4.2948e7     
    25 │ GOOG  176.33   -0.0006  1.1404e7  
```
"""
struct SubXLSheet <: AbstractXLSheet
    parent::AbstractXLSheet
    table::SubArray
end

Base.parent(x::SubXLSheet) = isa(x.parent, SubXLSheet) ? parent(x.parent) : x.parent
xl_sheetname(x::SubXLSheet) = xl_sheetname(parent(x))
xl_table(x::SubXLSheet) = collect(x)

function Base.getindex(x::AbstractXLSheet, i::Int, j::Int)
    return x.table[i, j][]
end

function Base.getindex(x::AbstractXLSheet, inds::Vararg{Any,2})
    return SubXLSheet(x, view(x.table, inds...))
end

function range_to_indices(x::AbstractXLSheet, parts::NTuple{2,CellRange})
    l_part, r_part = parts
    num_row = xl_nrow(x)

    return if all(isnothing, (r_part.row, r_part.column))
        if isnothing(l_part.row)
            (:), (l_part.column:l_part.column)
        else
            l_part.row, l_part.column
        end
    else
        if isnothing(l_part.row) && isnothing(r_part.row)
            (:), (l_part.column:r_part.column)
        elseif isnothing(r_part.row)
            (l_part.row:num_row), (l_part.column:r_part.column)
        elseif isnothing(l_part.row)
            (r_part.row:num_row), (l_part.column:r_part.column)
        else
            (l_part.row:r_part.row), (l_part.column:r_part.column)
        end
    end
end

function Base.getindex(x::AbstractXLSheet, addr::AbstractString)
    inds = range_to_indices(x, parse_cell_range(addr))
    return getindex(x, inds...)
end

"""
    XLWorkbook <: AbstractVector{XLSheet}

Represents an Excel workbook containing [`XLSheet`](@ref).

## Fields
- `sheets::Vector{XLSheet}`: Workbook sheets.

## Accessors
- `xl_sheetnames(x::XLWorkbook)`: Workbook sheet names.

See also: [`xl_parse`](@ref).
"""
struct XLWorkbook <: AbstractVector{XLSheet}
    sheets::Vector{XLSheet}
end

Base.size(x::XLWorkbook) = size(x.sheets)

Base.getindex(x::XLWorkbook) = x.sheets
Base.getindex(x::XLWorkbook, i::Int) = x.sheets[i]

xl_sheetnames(x::XLWorkbook) = map(xl_sheetname, x.sheets)

function Base.getindex(x::XLWorkbook, k::AbstractString)
    sheets = x[]
    i = findfirst(sheet -> xl_sheetname(sheet) == k, sheets)
    i === nothing && throw(KeyError(k))
    return sheets[i]
end

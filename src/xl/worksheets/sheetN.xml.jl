module WorksheetXML

export Worksheet, read_worksheet, nrow, ncol

using ZipFile
using Serde

using ..OpenXL: Maybe, ExcelFile, parse_cell_addr

struct FormulaItem
    _::Maybe{String}
end

struct ValueItem
    _::Maybe{String}
end

struct TextItem
    _::String
end

Base.string(x::TextItem) = x._

function Serde.deser(::Type{TextItem}, ::Type{String}, x::Nothing)
    return ""
end

struct RichTextItem
    t::Maybe{TextItem}
end

Base.string(x::RichTextItem) = x.t._

struct InlineStringItem
    t::Maybe{TextItem}
    r::Vector{RichTextItem}
end

function Base.string(s::InlineStringItem)
    return isnothing(s.t) ? join(map(string, s.r)) : string(s.t)
end

function Serde.deser(
    ::Type{InlineStringItem},
    ::Type{Vector{T}},
    x::AbstractDict,
) where {T<:RichTextItem}
    return T[Serde.deser(T, x)]
end

function Serde.deser(
    ::Type{InlineStringItem},
    ::Type{Vector{T}},
    x::Nothing,
) where {T<:RichTextItem}
    return T[]
end

struct CellItem
    r::String
    s::Maybe{String}
    t::Maybe{String}
    v::Maybe{ValueItem}
    is::Maybe{InlineStringItem}
    f::Maybe{FormulaItem}
end

struct RowItem
    r::Int
    c::Vector{CellItem}
end

function Serde.deser(
    ::Type{RowItem},
    ::Type{Vector{T}},
    x::AbstractDict,
) where {T<:CellItem}
    return T[Serde.deser(T, x)]
end

function Serde.deser(
    ::Type{RowItem},
    ::Type{Vector{T}},
    x::Nothing,
) where {T<:CellItem}
    return T[]
end

struct SheetDataItem
    row::Vector{RowItem}
end

function Serde.deser(
    ::Type{SheetDataItem},
    ::Type{Vector{T}},
    x::AbstractDict,
) where {T<:RowItem}
    return T[Serde.deser(T, x)]
end

function Serde.deser(
    ::Type{SheetDataItem},
    ::Type{Vector{T}},
    x::Nothing,
) where {T<:RowItem}
    return T[]
end

struct Worksheet <: ExcelFile
    sheetData::SheetDataItem
end

function nrow(worksheet::Worksheet)
    return maximum(row -> row.r, worksheet.sheetData.row, init = 0)
end

function ncol(worksheet::Worksheet)
    return maximum([parse_cell_addr(last(x.c).r).column for x in worksheet.sheetData.row if !isempty(x.c)], init = 0)
end

end

module WorksheetXML

export Worksheet, nrow, ncol

using Serde
import ..ExcelFile, ..parse_cell_addr

struct FormulaItem
    _::Union{Nothing,String}
end

struct ValueItem
    _::Union{Nothing,String}
end

struct TextItem
    _::String
end

Base.string(x::TextItem) = x._

function Serde.deser(::Type{TextItem}, ::Type{String}, x::Nothing)
    return ""
end

struct RichTextItem
    t::Union{Nothing,TextItem}
end

Base.string(x::RichTextItem) = x.t._

struct InlineStringItem
    t::Union{Nothing,TextItem}
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
    s::Union{Nothing,Int64}
    t::Union{Nothing,String}
    v::Union{Nothing,ValueItem}
    is::Union{Nothing,InlineStringItem}
    f::Union{Nothing,FormulaItem}
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

# <worksheet>
#     <sheetData>
#         <row r="1" ht="15.75" customHeight="1">
#             <c r="A1" s="1">
#                 <v>1.0</v>
#             </c>
#             <c r="B1" s="2" t="s">
#                 <v>0</v>
#             </c>
#             <c r="C1" s="3">
#                 <v>12345.12345</v>
#             </c>
#         </row>
#         ...
#     </sheetData>
# </worksheet>
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

module StylesXML

export Styles

using Serde
import ..ExcelFile

struct XFItem
    numFmtId::Int64
    fontId::Union{Nothing,Int64}
    fillId::Union{Nothing,Int64}
    borderId::Union{Nothing,Int64}
    xfId::Union{Nothing,Int64}
    applyNumberFormat::Union{Nothing,Int64}
    quotePrefix::Union{Nothing,Int64}
end

struct CellXfsItem
    count::Int64
    xf::Vector{XFItem}
end

function Serde.deser(
    ::Type{CellXfsItem},
    ::Type{Vector{T}},
    x::AbstractDict,
) where {T<:XFItem}
    return T[Serde.deser(T, x)]
end

function Serde.deser(
    ::Type{CellXfsItem},
    ::Type{Vector{T}},
    x::Nothing,
) where {T<:XFItem}
    return T[]
end

struct numFmtItem
    numFmtId::Int64
    formatCode::String
end

struct CellnumFmtsItem
    count::Int64
    numFmt::Vector{numFmtItem}
end

function Serde.deser(
    ::Type{CellnumFmtsItem},
    ::Type{Vector{T}},
    x::AbstractDict,
) where {T<:numFmtItem}
    return T[Serde.deser(T, x)]
end

function Serde.deser(
    ::Type{CellnumFmtsItem},
    ::Type{Vector{T}},
    x::Nothing,
) where {T<:numFmtItem}
    return T[]
end

# <styleSheet>
#     <cellXfs count="6">
#         <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" .../>
#         ...
#     </cellXfs>
#     <numFmts count="3">
#         <numFmt numFmtId="164" formatCode="M/d/yyyy" />
#         ...
#     </numFmts>
# </styleSheet>
struct Styles <: ExcelFile
    cellXfs::CellXfsItem
    numFmts::CellnumFmtsItem
end

const default_format_xf = XFItem(0, nothing, nothing, nothing, nothing, nothing, nothing)

Base.getindex(x::CellXfsItem, s::Int64) = get(x.xf, s + 1, default_format_xf)
Base.getindex(x::CellXfsItem, s::Nothing) = default_format_xf

const default_format_fmt = numFmtItem(0, "")
const default_format_fmts = CellnumFmtsItem(0, numFmtItem[])

Base.getindex(x::CellnumFmtsItem, s::Int64) = get(x.numFmt, s + 1, default_format_fmt)
Base.getindex(x::CellnumFmtsItem, s::Nothing) = default_format_fmt

function Serde.deser(
    ::Type{Styles},
    ::Type{CellnumFmtsItem},
    x::Nothing,
)
    return default_format_fmts
end

end

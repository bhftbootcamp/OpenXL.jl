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

# <styleSheet>
#     <cellXfs count="6">
#         <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" .../>
#         ...
#     </cellXfs>
# </styleSheet>
struct Styles <: ExcelFile
    cellXfs::CellXfsItem
end

const default_format = XFItem(0, nothing, nothing, nothing, nothing, nothing, nothing)

Base.getindex(x::Styles, s::Int64) = get(x.cellXfs.xf, s + 1, default_format)
Base.getindex(x::Styles, s::Nothing) = default_format

end

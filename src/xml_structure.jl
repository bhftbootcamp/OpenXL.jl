# xml_structure

abstract type AbstractXMLNode end

#__ XMLWorkbook

struct XMLWorksheet <: AbstractXMLNode
    sheetId::Int64
    name::String
    id::String
    state::Maybe{String}
end

struct XMLWorksheets <: AbstractXMLNode
    sheet::Vector{XMLWorksheet}
end

function Serde.deser(::Type{XMLWorksheets}, ::Type{Vector{T}}, x::AbstractDict) where {T<:XMLWorksheet}
    return T[Serde.deser(T, x)]
end

struct XMLWorkbook <: AbstractXMLNode
    sheets::XMLWorksheets
end

#__ XMLSheet

struct XMLCellFormula <: AbstractXMLNode
    _::Maybe{String}
end

struct XMLCellValue <: AbstractXMLNode
    _::Maybe{String}
end

struct XMLCellText <: AbstractXMLNode
    _::String
end

function Serde.deser(::Type{XMLCellText}, ::Type{String}, x::Nothing)
    return ""
end

struct XMLRichText <: AbstractXMLNode
    t::Maybe{XMLCellText}
end

struct ISSI <: AbstractXMLNode
    t::Maybe{XMLCellText}
    r::Vector{XMLRichText}
end

function Base.string(s::ISSI)
    return isnothing(s.t) ? join(map(r -> r.t._, s.r)) : s.t._
end

function Serde.deser(::Type{ISSI}, ::Type{Vector{T}}, x::AbstractDict) where {T<:XMLRichText}
    return T[Serde.deser(T, x)]
end

function Serde.deser(::Type{ISSI}, ::Type{Vector{T}}, x::Nothing) where {T<:XMLRichText}
    return T[]
end

struct XMLCell <: AbstractXMLNode
    v::Maybe{XMLCellValue}
    t::Maybe{String}
    r::String
    s::Maybe{Int64}
    is::Maybe{ISSI}
    f::Maybe{XMLCellFormula}
end

struct XMLRow <: AbstractXMLNode
    r::Int64
    c::Vector{XMLCell}
end

function Serde.deser(::Type{XMLRow}, ::Type{Vector{T}}, x::AbstractDict) where {T<:XMLCell}
    return T[Serde.deser(T, x)]
end

function Serde.deser(::Type{XMLRow}, ::Type{Vector{T}}, x::Nothing) where {T<:XMLCell}
    return T[]
end

struct XMLRows <: AbstractXMLNode
    row::Vector{XMLRow}
end

function Serde.deser(::Type{XMLRows}, ::Type{Vector{T}}, x::AbstractDict) where {T<:XMLRow}
    return T[Serde.deser(T, x)]
end

function Serde.deser(::Type{XMLRows}, ::Type{Vector{T}}, x::Nothing) where {T<:XMLRow}
    return T[]
end

struct XMLSheet <: AbstractXMLNode
    sheetData::XMLRows
end

Base.isempty(sheet::XMLSheet) = isempty(sheet.sheetData.row)

function nrow(s::XMLSheet)
    return maximum(x -> x.r, s.sheetData.row, init = 0)
end

function ncol(s::XMLSheet)
    return maximum([parse_cell_addr(last(x.c).r).column for x in s.sheetData.row if !isempty(x.c)], init = 0)
end

#__ XMLRelationships

struct XMLRelationship <: AbstractXMLNode
    Id::String
    Type::String
    Target::String
end

struct XMLRelationships <: AbstractXMLNode
    Relationship::Vector{XMLRelationship}
end

function Serde.deser(::Type{XMLRelationships}, ::Type{Vector{T}}, x::AbstractDict) where {T<:XMLRelationship}
    return T[Serde.deser(T, x)]
end

function Serde.deser(::Type{XMLRelationships}, ::Type{Vector{T}}, x::Nothing) where {T<:XMLRelationship}
    return T[]
end

@serde @de_name struct XMLWorkbookRels
    workbook_xml_rels::XMLRelationships | "workbook.xml.rels"
end

#__ XMLSharedStrings

struct XMLSharedStrings <: AbstractXMLNode
    uniqueCount::Int64
    si::Maybe{Vector{ISSI}}
    count::Maybe{Int64}
end

function Serde.deser(::Type{XMLSharedStrings}, ::Type{Vector{T}}, x::AbstractDict) where {T<:ISSI}
    return T[Serde.deser(T, x)]
end

@serde @de_name struct XL <: AbstractXMLNode
    _rels::XMLWorkbookRels                     | "_rels"
    worksheets::Dict{String,XMLSheet}          | "worksheets"
    workbook_xml::XMLWorkbook                  | "workbook.xml"
    sharedStrings_xml::Maybe{XMLSharedStrings} | "sharedStrings.xml"
end

function Base.getindex(x::XL, w::XMLWorksheet)
    relationships = x._rels.workbook_xml_rels.Relationship
    rel_ind = findfirst(x -> x.Id == w.id, relationships)
    sheet_path = basename(relationships[rel_ind].Target)
    return x.worksheets[sheet_path]
end

function Base.getindex(x::XL, cell::XMLCell)
    return if cell.t == "inlineStr"
        string(cell.is)
    elseif !isnothing(cell.f)
        isnothing(cell.v) ? cell.f._ : cell.v._
    elseif isnothing(cell.v) || isnothing(cell.v._)
        nothing
    elseif cell.t == "s"
        string(x.sharedStrings_xml.si[parse(Int64, cell.v._)+1])
    elseif cell.t == "b"
        cell.v._ == "1"
    elseif cell.t == "str" || cell.t == "e"
        cell.v._
    else
        parse(Float64, cell.v._)
    end
end

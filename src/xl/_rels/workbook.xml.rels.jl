module WorkbookRelsXML

export WorkbookRelsFile, read_workbookrels

using ZipFile
using Serde

using ..OpenXL: Maybe, read_zipfile

struct RelationshipItem
    Id::String
    Type::String
    Target::String
end

struct WorkbookRelsFile
    Relationship::Vector{RelationshipItem}
end

function Serde.deser(
    ::Type{WorkbookRelsFile},
    ::Type{Vector{T}},
    x::AbstractDict,
) where {T<:RelationshipItem}
    return T[Serde.deser(T, x)]
end

function Serde.deser(
    ::Type{WorkbookRelsFile},
    ::Type{Vector{T}},
    x::Nothing,
) where {T<:RelationshipItem}
    return T[]
end

function Base.getindex(x::WorkbookRelsFile, Id::AbstractString)
    ind = findfirst(item -> item.Id == Id, x.Relationship)
    return if !isnothing(ind)
        x.Relationship[ind]
    else
        throw(XLError("Relationship with Id = \"$Id\" not found."))
    end
end

function read_workbookrels(x::ZipFile.Reader)
    file = read_zipfile(x, "xl/_rels/workbook.xml.rels")
    return Serde.to_deser(WorkbookRelsFile, parse_xml(file))
end

end
module WorkbookRelsXML

export WorkbookRels

using Serde
import ..ExcelFile

struct RelationshipItem
    Id::String
    Type::String
    Target::String
end

# <Relationships>
#     <Relationship Id="rId1" Type="http://schemas.openxmlformats..." Target="xl/workbook.xml" />
#     ...
# </Relationships>
struct WorkbookRels <: ExcelFile
    Relationship::Vector{RelationshipItem}
end

function Serde.deser(
    ::Type{WorkbookRels},
    ::Type{Vector{T}},
    x::AbstractDict,
) where {T<:RelationshipItem}
    return T[Serde.deser(T, x)]
end

function Serde.deser(
    ::Type{WorkbookRels},
    ::Type{Vector{T}},
    x::Nothing,
) where {T<:RelationshipItem}
    return T[]
end

function Base.getindex(x::WorkbookRels, Id::AbstractString)
    ind = findfirst(item -> item.Id == Id, x.Relationship)
    return if !isnothing(ind)
        x.Relationship[ind]
    else
        throw(XLError("Relationship with Id = \"$Id\" not found."))
    end
end

end

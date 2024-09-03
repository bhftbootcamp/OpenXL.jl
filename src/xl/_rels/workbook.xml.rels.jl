module WorkbookRelsXML

export WorkbookRels

using Serde

using ..OpenXL: ExcelFile

struct RelationshipItem
    Id::String
    Type::String
    Target::String
end

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

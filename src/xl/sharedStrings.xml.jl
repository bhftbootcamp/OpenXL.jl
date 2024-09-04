module sharedStringsXML

export SharedStrings

using Serde

using ..OpenXL: ExcelFile

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

struct SharedItem
    t::Union{Nothing,TextItem}
    r::Vector{RichTextItem}
end

function Base.string(s::SharedItem)
    return isnothing(s.t) ? join(map(string, s.r)) : string(s.t)
end

function Serde.deser(
    ::Type{SharedItem},
    ::Type{Vector{T}},
    x::AbstractDict,
) where {T<:RichTextItem}
    return T[Serde.deser(T, x)]
end

function Serde.deser(
    ::Type{SharedItem},
    ::Type{Vector{T}},
    x::Nothing,
) where {T<:RichTextItem}
    return T[]
end

struct SharedStrings <: ExcelFile
    count::String
    uniqueCount::String
    si::Union{Nothing,Vector{SharedItem}}
end

function Serde.deser(
    ::Type{SharedStrings},
    ::Type{Vector{T}},
    x::AbstractDict,
) where {T<:SharedItem}
    return T[Serde.deser(T, x)]
end

end

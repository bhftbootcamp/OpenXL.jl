module sharedStringsXML

export SharedStringsFile, read_shared_strings

using ZipFile
using Serde

using ..OpenXL: Maybe, read_zipfile

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

struct SharedItem
    t::Maybe{TextItem}
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

struct SharedStringsFile
    count::String
    uniqueCount::String
    si::Maybe{Vector{SharedItem}}
end

function Serde.deser(
    ::Type{SharedStringsFile},
    ::Type{Vector{T}},
    x::AbstractDict,
) where {T<:SharedItem}
    return T[Serde.deser(T, x)]
end

function read_shared_strings(x::ZipFile.Reader)
    file = read_zipfile(x, "xl/sharedStrings.xml")
    return isnothing(file) ? nothing : Serde.to_deser(SharedStringsFile, parse_xml(file))
end

end

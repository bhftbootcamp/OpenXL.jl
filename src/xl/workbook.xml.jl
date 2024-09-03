module WorkbookXML

export WorkbookFile

using Serde

using ..OpenXL: Maybe, ExcelFile

struct SheetItem
    name::String
    sheetId::Int
    id::String
end

struct SheetsItem
    sheet::Vector{SheetItem}
end

function Serde.deser(
    ::Type{SheetsItem},
    ::Type{Vector{T}},
    x::AbstractDict,
) where {T<:SheetItem}
    return T[Serde.deser(T, x)]
end

struct WorkbookFile <: ExcelFile
    sheets::SheetsItem
end

end

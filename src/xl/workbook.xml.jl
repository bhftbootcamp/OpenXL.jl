module WorkbookXML

export WorkbookFile, read_workbook

using ZipFile
using Serde

using ..OpenXL: Maybe, read_zipfile

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

struct WorkbookFile
    sheets::SheetsItem
end

function read_workbook(x::ZipFile.Reader)
    file = read_zipfile(x, "xl/workbook.xml")
    return Serde.to_deser(WorkbookFile, parse_xml(file))
end

end

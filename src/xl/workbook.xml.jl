# <workbook>
#     <sheets>
#         <sheet state="visible" name="Sheet1" sheetId="1" r:id="rId1" />
#         ...
#     </sheets>
# </workbook>

module WorkbookXML

export Workbook

using Serde
import ..ExcelFile

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

struct Workbook <: ExcelFile
    sheets::SheetsItem
end

end

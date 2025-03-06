#__ xl_parser

abstract type ExcelFile end

include("xl/workbook.xml.jl")
using .WorkbookXML

include("xl/styles.xml.jl")
using .StylesXML

include("xl/_rels/workbook.xml.rels.jl")
using .WorkbookRelsXML

include("xl/sharedStrings.xml.jl")
using .sharedStringsXML

include("xl/worksheets/sheetN.xml.jl")
using .WorksheetXML

struct XLDocument
    rels::WorkbookRels
    worksheets::Vector{Worksheet}
    sharedStrings::Union{Nothing,SharedStrings}
    workbook::Workbook
    styles::Styles
end

function XLDocument(x::Vector{UInt8})
    xl_archive = ZipArchive(x)
    try
        rels = xl_archive(WorkbookRels, "xl/_rels/workbook.xml.rels")
        shared_strings = xl_archive(Union{Nothing,SharedStrings}, "xl/sharedStrings.xml")
        workbook = xl_archive(Workbook, "xl/workbook.xml")
        styles = xl_archive(Styles, "xl/styles.xml")

        worksheets = map(workbook.sheets.sheet) do sheet
            xl_archive(Worksheet, joinpath("xl", rels[sheet.id].Target))
        end

        XLDocument(rels, worksheets, shared_strings, workbook, styles)
    finally
        zip_discard(xl_archive)
    end
end

function (x::ZipArchive)(::Type{T}, name::String) where {T<:Union{Nothing,ExcelFile}}
    has_entry = any(file -> file.name == name, collect(x))
    if !has_entry && Nothing <: T
        nothing
    elseif !has_entry
        throw(ArgumentError("File $name not found in the ZIP archive."))
    else
        deser_xml(T, read(x, name))
    end
end

function Base.getindex(x::XLDocument, cell::WorksheetXML.CellItem)
    return if cell.t == "inlineStr"
        # Handle inline strings
        string(cell.is)
    elseif cell.t == "s"
        # Handle shared strings
        string(x.sharedStrings.si[parse(Int64, cell.v._)+1])
    elseif cell.t == "str" || !isnothing(cell.f)
        # Handle formulas
        isnothing(cell.v) ? cell.f._ : cell.v._
    elseif isnothing(cell.v) || isnothing(cell.v._)
        # Handle empty cells
        nothing
    elseif cell.t == "b"
        # Handle boolean values
        cell.v._ == "1"
    elseif cell.t == "e"
        # Handle errors
        cell.v._
    else
        # Handle numbers or dates
        fmt_id = x.styles.cellXfs[cell.s].numFmtId
        fmt_code = x.styles.numFmts[cell.s].formatCode
        number = parse(Float64, cell.v._)
        isdatetime(fmt_id, fmt_code, cell.t) ? xl_num2datetime(number) : number
    end
end

function Base.convert(::Type{XLWorkbook}, xl::XLDocument)
    r = map(zip(xl.workbook.sheets.sheet, xl.worksheets)) do (sheet_info, sheet_data)
        result = Matrix{XLCell}(nothing, nrow(sheet_data), ncol(sheet_data))
        for row in sheet_data.sheetData.row
            for cell in row.c
                column_addr = parse_cell_addr(cell.r).column
                val = xl[cell]
                result[row.r, column_addr] = XLCell(val, xl.styles.cellXfs[cell.s].numFmtId)
            end
        end
        XLSheet(sheet_info.name, sheet_info.sheetId, result)
    end
    return XLWorkbook(r)
end

"""
    xl_parse(x::AbstractString) -> XLWorkbook
    xl_parse(x::Vector{UInt8}) -> XLWorkbook

Parse Excel file into [`XLWorkbook`](@ref) object.

## Examples
```julia-repl
julia> raw_xlsx = xl_sample_employee_xlsx()
77626-element Vector{UInt8}:
 0x50
 0x4b
    ⋮
 0x00

julia> xl_parse(raw_xlsx)
1-element XLWorkbook:
 1001x13 XLSheet("Employee")
```
"""
function xl_parse(x::Vector{UInt8})
    return convert(XLWorkbook, XLDocument(x))
end

function xl_parse(x::AbstractString)
    return xl_parse(collect(UInt8, x))
end

"""
    xl_open(file::AbstractString) -> XLWorkbook

Read the specified XL `file` and parse it into [`XLWorkbook`](@ref).
"""
function xl_open(file::AbstractString)
    return xl_parse(read(file))
end

"""
    xl_open(io::IO) -> XLWorkbook

Read data from the specified `IO` object and parse it into an [`XLWorkbook`](@ref).
"""
function xl_open(io::IO)
    return xl_parse(read(io))
end

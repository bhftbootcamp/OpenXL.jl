#__ xl_parser

using ZipFile: Reader

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

function XLDocument(io::IOBuffer)
    xl_reader = Reader(io)
    try
        rels = xl_reader(WorkbookRels, "xl/_rels/workbook.xml.rels")
        shared_strings = xl_reader(Union{Nothing,SharedStrings}, "xl/sharedStrings.xml")
        workbook = xl_reader(Workbook, "xl/workbook.xml")
        styles = xl_reader(Styles, "xl/styles.xml")

        worksheets = map(workbook.sheets.sheet) do sheet
            xl_reader(Worksheet, joinpath("xl", rels[sheet.id].Target))
        end

        XLDocument(rels, worksheets, shared_strings, workbook, styles)
    finally
        close(xl_reader)
    end
end

function (x::Reader)(::Type{T}, name::String) where {T<:Union{Nothing,ExcelFile}}
    entry = findfirst(file -> file.name == name, x.files)
    if isnothing(entry) && Nothing <: T
        nothing
    elseif isnothing(entry)
        throw(ArgumentError("File $name not found in the ZIP archive."))
    else
        deser_xml(T, read(x.files[entry]))
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
        number = parse(Float64, cell.v._)
        fmt_id = x.styles[cell.s].numFmtId
        if cell.t == "d" || 14 <= fmt_id <= 22 || 45 <= fmt_id <= 47
            xl_num2datetime(number)
        else
            number
        end
    end
end

function Base.convert(::Type{XLWorkbook}, xl::XLDocument)
    r = map(zip(xl.workbook.sheets.sheet, xl.worksheets)) do (sheet_info, sheet_data)
        result = Matrix{XLCell}(nothing, nrow(sheet_data), ncol(sheet_data))
        for row in sheet_data.sheetData.row
            for cell in row.c
                column_addr = parse_cell_addr(cell.r).column
                val = xl[cell]
                result[row.r, column_addr] = XLCell(val, xl.styles[cell.s].numFmtId)
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
    io = IOBuffer(x)
    try
        convert(XLWorkbook, XLDocument(io))
    finally
        close(io)
    end
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

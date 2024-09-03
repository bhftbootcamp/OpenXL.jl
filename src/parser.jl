#__ parser

function read_zipfile(x::ZipFile.Reader, name::AbstractString)
    ind = findfirst(file -> file.name == name, x.files)
    return isnothing(ind) ? nothing : read(x.files[ind])
end

include("xl/workbook.xml.jl")
using .WorkbookXML

include("xl/_rels/workbook.xml.rels.jl")
using .WorkbookRelsXML

include("xl/sharedStrings.xml.jl")
using .sharedStringsXML

include("xl/worksheets/sheetN.xml.jl")
using .WorksheetXML

struct XL
    rels::WorkbookRelsFile
    worksheets::Vector{WorksheetFile}
    sharedStrings::Maybe{SharedStringsFile}
    workbook::WorkbookFile
end

function Base.getindex(x::XL, cell::WorksheetXML.CellItem)
    return if cell.t == "inlineStr"
        string(cell.is)
    elseif !isnothing(cell.f)
        isnothing(cell.v) ? cell.f._ : cell.v._
    elseif isnothing(cell.v) || isnothing(cell.v._)
        nothing
    elseif cell.t == "s"
        string(x.sharedStrings.si[parse(Int64, cell.v._)+1])
    elseif cell.t == "b"
        cell.v._ == "1"
    elseif cell.t == "str" || cell.t == "e"
        cell.v._
    else
        parse(Float64, cell.v._)
    end
end

function unzip_xl(io::IOBuffer)
    _zip = ZipFile.Reader(io)
    xl_workbook = read_workbook(_zip)
    workbook_rels = read_workbookrels(_zip)
    shared_strings = read_shared_strings(_zip)

    sheets = map(xl_workbook.sheets.sheet) do sheet
        sheet_path = joinpath("xl", workbook_rels[sheet.id].Target)
        return read_worksheet(_zip, sheet_path)
    end

    return XL(workbook_rels, sheets, shared_strings, xl_workbook)
end

function Base.convert(::Type{XLWorkbook}, xl::XL)
    r = map(zip(xl.workbook.sheets.sheet, xl.worksheets)) do (sheet_info, sheet_data)
        result = Matrix{Any}(nothing, nrow(sheet_data), ncol(sheet_data))
        for row in sheet_data.sheetData.row
            for cell in row.c
                column_addr = parse_cell_addr(cell.r).column
                result[row.r, column_addr] = xl[cell]
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
48378-element Vector{UInt8}:
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
        convert(XLWorkbook, unzip_xl(io))
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

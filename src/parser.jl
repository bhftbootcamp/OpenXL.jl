# parser

function unzip_xl(data::AbstractVector{UInt8})
    workbook = Dict{String,Any}()
    zip_reader = ZipFile.Reader(IOBuffer(data))

    for file_entry in zip_reader.files
        startswith(file_entry.name, "xl/worksheets/_rels/") && continue
        path_parts = splitpath(file_entry.name)
        xml_name = path_parts[end]
        end_node = workbook
        for part in path_parts[1:end-1]
            end_node = get!(end_node, part, Dict{String,Any}())
        end
        file = read(file_entry)
        if !isempty(file)
            end_node[xml_name] = parse_xml(file)
        end
    end

    return workbook
end

function deser_xl(::Type{XL}, data::AbstractVector{UInt8})
    return Serde.to_deser(XL, unzip_xl(data)["xl"])
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
    return XLWorkbook(deser_xl(XL, x))
end

function xl_parse(x::AbstractString)
    return xl_parse(collect(UInt8, x))
end

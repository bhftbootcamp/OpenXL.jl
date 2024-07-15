# xl_parser

function unzip_xl(buff::IOBuffer)
    outs = Dict{String,Any}()
    zip_reader = ZipFile.Reader(buff)
    for file_entry in zip_reader.files
        startswith(file_entry.name, "xl/worksheets/_rels/") && continue
        path_parts = splitpath(file_entry.name)
        end_node = outs
        for part in path_parts[1:end-1]
            end_node = get!(end_node, part, Dict{String,Any}())
        end
        file = read(file_entry)
        if !isempty(file)
            end_node[path_parts[end]] = parse_xml(file)
        end
    end
    close(zip_reader)
    return outs
end

function deser_xl(::Type{XL}, buff::IOBuffer)
    return Serde.to_deser(XL, unzip_xl(buff)["xl"])
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
    return XLWorkbook(deser_xl(XL, IOBuffer(x)))
end

function xl_parse(x::AbstractString)
    return xl_parse(collect(UInt8, x))
end

function xl_open(x::AbstractString)
    return xl_parse(collect(UInt8, read(x)))
end

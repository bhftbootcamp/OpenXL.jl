# xl_parser

function is_forbidden_path(path::String)
    return startswith(path, "xl/worksheets/_rels/") || startswith(path, "xl/media/")
end

function unzip_xl(buff::IOBuffer)
    outs = Dict{String,Any}()
    zip_reader = ZipFile.Reader(buff)
    for file_entry in zip_reader.files
        is_forbidden_path(file_entry.name) && continue
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
    buff = IOBuffer(x)
    try
        XLWorkbook(deser_xl(XL, buff))
    finally
        close(buff)
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

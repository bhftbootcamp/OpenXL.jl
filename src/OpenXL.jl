module OpenXL

export xl_parse,
    xl_open,
    xl_rowtable,
    xl_columntable,
    xl_print,
    xl_table,
    xl_sheetname,
    xl_sheetnames,
    xl_nrow,
    xl_ncol

export AbstractXLSheet,
    XLError,
    XLSheet,
    SubXLSheet,
    XLWorkbook

using Serde
using ZipFile

struct XLError <: Exception
    message::String
end

function Base.showerror(io::IO, e::XLError)
    print(io, "XLError: ", e.message)
end

const Maybe{T} = Union{Nothing,T}

include("utils.jl")
include("xml_structure.jl")
include("parser.jl")
include("interface.jl")
include("table.jl")
include("print.jl")
include("sample_data.jl")

end

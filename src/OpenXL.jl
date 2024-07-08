module OpenXL

export xl_parse,
    xl_rowtable,
    xl_columntable,
    xl_print_sheet,
    xl_name,
    xl_names,
    xl_nrow,
    xl_ncol

export XLError,
    XLTable,
    XLSheet,
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
include("interface.jl")
include("parser.jl")
include("print.jl")
include("sample_data.jl")

end

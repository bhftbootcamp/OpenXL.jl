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
    XLCell,
    XLSheet,
    SubXLSheet,
    XLWorkbook

using Dates
using Printf
using Serde
using LibZip

struct XLError <: Exception
    message::String
end

function Base.showerror(io::IO, e::XLError)
    print(io, "XLError: ", e.message)
end

include("utils.jl")
include("interface.jl")
include("parser.jl")
include("table.jl")
include("print.jl")

end

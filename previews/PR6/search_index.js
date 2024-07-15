var documenterSearchIndex = {"docs":
[{"location":"#OpenXL.jl","page":"Home","title":"OpenXL.jl","text":"","category":"section"},{"location":"","page":"Home","title":"Home","text":"OpenXL is a lightweight package designed to easily read Excel 2010 xlsx/xlsm/xltx/xltm files.","category":"page"},{"location":"#Installation","page":"Home","title":"Installation","text":"","category":"section"},{"location":"","page":"Home","title":"Home","text":"To install OpenXL, simply use the Julia package manager:","category":"page"},{"location":"","page":"Home","title":"Home","text":"] add OpenXL","category":"page"},{"location":"#Usage","page":"Home","title":"Usage","text":"","category":"section"},{"location":"","page":"Home","title":"Home","text":"Here is how you can use the basic interface for parsing and printing XL tables:","category":"page"},{"location":"","page":"Home","title":"Home","text":"using OpenXL\n\njulia> raw_xlsx = xl_sample_ticker24h_xlsx()\n64075-element Vector{UInt8}:\n 0x50\n 0x4b\n    ⋮\n 0x00\n\njulia> xlsx = xl_parse(raw_xlsx)\n1-element XLWorkbook:\n 2682x19 XLSheet(\"Ticker24h\")\n\njulia> sheet = xlsx[\"Ticker24h\"]\n2682x19 XLSheet(\"Ticker24h\")\n Sheet │ A        B         C        ⋯  Q         R            S                 \n───────┼────────────────────────────────────────────────────────────────────────\n     1 │ symbol   askPrice  askQty   ⋯  count     volume       weightedAvgPrice  \n     2 │ ETHBTC   0.05296   8.1061   ⋯  473424.0  86904.9028   0.05347515        \n     3 │ LTCBTC   0.001072  308.762  ⋯  43966.0   130937.575   0.00110825        \n     4 │ BNBBTC   0.008633  1.036    ⋯  277360.0  99484.88     0.00883183        \n     ⋮ │ ⋮        ⋮         ⋮        ⋯  ⋮         ⋮            ⋮                      \n  2681 │ ZKUSDC   0.1386    3612.7   ⋯  1572.0    1.3895516e6  0.15005404        \n  2682 │ ZROUSDC  2.925     437.83   ⋯  7957.0    356187.29    3.07800556   ","category":"page"},{"location":"","page":"Home","title":"Home","text":"You can slice a table using address indexing and then convert the data to a row representation:","category":"page"},{"location":"","page":"Home","title":"Home","text":"using OpenXL\n\njulia> xlsx = xl_parse(xl_sample_ticker24h_xlsx())\n1-element XLWorkbook:\n 2682x19 XLSheet(\"Ticker24h\")\n\njulia> sheet = xlsx[\"Ticker24h\"][\"A1:E\"]\n2682x5 XLTable\n Sheet │ A        B         C        D         E        \n───────┼────────────────────────────────────────────────\n     1 │ symbol   askPrice  askQty   bidPrice  bidQty \n     2 │ ETHBTC   0.05296   8.1061   0.05295   50.5655\n     3 │ LTCBTC   0.00107   308.762  0.00107   1433.70\n     4 │ BNBBTC   0.00863   1.036    0.00863   8.139  \n     ⋮ │  ⋮        ⋮          ⋮        ⋮         ⋮        \n  2681 │ ZKUSDC   0.1386    3612.7   0.138     11976.9\n  2682 │ ZROUSDC  2.925     437.83   2.922     353.730\n\njulia> xl_rowtable(sheet; header = true)\n2681-element Vector{NamedTuple{(:symbol, :askPrice, ...), Tuple{String, Vararg{Float64, 4}}}}:\n (symbol = \"ETHBTC\", askPrice = 0.0529, askQty = 8.1061, bidPrice = 0.0529, bidQty = 50.565)\n (symbol = \"LTCBTC\", askPrice = 0.0010, askQty = 308.76, bidPrice = 0.0010, bidQty = 1433.7)\n (symbol = \"BNBBTC\", askPrice = 0.0086, askQty = 1.036, bidPrice = 0.00863, bidQty = 8.1390)\n (symbol = \"NEOBTC\", askPrice = 0.0001, askQty = 6.52, bidPrice = 0.000160, bidQty = 318.15)\n ⋮\n (symbol = \"ZKUSDC\", askPrice = 0.1386, askQty = 3612.7, bidPrice = 0.138, bidQty = 11976.9)\n (symbol = \"ZROUSDC\", askPrice = 2.925, askQty = 437.83, bidPrice = 2.922, bidQty = 353.730)","category":"page"},{"location":"","page":"Home","title":"Home","text":"Table slices can be obtained in the same way as with a regular matrix, which can then also be converted to a column representation:","category":"page"},{"location":"","page":"Home","title":"Home","text":"using OpenXL\n\njulia> xlsx = xl_parse(xl_sample_ticker24h_xlsx())\n1-element XLWorkbook:\n 2682x19 XLSheet(\"Ticker24h\")\n\njulia> sheet = xlsx[\"Ticker24h\"][2:500, 7:10]\n499x4 XLTable\n Sheet │ A          B        C          D         \n───────┼─────────────────────────────────────────\n     1 │ 0.05296    8.1061   0.05295    50.5655   \n     2 │ 0.001072   308.762  0.001071   1433.702  \n     3 │ 0.008633   1.036    0.008632   8.139     \n     4 │ 0.0001604  6.52     0.0001601  318.15    \n     ⋮ │ ⋮           ⋮         ⋮          ⋮             \n   498 │ 0.000936   4859.3   0.000932   6302.9    \n   499 │ 8.06e-6    6002.8   8.05e-6    7715.9\n\n\njulia> xl_columntable(sheet; alt_keys = Dict(\"A\" => \"O\", \"B\" => \"H\", \"C\" => \"L\", \"D\" => \"C\"))\n(\n  O = Any[0.05477, 0.001197  …  0.00096, 8.86e-6],\n  H = Any[0.05501, 0.001213  …  0.00098, 8.98e-6],\n  L = Any[0.05216, 0.001029  …  0.00091, 7.85e-6],\n  C = Any[0.05295, 0.001071  …  0.00093, 8.06e-6],\n)","category":"page"},{"location":"","page":"Home","title":"Home","text":"If necessary, you can make a DataFrame object using a row representation of the table.","category":"page"},{"location":"","page":"Home","title":"Home","text":"using OpenXL\nusing DataFrames\n\njulia> xlsx = xl_parse(xl_sample_ticker24h_xlsx())\n1-element XLWorkbook:\n 2682x19 XLSheet(\"Ticker24h\")\n\njulia> DataFrame(xl_rowtable(xlsx[\"Ticker24h\"], header = true))\n2681×19 DataFrame\n  Row │ symbol    askPrice    askQty      bidPrice    bidQty      lastQty    openPrice    highPrice    lowPrice    lastPrice   openTime                       closeTime                      prevClosePrice  priceChange  priceChangePercent  quoteVolume    count     volume         weightedAvgPrice \n      │ String    Float64     Float64     Float64     Float64     Float64    Float64      Float64      Float64     Float64     String                         String                         Float64         Float64      Float64             Float64        Float64   Float64        Float64          \n──────┼────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────\n    1 │ ETHBTC     0.05296        8.1061   0.05295       50.5655      1.15     0.05477      0.05501     0.05216     0.05295    2024-07-04T08:24:48.264999936  2024-07-05T08:24:48.264999936       0.05478      -0.00182               -3.323   4647.25       473424.0  86904.9              0.0534751\n    2 │ LTCBTC     0.001072     308.762    0.001071    1433.7        25.421    0.001197     0.001213    0.001029    0.001071   2024-07-04T08:24:45.672        2024-07-05T08:24:45.672             0.001198     -0.000126             -10.526    145.111       43966.0      1.30938e5        0.00110825\n    3 │ BNBBTC     0.008633       1.036    0.008632       8.139       0.013    0.009138     0.009208    0.008422    0.008632   2024-07-04T08:24:47.456999936  2024-07-05T08:24:47.456999936       0.009139     -0.000506              -5.537    878.633      277360.0  99484.9              0.00883183\n    4 │ NEOBTC     0.0001604      6.52     0.0001601    318.15       33.69     0.000181     0.0001832   0.000154    0.0001603  2024-07-04T08:24:46.680        2024-07-05T08:24:46.680             0.0001812    -2.07e-5              -11.436      6.82048      3096.0  41199.0              0.00016555\n  ⋮   │    ⋮          ⋮           ⋮           ⋮           ⋮           ⋮           ⋮            ⋮           ⋮           ⋮                     ⋮                              ⋮                      ⋮              ⋮               ⋮                 ⋮           ⋮            ⋮               ⋮\n 2678 │ BAKETRY    6.72        3641.0      6.71        3071.0     11350.0      8.79         8.91        6.29        6.71       2024-07-04T08:24:40.859000064  2024-07-05T08:24:40.859000064       8.79         -2.08                 -23.663      3.92568e7   10195.0      5.27997e6        7.43504\n 2679 │ WIFBRL     9.59          15.4      8.81        1595.9         4.5     10.05        11.96        8.72        8.84       2024-07-04T08:24:48.124        2024-07-05T08:24:48.124            10.01         -1.21                 -12.04   87665.9           111.0   9085.4              9.6491\n 2680 │ ZKUSDC     0.1386      3612.7      0.138      11976.9       812.0      0.1628       0.1679      0.1275      0.1387     2024-07-04T08:24:45.328        2024-07-05T08:24:45.328             0.1602       -0.0241               -14.803      2.08508e5    1572.0      1.38955e6        0.150054\n 2681 │ ZROUSDC    2.925        437.83     2.922        353.73       16.08     3.24         3.27        2.735       2.922      2024-07-04T08:24:44.710000128  2024-07-05T08:24:44.710000128       3.221        -0.318                 -9.815      1.09635e6    7957.0      3.56187e5        3.07801\n                                                                                                                                                                                                                                                                                      2673 rows omitted","category":"page"},{"location":"pages/api_reference/#API-Reference","page":"API Reference","title":"API Reference","text":"","category":"section"},{"location":"pages/api_reference/","page":"API Reference","title":"API Reference","text":"xl_parse","category":"page"},{"location":"pages/api_reference/#OpenXL.xl_parse","page":"API Reference","title":"OpenXL.xl_parse","text":"xl_parse(x::AbstractString) -> XLWorkbook\nxl_parse(x::Vector{UInt8}) -> XLWorkbook\n\nParse Excel file into XLWorkbook object.\n\nExamples\n\njulia> raw_xlsx = xl_sample_employee_xlsx()\n48378-element Vector{UInt8}:\n 0x50\n 0x4b\n    ⋮\n 0x00\n\njulia> xl_parse(raw_xlsx)\n1-element XLWorkbook:\n 1001x13 XLSheet(\"Employee\")\n\n\n\n\n\n","category":"function"},{"location":"pages/api_reference/#Types","page":"API Reference","title":"Types","text":"","category":"section"},{"location":"pages/api_reference/","page":"API Reference","title":"API Reference","text":"XLWorkbook\nOpenXL.AbstractXLSheet\nXLSheet\nXLTable","category":"page"},{"location":"pages/api_reference/#OpenXL.XLWorkbook","page":"API Reference","title":"OpenXL.XLWorkbook","text":"XLWorkbook <: AbstractVector{XLSheet}\n\nRepresents an Excel workbook containing XLSheet.\n\nFields\n\nsheets::Vector{XLSheet}: Workbook sheets.\n\nAccessors\n\nxl_sheetnames(x::XLWorkbook): Workbook sheet names.\n\nSee also: xl_parse.\n\n\n\n\n\n","category":"type"},{"location":"pages/api_reference/#OpenXL.AbstractXLSheet","page":"API Reference","title":"OpenXL.AbstractXLSheet","text":"AbstractXLSheet <: AbstractArray{Any,2}\n\nAbstract supertype for XLSheet and XLTable.\n\n\n\n\n\n","category":"type"},{"location":"pages/api_reference/#OpenXL.XLSheet","page":"API Reference","title":"OpenXL.XLSheet","text":"XLSheet <: AbstractXLSheet\n\nSheet of the XLWorkbook. Supports indexing like a regular Matrix, as well as address indexing (e.g. A, A1, AB3 or range D:E, B1:C10, etc.).\n\nThe sheet slice will be converted into a separate XLTable.\n\nFields\n\nname::String: Sheet name.\ntable::XLTable: Table representation.\n\nAccessors\n\nxl_sheetname(x::XLSheet): Sheet name.\nxl_nrow(x::XLTable): Number of rows.\nxl_ncol(x::XLTable): Number of columns.\n\nSee also: xl_rowtable, xl_columntable, xl_print.\n\nExamples\n\njulia> xlsx = xl_parse(xl_sample_stock_xlsx())\n1-element XLWorkbook:\n 41x6 XLSheet(\"Stock\")\n\njulia> sheet = xlsx[\"Stock\"]\n41x6 XLSheet(\"Stock\")\n Sheet │ A      B         C        D         E              F\n───────┼─────────────────────────────────────────────────────────────────────\n     1 │ name   price     h24      volume    mkt            sector\n     2 │ MSFT   430.16    0.0007   11855456  3197000000000  Technology Serv…\n     3 │ AAPL   189.98    -0.0005  36327000  2913000000000  Electronic Tech…\n     4 │ NVDA   1064.69   0.0045   42948000  2662000000000  Electronic Tech…\n     ⋮ │ ⋮      ⋮         ⋮        ⋮         ⋮              ⋮\n    40 │ JNJ    146.97    0.0007   7.173e6   3.5371e11      Health Technolo…\n    41 │ ORCL   122.91    -0.0003  5.984e6   3.3782e11      Technology Serv…\n\n\n\n\n\n","category":"type"},{"location":"pages/api_reference/#OpenXL.XLTable","page":"API Reference","title":"OpenXL.XLTable","text":"XLTable <: AbstractXLSheet\n\nTable representation of the XLSheet object. Supports the same table operations as XLSheet.\n\nFields\n\ndata::Matrix: Table data.\n\nAccessors\n\nxl_nrow(x::XLTable): Number of rows.\nxl_ncol(x::XLTable): Number of columns.\n\nSee also: xl_rowtable, xl_columntable, xl_print.\n\nExamples\n\njulia> xlsx = xl_parse(xl_sample_stock_xlsx())\n1-element XLWorkbook:\n 41x6 XLSheet(\"Stock\")\n\njulia> xlsx[\"Stock\"][\"A1:D25\"]\n25x4 XLTable\n Sheet │ A      B         C        D\n───────┼───────────────────────────────────────\n     1 │ name   price     h24      volume\n     2 │ MSFT   430.16    0.0007   11855456\n     3 │ AAPL   189.98    -0.0005  36327000\n     4 │ NVDA   1064.69   0.0045   42948000\n     ⋮ │ ⋮      ⋮         ⋮        ⋮\n    24 │ NVDA   1064.69   0.0045   4.2948e7\n    25 │ GOOG   176.33    -0.0006  1.1404e7\n\n\n\n\n\n","category":"type"},{"location":"pages/api_reference/#Methods","page":"API Reference","title":"Methods","text":"","category":"section"},{"location":"pages/api_reference/","page":"API Reference","title":"API Reference","text":"xl_rowtable\nxl_columntable\nxl_print","category":"page"},{"location":"pages/api_reference/#OpenXL.xl_rowtable","page":"API Reference","title":"OpenXL.xl_rowtable","text":"xl_rowtable(sheet::AbstractXLSheet; kw...) -> Vector{NamedTuple}\n\nConverts sheet rows to a Vector of NamedTuples.\n\nKeyword arguments\n\nalt_keys::Dict{String,String}: Alternative custom column headers.\nheader::Bool = false: Use first row elements as column headers.\n\nExamples\n\njulia> xlsx = xl_parse(xl_sample_stock_xlsx())\n1-element XLWorkbook:\n 41x6 XLSheet(\"Stock\")\n\njulia> xl_rowtable(xlsx[\"Stock\"][\"A1:C30\"], header = true)\n29-element Vector{NamedTuple{(:name, :price, :h24)}}:\n (name = \"MSFT\", price = \"430.16\", h24 = \"0.0007\")\n (name = \"AAPL\", price = \"189.98\", h24 = \"-0.0005\")\n (name = \"NVDA\", price = \"1064.69\", h24 = \"0.0045\")\n (name = \"GOOG\", price = \"176.33\", h24 = \"-0.0006\")\n ⋮\n (name = \"LLY\", price = 807.43, h24 = -0.0024)\n (name = \"AVGO\", price = 1407.84, h24 = 0.0036)\n\n\n\n\n\n","category":"function"},{"location":"pages/api_reference/#OpenXL.xl_columntable","page":"API Reference","title":"OpenXL.xl_columntable","text":"xl_columntable(sheet::AbstractXLSheet; kw...) -> Vector{NamedTuple}\n\nConverts sheet columns to a Vector of NamedTuples.\n\nKeyword arguments\n\nalt_keys::Dict{String,String}: Alternative custom column headers.\nheader::Bool = false: Use first row elements as column headers.\n\nExamples\n\njulia> xlsx = xl_parse(xl_sample_stock_xlsx())\n1-element XLWorkbook:\n 41x6 XLSheet(\"Stock\")\n\njulia> alt_keys = Dict(\"A\" => \"Name\", \"B\" => \"Price\", \"C\" => \"H24\");\n\njulia> xl_columntable(xlsx[\"Stock\"][2:end, 1:3]; alt_keys)\n(\n    Name = Any[\"MSFT\", \"AAPL\"  …  \"JNJ\", \"ORCL\"]\n    Price = Any[430.16, 189.98  …  146.97, 122.91],\n    H24 = Any[0.0007, -0.0005  …  0.0007, -0.0003],\n)\n\n\n\n\n\n","category":"function"},{"location":"pages/api_reference/#OpenXL.xl_print","page":"API Reference","title":"OpenXL.xl_print","text":"xl_print([io::IO], sheet::AbstractXLSheet; kw...)\n\nPrint a sheet as a table representation.\n\nKeyword arguments\n\ntitle::AbstractString = \"Sheet\": Table title in upper left corner.\nheader::Bool = false: Use first row elements as column headers.\nmax_len::Int = 16: Maximum length of an element in a cell.\ncompact::Bool = true: Omit rows and columns to save space.\n\nExamples\n\njulia> xlsx = xl_parse(xl_sample_employee_xlsx())\n1-element XLWorkbook:\n 1001x13 XLSheet(\"Employee\")\n\njulia> xl_print(xlsx[\"Employee\"]; header = true)\n Sheet │ eeid    full_name        job_title         ⋯  country        city       exit_date\n───────┼───────────────────────────────────────────────────────────────────────────────────\n     2 │ E02387  Emily Davis      Sr. Manger        ⋯  United States  Seattle    44485.0\n     3 │ E04105  Theodore Dinh    Technical Archi…  ⋯  China          Chongqing  nothing\n     4 │ E02572  Luna Sanders     Director          ⋯  United States  Chicago    nothing\n     5 │ E02832  Penelope Jordan  Computer System…  ⋯  United States  Chicago    nothing\n     ⋮ │ ⋮       ⋮                ⋮                 ⋯  ⋮              ⋮          ⋮\n  1000 │ E02521  Lily Nguyen      Sr. Analyst       ⋯  China          Chengdu    nothing\n  1001 │ E03545  Sofia Cheng      Vice President    ⋯  United States  Miami      nothing\n\n\n\n\n\n","category":"function"}]
}

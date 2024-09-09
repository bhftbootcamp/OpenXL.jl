# OpenXL.jl

[![Stable](https://img.shields.io/badge/docs-stable-blue.svg)](https://bhftbootcamp.github.io/OpenXL.jl/stable/)
[![Dev](https://img.shields.io/badge/docs-dev-blue.svg)](https://bhftbootcamp.github.io/OpenXL.jl/dev/)
[![Build Status](https://github.com/bhftbootcamp/OpenXL.jl/actions/workflows/CI.yml/badge.svg?branch=master)](https://github.com/bhftbootcamp/OpenXL.jl/actions/workflows/CI.yml?query=branch%3Amaster)
[![Coverage](https://codecov.io/gh/bhftbootcamp/OpenXL.jl/branch/master/graph/badge.svg)](https://codecov.io/gh/bhftbootcamp/OpenXL.jl)
[![Registry](https://img.shields.io/badge/registry-General-4063d8)](https://github.com/JuliaRegistries/General)

OpenXL is a lightweight package designed to easily read Excel 2010 xlsx/xlsm/xltx/xltm files.

## Installation

To install OpenXL, simply use the Julia package manager:

```julia
] add OpenXL
```

## Usage

Here is how you can use the basic interface for parsing and printing XL tables:

```julia
using OpenXL

julia> raw_xlsx = xl_sample_ticker24h_xlsx()
266033-element Vector{UInt8}:
 0x50
 0x4b
    ⋮
 0x00

julia> xlsx = xl_parse(raw_xlsx)
1-element XLWorkbook:
 2682x19 XLSheet("Ticker24h")

julia> sheet = xlsx["Ticker24h"]
2682x19 XLSheet("Ticker24h")
 Sheet │ A         B          C         ⋯  Q           R             S                 
───────┼──────────────────────────────────────────────────────────────────────────────
     1 │ "symbol"  "askPrice" "askQty"  ⋯  "count"     "volume"      "weightedAvgPri…  
     2 │ "ETHBTC"  0.05       8.11      ⋯  473,424.00  86,904.90     0.05              
     3 │ "LTCBTC"  0.00       308.76    ⋯  43,966.00   130,937.57    0.00              
     4 │ "BNBBTC"  0.01       1.04      ⋯  277,360.00  99,484.88     0.01              
     ⋮ │  ⋮         ⋮          ⋮         ⋯   ⋮           ⋮             ⋮                 
  2681 │ "ZKUSDC"  0.14       3,612.70  ⋯  1,572.00    1,389,551.60  0.15              
  2682 │ "ZROUSDC" 2.92       437.83    ⋯  7,957.00    356,187.29    3.08        
```

You can slice a table using address indexing and then convert the data to a row representation:

```julia
using OpenXL

julia> xlsx = xl_parse(xl_sample_ticker24h_xlsx())
1-element XLWorkbook:
 2682x19 XLSheet("Ticker24h")

julia> sheet = xlsx["Ticker24h"]["A1:E"]
2682x5 SubXLSheet("Ticker24h")
 Sheet │ A          B           C         D           E          
───────┼────────────────────────────────────────────────────────
     1 │ "symbol"   "askPrice"  "askQty"  "bidPrice"  "bidQty"   
     2 │ "ETHBTC"   0.05        8.11      0.05        50.57      
     3 │ "LTCBTC"   0.00        308.76    0.00        1,433.70   
     4 │ "BNBBTC"   0.01        1.04      0.01        8.14       
     ⋮ │  ⋮          ⋮            ⋮         ⋮           ⋮          
  2681 │ "ZKUSDC"   0.14        3,612.70  0.14        11,976.90  
  2682 │ "ZROUSDC"  2.92        437.83    2.92        353.73     

julia> xl_rowtable(sheet; header = true)
2681-element Vector{NamedTuple{(:symbol, :askPrice, :askQty, :bidPrice, :bidQty)}}:
 (symbol = "ETHBTC", askPrice = 0.05296, askQty = 8.1061, ...)
 (symbol = "LTCBTC", askPrice = 0.001072, askQty = 308.762, ...)
 (symbol = "BNBBTC", askPrice = 0.008633, askQty = 1.036, ...)
 (symbol = "NEOBTC", askPrice = 0.0001604, askQty = 6.52, ...)
 ⋮
 (symbol = "ZKUSDC", askPrice = 0.1386, askQty = 3612.7, ...)
 (symbol = "ZROUSDC", askPrice = 2.925, askQty = 437.83, ...)
```

Table slices can be obtained in the same way as with a regular matrix, which can then also be converted to a column representation:

```julia
using OpenXL

julia> xlsx = xl_parse(xl_sample_ticker24h_xlsx())
1-element XLWorkbook:
 2682x19 XLSheet("Ticker24h")

julia> sheet = xlsx["Ticker24h"][2:500, 7:10]
499x4 SubXLSheet("Ticker24h")
 Sheet │ A     B     C     D     
───────┼────────────────────────
     1 │ 0.05  0.06  0.05  0.05  
     2 │ 0.00  0.00  0.00  0.00  
     3 │ 0.01  0.01  0.01  0.01  
     4 │ 0.00  0.00  0.00  0.00  
     ⋮ │  ⋮     ⋮      ⋮     ⋮     
   498 │ 0.00  0.00  0.00  0.00  
   499 │ 0.00  0.00  0.00  0.00  

julia> xl_columntable(sheet; alt_keys = Dict("A" => "O", "B" => "H", "C" => "L", "D" => "C"))
(
  O = [0.05477, 0.001197  …  0.000968, 8.86e-6],
  H = [0.05501, 0.001213  …  0.00098, 8.98e-6],
  L = [0.05216, 0.001029  …  0.00091, 7.85e-6],
  C = [0.05295, 0.001071  …  0.000932, 8.06e-6],
)
```

If necessary, you can make a `DataFrame` object using a row representation of the table.

```julia
using OpenXL
using DataFrames

julia> xlsx = xl_parse(xl_sample_ticker24h_xlsx())
1-element XLWorkbook:
 2682x19 XLSheet("Ticker24h")

julia> DataFrame(eachrow(xlsx["Ticker24h"], header = true))
2681×19 DataFrame
  Row │ symbol    askPrice    askQty      bidPrice    bidQty      lastQty    openPrice  ⋯
      │ String    Float64     Float64     Float64     Float64     Float64    Float64    ⋯
──────┼──────────────────────────────────────────────────────────────────────────────────
    1 │ ETHBTC     0.05296        8.1061   0.05295       50.5655      1.15     0.05477  ⋯
    2 │ LTCBTC     0.001072     308.762    0.001071    1433.7        25.421    0.001197
    3 │ BNBBTC     0.008633       1.036    0.008632       8.139       0.013    0.009138
    4 │ NEOBTC     0.0001604      6.52     0.0001601    318.15       33.69     0.000181
    5 │ QTUMETH    0.000696     100.3      0.000692     792.4       157.3      0.000738 ⋯
  ⋮   │    ⋮          ⋮           ⋮           ⋮           ⋮           ⋮           ⋮     ⋱
 2678 │ BAKETRY    6.72        3641.0      6.71        3071.0     11350.0      8.79
 2679 │ WIFBRL     9.59          15.4      8.81        1595.9         4.5     10.05
 2680 │ ZKUSDC     0.1386      3612.7      0.138      11976.9       812.0      0.1628
 2681 │ ZROUSDC    2.925        437.83     2.922        353.73       16.08     3.24     ⋯
                                                         13 columns and 2672 rows omitted
```

## Contributing

Contributions to OpenXL are welcome! If you encounter a bug, have a feature request, or would like to contribute code, please open an issue or a pull request on GitHub.

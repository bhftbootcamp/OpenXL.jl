# OpenXL.jl

OpenXL is a lightweight package designed to easily read Excel 2010 xlsx/xlsm/xltx/xltm files.

## Installation

To install OpenXL, simply use the Julia package manager:

```julia
] add OpenXL
```

## Usage

Here is how you can use the basic interface for parsing and printing XL tables:

```julia-repl
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
 Sheet │ A        B         C        ⋯  Q         R            S                 
───────┼────────────────────────────────────────────────────────────────────────
     1 │ symbol   askPrice  askQty   ⋯  count     volume       weightedAvgPrice  
     2 │ ETHBTC   0.05296   8.1061   ⋯  473424.0  86904.9028   0.05347515        
     3 │ LTCBTC   0.001072  308.762  ⋯  43966.0   130937.575   0.00110825        
     4 │ BNBBTC   0.008633  1.036    ⋯  277360.0  99484.88     0.00883183        
     ⋮ │ ⋮        ⋮         ⋮        ⋯  ⋮         ⋮            ⋮                 
  2681 │ ZKUSDC   0.1386    3612.7   ⋯  1572.0    1.3895516e6  0.15005404        
  2682 │ ZROUSDC  2.925     437.83   ⋯  7957.0    356187.29    3.07800556
```

You can slice a table using address indexing and then convert the data to a row representation:

```julia-repl
using OpenXL

julia> xlsx = xl_parse(xl_sample_ticker24h_xlsx())
1-element XLWorkbook:
 2682x19 XLSheet("Ticker24h")

julia> sheet = xlsx["Ticker24h"]["A1:E"]
2682x5 SubXLSheet("Ticker24h")
 Sheet │ A        B         C        D         E         
───────┼────────────────────────────────────────────────
     1 │ symbol   askPrice  askQty   bidPrice  bidQty    
     2 │ ETHBTC   0.05296   8.1061   0.05295   50.5655   
     3 │ LTCBTC   0.001072  308.762  0.001071  1433.702  
     4 │ BNBBTC   0.008633  1.036    0.008632  8.139     
     ⋮ │ ⋮        ⋮         ⋮        ⋮         ⋮         
  2681 │ ZKUSDC   0.1386    3612.7   0.138     11976.9   
  2682 │ ZROUSDC  2.925     437.83   2.922     353.73  

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

```julia-repl
using OpenXL

julia> xlsx = xl_parse(xl_sample_ticker24h_xlsx())
1-element XLWorkbook:
 2682x19 XLSheet("Ticker24h")

julia> sheet = xlsx["Ticker24h"][2:500, 7:10]
499x4 SubXLSheet("Ticker24h")
 Sheet │ A          B          C          D          
───────┼────────────────────────────────────────────
     1 │ 0.05477    0.05501    0.05216    0.05295    
     2 │ 0.001197   0.001213   0.001029   0.001071   
     3 │ 0.009138   0.009208   0.008422   0.008632   
     4 │ 0.000181   0.0001832  0.000154   0.0001603  
     ⋮ │ ⋮          ⋮          ⋮          ⋮           
   498 │ 0.000968   0.00098    0.00091    0.000932   
   499 │ 8.86e-6    8.98e-6    7.85e-6    8.06e-6   

julia> xl_columntable(sheet; alt_keys = Dict("A" => "O", "B" => "H", "C" => "L", "D" => "C"))
(
  O = [0.05477, 0.001197  …  0.000968, 8.86e-6],
  H = [0.05501, 0.001213  …  0.00098, 8.98e-6],
  L = [0.05216, 0.001029  …  0.00091, 7.85e-6],
  C = [0.05295, 0.001071  …  0.000932, 8.06e-6],
)
```

If necessary, you can make a `DataFrame` object using a row representation of the table.

```julia-repl
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

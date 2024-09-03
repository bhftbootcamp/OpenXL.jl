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
64075-element Vector{UInt8}:
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
     3 │ LTCBTC   0.00107   308.762  0.00107   1433.70
     4 │ BNBBTC   0.00863   1.036    0.00863   8.139  
     ⋮ │  ⋮        ⋮          ⋮        ⋮         ⋮        
  2681 │ ZKUSDC   0.1386    3612.7   0.138     11976.9
  2682 │ ZROUSDC  2.925     437.83   2.922     353.730

julia> xl_rowtable(sheet; header = true)
2681-element Vector{NamedTuple{(:symbol, :askPrice, ...), Tuple{String, Vararg{Float64, 4}}}}:
 (symbol = "ETHBTC", askPrice = 0.0529, askQty = 8.1061, bidPrice = 0.0529, bidQty = 50.565)
 (symbol = "LTCBTC", askPrice = 0.0010, askQty = 308.76, bidPrice = 0.0010, bidQty = 1433.7)
 (symbol = "BNBBTC", askPrice = 0.0086, askQty = 1.036, bidPrice = 0.00863, bidQty = 8.1390)
 (symbol = "NEOBTC", askPrice = 0.0001, askQty = 6.52, bidPrice = 0.000160, bidQty = 318.15)
 ⋮
 (symbol = "ZKUSDC", askPrice = 0.1386, askQty = 3612.7, bidPrice = 0.138, bidQty = 11976.9)
 (symbol = "ZROUSDC", askPrice = 2.925, askQty = 437.83, bidPrice = 2.922, bidQty = 353.730)
```

Table slices can be obtained in the same way as with a regular matrix, which can then also be converted to a column representation:

```julia-repl
using OpenXL

julia> xlsx = xl_parse(xl_sample_ticker24h_xlsx())
1-element XLWorkbook:
 2682x19 XLSheet("Ticker24h")

julia> sheet = xlsx["Ticker24h"][2:500, 7:10]
499x4 SubXLSheet("Ticker24h")
 Sheet │ A         B          C         D          
───────┼──────────────────────────────────────────
     1 │ 0.05477   0.05501    0.05216   0.05295    
     2 │ 0.001197  0.001213   0.001029  0.001071   
     3 │ 0.009138  0.009208   0.008422  0.008632   
     4 │ 0.000181  0.0001832  0.000154  0.0001603  
     ⋮ │  ⋮         ⋮          ⋮          ⋮          
   498 │ 0.000968  0.00098    0.00091   0.000932   
   499 │ 8.86e-6   8.98e-6    7.85e-6   8.06e-6 


julia> xl_columntable(sheet; alt_keys = Dict("A" => "O", "B" => "H", "C" => "L", "D" => "C"))
(
  O = Any[0.05477, 0.001197  …  0.00096, 8.86e-6],
  H = Any[0.05501, 0.001213  …  0.00098, 8.98e-6],
  L = Any[0.05216, 0.001029  …  0.00091, 7.85e-6],
  C = Any[0.05295, 0.001071  …  0.00093, 8.06e-6],
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
  Row │ symbol    askPrice    askQty      bidPrice    bidQty      lastQty    openPrice    highPrice    lowPrice    lastPrice   openTime                       closeTime                      prevClosePrice  priceChange  priceChangePercent  quoteVolume    count     volume         weightedAvgPrice 
      │ String    Float64     Float64     Float64     Float64     Float64    Float64      Float64      Float64     Float64     String                         String                         Float64         Float64      Float64             Float64        Float64   Float64        Float64          
──────┼────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
    1 │ ETHBTC     0.05296        8.1061   0.05295       50.5655      1.15     0.05477      0.05501     0.05216     0.05295    2024-07-04T08:24:48.264999936  2024-07-05T08:24:48.264999936       0.05478      -0.00182               -3.323   4647.25       473424.0  86904.9              0.0534751
    2 │ LTCBTC     0.001072     308.762    0.001071    1433.7        25.421    0.001197     0.001213    0.001029    0.001071   2024-07-04T08:24:45.672        2024-07-05T08:24:45.672             0.001198     -0.000126             -10.526    145.111       43966.0      1.30938e5        0.00110825
    3 │ BNBBTC     0.008633       1.036    0.008632       8.139       0.013    0.009138     0.009208    0.008422    0.008632   2024-07-04T08:24:47.456999936  2024-07-05T08:24:47.456999936       0.009139     -0.000506              -5.537    878.633      277360.0  99484.9              0.00883183
    4 │ NEOBTC     0.0001604      6.52     0.0001601    318.15       33.69     0.000181     0.0001832   0.000154    0.0001603  2024-07-04T08:24:46.680        2024-07-05T08:24:46.680             0.0001812    -2.07e-5              -11.436      6.82048      3096.0  41199.0              0.00016555
  ⋮   │    ⋮          ⋮           ⋮           ⋮           ⋮           ⋮           ⋮            ⋮           ⋮           ⋮                     ⋮                              ⋮                      ⋮              ⋮               ⋮                 ⋮           ⋮            ⋮               ⋮
 2678 │ BAKETRY    6.72        3641.0      6.71        3071.0     11350.0      8.79         8.91        6.29        6.71       2024-07-04T08:24:40.859000064  2024-07-05T08:24:40.859000064       8.79         -2.08                 -23.663      3.92568e7   10195.0      5.27997e6        7.43504
 2679 │ WIFBRL     9.59          15.4      8.81        1595.9         4.5     10.05        11.96        8.72        8.84       2024-07-04T08:24:48.124        2024-07-05T08:24:48.124            10.01         -1.21                 -12.04   87665.9           111.0   9085.4              9.6491
 2680 │ ZKUSDC     0.1386      3612.7      0.138      11976.9       812.0      0.1628       0.1679      0.1275      0.1387     2024-07-04T08:24:45.328        2024-07-05T08:24:45.328             0.1602       -0.0241               -14.803      2.08508e5    1572.0      1.38955e6        0.150054
 2681 │ ZROUSDC    2.925        437.83     2.922        353.73       16.08     3.24         3.27        2.735       2.922      2024-07-04T08:24:44.710000128  2024-07-05T08:24:44.710000128       3.221        -0.318                 -9.815      1.09635e6    7957.0      3.56187e5        3.07801
                                                                                                                                                                                                                                                                                      2673 rows omitted
```

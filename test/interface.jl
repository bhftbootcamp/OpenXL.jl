# interface

xlsx = xl_parse(read("xl_data/general_tables.xlsx"))

@testset "Workbook interface" begin
    @testset "Case №1: Accessors" begin
        @test length(xlsx) == 13
        @test xl_sheetnames(xlsx) == [
            "general",
            "table3",
            "table4",
            "table",
            "table2",
            "empty",
            "table5",
            "table6",
            "table7",
            "lookup",
            "header_error",
            "named_ranges_2",
            "named_ranges",
        ]
    end

    @testset "Case №2: Indexing" begin
        @test xlsx[1]  == xlsx["general"]
        @test xlsx[2]  == xlsx["table3"]
        @test xlsx[3]  == xlsx["table4"]
        @test xlsx[4]  == xlsx["table"]
        @test xlsx[5]  == xlsx["table2"]
        @test xlsx[6]  == xlsx["empty"]
        @test xlsx[7]  == xlsx["table5"]
        @test xlsx[8]  == xlsx["table6"]
        @test xlsx[9]  == xlsx["table7"]
        @test xlsx[10] == xlsx["lookup"]
        @test xlsx[11] == xlsx["header_error"]
        @test xlsx[12] == xlsx["named_ranges_2"]
        @test xlsx[13] == xlsx["named_ranges"]
    end
end

@testset "Sheet interface" begin
    sheet = xlsx["general"]

    @testset "Case №1: Accessors" begin
        @test xl_nrow(sheet) == 10
        @test xl_ncol(sheet) == 2
        @test xl_sheetname(sheet) == "general"
        @test xl_table(sheet) == [
            "text"             "regular_text"
            "integer"       102.0
            "float"         102.2
            "date"        30422.0
            "hour"            1.8227199074074074
            "datetime"    43206.805451388886
            "float cient"  -220.0
            "integer neg" -2000.0
            "bigint"          1.0e14
            "bigint neg"       "-100000000000000"
        ]
    end

    @testset "Case №2: Indexing" begin
        @test sheet[1] == "text"
        @test sheet[2] == "integer"

        @test sheet[1, 1] == "text"
        @test sheet[1, 2] == "regular_text"
        @test sheet[1, :] == ["text", "regular_text"]

        @test sheet["A1"] == "text"
        @test sheet["B1"] == "regular_text"

        @test sheet["A"] == [
            "text",
            "integer",
            "float",
            "date",
            "hour",
            "datetime",
            "float cient",
            "integer neg",
            "bigint",
            "bigint neg",
        ]

        @test xl_table(sheet["A8:B"]) ==
              ["integer neg" -2000.0; "bigint" 1.0e14; "bigint neg" "-100000000000000"]

        @test sheet["A:B8"] ==
            [ "integer neg" -2000.0; "bigint" 1.0e14; "bigint neg" "-100000000000000"]

        @test sheet["A1:B2"] == ["text" "regular_text"; "integer" 102.0]
        @test sheet["A:B"] == sheet[:, :]

        @test_throws XLError sheet["a1"]
        @test_throws XLError sheet["A:A:B"]
        @test_throws XLError sheet["A::B"]
        @test_throws XLError sheet["A1:1A"]
        @test_throws XLError sheet["a:B"]
        @test_throws XLError sheet["A:b"]
        @test_throws XLError sheet[":A"]
        @test_throws XLError sheet["A:"]
        @test_throws XLError sheet[":A:"]
        @test_throws XLError sheet["1:A:"]
        @test_throws XLError sheet[":A:1"]
        @test_throws XLError sheet[":A:B"]
        @test_throws XLError sheet["A:B:"]
        @test_throws XLError sheet["1A:B2"]
        @test_throws XLError sheet["A1B2"]
        @test_throws XLError sheet["A101:B02"]
    end

    @testset "Case №3: Utils" begin
        @test xl_rowtable(sheet) == [
            (A = "text", B = "regular_text"),
            (A = "integer", B = 102.0),
            (A = "float", B = 102.2),
            (A = "date", B = 30422.0),
            (A = "hour", B = 1.8227199074074074),
            (A = "datetime", B = 43206.805451388886),
            (A = "float cient", B = -220.0),
            (A = "integer neg", B = -2000.0),
            (A = "bigint", B = 1.0e14),
            (A = "bigint neg", B = "-100000000000000"),
        ]

        @test xl_rowtable(sheet; header = true) == [
            (text = "integer", regular_text = 102.0),
            (text = "float", regular_text = 102.2),
            (text = "date", regular_text = 30422.0),
            (text = "hour", regular_text = 1.8227199074074074),
            (text = "datetime", regular_text = 43206.805451388886),
            (text = "float cient", regular_text = -220.0),
            (text = "integer neg", regular_text = -2000.0),
            (text = "bigint", regular_text = 1.0e14),
            (text = "bigint neg", regular_text = "-100000000000000"),
        ]

        @test xl_rowtable(sheet; header = true, alt_keys = Dict("text" => "alt_text")) == [
            (alt_text = "integer", regular_text = 102.0),
            (alt_text = "float", regular_text = 102.2),
            (alt_text = "date", regular_text = 30422.0),
            (alt_text = "hour", regular_text = 1.8227199074074074),
            (alt_text = "datetime", regular_text = 43206.805451388886),
            (alt_text = "float cient", regular_text = -220.0),
            (alt_text = "integer neg", regular_text = -2000.0),
            (alt_text = "bigint", regular_text = 1.0e14),
            (alt_text = "bigint neg", regular_text = "-100000000000000"),
        ]

        @test xl_rowtable(sheet; alt_keys = ["alt_text", "regular_text"]) == [
            (alt_text = "text", regular_text = "regular_text"),
            (alt_text = "integer", regular_text = 102.0),
            (alt_text = "float", regular_text = 102.2),
            (alt_text = "date", regular_text = 30422.0),
            (alt_text = "hour", regular_text = 1.8227199074074074),
            (alt_text = "datetime", regular_text = 43206.805451388886),
            (alt_text = "float cient", regular_text = -220.0),
            (alt_text = "integer neg", regular_text = -2000.0),
            (alt_text = "bigint", regular_text = 1.0e14),
            (alt_text = "bigint neg", regular_text = "-100000000000000"),
        ]

        @test xl_rowtable(sheet; header = true, alt_keys = ["alt_text", "regular_text"]) == [
            (alt_text = "integer", regular_text = 102.0),
            (alt_text = "float", regular_text = 102.2),
            (alt_text = "date", regular_text = 30422.0),
            (alt_text = "hour", regular_text = 1.8227199074074074),
            (alt_text = "datetime", regular_text = 43206.805451388886),
            (alt_text = "float cient", regular_text = -220.0),
            (alt_text = "integer neg", regular_text = -2000.0),
            (alt_text = "bigint", regular_text = 1.0e14),
            (alt_text = "bigint neg", regular_text = "-100000000000000"),
        ]

        @test xl_columntable(sheet) == (
            A = Any[
                "text",
                "integer",
                "float",
                "date",
                "hour",
                "datetime",
                "float cient",
                "integer neg",
                "bigint",
                "bigint neg",
            ],
            B = Any[
                "regular_text",
                102.0,
                102.2,
                30422.0,
                1.8227199074074074,
                43206.805451388886,
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )

        @test xl_columntable(sheet; header = true) == (
            text = Any[
                "integer",
                "float",
                "date",
                "hour",
                "datetime",
                "float cient",
                "integer neg",
                "bigint",
                "bigint neg",
            ],
            regular_text = Any[
                102.0,
                102.2,
                30422.0,
                1.8227199074074074,
                43206.805451388886,
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )

        @test xl_columntable(sheet; header = true, alt_keys = Dict("text" => "alt_text")) == (
            alt_text = Any[
                "integer",
                "float",
                "date",
                "hour",
                "datetime",
                "float cient",
                "integer neg",
                "bigint",
                "bigint neg",
            ],
            regular_text = Any[
                102.0,
                102.2,
                30422.0,
                1.8227199074074074,
                43206.805451388886,
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )

        @test xl_columntable(sheet; alt_keys = ["alt_text", "regular_text"]) == (
            alt_text = Any[
                "text",
                "integer",
                "float",
                "date",
                "hour",
                "datetime",
                "float cient",
                "integer neg",
                "bigint",
                "bigint neg",
            ],
            regular_text = Any[
                "regular_text",
                102.0,
                102.2,
                30422.0,
                1.8227199074074074,
                43206.805451388886,
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )

        @test xl_columntable(sheet; header = true, alt_keys = ["alt_text", "regular_text"]) == (
            alt_text = Any[
                "integer",
                "float",
                "date",
                "hour",
                "datetime",
                "float cient",
                "integer neg",
                "bigint",
                "bigint neg",
            ],
            regular_text = Any[
                102.0,
                102.2,
                30422.0,
                1.8227199074074074,
                43206.805451388886,
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )
    end
end

@testset "Table interface" begin
    slice = xlsx["general"][:, :]

    @testset "Case №1: Accessors" begin
        @test xl_nrow(slice) == 10
        @test xl_ncol(slice) == 2
    end

    @testset "Case №2: Indexing" begin
        @test slice[1] == "text"
        @test slice[2] == "integer"

        @test slice[1, 1] == "text"
        @test slice[1, 2] == "regular_text"
        @test slice[1, :] == ["text", "regular_text"]

        @test slice["A1"] == "text"
        @test slice["B1"] == "regular_text"

        @test slice["A"] == [
            "text",
            "integer",
            "float",
            "date",
            "hour",
            "datetime",
            "float cient",
            "integer neg",
            "bigint",
            "bigint neg",
        ]

        @test slice["A8:B"] ==
            [ "integer neg" -2000.0; "bigint" 1.0e14; "bigint neg" "-100000000000000"]

        @test slice["A:B8"] ==
            [ "integer neg" -2000.0; "bigint" 1.0e14; "bigint neg" "-100000000000000"]

        @test slice["A1:B2"] == ["text" "regular_text"; "integer" 102.0]
        @test slice["A:B"] == slice[:, :]

        @test_throws XLError slice["a1"]
        @test_throws XLError slice["A:A:B"]
        @test_throws XLError slice["A::B"]
        @test_throws XLError slice["A1:1A"]
        @test_throws XLError slice["a:B"]
        @test_throws XLError slice["A:b"]
        @test_throws XLError slice[":A"]
        @test_throws XLError slice["A:"]
        @test_throws XLError slice[":A:"]
        @test_throws XLError slice["1:A:"]
        @test_throws XLError slice[":A:1"]
        @test_throws XLError slice[":A:B"]
        @test_throws XLError slice["A:B:"]
        @test_throws XLError slice["1A:B2"]
        @test_throws XLError slice["A1B2"]
        @test_throws XLError slice["A101:B02"]
    end

    @testset "Case №3: Utils" begin
        @test xl_rowtable(slice) == [
            (A = "text", B = "regular_text"),
            (A = "integer", B = 102.0),
            (A = "float", B = 102.2),
            (A = "date", B = 30422.0),
            (A = "hour", B = 1.8227199074074074),
            (A = "datetime", B = 43206.805451388886),
            (A = "float cient", B = -220.0),
            (A = "integer neg", B = -2000.0),
            (A = "bigint", B = 1.0e14),
            (A = "bigint neg", B = "-100000000000000"),
        ]

        @test xl_rowtable(slice; header = true) == [
            (text = "integer", regular_text = 102.0),
            (text = "float", regular_text = 102.2),
            (text = "date", regular_text = 30422.0),
            (text = "hour", regular_text = 1.8227199074074074),
            (text = "datetime", regular_text = 43206.805451388886),
            (text = "float cient", regular_text = -220.0),
            (text = "integer neg", regular_text = -2000.0),
            (text = "bigint", regular_text = 1.0e14),
            (text = "bigint neg", regular_text = "-100000000000000"),
        ]

        @test xl_rowtable(slice; header = true, alt_keys = Dict("text" => "alt_text")) == [
            (alt_text = "integer", regular_text = 102.0),
            (alt_text = "float", regular_text = 102.2),
            (alt_text = "date", regular_text = 30422.0),
            (alt_text = "hour", regular_text = 1.8227199074074074),
            (alt_text = "datetime", regular_text = 43206.805451388886),
            (alt_text = "float cient", regular_text = -220.0),
            (alt_text = "integer neg", regular_text = -2000.0),
            (alt_text = "bigint", regular_text = 1.0e14),
            (alt_text = "bigint neg", regular_text = "-100000000000000"),
        ]

        @test xl_rowtable(slice; alt_keys = ["alt_text", "regular_text"]) == [
            (alt_text = "text", regular_text = "regular_text"),
            (alt_text = "integer", regular_text = 102.0),
            (alt_text = "float", regular_text = 102.2),
            (alt_text = "date", regular_text = 30422.0),
            (alt_text = "hour", regular_text = 1.8227199074074074),
            (alt_text = "datetime", regular_text = 43206.805451388886),
            (alt_text = "float cient", regular_text = -220.0),
            (alt_text = "integer neg", regular_text = -2000.0),
            (alt_text = "bigint", regular_text = 1.0e14),
            (alt_text = "bigint neg", regular_text = "-100000000000000"),
        ]

        @test xl_rowtable(slice; header = true, alt_keys = ["alt_text", "regular_text"]) == [
            (alt_text = "integer", regular_text = 102.0),
            (alt_text = "float", regular_text = 102.2),
            (alt_text = "date", regular_text = 30422.0),
            (alt_text = "hour", regular_text = 1.8227199074074074),
            (alt_text = "datetime", regular_text = 43206.805451388886),
            (alt_text = "float cient", regular_text = -220.0),
            (alt_text = "integer neg", regular_text = -2000.0),
            (alt_text = "bigint", regular_text = 1.0e14),
            (alt_text = "bigint neg", regular_text = "-100000000000000"),
        ]

        @test xl_columntable(slice) == (
            A = Any[
                "text",
                "integer",
                "float",
                "date",
                "hour",
                "datetime",
                "float cient",
                "integer neg",
                "bigint",
                "bigint neg",
            ],
            B = Any[
                "regular_text",
                102.0,
                102.2,
                30422.0,
                1.8227199074074074,
                43206.805451388886,
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )

        @test xl_columntable(slice; header = true) == (
            text = Any[
                "integer",
                "float",
                "date",
                "hour",
                "datetime",
                "float cient",
                "integer neg",
                "bigint",
                "bigint neg",
            ],
            regular_text = Any[
                102.0,
                102.2,
                30422.0,
                1.8227199074074074,
                43206.805451388886,
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )

        @test xl_columntable(slice; header = true, alt_keys = Dict("text" => "alt_text")) == (
            alt_text = Any[
                "integer",
                "float",
                "date",
                "hour",
                "datetime",
                "float cient",
                "integer neg",
                "bigint",
                "bigint neg",
            ],
            regular_text = Any[
                102.0,
                102.2,
                30422.0,
                1.8227199074074074,
                43206.805451388886,
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )

        @test xl_columntable(slice; alt_keys = ["alt_text", "regular_text"]) == (
            alt_text = Any[
                "text",
                "integer",
                "float",
                "date",
                "hour",
                "datetime",
                "float cient",
                "integer neg",
                "bigint",
                "bigint neg",
            ],
            regular_text = Any[
                "regular_text",
                102.0,
                102.2,
                30422.0,
                1.8227199074074074,
                43206.805451388886,
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )

        @test xl_columntable(slice; header = true, alt_keys = ["alt_text", "regular_text"]) == (
            alt_text = Any[
                "integer",
                "float",
                "date",
                "hour",
                "datetime",
                "float cient",
                "integer neg",
                "bigint",
                "bigint neg",
            ],
            regular_text = Any[
                102.0,
                102.2,
                30422.0,
                1.8227199074074074,
                43206.805451388886,
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )
    end
end

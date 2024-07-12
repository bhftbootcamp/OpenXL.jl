# interface

xlsx = xl_parse(read("xl_data/general_tables.xlsx"))

@testset "Workbook interface" begin
    @testset "Case №1: Accessors" begin
        @test length(xlsx) == 13
        @test xl_names(xlsx) == [
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
        @test xl_name(sheet) == "general"
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

        @test sheet["A8:B"] == XLTable(
            [ "integer neg" -2000.0; "bigint" 1.0e14; "bigint neg" "-100000000000000"],
        )

        @test sheet["A:B8"] == XLTable(
            [ "integer neg" -2000.0; "bigint" 1.0e14; "bigint neg" "-100000000000000"],
        )

        @test sheet["A1:B2"] == XLTable(["text" "regular_text"; "integer" 102.0])
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
    table = xlsx["general"][:, :]

    @testset "Case №1: Accessors" begin
        @test xl_nrow(table) == 10
        @test xl_ncol(table) == 2
    end

    @testset "Case №2: Indexing" begin
        @test table[1] == "text"
        @test table[2] == "integer"

        @test table[1, 1] == "text"
        @test table[1, 2] == "regular_text"
        @test table[1, :] == ["text", "regular_text"]

        @test table["A1"] == "text"
        @test table["B1"] == "regular_text"

        @test table["A"] == [
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

        @test table["A8:B"] == XLTable(
            [ "integer neg" -2000.0; "bigint" 1.0e14; "bigint neg" "-100000000000000"],
        )

        @test table["A:B8"] == XLTable(
            [ "integer neg" -2000.0; "bigint" 1.0e14; "bigint neg" "-100000000000000"],
        )

        @test table["A1:B2"] == XLTable(["text" "regular_text"; "integer" 102.0])
        @test table["A:B"] == table[:, :]

        @test_throws XLError table["a1"]
        @test_throws XLError table["A:A:B"]
        @test_throws XLError table["A::B"]
        @test_throws XLError table["A1:1A"]
        @test_throws XLError table["a:B"]
        @test_throws XLError table["A:b"]
        @test_throws XLError table[":A"]
        @test_throws XLError table["A:"]
        @test_throws XLError table[":A:"]
        @test_throws XLError table["1:A:"]
        @test_throws XLError table[":A:1"]
        @test_throws XLError table[":A:B"]
        @test_throws XLError table["A:B:"]
        @test_throws XLError table["1A:B2"]
        @test_throws XLError table["A1B2"]
        @test_throws XLError table["A101:B02"]
    end

    @testset "Case №3: Utils" begin
        @test xl_rowtable(table) == [
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

        @test xl_rowtable(table; header = true) == [
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

        @test xl_rowtable(table; header = true, alt_keys = Dict("text" => "alt_text")) == [
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

        @test xl_rowtable(table; alt_keys = ["alt_text", "regular_text"]) == [
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

        @test xl_rowtable(table; header = true, alt_keys = ["alt_text", "regular_text"]) == [
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

        @test xl_columntable(table) == (
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

        @test xl_columntable(table; header = true) == (
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

        @test xl_columntable(table; header = true, alt_keys = Dict("text" => "alt_text")) == (
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

        @test xl_columntable(table; alt_keys = ["alt_text", "regular_text"]) == (
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

        @test xl_columntable(table; header = true, alt_keys = ["alt_text", "regular_text"]) == (
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

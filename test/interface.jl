# interface

xlsx = xl_parse(read("xl_data/general_tables.xlsx"))

@testset "Workbook" begin
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

@testset "Cell interface" begin
    numbers = xl_parse(read("xl_data/general_formats.xlsx"))[1]

    @testset "Case №1: Accessors" begin
        # @test string(numbers[1, 3])  == "12345"
        # @test string(numbers[2, 3])  == "12345.12"
        # @test string(numbers[3, 3])  == "12,345"
        # @test string(numbers[4, 3])  == "12,345.12"
        # @test string(numbers[5, 3])  == "12345%"
        # @test string(numbers[6, 3])  == "12345.12%"
        # @test string(numbers[7, 3])  == "1.23E+04"
        # @test string(numbers[8, 3])  == "12345 1/8"
        # @test string(numbers[9, 3])  == "12345 1/8"
        # @test string(numbers[10, 3]) == "10-18-33"
        # @test string(numbers[11, 3]) == "18-Oct-33"
        # @test string(numbers[12, 3]) == "18-Oct"
        # @test string(numbers[13, 3]) == "Oct-33"
        # @test string(numbers[14, 3]) == "2:57 AM"
        # @test string(numbers[15, 3]) == "2:57:46 AM"
        # @test string(numbers[16, 3]) == "2:57"
        # @test string(numbers[17, 3]) == "2:57:46"
        # @test string(numbers[18, 3]) == "10/18/33 2:57"
        # @test string(numbers[19, 3]) == "12,345"
        # @test string(numbers[20, 3]) == "12,345"
        # @test string(numbers[21, 3]) == "12,345.12"
        # @test string(numbers[22, 3]) == "12,345.12"
        # @test string(numbers[23, 3]) == "57:46"
        # @test string(numbers[24, 3]) == "296282:57:46"
        # @test string(numbers[25, 3]) == "5746.079"
        # @test string(numbers[26, 3]) == "1.23E+04"
        # @test string(numbers[27, 3]) == "\"12345.12345\""

        @test numbers[1, 3]  == 12345.12345
        @test numbers[2, 3]  == 12345.12345
        @test numbers[3, 3]  == 12345.12345
        @test numbers[4, 3]  == 12345.12345
        @test numbers[5, 3]  == 12345.12345
        @test numbers[6, 3]  == 12345.12345
        @test numbers[7, 3]  == 12345.12345
        @test numbers[8, 3]  == 12345.12345
        @test numbers[9, 3]  == 12345.12345
        @test numbers[10, 3] == DateTime("1933-10-18T02:57:46.079")
        @test numbers[11, 3] == DateTime("1933-10-18T02:57:46.079")
        @test numbers[12, 3] == DateTime("1933-10-18T02:57:46.079")
        @test numbers[13, 3] == DateTime("1933-10-18T02:57:46.079")
        @test numbers[14, 3] == DateTime("1933-10-18T02:57:46.079")
        @test numbers[15, 3] == DateTime("1933-10-18T02:57:46.079")
        @test numbers[16, 3] == DateTime("1933-10-18T02:57:46.079")
        @test numbers[17, 3] == DateTime("1933-10-18T02:57:46.079")
        @test numbers[18, 3] == DateTime("1933-10-18T02:57:46.079")
        @test numbers[19, 3] == 12345.12345
        @test numbers[20, 3] == 12345.12345
        @test numbers[21, 3] == 12345.12345
        @test numbers[22, 3] == 12345.12345
        @test numbers[23, 3] == DateTime("1933-10-18T02:57:46.079")
        @test numbers[24, 3] == DateTime("1933-10-18T02:57:46.079")
        @test numbers[25, 3] == DateTime("1933-10-18T02:57:46.079")
        @test numbers[26, 3] == 12345.12345
        @test numbers[27, 3] == "12345.12345"
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
            "date"             DateTime("1983-04-16T00:00:00")
            "hour"             DateTime("1899-12-31T19:44:43")
            "datetime"         DateTime("2018-04-16T19:19:50.999")
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
        @test xl_table(sheet[1, :]) == hcat(["text", "regular_text"])

        @test sheet["A1"] == "text"
        @test sheet["B1"] == "regular_text"

        @test xl_table(sheet["A"]) == hcat([
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
        ])

        @test xl_table(sheet["A8:B"]) ==
              ["integer neg" -2000.0; "bigint" 1.0e14; "bigint neg" "-100000000000000"]

        @test xl_table(sheet["A:B8"]) ==
            [ "integer neg" -2000.0; "bigint" 1.0e14; "bigint neg" "-100000000000000"]

        @test xl_table(sheet["A1:B2"]) == ["text" "regular_text"; "integer" 102.0]
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
            (A = "date", B = DateTime("1983-04-16T00:00:00")),
            (A = "hour", B = DateTime("1899-12-31T19:44:43")),
            (A = "datetime", B = DateTime("2018-04-16T19:19:50.999")),
            (A = "float cient", B = -220.0),
            (A = "integer neg", B = -2000.0),
            (A = "bigint", B = 1.0e14),
            (A = "bigint neg", B = "-100000000000000"),
        ]

        @test xl_rowtable(sheet; header = true) == [
            (text = "integer", regular_text = 102.0),
            (text = "float", regular_text = 102.2),
            (text = "date", regular_text = DateTime("1983-04-16T00:00:00")),
            (text = "hour", regular_text = DateTime("1899-12-31T19:44:43")),
            (text = "datetime", regular_text = DateTime("2018-04-16T19:19:50.999")),
            (text = "float cient", regular_text = -220.0),
            (text = "integer neg", regular_text = -2000.0),
            (text = "bigint", regular_text = 1.0e14),
            (text = "bigint neg", regular_text = "-100000000000000"),
        ]

        @test xl_rowtable(sheet; header = true, alt_keys = Dict("text" => "alt_text")) == [
            (alt_text = "integer", regular_text = 102.0),
            (alt_text = "float", regular_text = 102.2),
            (alt_text = "date", regular_text = DateTime("1983-04-16T00:00:00")),
            (alt_text = "hour", regular_text = DateTime("1899-12-31T19:44:43")),
            (alt_text = "datetime", regular_text = DateTime("2018-04-16T19:19:50.999")),
            (alt_text = "float cient", regular_text = -220.0),
            (alt_text = "integer neg", regular_text = -2000.0),
            (alt_text = "bigint", regular_text = 1.0e14),
            (alt_text = "bigint neg", regular_text = "-100000000000000"),
        ]

        @test xl_rowtable(sheet; alt_keys = ["alt_text", "regular_text"]) == [
            (alt_text = "text", regular_text = "regular_text"),
            (alt_text = "integer", regular_text = 102.0),
            (alt_text = "float", regular_text = 102.2),
            (alt_text = "date", regular_text = DateTime("1983-04-16T00:00:00")),
            (alt_text = "hour", regular_text = DateTime("1899-12-31T19:44:43")),
            (alt_text = "datetime", regular_text = DateTime("2018-04-16T19:19:50.999")),
            (alt_text = "float cient", regular_text = -220.0),
            (alt_text = "integer neg", regular_text = -2000.0),
            (alt_text = "bigint", regular_text = 1.0e14),
            (alt_text = "bigint neg", regular_text = "-100000000000000"),
        ]

        @test xl_rowtable(sheet; header = true, alt_keys = ["alt_text", "regular_text"]) == [
            (alt_text = "integer", regular_text = 102.0),
            (alt_text = "float", regular_text = 102.2),
            (alt_text = "date", regular_text = DateTime("1983-04-16T00:00:00")),
            (alt_text = "hour", regular_text = DateTime("1899-12-31T19:44:43")),
            (alt_text = "datetime", regular_text = DateTime("2018-04-16T19:19:50.999")),
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
                DateTime("1983-04-16T00:00:00"),
                DateTime("1899-12-31T19:44:43"),
                DateTime("2018-04-16T19:19:50.999"),
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
                DateTime("1983-04-16T00:00:00"),
                DateTime("1899-12-31T19:44:43"),
                DateTime("2018-04-16T19:19:50.999"),
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
                DateTime("1983-04-16T00:00:00"),
                DateTime("1899-12-31T19:44:43"),
                DateTime("2018-04-16T19:19:50.999"),
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
                DateTime("1983-04-16T00:00:00"),
                DateTime("1899-12-31T19:44:43"),
                DateTime("2018-04-16T19:19:50.999"),
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
                DateTime("1983-04-16T00:00:00"),
                DateTime("1899-12-31T19:44:43"),
                DateTime("2018-04-16T19:19:50.999"),
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )
    end
end

@testset "SubXLSheet interface" begin
    sub_table = xlsx["general"][:, :]

    @testset "Case №1: Accessors" begin
        @test xl_nrow(sub_table) == 10
        @test xl_ncol(sub_table) == 2
    end

    @testset "Case №2: Indexing" begin
        @test sub_table[1] == "text"
        @test sub_table[2] == "integer"

        @test sub_table[1, 1] == "text"
        @test sub_table[1, 2] == "regular_text"
        @test xl_table(sub_table[1, :]) == hcat(["text", "regular_text"])

        @test sub_table["A1"] == "text"
        @test sub_table["B1"] == "regular_text"

        @test xl_table(sub_table["A"]) == hcat([
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
        ])

        @test xl_table(sub_table["A8:B"]) ==
            [ "integer neg" -2000.0; "bigint" 1.0e14; "bigint neg" "-100000000000000"]

        @test xl_table(sub_table["A:B8"]) ==
            [ "integer neg" -2000.0; "bigint" 1.0e14; "bigint neg" "-100000000000000"]

        @test xl_table(sub_table["A1:B2"]) == ["text" "regular_text"; "integer" 102.0]
        @test sub_table["A:B"] == sub_table[:, :]

        @test_throws XLError sub_table["a1"]
        @test_throws XLError sub_table["A:A:B"]
        @test_throws XLError sub_table["A::B"]
        @test_throws XLError sub_table["A1:1A"]
        @test_throws XLError sub_table["a:B"]
        @test_throws XLError sub_table["A:b"]
        @test_throws XLError sub_table[":A"]
        @test_throws XLError sub_table["A:"]
        @test_throws XLError sub_table[":A:"]
        @test_throws XLError sub_table["1:A:"]
        @test_throws XLError sub_table[":A:1"]
        @test_throws XLError sub_table[":A:B"]
        @test_throws XLError sub_table["A:B:"]
        @test_throws XLError sub_table["1A:B2"]
        @test_throws XLError sub_table["A1B2"]
        @test_throws XLError sub_table["A101:B02"]
    end

    @testset "Case №3: Sub-...-SubXLSheet" begin
        sub_sub_table = sub_table["A1:B4"]

        parent(sub_table) == xlsx["general"]
        parent(sub_sub_table) == xlsx["general"]
        parent(sub_sub_table) == parent(sub_table)

        xl_table(sub_sub_table) == [
            "text"          "regular_text"
            "integer"    102.0
            "float"      102.2
            "date"     DateTime("1983-04-16T00:00:00")
        ]
    end

    @testset "Case №4: Utils" begin
        @test xl_rowtable(sub_table) == [
            (A = "text", B = "regular_text"),
            (A = "integer", B = 102.0),
            (A = "float", B = 102.2),
            (A = "date", B = DateTime("1983-04-16T00:00:00")),
            (A = "hour", B = DateTime("1899-12-31T19:44:43")),
            (A = "datetime", B = DateTime("2018-04-16T19:19:50.999")),
            (A = "float cient", B = -220.0),
            (A = "integer neg", B = -2000.0),
            (A = "bigint", B = 1.0e14),
            (A = "bigint neg", B = "-100000000000000"),
        ]

        @test xl_rowtable(sub_table; header = true) == [
            (text = "integer", regular_text = 102.0),
            (text = "float", regular_text = 102.2),
            (text = "date", regular_text = DateTime("1983-04-16T00:00:00")),
            (text = "hour", regular_text = DateTime("1899-12-31T19:44:43")),
            (text = "datetime", regular_text = DateTime("2018-04-16T19:19:50.999")),
            (text = "float cient", regular_text = -220.0),
            (text = "integer neg", regular_text = -2000.0),
            (text = "bigint", regular_text = 1.0e14),
            (text = "bigint neg", regular_text = "-100000000000000"),
        ]

        @test xl_rowtable(sub_table; header = true, alt_keys = Dict("text" => "alt_text")) == [
            (alt_text = "integer", regular_text = 102.0),
            (alt_text = "float", regular_text = 102.2),
            (alt_text = "date", regular_text = DateTime("1983-04-16T00:00:00")),
            (alt_text = "hour", regular_text = DateTime("1899-12-31T19:44:43")),
            (alt_text = "datetime", regular_text = DateTime("2018-04-16T19:19:50.999")),
            (alt_text = "float cient", regular_text = -220.0),
            (alt_text = "integer neg", regular_text = -2000.0),
            (alt_text = "bigint", regular_text = 1.0e14),
            (alt_text = "bigint neg", regular_text = "-100000000000000"),
        ]

        @test xl_rowtable(sub_table; alt_keys = ["alt_text", "regular_text"]) == [
            (alt_text = "text", regular_text = "regular_text"),
            (alt_text = "integer", regular_text = 102.0),
            (alt_text = "float", regular_text = 102.2),
            (alt_text = "date", regular_text = DateTime("1983-04-16T00:00:00")),
            (alt_text = "hour", regular_text = DateTime("1899-12-31T19:44:43")),
            (alt_text = "datetime", regular_text = DateTime("2018-04-16T19:19:50.999")),
            (alt_text = "float cient", regular_text = -220.0),
            (alt_text = "integer neg", regular_text = -2000.0),
            (alt_text = "bigint", regular_text = 1.0e14),
            (alt_text = "bigint neg", regular_text = "-100000000000000"),
        ]

        @test xl_rowtable(sub_table; header = true, alt_keys = ["alt_text", "regular_text"]) == [
            (alt_text = "integer", regular_text = 102.0),
            (alt_text = "float", regular_text = 102.2),
            (alt_text = "date", regular_text = DateTime("1983-04-16T00:00:00")),
            (alt_text = "hour", regular_text = DateTime("1899-12-31T19:44:43")),
            (alt_text = "datetime", regular_text = DateTime("2018-04-16T19:19:50.999")),
            (alt_text = "float cient", regular_text = -220.0),
            (alt_text = "integer neg", regular_text = -2000.0),
            (alt_text = "bigint", regular_text = 1.0e14),
            (alt_text = "bigint neg", regular_text = "-100000000000000"),
        ]

        @test xl_columntable(sub_table) == (
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
                DateTime("1983-04-16T00:00:00"),
                DateTime("1899-12-31T19:44:43"),
                DateTime("2018-04-16T19:19:50.999"),
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )

        @test xl_columntable(sub_table; header = true) == (
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
                DateTime("1983-04-16T00:00:00"),
                DateTime("1899-12-31T19:44:43"),
                DateTime("2018-04-16T19:19:50.999"),
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )

        @test xl_columntable(sub_table; header = true, alt_keys = Dict("text" => "alt_text")) == (
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
                DateTime("1983-04-16T00:00:00"),
                DateTime("1899-12-31T19:44:43"),
                DateTime("2018-04-16T19:19:50.999"),
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )

        @test xl_columntable(sub_table; alt_keys = ["alt_text", "regular_text"]) == (
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
                DateTime("1983-04-16T00:00:00"),
                DateTime("1899-12-31T19:44:43"),
                DateTime("2018-04-16T19:19:50.999"),
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )

        @test xl_columntable(sub_table; header = true, alt_keys = ["alt_text", "regular_text"]) == (
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
                DateTime("1983-04-16T00:00:00"),
                DateTime("1899-12-31T19:44:43"),
                DateTime("2018-04-16T19:19:50.999"),
                -220.0,
                -2000.0,
                1.0e14,
                "-100000000000000",
            ],
        )
    end
end

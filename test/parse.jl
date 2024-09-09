# parse

@testset "Parsing" begin
    @testset "Case №1: Parsing xlsx files" begin
        @test_nowarn xl_parse(read("xl_data/blank.xlsx"))
        @test_nowarn xl_parse(read("xl_data/simple_table.xlsx"))
        @test_nowarn xl_parse(read("xl_data/sparse_table.xlsx"))
        @test_nowarn xl_parse(read("xl_data/styled_table.xlsx"))
        @test_nowarn xl_parse(read("xl_data/general_tables.xlsx"))
        @test_nowarn xl_parse(read("xl_data/general_formats.xlsx"))
        @test_nowarn xl_parse(read("xl_data/image_table.xlsx"))
    end

    @testset "Case №2: Parsing xlsm/xltx/xltm files" begin
        @test_nowarn xl_parse(read("xl_data/simple_table.xlsm"))
        @test_nowarn xl_parse(read("xl_data/simple_table.xltx"))
        @test_nowarn xl_parse(read("xl_data/simple_table.xltm"))
    end
end

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

    @testset "Case №3: Reading xlsx files" begin
        @test_nowarn xl_open("xl_data/simple_table.xlsx")
        @test_nowarn xl_open(IOBuffer(read("xl_data/simple_table.xlsx")))
    end

    @testset "Case №3: Parsing date time" begin
        datetimes = xl_parse(read("xl_data/datetime_types.xlsx"))[1]

        @test datetimes[1,1] == DateTime("2025-03-06T00:00:00")
        @test datetimes[1,2] == DateTime("2025-03-06T00:00:00")
        @test datetimes[1,3] == DateTime("2025-03-06T00:00:00")
        @test datetimes[1,4] == DateTime("2025-03-06T00:00:00")
        @test datetimes[2,1] == DateTime("1899-12-30T11:02:10.614")
        @test datetimes[2,2] == DateTime("1899-12-30T11:02:10.614")
        @test datetimes[2,3] == DateTime("1899-12-30T11:02:10.614")
        @test datetimes[2,4] == DateTime("1899-12-30T11:02:10.614")
        @test datetimes[3,1] == DateTime("2025-03-06T11:02:19.704")
        @test datetimes[3,2] == DateTime("2025-03-06T11:02:19.704")
        @test datetimes[3,3] == DateTime("2025-03-06T11:02:19.704")
        @test datetimes[3,4] == DateTime("2025-03-06T11:02:19.704")
    end
end

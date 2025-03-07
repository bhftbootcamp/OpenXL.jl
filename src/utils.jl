# xl_utils

"""
    xl_num2date(number::Number) -> Date

Converts a number to a date in Excel style (number of days from `1899-12-30T00:00:00`).

```julia-repl
julia> xl_num2date(42440.0)
2016-03-11

julia> xl_num2date(1000.0)
1902-09-26
```
"""
function xl_num2date(number::Float64)
    return Date(1899, 12, 30) + Day(floor(Int, number))
end

xl_num2date(number::Number) = xl_num2date(convert(Float64, number))

function xl_num2time(number::Float64)
    h = 24 * number
    m = 60 * (h % 1)
    s = 60 * (m % 1)
    ms = 1000 * (s % 1)
    return Time(h ÷ 1, m ÷ 1, s ÷ 1, ms ÷ 1)
end

"""
    xl_num2time(number::Number) -> Time

Converts a number to a time in Excel style.

```julia-repl
julia> xl_num2time(0.125)
03:00:00

julia> xl_num2time(0.560)
13:26:24
```
"""
xl_num2time(number::Number) = xl_num2time(convert(Float64, number))

function xl_num2datetime(number::Float64)
    return DateTime(xl_num2date(number ÷ 1), xl_num2time(number % 1))
end

"""
    xl_num2datetime(number::Number) -> DateTime

Converts a number to a datetime in Excel style.

```julia-repl
julia> xl_num2datetime(32400.892)
1988-09-14T21:24:28.799

julia> xl_num2datetime(42125.5345)
2015-05-01T12:49:40.800
```
"""
xl_num2datetime(number::Number) = xl_num2datetime(convert(Float64, number))

function separate_number(number::String)
    return replace(string(number), r"(?<=[0-9])(?=(?:[0-9]{3})+(?![0-9]))" => ",")
end

function format_fraction(number::Number)
    frac = rationalize(number % 1, tol = 1e-2)
    return string(floor(Int, number), " ", numerator(frac), "/", denominator(frac))
end

format_integer(number::Number) = @sprintf("%i", number)
format_float(number::Number) = @sprintf("%.2f", number)
format_scientific(number::Number) = @sprintf("%.2E", number)

function format_number(x::Number, fmt_code::Int64)
    return if fmt_code == 0
        # General
        string(x)
    elseif fmt_code == 1
        # 0
        format_integer(x)
    elseif fmt_code == 2
        # 0.00
        format_float(x)
    elseif fmt_code == 3
        # #,##0
        separate_number(format_integer(x))
    elseif fmt_code == 4
        # #,##0.00
        separate_number(format_float(x))
    elseif fmt_code == 9
        # 0%
        string(format_integer(x), "%")
    elseif fmt_code == 10
        # 0.00%
        string(format_float(x), "%")
    elseif fmt_code == 11
        # 0.00E+00
        format_scientific(x)
    elseif fmt_code == 12
        # ?/?
        format_fraction(x)
    elseif fmt_code == 13
        # ??/??
        format_fraction(x)
    elseif fmt_code == 37
        # #,##0 ;(#,##0)
        num = separate_number(format_integer(x))
        x > 0 ? num : "($num)"
    elseif fmt_code == 38
        # #,##0 ;[Red](#,##0)
        num = separate_number(format_integer(x))
        x > 0 ? num : "($num)"
    elseif fmt_code == 39
        # #,##0.00;(#,##0.00)
        num = separate_number(format_float(x))
        x > 0 ? num : "($num)"
    elseif fmt_code == 40
        # #,##0.00;[Red](#,##0.00)
        num = separate_number(format_float(x))
        x > 0 ? num : "($num)"
    elseif fmt_code == 48
        # ##0.0E+0
        format_scientific(x)
    else
        # @ or Custom Formats
        string(x)
    end
end

function format_datetime(x::DateTime, fmt_code::Int64)
    return if fmt_code == 0
        # General
        string(x)
    elseif fmt_code == 14
        # mm-dd-yy
        Dates.format(x, "mm-dd-yy")
    elseif fmt_code == 15
        # d-mmm-yy
        Dates.format(x, "d-u-yy")
    elseif fmt_code == 16
        # d-mmm
        Dates.format(x, "d-u")
    elseif fmt_code == 17
        # mmm-yy
        Dates.format(x, "u-yy")
    elseif fmt_code == 18
        # h:mm AM/PM
        Dates.format(x, "H:MM p")
    elseif fmt_code == 19
        # h:mm:ss AM/PM
        Dates.format(x, "H:MM:SS p")
    elseif fmt_code == 20
        # h:mm
        Dates.format(x, "H:MM")
    elseif fmt_code == 21
        # h:mm:ss
        Dates.format(x, "H:MM:SS")
    elseif fmt_code == 22
        # m/d/yy h:mm
        Dates.format(x, "m/d/yy H:MM")
    elseif fmt_code == 45
        # mm:ss
        Dates.format(x, "MM:SS")
    elseif fmt_code == 46
        # [h]:mm:ss
        h = Dates.value(Hour(Date(x) - Date(1899, 12, 30))) + hour(x)
        m = minute(x)
        s = second(x)
        "$h:$m:$s"
    elseif fmt_code == 47
        # mmss.0
        Dates.format(x, "MMSS.s")
    else
        # @ or Custom Formats
        string(x)
    end
end

"""
    format_description(code::Int) -> String

Returns a formatting description by its `code` (see [Number Format](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-3.0.1)).

```julia-repl
julia> format_description(0)
"General"

julia> format_description(3)
"#,##0"

julia> format_description(11)
"0.00E+00"

julia> format_description(19)
"h:mm:ss AM/PM"
```
"""
function format_description(x::Int)
    return if x == 0
        "General"
    elseif x == 1
        "0"
    elseif x == 2
        "0.00"
    elseif x == 3
        "#,##0"
    elseif x == 4
        "#,##0.00"
    elseif x == 9
        "0%"
    elseif x == 10
        "0.00%"
    elseif x == 11
        "0.00E+00"
    elseif x == 12
        "?/?"
    elseif x == 13
        "??/??"
    elseif x == 14
        "mm-dd-yy"
    elseif x == 15
        "d-mmm-yy"
    elseif x == 16
        "d-mmm"
    elseif x == 17
        "mmm-yy"
    elseif x == 18
        "h:mm AM/PM"
    elseif x == 19
        "h:mm:ss AM/PM"
    elseif x == 20
        "h:mm"
    elseif x == 21
        "h:mm:ss"
    elseif x == 22
        "m/d/yy h:mm"
    elseif x == 37
        "#,##0 ;(#,##0)"
    elseif x == 38
        "#,##0 ;[Red](#,##0)"
    elseif x == 39
        "#,##0.00;(#,##0.00)"
    elseif x == 40
        "#,##0.00;[Red](#,##0.00)"
    elseif x == 45
        "mm:ss"
    elseif x == 46
        "[h]:mm:ss"
    elseif x == 47
        "mmss.0"
    elseif x == 48
        "##0.0E+0"
    elseif x == 49
        "@"
    else
        "Custom"
    end
end

"""
    column_letter_to_index(letter::AbstractString) -> Int

Converts an Excel column letter ("A", "B", ..., "Z", "AA", etc.) to its corresponding numerical index.

## Examples
```julia
julia> column_letter_to_index("A")
1

julia> column_letter_to_index("Z")
26

julia> column_letter_to_index("AA")
27
```
"""
function column_letter_to_index(letter::AbstractString)
    idx = 0
    for c in letter
        idx = (c - 'A' + 1) + idx * 26
    end
    return idx
end

"""
    index_to_column_letter(inx::Int) -> String

Converts a numerical index into its corresponding Excel column letter ("A", "B", ..., "Z", "AA", etc.).

## Examples
```julia
julia> index_to_column_letter(1)
"A"

julia> index_to_column_letter(26)
"Z"

julia> index_to_column_letter(27)
"AA"
```
"""
function index_to_column_letter(inx::Int)
    inx, rem1 = divrem(inx - 1, 26)
    iszero(inx) && return string(Char(rem1 + 65))

    inx, rem2 = divrem(inx - 1, 26)
    iszero(inx) && return string(Char(rem2 + 65), Char(rem1 + 65))

    _, rem3 = divrem(inx - 1, 26)
    return string(Char(rem3 + 65), Char(rem2 + 65), Char(rem1 + 65))
end

struct CellRange
    column::Union{Nothing,Int}
    row::Union{Nothing,Int}
end

function parse_cell_addr(addr::AbstractString)
    match_result = match(r"^([A-Z]+)([1-9]\d*)?$", addr)

    isnothing(match_result) && throw(
        XLError("Invalid cell address format. Expected 'A1', 'AB12', etc., got $addr."),
    )

    column_letter, row_number = match_result

    return CellRange(
        column_letter_to_index(column_letter),
        isnothing(row_number) ? nothing : parse(Int64, row_number),
    )
end

function parse_cell_range(addr::AbstractString)
    if isempty(addr)
        throw(XLError("Empty cell range. Expected 'A', 'A:B', 'A1:B2', 'AB12:CD34', etc."))
    end

    range_parts = split(addr, ':'; limit = 2)

    any(isempty, range_parts) && throw(
        XLError("Incomplete address range. Expected 'A', 'A:B', 'A1:B2', 'AB12:CD34', etc., got $addr.")
    )

    cell_ranges = parse_cell_addr.(range_parts)

    return (cell_ranges[1], get(cell_ranges, 2, CellRange(nothing, nothing)))
end

function isdatetime(fmt_id::Int64, fmt_code::String, cell_type::Union{Nothing,String})
    date_tokens = r"[dD]{1,4}|[mM]{1,4}|[yY]{2,4}"
    time_tokens = r"[hH]{1,2}|[mM]{1,2}|[sS]{1,2}"
    
    has_date = occursin(date_tokens, fmt_code)
    has_time = occursin(time_tokens, fmt_code)
    
    return cell_type == "d" || 14 <= fmt_id <= 22 || 45 <= fmt_id <= 47 || has_date || has_time 
end
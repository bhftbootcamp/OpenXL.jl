# xl_utils

"""
    column_letter_to_index(letter::AbstractString) -> Int

Converts an Excel column letter to its corresponding numerical index.

## Arguments
- `letter::AbstractString`: The column letter(s) (e.g., "A", "B", ..., "Z", "AA", etc.).

## Returns
- `Int`: The numerical index corresponding to the column letter.

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

Converts a numerical index into its corresponding Excel column letter.

## Arguments
- `inx::Int`: The numerical index of the column. The index should be a positive integer, where 1 corresponds to "A", 2 to "B", ..., 26 to "Z", 27 to "AA", and so forth.

## Returns
- `String`: The Excel column letter corresponding to the given index.

## Examples
```julia
julia> index_to_column_letter(1)
"A"

julia> index_to_column_letter(26)
"Z"

julia> index_to_column_letter(27)
"AA"

julia> index_to_column_letter(703)
"AAA"
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
    column::Maybe{Int}
    row::Maybe{Int}
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

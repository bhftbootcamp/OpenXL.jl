# utils

#__ Indexing

function column_letter_to_index(key::AbstractString)
    i = 0
    for c in key
        i = (c - 'A' + 1) + i * 26
    end
    return i
end

function index_to_column_letter(i::Int)
    i, r1 = divrem(i - 1, 26)
    iszero(i) && return string(Char(r1 + 65))

    i, r2 = divrem(i - 1, 26)
    iszero(i) && return string(Char(r2 + 65), Char(r1 + 65))

    _, r3 = divrem(i - 1, 26)
    return string(Char(r3 + 65), Char(r2 + 65), Char(r1 + 65))
end

function gen_column_keys(n::Int, alt_keys::Dict{String,String})
    return map(1:n) do i
        key = index_to_column_letter(i)
        return Symbol(get(alt_keys, key, key))
    end
end

function gen_column_keys(n::Int, alt_keys::AbstractVector{String})
    return map(1:n) do i
        key = index_to_column_letter(i)
        return Symbol(get(alt_keys, i, key))
    end
end

struct CellRange
    column::Maybe{Int}
    row::Maybe{Int}
end

function parse_cell_addr(addr::AbstractString)
    match_result = match(r"^([A-Z]+)([1-9]\d*)?$", addr)

    if isnothing(match_result)
        throw(
            XLError("Invalid cell address format. Expected 'A1', 'AB12', etc., got $addr."),
        )
    end

    column_letter, row_number = match_result
    column_index = column_letter_to_index(column_letter)
    row_index = isnothing(row_number) ? nothing : parse(Int64, row_number)

    return CellRange(column_index, row_index)
end

function parse_cell_range(addr::AbstractString)
    if isempty(addr)
        throw(XLError("Empty cell range. Expected 'A', 'A:B', 'A1:B2', 'AB12:CD34', etc."))
    end

    range_parts = split(addr, ':', limit = 2)

    if any(isempty, range_parts)
        throw(
            XLError(
                "Incomplete address range. Expected 'A', 'A:B', 'A1:B2', 'AB12:CD34', etc., got $addr.",
            ),
        )
    end

    cell_ranges = parse_cell_addr.(range_parts)

    return (cell_ranges[1], get(cell_ranges, 2, CellRange(nothing, nothing)))
end

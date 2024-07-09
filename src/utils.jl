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

struct CellRange
    column::Maybe{Int}
    row::Maybe{Int}
end

function parse_cell_addr(addr::AbstractString)
    result = match(r"^([A-Z]+)([1-9]\d*)?$", addr)

    if isnothing(result)
        throw(XLError("Invalid cell address format. Expected 'A1', 'AB12', etc., got $addr."))
    end

    col_key, row_index = result

    col_ind = column_letter_to_index(col_key)
    row_ind = isnothing(row_index) ? nothing : parse(Int64, row_index)

    return CellRange(col_ind, row_ind)
end

function parse_cell_range(addr::AbstractString)
    if isempty(addr)
        throw(XLError("Empty cell range."))
    end

    parts = split(addr, ':', limit = 2)

    if any(isempty, parts)
        throw(XLError("Incomplete address range. Expected 'A', 'A:B', 'A1:B2', 'AB12:CD34', etc., got $addr."))
    end

    vals = parse_cell_addr.(parts)

    return (vals[1], length(vals) == 1 ? CellRange(nothing, nothing) : vals[2])
end

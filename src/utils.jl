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

function parse_cell_addr(addr::AbstractString)
    result = match(r"^([A-Z]+)([1-9]+[0-9]*)?$", addr)
    if isnothing(result)
        throw(XLError("Invalid cell address format. Expected 'A1', 'AB12', etc., got $addr."))
    else
        column_key, row_index = result

        col_ind = if isnothing(column_key)
            throw(XLError("Invalid column address symbol. Expected 'A', 'A1', 'AB12', etc., got $addr."))
        else
            column_letter_to_index(column_key)
        end

        row_ind = if isnothing(row_index)
            nothing
        else
            parse(Int64, row_index)
        end

        return col_ind, row_ind
    end
end


function parse_cell_range(addr::AbstractString)
    if isempty(addr)
        throw(XLError("Empty cell range."))
    end

    parts = split(addr, ':', limit = 2)
    ranges = NTuple{2,Maybe{Int}}[
        (nothing, nothing),
        (nothing, nothing),
    ]

    for (index, part) in enumerate(parts)
        if isempty(part)
            throw(XLError("Incompleted address range. Expected 'A', 'A:B', 'A1:B2', 'AB12:CD34', etc., got $addr."))
        end
        ranges[index] = parse_cell_addr(part)
    end
    return tuple(ranges...)
end

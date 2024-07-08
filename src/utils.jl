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
    result = match(r"^([A-Z]+)([1-9]+[0-9]*)$", addr)
    return if isnothing(result)
        throw(XLError("Invalid cell address format. Expected 'A1', 'AB12', etc., got $addr."))
    else
        column_key, row_index = result
        column_letter_to_index(column_key), parse(Int64, row_index)
    end
end

function try_parse_range_of_cells(x::AbstractString)
    result = match(r"^([A-Z]+[1-9]+[0-9]*):([A-Z]+[1-9]+[0-9]*)$", x)
    return if !isnothing(result)
        l_addr, r_addr = result
        l_col_addr, l_row_addr = parse_cell_addr(l_addr)
        r_col_addr, r_row_addr = parse_cell_addr(r_addr)
        l_row_addr:r_row_addr, l_col_addr:r_col_addr
    end
end

function try_parse_range_of_cells(x::AbstractString, dims::Dims)
    result = match(r"^([A-Z]+[1-9]+[0-9]*):([A-Z]+)$", x)
    return if !isnothing(result)
        l_addr, r_col_addr = result
        r_row_addr = dims[1]
        l_col_addr, l_row_addr = parse_cell_addr(l_addr)
        l_row_addr:r_row_addr, l_col_addr:column_letter_to_index(r_col_addr)
    end
end

function try_parse_single_cell(x::AbstractString)
    result = match(r"^([A-Z]+)([1-9]+[0-9]*)$", x)
    return if !isnothing(result)
        column_key, row_index = result
        parse(Int64, row_index), column_letter_to_index(column_key)
    end
end

function try_parse_range_of_columns(x::AbstractString)
    result = match(r"^([A-Z]+):([A-Z]+)$", x)
    return if !isnothing(result)
        l_addr, r_addr = result
        return (:), column_letter_to_index(l_addr):column_letter_to_index(r_addr)
    end
end

function try_parse_single_column(x::AbstractString)
    result = match(r"^([A-Z]+)$", x)
    return if !isnothing(result)
        col_addr = result[1]
        (:), column_letter_to_index(col_addr)
    end
end

function parse_cell_range(addr::AbstractString, dims::Dims)
    result = try_parse_range_of_cells(addr)
    !isnothing(result) && return result

    result = try_parse_range_of_cells(addr, dims)
    !isnothing(result) && return result

    result = try_parse_single_cell(addr)
    !isnothing(result) && return result

    result = try_parse_range_of_columns(addr)
    !isnothing(result) && return result

    result = try_parse_single_column(addr)
    !isnothing(result) && return result

    throw(XLError("Invalid table slice addresses. Expected 'A', 'A:B', 'A1:B2', 'AB12:CD34', etc., got $addr."))
end

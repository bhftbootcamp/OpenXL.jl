# xl_print

function compact_length(str::AbstractString; max_len::Int)
    return min(length(str), max_len)
end

function compact_string(str::AbstractString; max_len::Int)
    return if length(str) > max_len
        chars = collect(str)
        string(chars[1:max_len-1]..., "…")
    else
        str
    end
end

function head_tail(x::AbstractVector, head::Int, tail::Int)
    len = length(x)
    return if head >= len - tail
        x
    else
        view(x, [1:head; len-tail+1:len])
    end
end

"""
    xl_print([io::IO], sheet::AbstractXLSheet; kw...)

Print a `sheet` as a table representation.

## Keyword arguments
- `title::AbstractString = "Sheet"`: Table title in upper left corner.
- `header::Bool = false`: Use first row elements as column headers.
- `max_len::Int = 16`: Maximum length of an element in a cell.
- `compact::Bool = true`: Omit rows and columns to save space.

## Examples
```julia-repl
julia> xlsx = xl_parse(xl_sample_employee_xlsx())
1-element XLWorkbook:
 1001x13 XLSheet("Employee")

julia> xl_print(xlsx["Employee"]; header = true)
 Sheet │ eeid    full_name        job_title         ⋯  country        city       exit_date
───────┼───────────────────────────────────────────────────────────────────────────────────
     2 │ E02387  Emily Davis      Sr. Manger        ⋯  United States  Seattle    44485.0
     3 │ E04105  Theodore Dinh    Technical Archi…  ⋯  China          Chongqing  nothing
     4 │ E02572  Luna Sanders     Director          ⋯  United States  Chicago    nothing
     5 │ E02832  Penelope Jordan  Computer System…  ⋯  United States  Chicago    nothing
     ⋮ │ ⋮       ⋮                ⋮                 ⋯  ⋮              ⋮          ⋮
  1000 │ E02521  Lily Nguyen      Sr. Analyst       ⋯  China          Chengdu    nothing
  1001 │ E03545  Sofia Cheng      Vice President    ⋯  United States  Miami      nothing
```
"""
function xl_print(
    io::IO,
    sheet::AbstractXLSheet;
    title::AbstractString = "Sheet",
    header::Bool = false,
    compact::Bool = true,
    max_len::Int = compact ? 16 : typemax(Int),
)
    table = xl_table(sheet)
    isempty(table) && return nothing
    display_height, display_width = displaysize(io)
    display_elements = floor(Int64, display_height/2 - 4)
    num_rows, num_cols = size(table)
    title_width = max(length(title), ndigits(num_rows))
    cols, row_start = if header
        view(table, 1, :), 2
    else
        (index_to_column_letter.(1:num_cols), 1)
    end
    col_widths = map(enumerate(eachcol(table))) do (i, col)
        return max(
            maximum(
                cell -> compact_length(string(cell), max_len = max_len),
                head_tail(col, display_elements, display_elements);
                init = 0,
            ),
            length(string(cols[i])),
        )
    end
    print(io, " ", rpad(title, title_width), " │ ")
    col_idx = 1
    omitted = false
    sep_width = 0
    while col_idx <= num_cols
        if compact && !omitted && col_idx > 4 && num_cols > 8
            print(io, " ⋯  ")
            omitted = true
            sep_width += 4
            col_idx = max(5, num_cols - 3)
        else
            value = rpad(cols[col_idx], col_widths[col_idx])
            sep_width += length(value) + 2
            print(io, value, "  ")
            col_idx += 1
        end
    end
    print(io, "\n", repeat('─', title_width + 2), "┼", repeat('─', sep_width), "\n")
    row_idx = row_start
    omitted_rows = false
    while row_idx <= num_rows
        if compact && !omitted_rows && row_idx > display_elements && num_rows > display_height
            print(io, " ", lpad("⋮", title_width), " │ ")
            col_idx = 1
            omitted = false
            while col_idx <= num_cols
                if !omitted && col_idx > 4 && num_cols > 8
                    print(io, " ⋯  ")
                    omitted = true
                    col_idx = max(5, num_cols - 3)
                end
                print(io, rpad("⋮", col_widths[col_idx]), "  ")
                col_idx += 1
            end
            print(io, "\n")
            omitted_rows = true
            row_idx = num_rows - display_elements
        end
        print(io, " ", lpad(row_idx, title_width), " │ ")
        col_idx = 1
        omitted = false
        while col_idx <= num_cols
            if compact && !omitted && col_idx > 4 && num_cols > 8
                print(io, " ⋯  ")
                omitted = true
                col_idx = max(5, num_cols - 3)
            end
            value = compact_string(string(table[row_idx, col_idx]), max_len = max_len)
            print(io, rpad(value, col_widths[col_idx]), "  ")
            col_idx += 1
        end
        print(io, "\n")
        row_idx += 1
    end
end

function xl_print(sheet::AbstractXLSheet; kw...)
    return xl_print(stdout, sheet; kw...)
end

function xl_print(x...; kw...)
    return print(x...; kw...)
end

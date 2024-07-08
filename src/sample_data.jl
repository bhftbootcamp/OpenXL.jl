# sample_data

export xl_sample_ticker24h_xlsx,
    xl_sample_stock_xlsx,
    xl_sample_employee_xlsx

function xl_sample_ticker24h_xlsx()
    return read(joinpath(@__DIR__, "../assets/ticker24h_sample_data.xlsx"))
end

function xl_sample_stock_xlsx()
    return read(joinpath(@__DIR__, "../assets/stock_sample_data.xlsx"))
end

function xl_sample_employee_xlsx()
    return read(joinpath(@__DIR__, "../assets/employee_sample_data.xlsx"))
end

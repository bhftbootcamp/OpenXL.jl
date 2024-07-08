using OpenXL
using Documenter

DocMeta.setdocmeta!(OpenXL, :DocTestSetup, :(using OpenXL); recursive = true)

makedocs(;
    modules = [OpenXL],
    sitename = "OpenXL.jl",
    format = Documenter.HTML(;
        repolink = "https://github.com/bhftbootcamp/OpenXL.jl",
        canonical = "https://bhftbootcamp.github.io/OpenXL.jl",
        edit_link = "master",
        assets = ["assets/favicon.ico"],
        sidebar_sitename = true,
    ),
    pages = [
        "Home" => "index.md",
        "API Reference" => "pages/api_reference.md",
    ],
    warnonly = [:doctest, :missing_docs],
)

deploydocs(;
    repo = "github.com/bhftbootcamp/OpenXL.jl",
    devbranch = "master",
    push_preview = true,
)

# OpenXL.jl Changelog

The latest version of this file can be found at the master branch of the [OpenXL.jl repository](https://github.com/bhftbootcamp/OpenXL.jl).

## 2.0.0 (15/07/2024)

### Added

- Added the ability to set headers using a vector for methods `xl_rowtable` and `xl_columntable`.
- Added support for rows iteration by `Base.eachrow` method.
- Added method `xl_open` for reading and parsing XL files.
- Added methods `column_letter_to_index` and `index_to_column_letter` to the documentation.
- General fixes and improvements.

### Changed

- Renamed accessor methods `xl_name` and `xl_names` to `xl_sheetname` and `xl_sheetnames`.
- Renamed print method `xl_print_sheet` to `xl_print`.

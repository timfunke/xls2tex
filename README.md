# `xls2tex.py`

`xls2tex.py` takes an Excel file as first argument and outputs the first sheet 
as LaTeX-formatted table to `STDOUT` (to ultimately turn it into a PDF).
No formatting in the Excel file will be transferred, only the cell contents with their default 
It is expected that the output will need some manual editing / tuning.
For example, the `tabularx` column format specifiers will need to be updated as they are all specified as `'X'`.
The output of the table should be relatively easy to read as it is visually formatted.
The LaTeX preamble is read from a resource file called `preamble.tex`.
The script is quite simplistic in the sense that it, for instance, does not do any checking on its parameters.

A simple `makefile` demonstrates the usage and workflow (which assumes an Excel input file `input.xlsx` is present, which is not part of this repo).


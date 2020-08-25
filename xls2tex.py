#!/bin/python

"""
This script takes an Excel file as first argument and outputs the first sheet 
as LaTeX-formatted table to STDOUT (to ultimately turn it into a PDF).
It is expected that the output will need some manual editing / tuning.
For exampe, the `tabularx` column format specifiers will need to be updated.
The output of the table should be relatively easy to read as it is visually formatted.
The LaTeX preamble is read from a resource file called `preamble.tex`.
The script is quite simplistic in the sense that it, for instance, does not do any checking on its parameters.
"""

import xlrd         # to read the data from the Excel sheet
import sys          # to read command-line parameters (a single one: the Excel file)
import itertools    # function "zip_longest" is needed twice to fill up empty columns with no values

maxLength = 0
maxColLengths = []
postamble = "\\end{document}"

### read the contents of the Excel file into a list of lists of strings ###########################

# open the Excel file ...
wb = xlrd.open_workbook(sys.argv[1]) 
# ... and pick the first sheet:
sheet = wb.sheet_by_index(0) 

lines = []
for i in range(sheet.nrows):
  # turn each line into a list of strings:
  line = [str(elem) for elem in sheet.row_values(i)]
  # record the length of the longest list:
  if len(line) > maxLength:
    maxLength = len(line)
  # append it to our list:
  lines.append(line)
  # record the max width of each column -- we do this only for nice (but potentially impractical) formatting of the output table
  colLengths = [len(col) for col in line]
  maxColLengths = [max(a, b) for (a, b) in itertools.zip_longest(colLengths, maxColLengths, fillvalue=0)]
# now, the list `lines` holds all the data from the Excel file, all values standard-converted to strings

### output the lines in memory formatted as LaTeX table ###########################

# put the standard LaTeX header formatting information in the `preamble.tex` resource file:
preamble = ''
with open('preamble.tex', encoding="utf8") as f:
  preamble = ''.join(f.readlines())
# I use '<>' instead of '{}' in LaTeX template resources because the LaTeX standard '{}' 
# mess up variable formatting with `format`, so '<>' need to be replaced by '{}' now:
print(preamble.replace('<', '{').replace('>', '}'))

# specify the formatters so they accommodate for the longest value in each column:
formatters = ["<:{}s>".format(a).replace('<', '{').replace('>', '}') for a in maxColLengths]
# output some diagnostic information in the LaTeX source code:
print("% DEBUG maxiumum column lengths: ", maxColLengths)
print("% DEBUG column formatters: ", formatters)
# there can be only one "X" column in a tabularx environment, so it needs to be decided later which one it's going to be:
print("\\begin{tabularx}{\\textwidth}{|" + ("X|" * maxLength) + "} \\hline")
# use "scriptsize" font because very likely space will be tight!
for line in lines:
  print("\\scriptsize{}" \
        + " & \\scriptsize{}".join([a.format(b) for (a, b) in itertools.zip_longest(formatters, line, fillvalue="")]) \
        + "\\\\    \\hline")
print("\\end{tabularx}")
print(postamble)

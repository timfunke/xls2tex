all: output.pdf

output.tex: xls2tex.py preamble.tex input.xlsx
	python3 xls2tex.py input.xlsx > output.tex

output.pdf: output.tex
	xelatex output.tex
	xdg-open output.pdf

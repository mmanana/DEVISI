latex caso001
bibtex caso001
makeindex -s caso001.ist -t caso001.alg -o caso001.acr caso001.acn
makeindex caso001.nlo -s nomencl.ist -o caso001.nls
latex caso001
latex caso001
dvips -O 0cm,0cm caso001

"C:\Program Files (x86)\Adobe\Acrobat 10.0\Acrobat\acrodist" "C:\Alberto\Latex\Innovacion\caso001\caso001.ps"

del caso001.ps
caso001.pdf
del caso001.dvi
del caso001.mtc
del caso001.aux
del caso001.bbl
del caso001.blg

#!/bin/bash

# check for Pymetrics flag
pymetrics=false
if [ "$1" = "-p" ] || [ "$1" = "--pymetrics" ]
then
	pymetrics=true
fi

# read company name
if [ $pymetrics = true ];
then
	name=$2
else
	name=$1
fi

# create hyphen-separated company name
if [ $pymetrics == true ]
then
	for arg in $*
	do
		if [ $arg != $1 ] && [ $arg != $2 ]
		then
			name+="-"
			name+=$arg
		fi
	done
else
	for arg in $*
	do
		if [ $arg != $1 ]
		then
			name+="-"
			name+=$arg
		fi
	done
fi

# create the document
if [ $pymetrics = true ];
then
	python create_letter.py -p $name
else
	python create_letter.py $name
fi
name="cover-letter-"$name

# move pdf file to dropbox
mv $name.docx ~/Dropbox/Jobs/

# convert to pdf using LibreOffice writer
/Applications/LibreOffice.app/Contents/MacOS/soffice --headless --convert-to pdf --outdir ~/Dropbox/Jobs/ ~/Dropbox/Jobs/$name.docx

# rename the pdf file
mv ~/Dropbox/Jobs/$name.pdf ~/Dropbox/Jobs/Samuel-Hage-Cover-Letter.pdf

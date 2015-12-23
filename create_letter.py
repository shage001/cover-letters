'''
Sam Hage
Script to write cover letters for me
12/2015
'''

import docx
import shutil
import sys
import os
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt

TEMPLATE = 'TEMPLATE.docx'

def main():
	"""
	**********************************************************************************************************************
	Main code to run on program execution
	"""
	# company_name_list = sys.argv[1:]
	pymetrics = False

	if len( sys.argv ) > 2 and sys.argv[1] in [ '-p', '--pymetrics' ]:
		pymetrics = True

	if pymetrics:
		company_name = sys.argv[2]
	else:
		company_name = sys.argv[1]
	create_document( company_name, pymetrics )


def create_document( company_name, pymetrics ):
	"""
	**********************************************************************************************************************
	Copy the template document and change the appropriate information

	@param: {string[]} company_name_list A list of words in the company name
	@param: {boolean} pymetrics Whether to include the line about Pymetrics
	"""
	## read company name and position ##
	# company_name_list = raw_input( 'Company name: ' ).lower().split()
	position = raw_input( 'Position: ' ).lower()
	changing_the_way = raw_input( '...changing the way ' )

	## construct the file name ##
	# file_name = 'cover-letter-' + '-'.join( company_name_list ) + '.docx'
	file_name = 'cover-letter-' + company_name + '.docx'
	company_name_list = company_name.split( '-' )
	company_name = ' '.join( company_name_list ).title()

	## copy the template ##
	shutil.copy( TEMPLATE, file_name )
	document = docx.Document( file_name )

	## set font ##
	obj_styles = document.styles
	obj_charstyle = obj_styles.add_style( 'Better', WD_STYLE_TYPE.PARAGRAPH )
	obj_charstyle.font.size = Pt(12)
	obj_charstyle.font.name = 'Times New Roman'

	## replace the text as appropriate ##
	i = 0
	for p in document.paragraphs:

		p.style = obj_charstyle # set paragraph style
		p.text = p.text.replace( '$$$$$', company_name )
		p.text = p.text.replace( '#####', changing_the_way )
		p.text = p.text.replace( '@@@@@', position )

		if i == 4 and not pymetrics: # remove the sentence about pymetrics
			start = p.text.find( 'I discovered' )
			end = p.text.find( 'evaluation', start ) + 12
			p.text = p.text[ : start ] + p.text[ end : ]
		i += 1

	document.save( file_name )


def print_usage():
	"""
	**********************************************************************************************************************
	Print the correct usage of the program if arguments are entered incorrectly
	"""
	print( 'python write_letter.py <company-name> <pymetrics=False>' )

if __name__ == '__main__':
	main()

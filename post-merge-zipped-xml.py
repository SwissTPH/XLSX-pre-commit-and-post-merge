'''
pre-commit-XML.py

Last updated 23/02/2019

Script extracts customUI.xml/customUI14.xml files into a new 'FILENAME.XML' subdirectory within the repository root folder (by default)

With the standard .gitignore entries, only Excel files located within the root directory of the repository will be processed & added to a commit
'''

import os
import sys
import getopt
import shutil
import xml.dom.minidom
from zipfile import ZipFile

# list excel extensions that will be processed and have VBA modules extracted
excel_file_extensions = ('xlsx', 'docx' )

import os, zipfile

prefix = '--unzipped--'

extension = ('xlsb', 'xlsx', 'xlsm', 'xlam', 'xltm')
	
# remove pretty print	
def recursive_prettyprint_removal(dir):
	for item in os.listdir(dir):		
		if item.endswith('.xml'): 
			file_name = os.path.abspath(dir+'/'+item) # get full path of files
			print(file_name)
			xmlContent = xml.dom.minidom.parse(file_name)
			Content = xmlContent.toprettyxml('','','utf-8')
			if isinstance(Content, str):
				print("pretty print string "+item)
				fout = open(file_name, "w")
				fout.write(Content)
				fout.close()
			elif isinstance(Content, bytes):
				print("pretty print bytes "+item)
				fout = open(file_name, "wb")
				fout.write(Content)
				fout.close()
			else:
				print(type(Content)+' is not supported')
		if os.path.isdir(dir+'/'+item):
			recursive_prettyprint_removal(dir+'/'+item)
def usage():
	print("specify prefix with -p or --prefix")
	print("specify extension with -e or --extension")
	print("specify root directory with -d or --directory")

#ZIP file back		

def zipdir(path, ziph):
    # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root,file))
			
def main(argv):                          
	prefix = '!unzipped!'  
	extension = ('xlsb', 'xlsx', 'xlsm', 'xlam', 'xltm')
	main_dir = '.'	
	try:
		opts, args = getopt.getopt(argv, "hg:d", ["help", "prefix=","extension="])
	except getopt.GetoptError:		 
		usage()
		sys.exit(2)
	for opt, arg in opts:
		if opt in ("-h", "--help"):
			usage()
			sys.exit()
		elif opt == '-d':
			global _debug
			_debug = 1
		elif opt in ("-p", "--prefix"):
			prefix = arg
		elif opt in ("-e", "--extension"):
			extension = arg
		elif opt in ("-d", "--directory"):
			main_dir = arg
			
	for dir in os.listdir(main_dir): # loop through items in dir
		if os.path.isdir(dir) and dir.startswith(prefix): 
			recursive_prettyprint_removal(dir)

	rootPath = os.getcwd()
	for item in os.listdir(main_dir): # loop through items in dir
		if os.path.isdir(item) and item.startswith(prefix): 
			zip_ref = zipfile.ZipFile(item[len(prefix):], "w")
			file_name = os.path.abspath(item)
			os.chdir(file_name)
			zipdir('./',zip_ref)
			os.chdir(rootPath)
		
if __name__ == "__main__":
	main(sys.argv[1:])		
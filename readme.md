# Excel versioning using Git

Those script act for commit and merge, the data stored on git hub are the unzipped files, the scpits can be used in stand alone mode and parameters:

-d, --directory to specify where to look for the xlsx
-e, --extension to specify the extension to zip/unzip
-p, --prefix to specify the prefix of the folder containing the zip file content

## git commit
Before the actual git commit, it will unzip the xlsx (should work with other format like docx) and will turn all the xml in pretty print so the diff can show the lines (Excel XML are on 2 lines)

This is perfomed by copiyng the files "pre-commit" and "pre-commit-zipped-xml" in the `.git/hooks/` folder

## git pull/merge
Before the actual git pull/merge, it will unzip the xlsx (should work with other format like docx) and will turn all the xml in pretty print so the diff can show the lines (Excel XML are on 2 lines)

After the actual git pull/merge, it will zip the xlsx (should work with other format like docx) after putting bremoving the pretty print so Excel can recognize the file

This is perfomed by copiyng the files "pre-merge", "pre-commit-zipped-xml", "post-merge" and "post-merge-zipped-xml.py"


## Credits

Those Scripts were created in the context of the Sysref project funded by the Stanley Thomas Johnson Fundation https://www.johnsonstiftung.ch/ to keep a version control on XLSForms used to make CDSS (Computerized decision support systems)

Those script require Python3, they were inspired by https://github.com/Agent6-6-6/Excel-VBA-XML-Export-Pre-Commit-Hook
https://paulhammant.com/2015/07/30/git-storing-unzipped-office-docs/
https://stackoverflow.com/questions/1855095/how-to-create-a-zip-archive-of-a-directory-in-python
Python and docs and other forums


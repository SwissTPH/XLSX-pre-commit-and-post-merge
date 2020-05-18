# Excel versioning using Git

Those script act for commit and merge, the data stored on git hub are the unzipped files, the scpits can be used in stand alone mode and parameters:
```
-d, --directory to specify where to look for the xlsx

-e, --extension to specify the extension to zip/unzip

-p, --prefix to specify the prefix of the folder containing the zip file content

```

## Git clone

After a doing the git-clone 

```
git clone <your repository using those script>
```
Copy the files "pre-commit", "pre-merge",  "post-merge", "pre-commit-zipped-xml.py" and "post-merge-zipped-xml.py" in the `.git/hooks/` folder

Then the "post-merge-zipped-xml.py" must be called fron the rool to create the xlsx file

```
pynthon ./.git/hooks/post-merge-zipped-xml.py
```

## git commit
Before the actual git commit, it will unzip the xlsx (should work with other format like docx) and will turn all the xml in pretty print so the diff can show the lines (Excel XML are on 2 lines)

This is perfomed by the files "pre-commit" and "pre-commit-zipped-xml.py"

## git pull/merge
Before the actual git pull/merge, it will unzip the xlsx (should work with other format like docx) and will turn all the xml in pretty print so the diff can show the lines (Excel XML are on 2 lines)

After the actual git pull/merge, it will zip the xlsx (should work with other format like docx) after putting bremoving the pretty print so Excel can recognize the file

This is perfomed by the files "pre-merge", "pre-commit-zipped-xml.py", "post-merge" and "post-merge-zipped-xml.py"


## Credits

Those Scripts were created in the context of the Sysref project funded by the Stanley Thomas Johnson Fundation https://www.johnsonstiftung.ch/ to keep a version control on XLSForms used to make CDSS (Computerized decision support systems)

Those script require Python3, they were inspired by
- https://github.com/Agent6-6-6/Excel-VBA-XML-Export-Pre-Commit-Hook
- https://paulhammant.com/2015/07/30/git-storing-unzipped-office-docs/
- https://stackoverflow.com/questions/1855095/how-to-create-a-zip-archive-of-a-directory-in-python
- Python docs 
- Other forums


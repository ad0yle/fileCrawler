# fileCrawler

_run italicized lines in the command prompt window by copying & pasting them in and hitting enter_

__bold items denote files and folders__
 
1. download __imperative things__ folder

set up proxy
**skip steps 2&3 if you are not connected to BLUESSO while following these instructions**

2. _set HTTP_PROXY=http://iss-americas-pitc-cincinnatiz.proxy.corporate.ge.com:80_
3. _set HTTPS_PROXY=http://iss-americas-pitc-cincinnatiz.proxy.corporate.ge.com:80_

install python

4. double click on the file python-2.7.14.amd64.msi in the __imperative things/install files__ folder and follow the default setup instructions
5. Open command prompt (click the windows logo in the bottom left corner and type command prompt to open the application)
6. _cd C:\Python27_
7. _python_ - make sure this does not return an error
8. move get-pip.py from the __imperative things__ folder to the __Python27__ folder
install pip
9. _cd C:\Python27
10. python get-pip.py
11. cd C:\Python27\Scripts
12. pip freeze
13. pip install pdfminer
14. pip install python-docx
15. pip install python-pptx_
--------------------------------------------------
16. move fileCrawler.py from the __imperative things__ folder to the __Python27__ folder
17. create a folder in the C:\Python27 folder called __files_to_parse__ and put all docx, ppt, pptx, & PDF files you would like to scan in there
18. move keywords.csv from the __imperative things__ folder to the __Python27__ folder and replace all sample key words with any number of  key words you would like to search for in the documents
_19. cd C:\Python27\
20. python fileCrawler.py_
21. open "crawlerOutput.txt"
 
notes
- PDF files cannot start with numbers
- parser cannot handle .doc files, only .docx - if this is a problem(aka all files are .doc), I can write a workaround for this
- crawlerOutput.txt is the output file that will show the keyword search report - this will be written over each time the program (fileCrawler.py) is run, so if you would like to run this program on different segments of files, save the output file in a new location, so it does not get overwritten

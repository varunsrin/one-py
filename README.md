onepy
=====

COM Object Model for OneNote 2013 in Python




#### What are the requirements for onepy?

* Windows 7 with Python 3.x
* OneNote 2013 or 2010 with your notebooks open


#### How do I setup my environment?

* Install Python 3.4 x86 from [here](https://www.python.org/download/releases/3.4.0/) 
* Install PyWin32 for Python 3.4 x86 from [here](http://sourceforge.net/projects/pywin32/files/pywin32/) 
* Add `C:\Python34\` to your PATH variable
* Run `C:\Python34\Lib\site-packages\win32com\client\makepy.py`
* Select `Microsoft OneNote 15.0 Extended Type Library`


#### How do I submit a new version to the Package Manager?

* From the repo, run `python.exe setup.py register sdist bdist_wininst upload`


#### How do I install onepy?

`pip install onepy`


#### How do I use onepy?

onepy exposes two main classes - OneNote and ONProcess. 

**OneNote**

OneNote is an object model class that lets you read content and hierarchy 
from the OneNote application. It exposes them as native python types so you
can read OneNote data without having to muck around with the underlying
COM interfaces.

Updating content via the object model is possible, but not implemented today.

Use OneNote to read content from notebooks:
```python
import onepy
  
on = onepy.OneNote()
  
# print a list of notebooks open in the OneNote 2013 client
for notebook in on.hierarchy:
  print (notebook)
```


**ONProcess**

ONProcess is a thin python wrapper around the COM interfaces for the [OneNote API](https://msdn.microsoft.com/en-us/library/office/jj680118\(v=office.15\).aspx).
It simplifies starting up the process, choosing the right process when multiple
versions are available and provides more pythonic interfaces for the OneNote
process.

You'll need to use ONProcess to do anything outside of reading content. Read the
source for onmanager.py for a list of available API calls.

For example, you can export onenote sections to PDF:
```python
import onepy
  
on = onepy.OneNote()
proc = on.process

def first_section_id():
  for notebook in on.hierarchy:
    for section in notebook:
      return section.id

proc.publish(first_section_id(), "C:\\Users\<account name>\Desktop\onepy-test.pdf", 3)

```



#### Common Errors

```
(Office 2013) This COM object can not automate the makepy process - please run makepy manually for this object
```

To work around this, run regedit.exe, and navigate to 
```
HKEY_CLASSES_ROOT\TypeLib\{0EA692EE-BB50-4E3C-AEF0-356D91732725}
```

There should only be one subfolder in this class called 1.1. If you see 1.0 or any other folders, you'll need to delete them. The final hierarchy should look like this: 

```
|- {0EA692EE-BB50-4E3C-AEF0-356D91732725}
|     |- 1.1
|         |-0
|         | |- win32
|         |- FLAGDS
|         |- HELPDIR
```

Source: [Stack Overflow](http://stackoverflow.com/questions/16287432/python-pywin-onenote-com-onenote-application-15-cannot-automate-the-makepy-p)

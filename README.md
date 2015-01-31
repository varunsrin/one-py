onepy
=====

COM Object Model for OneNote 2013 in Python




#### What are the requirements for onepy?

* Windows 7 with Python 3.x
* OneNote 2013 or 2010 with your notebooks open


### How do I setup my environment?

* Install Python 3.4 x86 from [here](https://www.python.org/download/releases/3.4.0/) 
* Install PyWin32 for Python 3.4 x86 from [here](http://sourceforge.net/projects/pywin32/files/pywin32/) 
* Add `C:\Python34\` to your PATH variable
* Run `C:\Python34\Lib\site-packages\win32com\client\makepy.py`
* Select `Microsoft OneNote 15.0 Extended Type Library`


### How do I build onepy?

* From the repo, run `python.exe setup.py register sdist bdist_wininst upload`


### How do I install onepy?

`pip install onepy`


### How do I use onepy?

Create a new file called `nb_printer.py` and type the following into it: 

    import onepy
  
    on = onepy.OneNote()
  
    # print a list of notebooks open in the OneNote 2013 client
    for notebook in on.hierarchy:
      print (notebook)

Save the file, and run `nb_printer.py` from the cmd prompt


### Common Errors

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


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
    for notebook in on.hierarchy
      print (notebook)

Save the file, and run `nb_printer.py` from the cmd prompt

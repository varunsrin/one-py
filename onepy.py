import json
import win32com.client
from xml.etree import ElementTree

onapp = win32com.client.gencache.EnsureDispatch('OneNote.Application')
NS = "{http://schemas.microsoft.com/office/onenote/2010/onenote}"



#Returns the Notebook Hierarchy as JSON
def getHierarchyJson():
    return(json.dumps(getHierarchy(), indent=4))



# Returns the Notebook Hierarchy as a Dictionary Array  
def getHierarchy():
    oneTree = ElementTree.fromstring(onapp.GetHierarchy("",win32com.client.constants.hsPages))

    notebooks = []

    for notebook in oneTree:
        nbk = parseAttributes(notebook)
        nbk['sections'], nbk['sectionGroups'], nbk['recycleBin'] = getSections(notebook)
        notebooks.append(nbk)

    return notebooks



# Takes in a Notebook or SectionGroup  and returns a Dict Array of its Sections & Section Groups
def getSections(notebook):
    sections = []
    sectionGroups = []
    recycleBin = ""
    for section in notebook:
        if (section.tag == NS + "SectionGroup"):               
            newSectionGroup = parseAttributes(section)
            newSectionGroup['sections'], newSectionGroup['sectionGroups'], newSectionGroup["recycleBin"] = getSections(section)
            if (section.get("isRecycleBin")):
               recycleBin = newSectionGroup
            else: sectionGroups.append(newSectionGroup)   
            
        if (section.tag == NS + "Section"):
            newSection = parseAttributes(section)
            newSection['pages'] = getPages(section)
            sections.append(newSection)
            
    return sections, sectionGroups, recycleBin



# Takes in a Section and returns a Dict Array of its Pages
def getPages(section):
     pages =[]
     for page in section:
         newPage = parseAttributes(page)
         newPage['meta'] = getMeta(page)
         pages.append(newPage)
     return pages




# Takes in a Page and returns a Dict Array of its Meta properties
def getMeta (page):
    metas = []
    for meta in page:
        metas.append(parseAttributes(meta))
    return metas



# Takes in an object and returns a dictionary of its values
def parseAttributes(obj):
        tempDict = {}
        for key,value in obj.items():
            tempDict[key] = value
        return tempDict







#Gets the Hierarchy as an Array of Python Dictionaries
notebooks = getHierarchy()

#print (notebooks[0]["sections"][1]["pages"][1]['ID'])

print (notebooks)



#xmlPage = onapp.GetPageContent("{05BC91ED-2B61-03AA-0BEB-24E475D27964}{1}{B0}")
#print (xmlPage)




#need to write something that takes a page, strips it of its XML contents and prints it


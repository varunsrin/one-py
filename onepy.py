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
        nbk['sections'], nbk['sectionGroups'] = getSections(notebook)
        notebooks.append(nbk)

    return notebooks



# Takes in a Notebook or SectionGroup  and returns a Dict Array of its Sections & Section Groups
def getSections(notebook):
    sections = []
    sectionGroups = []
    for section in notebook:
        if (section.tag == NS + "SectionGroup"):
            newSectionGroup = parseAttributes(section)
            newSectionGroup['sections'], newSectionGroup['sectionGroups'] = getSections(section)
            sectionGroups.append(newSectionGroup)   
            
        if (section.tag == NS + "Section"):
            newSection = parseAttributes(section)
            newSection['pages'] = getPages(section)
            sections.append(newSection)
            
    return sections, sectionGroups



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

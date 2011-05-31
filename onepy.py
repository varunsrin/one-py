import win32com.client
from xml.dom import minidom
import json

onapp = win32com.client.gencache.EnsureDispatch('OneNote.Application')
nb = onapp.GetHierarchy("",win32com.client.constants.hsPages)
onetree = minidom.parseString(nb)



#Returns the Notebook Hierarchy as JSON
def getHierarchyJson():
    return(json.dumps(getHierarchy(), indent=4))


#Returns the Notebook Hierarchy as a Dictionary
def getHierarchy():
    notebooks = []
    for node in onetree.getElementsByTagName("one:Notebook"):
        book = parseAttributes(node)
        if node.hasChildNodes():
            book['sections'], book['sectionGroups'] = parseChildren(node.childNodes)
        notebooks.append(book)
    return notebooks




#Recursively returns Dictionary of Section Groups & Sections
def parseChildren (nodelist):
    sectionGroups = []
    sections = []
    for node in nodelist:
        if (node.nodeName == 'one:SectionGroup'):
            group = parseAttributes(node)
            if node.hasChildNodes():
                group['sections'],group['sectionGroups'] = parseChildren(node.childNodes)
            sectionGroups.append(group)

        if (node.nodeName == 'one:Section'):
            sec = parseAttributes(node)
            if node.hasChildNodes():
                sec['pages'] = getPages(node.childNodes)
            sections.append(sec)

    return sections, sectionGroups

#Takes in a node and returns a dictionary of its attributes
def parseAttributes (node):
        tempDict = {}
        for prop in node.attributes.values():
            tempDict[prop.name] = prop.value
        return tempDict





#Returns Dictionary of Pages from a nodelist
def getPages(nodelist):
    pages = []
    for node in nodelist:
        page = parseAttributes(node)
        if node.hasChildNodes():
            page['meta'] = getMeta(node.childNodes)
        pages.append(page)
    return pages



def getMeta (nodelist):
    meta = []
    for node in nodelist:
        meta.append(parseAttributes(node))
    return meta





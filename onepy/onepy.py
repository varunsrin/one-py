import onmanager
from xml.etree import cElementTree


NS = ""

# 
class OneNote():

    def __init__(self):
        self.process = onmanager.ONProcess()
        global NS
        NS = self.process.NS
        self.object_tree = cElementTree.fromstring(self.process.GetHierarchy("",4))
        self.hierarchy = Hierarchy()
        self.hierarchy.deserialize_from_xml(self.object_tree)
        
    def get_page_content(self, page_id):
        page_content_xml = cElementTree.fromstring(self.process.GetPageContent(page_id))
        return PageContent(page_content_xml)
        
#
#  HIERARCHY
#

class Hierarchy():

    def __init__(self):
        self._children = []

    def deserialize_from_xml(self, xml):
        self._children = [Notebook(n) for n in xml]
                
    def __iter__(self):
        for c in self._children:
            yield c


#
# HIERARCHY NODE
#

class HierarchyNode():

    def _init_(self, parent=None):
        self.name = ""
        self.path = ""
        self.id = ""
        self.last_modified_time = ""
        self.synchronized = ""

    def deserialize_from_xml(self, xml):
        self.name = xml.get("name")
        self.path = xml.get("path")
        self.id = xml.get("ID")
        self.last_modified_time = xml.get("lastModifiedTime")


#
#   NOTEBOOK CLASS
#

class Notebook(HierarchyNode):

    def __init__ (self, xml=None):
        HierarchyNode.__init__(self)
        self.nickname = ""
        self.color = ""
        self.is_currently_viewed = ""
        self.recycleBin = None
        self._children = []
        if (xml != None):
            self.deserialize_from_xml(xml)

    def deserialize_from_xml(self, xml):
        HierarchyNode.deserialize_from_xml(self, xml)
        self.nickname = xml.get("nickname")
        self.color = xml.get("color")
        self.is_currently_viewed = xml.get("isCurrentlyViewed")
        self.recycleBin = None
        for node in xml:
            if (node.tag == NS + "Section"):
                self._children.append(Section(node, self)) 

            elif (node.tag == NS + "SectionGroup"):
                if(node.get("isRecycleBin")):
                    self.recycleBin = SectionGroup(node, self)
                else:
                    self._children.append(SectionGroup(node, self))

    def __iter__(self):
        for c in self._children:
            yield c

    def __str__(self):
        return self.name 



#
# SECTION GROUP CLASS
#

class SectionGroup(HierarchyNode):

    def __init__ (self, xml=None, parent_node=None):
        HierarchyNode.__init__(self)
        self.is_recycle_Bin = False
        self._children = []
        self.parent = parent_node
        if (xml != None):
            self.deserialize_from_xml(xml)

    def __iter__(self):
        for c in self._children:
            yield c
    
    def __str__(self):
        return self.name 

    def deserialize_from_xml(self, xml):
        HierarchyNode.deserialize_from_xml(self, xml)
        self.is_recycle_Bin = xml.get("isRecycleBin")
        for node in xml:
            if (node.tag == NS + "SectionGroup"):
                self._children.append(SectionGroup(node, self))
            if (node.tag == NS + "Section"):
                self._children.append(Section(node, self))


#
# SECTION CLASS
#

class Section(HierarchyNode):
       
    def __init__ (self, xml=None, parent_node=None):
        HierarchyNode.__init__(self)
        self.color = ""
        self.read_only = False
        self.is_currently_viewed = False      
        self._children = []
        self.parent = parent_node
        if (xml != None):
            self.deserialize_from_xml(xml)


    def __iter__(self):
        for c in self._children:
            yield c
    
    def __str__(self):
        return self.name


    def deserialize_from_xml(self, xml):
        HierarchyNode.deserialize_from_xml(self, xml)
        self.color = xml.get("color")
        try:
            self.read_only = xml.get("readOnly")
        except:
            self.read_only = False
        try:
            self.is_currently_viewed = xml.get("isCurrentlyViewed")      
        except:
            self.is_currently_viewed = False

        self._children = [Page(node, self) for node in xml]



#
#  PAGE CLASS
#

class Page():
    
    def __init__ (self, xml=None, parent_node=None):
        self.name = ""
        self.id = ""
        self.date_time = ""
        self.last_modified_time = ""
        self.page_level = ""
        self.is_currently_viewed = ""
        self._children = []
        self.parent = parent_node
        if (xml != None):
            self.deserialize_from_xml(xml)

    def __iter__(self):
        for c in self._children:
            yield c

    def __str__(self):
        return self.name 

    # Get / Set Meta

    def deserialize_from_xml (self, xml):
        self.name = xml.get("name")
        self.id = xml.get("ID")
        self.date_time = xml.get("dateTime")
        self.last_modified_time = xml.get("lastModifiedTime")
        self.page_level = xml.get("pageLevel")
        self.is_currently_viewed = xml.get("isCurrentlyViewed")
        self._children = [Meta(node) for node in xml]



#
# META CLASS
#

class Meta():
    
    def __init__ (self, xml = None):
        self.name = ""
        self.content = ""
        if (xml!=None):
            self.deserialize_from_xml(xml)

    def __str__(self):
        return self.name 

    def deserialize_from_xml (self, xml):
        self.name = xml.get("name")
        self.id = xml.get("content")


#
# PAGE CONTENT CLASS
#

class PageContent():

    def __init__ (self, xml=None):
        self.name = ""
        self.id = ""
        self.date_time = ""
        self.last_modified_time = ""
        self.page_level = ""
        self.lang = ""
        self.is_currently_viewed = ""
        self._children= []
        self.files = []
        if (xml != None):
            self.deserialize_from_xml(xml)

    def __iter__(self):
        for c in self._children:
            yield c
    
    def __str__(self):
        return self.name 

    def deserialize_from_xml(self, xml):
            self.name = xml.get("name")
            self.id = xml.get("ID")
            self.date_time = xml.get("dateTime")
            self.last_modified_time = xml.get("lastModifiedTime")
            self.page_level = xml.get("pageLevel")
            self.lang = xml.get("lang")
            self.is_currently_viewed = xml.get("isCurrentlyViewed")
            for node in xml:
                if (node.tag == NS + "Outline"):
                   self._children.append(Outline(node))
                elif (node.tag == NS + "Ink"):
                    self.files.append(Ink(node))
                elif (node.tag == NS + "Image"):
                    self.files.append(Image(node))
                elif (node.tag == NS + "InsertedFile"):
                    self.files.append(InsertedFile(node))       
                elif (node.tag == NS + "Title"):
                    self._children.append(Title(node))       

#
# TITLE CLASS
#

class Title():

    def __init__ (self, xml=None):
        self.style = ""
        self.lang = ""
        self._children = []

    def __str__ (self):
        return "Page Title"

    def __iter__ (self):
        for c in self._children:
            yield c

    def deserialize_from_xml(self, xml):
        self.style = xml.get("style")
        self.lang = xml.get("lang")
        for node in xml:
            if (node.tag == NS + "OE"):
                self._children.append(OE(node, self))



#
# OUTLINE CLASS
#

class Outline():

    def __init__ (self, xml=None):
        self.author = ""
        self.author_initials = ""
        self.last_modified_by = ""
        self.last_modified_by_initials = ""
        self.last_modified_time = ""
        self.id = ""
        self._children = []
        if(xml != None):
            self.deserialize_from_xml(xml)

    def __iter__(self):
        for c in self._children:
            yield c

    def __str__(self):
        return "Outline"

    def deserialize_from_xml (self, xml):     
        self.author = xml.get("author")
        self.author_initials = xml.get("authorInitials")
        self.last_modified_by = xml.get("lastModifiedBy")
        self.last_modified_by_initials = xml.get("lastModifiedByInitials")
        self.last_modified_time = xml.get("lastModifiedTime")
        self.id = xml.get("objectID")
        append = self._children.append
        for node in xml:
            if (node.tag == NS + "OEChildren"):
                for childNode in node:
                    if (childNode.tag == NS + "OE"):
                        append(OE(childNode, self))     


#
# POSITION CLASS
#

class Position():

    def __init__ (self, xml=None, parent_node=None):
        self.x = ""
        self.y = ""
        self.z = ""
        self.parent = parent_node
        if (xml!=None):
            self.deserialize_from_xml(xml)

    def deserialize_from_xml(self, xml):
        self.x = xml.get("x")
        self.y = xml.get("y")
        self.z = xml.get("z")

#
# SIZE CLASS
#

class Size():

    def __init__ (self, xml=None, parent_node=None):
        self.width = ""
        self.height = ""
        self.parent = parent_node
        if (xml!=None):
            self.deserialize_from_xml(xml)

    def deserialize_from_xml(self, xml):
        self.width = xml.get("width")
        self.height = xml.get("height")



#
# OE CLASS
#

class OE():

    def __init__ (self, xml=None, parent_node=None):
        
        self.creation_time = ""
        self.last_modified_time = ""
        self.last_modified_by = ""
        self.id = ""
        self.alignment = ""
        self.quick_style_index = ""
        self.style = ""
        self.text = ""
        self._children = []
        self.parent = parent_node
        self.files = []
        if (xml != None):
            self.deserialize_from_xml(xml)

    def __iter__(self):
        for c in self._children:
            yield c
    
    def __str__(self):
        try:
            return self.text
        except AttributeError:
            return "Empty OE"

    def deserialize_from_xml(self, xml):
        self.creation_time = xml.get("creationTime")
        self.last_modified_time = xml.get("lastModifiedTime")
        self.last_modified_by = xml.get("lastModifiedBy")
        self.id = xml.get("objectID")
        self.alignment = xml.get("alignment")
        self.quick_style_index = xml.get("quickStyleIndex")
        self.style = xml.get("style")

        for node in xml:
            if (node.tag == NS + "T"):
                if (node.text != None):
                    self.text = node.text
                else:
                    self.text = "NO TEXT"

            elif (node.tag == NS + "OEChildren"):
                for childNode in node:
                    if (childNode.tag == NS + "OE"):
                        self._children.append(OE(childNode, self))

            elif (node.tag == NS + "Image"):
                self.files.append(Image(node, self))

            elif (node.tag == NS + "InkWord"):
                self.files.append(Ink(node, self))

            elif (node.tag == NS + "InsertedFile"):
                self.files.append(InsertedFile(node, self))
      



#
# INSERTED FILE CLASS
#

class InsertedFile():

    # need to add position data to this class

    def __init__ (self, xml=None, parent_node=None):
        self.path_cache = ""
        self.path_source = ""
        self.preferred_name = ""
        self.last_modified_time = ""
        self.last_modified_by = ""
        self.id = ""
        self.parent = parent_node
        if (xml != None):
            self.deserialize_from_xml(xml)

    def _iter_ (self):
        yield None
    
    def __str__(self):
        try:
            return self.preferredName
        except AttributeError:
            return "Unnamed File"

    def deserialize_from_xml(self, xml):
        self.path_cache = xml.get("pathCache")
        self.path_source = xml.get("pathSource")
        self.preferred_name = xml.get("preferredName")
        self.last_modified_time = xml.get("lastModifiedTime")
        self.last_modified_by = xml.get("lastModifiedBy")
        self.id = xml.get("objectID")   


#
# INK CLASS
#


class Ink():

    # need to add position data to this class

    def __init__ (self, xml=None, parent_node=None):   
        self.recognized_text = ""
        self.x = ""
        self.y = ""
        self.ink_origin_x = ""
        self.ink_origin_y = ""
        self.width = ""
        self.height = ""
        self.data = ""
        self.callback_id = ""
        self.parent = parent_node

        if (xml != None):
            self.deserialize_from_xml(xml)

    def _iter_ (self):
        yield None
    
    def __str__(self):
        try:
            return self.recognizedText
        except AttributeError:
            return "Unrecognized Ink"

    def deserialize_from_xml(self, xml):
        self.recognized_text = xml.get("recognizedText")
        self.x = xml.get("x")
        self.y = xml.get("y")
        self.ink_origin_x = xml.get("inkOriginX")
        self.ink_origin_y = xml.get("inkOriginY")
        self.width = xml.get("width")
        self.height = xml.get("height")
            
        for node in xml:
            if (node.tag == NS + "CallbackID"):
                self.callback_id = node.get("callbackID")
            elif (node.tag == NS + "Data"):
                self.data = node.text
                    
    

#
#  IMAGE CLASS
#


class Image():

    def __init__ (self, xml=None, parent_node=None):    
        self.format = ""
        self.original_page_number = ""
        self.last_modified_time = ""
        self.id = ""
        self.callback_id = None
        self.data = ""
        self.parent = parent_node
        if (xml != None):
            self.deserialize_from_xml(xml)

    def _iter_ (self):
        yield None
    
    def __str__(self):
        return self.format + " Image"

    def deserialize_from_xml(self, xml):
        self.format = xml.get("format")
        self.original_page_number = xml.get("originalPageNumber")
        self.last_modified_time = xml.get("lastModifiedTime")
        self.id = xml.get("objectID")
        for node in xml:
            if (node.tag == NS + "CallbackID"):
                self.callback_id = node.get("callbackID")
            elif (node.tag == NS + "Data"):
                if (node.text != None):
                    self.data = node.text
                
import win32com.client

if win32com.client.gencache.is_readonly == True:
    win32com.client.gencache.is_readonly = False
    win32com.client.gencache.Rebuild()

#
# OnePy
# Provides pythonic wrappers around OneNote COM interfaces
#

class ONProcess():

    def __init__(self):
        try:
            self.process = win32com.client.gencache.EnsureDispatch('OneNote.Application.15')
            self.NS = "{http://schemas.microsoft.com/office/onenote/2013/onenote}"
        except Exception as e:
            print (e)
            print("error starting OneNote 15")
            print("trying OneNote 14 instead")
            try:
                self.process = win32com.client.gencache.EnsureDispatch('OneNote.Application.14')
                self.NS = "{http://schemas.microsoft.com/office/onenote/2010/onenote}"
            except Exception as e: 
                print (e)
                print ("error starting OneNote 14")


    def GetHierarchy(self, StartNodeID="", HierarchyScope=4):
        # HierarchyScope
        # 0 - Gets just the start node specified and no descendants.
        # 1 - Gets the immediate child nodes of the start node, and no descendants in higher or lower subsection groups.
        # 2 - Gets all notebooks below the start node, or root.
        # 3 - Gets all sections below the start node, including sections in section groups and subsection groups.
        # 4 - Gets all pages below the start node, including all pages in section groups and subsection groups.
        return (self.process.GetHierarchy(StartNodeID, HierarchyScope))

    def UpdateHierarchy(self, ChangesXMLIn):
        try:
            self.process.UpdateHierarchy(ChangesXMLIn)
        except:
            print("Could not Update Hierarchy")

    def OpenHierarchy(self, Path, RelativeToObjectID, ObjectID, CreateFileType=0):
        # CreateFileType
        # 0 - Creates no new object.
        # 1 - Creates a notebook with the specified name at the specified location.
        # 2 - Creates a section group with the specified name at the specified location.
        # 3 - Creates a section with the specified name at the specified location.
        try:
            return(self.process.OpenHierarchy(Path, RelativeToObjectID, "", CreateFileType))
        except:
            print("Could not Open Hierarchy")


    def DeleteHierarchy (self, ObjectID, ExpectedLastModified=""):
        try:
            self.process.DeleteHierarchy(ObjectID, ExpectedLastModified)
        except:
            print("Could not Delete Hierarchy")

    def CreateNewPage (self, SectionID, NewPageStyle=0):
        # NewPageStyle
        # 0 - Create a Page that has Default Page Style
        # 1 - Create a blank page with no title
        # 2 - Createa blank page that has no title
        try:
            self.process.CreateNewPage(SectionID, "", NewPageStyle)
        except:
            print("Unable to create the page")
            
    def CloseNotebook(self, NotebookID):
        try:
            self.process.CloseNotebook(NotebookID)
        except:
            print("Could not Close Notebook")

    def GetPageContent(self, PageID, PageInfo=0):
        # PageInfo
        # 0 - Returns only basic page content, without selection markup and binary data objects. This is the standard value to pass.
        # 1 - Returns page content with no selection markup, but with all binary data.
        # 2 - Returns page content with selection markup, but no binary data.
        # 3 - Returns page content with selection markup and all binary data.
        try:
            return(self.process.GetPageContent(PageID, "", PageInfo))
        except:
            print("Could not get Page Content")
            
    def UpdatePageContent(self, PageChangesXMLIn, ExpectedLastModified=0):
        try:
            self.process.UpdatePageContent(PageChangesXMLIn, ExpectedLastModified)
        except:
            print("Could not Update Page Content")
            
    def GetBinaryPageContent(self, PageID, CallbackID):
        try:
            return(self.process.GetBinaryPageContent(PageID, CallbackID))
        except:
            print("Could not Get Binary Page Content")



    def DeletePageContent(self, PageID, ObjectID, ExpectedLastModified=0):
        try:
            self.process.DeletePageContent(PageID, ObjectID, ExpectedLastModified)
        except:
            print("Could not Delete Page Content")


    # Actions


    def NavigateTo(self, ObjectID, NewWindow=False):
        try:
            self.process.NavigateTo(ObjectID, "", NewWindow)
        except:
            print("Could not Navigate To")

    def Publish(self, HierarchyID, TargetFilePath, PublishFormat, CLSIDofExporter=0):
        #PublishFormat
        # 0 - Published page is in .one format.
        # 1 - Published page is in .onea format.
        # 2 - Published page is in .mht format.
        # 3 - Published page is in .pdf format.
        # 4 - Published page is in .xps format.
        # 5 - Published page is in .doc or .docx format.
        # 6 - Published page is in enhanced metafile (.emf) format.
        try:
            self.process.Publish(HierarchyID, TargetFilePath, PublishFormat, CLSIDofExporter)
        except:
            print("Could not Publish")

    def OpenPackage(self, PathPackage, PathDest):
        try:
            return(self.process.OpenPackage(PathPackage, PathDest))
        except:
            print("Could not Open Package")

    def GetHyperlinkToObject(self, HierarchyID, TargetFilePath=""):
        try:
            return(self.process.GetHyperlinkToObject(HierarchyID, TargetFilePath))
        except:
            print("Could not Get Hyperlink")

    def FindPages(self, StartNodeID, SearchString, Display):
        try:
            return(self.process.FindPages(StartNodeID, SearchString, "", False, Display))
        except:
            print("Could not Find Pages")

    def GetSpecialLocation(self, SpecialLocation=0):
        # SpecialLocation
        # 0 - Gets the path to the Backup Folders folder location.
        # 1 - Gets the path to the Unfiled Notes folder location.
        # 2 - Gets the path to the Default Notebook folder location.
        try:
            return(self.process.GetSpecialLocation(SpecialLocation))
        except:
            print("Could not retreive special location")
    
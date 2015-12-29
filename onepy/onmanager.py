import win32com.client

if win32com.client.gencache.is_readonly == True:
    win32com.client.gencache.is_readonly = False
    win32com.client.gencache.Rebuild()

"""
  OnePy
  Provides pythonic wrappers around OneNote COM interfaces
"""

ON15_APP_ID = 'OneNote.Application.15'
ON15_SCHEMA = "{http://schemas.microsoft.com/office/onenote/2013/onenote}"
ON14_APP_ID = 'OneNote.Application.14'
ON14_SCHEMA = "{http://schemas.microsoft.com/office/onenote/2010/onenote}"


class ONProcess():

    def __init__(self, version=15):

        try: 
            if (version == 15):
                self.process = win32com.client.gencache.EnsureDispatch(ON15_APP_ID)
                self.namespace = ON15_SCHEMA
            if (version == 14):
                self.process = win32com.client.gencache.EnsureDispatch(ON14_APP_ID)
                self.namespace = ON14_SCHEMA         
        except Exception as e:
            print (e)
            print("error starting onenote {}".format(version))


    def get_hierarchy(self, start_node_id="", hierarchy_scope=4):
        """
          HierarchyScope
          0 - Gets just the start node specified and no descendants.
          1 - Gets the immediate child nodes of the start node, and no descendants in higher or lower subsection groups.
          2 - Gets all notebooks below the start node, or root.
          3 - Gets all sections below the start node, including sections in section groups and subsection groups.
          4 - Gets all pages below the start node, including all pages in section groups and subsection groups.
        """
        return (self.process.GetHierarchy(start_node_id, hierarchy_scope))

    def update_hierarchy(self, changes_xml_in):
        try:
            self.process.UpdateHierarchy(changes_xml_in)
        except Exception as e:
            print(e) 
            print("Could not Update Hierarchy")

    def open_hierarchy(self, path, relative_to_object_id, object_id, create_file_type=0):
        """
          CreateFileType
          0 - Creates no new object.
          1 - Creates a notebook with the specified name at the specified location.
          2 - Creates a section group with the specified name at the specified location.
          3 - Creates a section with the specified name at the specified location.
        """
        try:
            return(self.process.OpenHierarchy(path, relative_to_object_id, "", create_file_type))
        except Exception as e: 
            print(e)
            print("Could not Open Hierarchy")


    def delete_hierarchy (self, object_id, excpect_last_modified=""):
        try:
            self.process.DeleteHierarchy(object_id, excpect_last_modified)
        except Exception as e: 
            print(e)
            print("Could not Delete Hierarchy")

    def create_new_page (self, section_id, new_page_style=0):
        """
          NewPageStyle
          0 - Create a Page that has Default Page Style
          1 - Create a blank page with no title
          2 - Createa blank page that has no title
        """
        try:
            self.process.CreateNewPage(section_id, "", new_page_style)
        except Exception as e: 
            print(e)
            print("Unable to create the page")
            
    def close_notebook(self, notebook_id):
        try:
            self.process.CloseNotebook(notebook_id)
        except Exception as e: 
            print(e)
            print("Could not Close Notebook")

    def get_page_content(self, page_id, page_info=0):
        """
          PageInfo
          0 - Returns only basic page content, without selection markup and binary data objects. This is the standard value to pass.
          1 - Returns page content with no selection markup, but with all binary data.
          2 - Returns page content with selection markup, but no binary data.
          3 - Returns page content with selection markup and all binary data.
        """
        try:
            return(self.process.GetPageContent(page_id, "", page_info))
        except Exception as e: 
            print(e)
            print("Could not get Page Content")
            
    def update_page_content(self, page_changes_xml_in, excpect_last_modified=0):
        try:
            self.process.UpdatePageContent(page_changes_xml_in, excpect_last_modified)
        except Exception as e: 
            print(e)
            print("Could not Update Page Content")
            
    def get_binary_page_content(self, page_id, callback_id):
        try:
            return(self.process.GetBinaryPageContent(page_id, callback_id))
        except Exception as e: 
            print(e)
            print("Could not Get Binary Page Content")

    def delete_page_content(self, page_id, object_id, excpect_last_modified=0):
        try:
            self.process.DeletePageContent(page_id, object_id, excpect_last_modified)
        except Exception as e: 
            print(e)
            print("Could not Delete Page Content")


      # Actions


    def navigate_to(self, object_id, new_window=False):
        try:
            self.process.NavigateTo(object_id, "", new_window)
        except Exception as e: 
            print(e)
            print("Could not Navigate To")

    def publish(self, hierarchy_id, target_file_path, publish_format, clsid_of_exporter=""):
        """
         PublishFormat
          0 - Published page is in .one format.
          1 - Published page is in .onea format.
          2 - Published page is in .mht format.
          3 - Published page is in .pdf format.
          4 - Published page is in .xps format.
          5 - Published page is in .doc or .docx format.
          6 - Published page is in enhanced metafile (.emf) format.
        """
        try:
            self.process.Publish(hierarchy_id, target_file_path, publish_format, clsid_of_exporter)
        except Exception as e: 
            print(e)
            print("Could not Publish")

    def open_package(self, path_package, path_dest):
        try:
            return(self.process.OpenPackage(path_package, path_dest))
        except Exception as e: 
            print(e)
            print("Could not Open Package")

    def get_hyperlink_to_object(self, hierarchy_id, target_file_path=""):
        try:
            return(self.process.GetHyperlinkToObject(hierarchy_id, target_file_path))
        except Exception as e: 
            print(e)
            print("Could not Get Hyperlink")

    def find_pages(self, start_node_id, search_string, display):
        try:
            return(self.process.FindPages(start_node_id, search_string, "", False, display))
        except Exception as e: 
            print(e)
            print("Could not Find Pages")

    def get_special_location(self, special_location=0):
        """
          SpecialLocation
          0 - Gets the path to the Backup Folders folder location.
          1 - Gets the path to the Unfiled Notes folder location.
          2 - Gets the path to the Default Notebook folder location.
        """
        try:
            return(self.process.GetSpecialLocation(special_location))
        except Exception as e: 
            print(e)
            print("Could not retreive special location")
    
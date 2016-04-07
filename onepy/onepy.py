from .onmanager import ONProcess
import lxml.etree as ET
from lxml.builder import ElementMaker
import time
import re

__all__ = ["OneNote", "PageEditor"]

namespace = ""

class OneNote():
    def __init__(self, version=14):
        self.process = ONProcess(version=version)
        global namespace
        namespace = self.process.namespace
        self.object_tree = ET.fromstring(self.process.get_hierarchy("",4))
        self.hierarchy = Hierarchy(self.object_tree)
        
    def get_page_content(self, page_id, page_info=0):
        page_content_xml = ET.fromstring(self.process.get_page_content(page_id, page_info))
        return PageContent(page_content_xml)

class PageEditor():
    def __init__(self, version=14):
        self._process = ONProcess(version=version)
        self._namespace = self._process.namespace
        self._page = None
        #ET.register_namespace("one", self._namespace)
        self._xml = None
        self._title = None
        self._flat_contents = []

    def create(self, section, title, lines=None):
        now = time.time()
        self._process.create_new_page(section.id)
        #get the newly created page, regenerate hierarchy
        refresh_section_xml = self._process.get_hierarchy(section.id,4)
        refresh_section = Section(ET.fromstring(refresh_section_xml))
        assert refresh_section.id == section.id
        self.open(refresh_section[-1])
        creation = time.mktime(time.strptime(self._page.date_time, "%Y-%m-%dT%H:%M:%S.%fZ"))
        if creation - now > 5:
            raise Exception("Page is too old to be freshly created, something went wrong?")
            
        self.update_title(title)
        
        if lines:
            self.add_lines(lines)

    def find_in_xml(self, patterns):
        xml = self._rawxml
        found = []
        if not isinstance(patterns, (list, tuple)):
            patterns = [patterns]
        for p in patterns:
            found.extend(re.findall(p, xml))
        return found

    def overwrite_content(self, newxml):
        self._process.update_page_content(newxml)
        
    def replace_in_xml(self, originals, replacements, dry_run=True, confirm=True):
        skipped = []
        applied = []
        xml = self._rawxml
        for orig, rep in zip(originals, replacements):
            if confirm:
                ok = input('\n\nReplace:\n{}\n\nwith:\n{}\n?[n for No]'.format(orig.replace('\r\n', ' '), rep))
            else:
                ok = 'Y'
            if ok.upper() == 'N':
                skipped.append(orig.replace('\r\n', ' '))
                continue
            else:
                applied.append((orig.replace('\r\n', ' '), rep))
            xml = re.sub(re.escape(orig), rep, xml)
        if not dry_run:
            self._process.update_page_content(b'<?xml version="1.0"?>\n' + ET.tostring(ET.fromstring(xml)))
        else:
            print('Dry run for page {} (changes not applied)'.format(self._page.name))

        return applied, skipped

        
    def add_lines(self, lines):
        maker = ElementMaker(namespace="one")
        root = maker.root()
        root.append(maker.Outline())
        root[0].append(maker.OEChildren())
        
        for index, line in enumerate(lines):
            root[0][0].append(maker.OE())
            root[0][0][index].append(maker.T(ET.CDATA(line)))        
        
        ns = {"one":"http://schemas.microsoft.com/office/onenote/2010/onenote"}
        page = ET.fromstring(self._process.get_page_content(self._page.id))
        oechildren = page.find(".//one:Outline/one:OEChildren", ns)        
        if oechildren is not None:
            for oe in root[0][0]:
                oechildren.append(oe)        
        else:
            page.append(root[0])
                 
        self._process.update_page_content( #bit of a HACK to clean the dammed "ns0"
            ET.tostring(page).replace(b"ns0:",b"one:").replace(b'xmlns:one="one"', b""))
                
        self._flatten()

    def open(self, page):
        self._page = page
        self._flatten()

    def _flatten(self):
        """Expose each line without the xml nesting"""
        self._rawxml = self._process.get_page_content(self._page.id)
        self._xml = ET.fromstring(self._rawxml)
        flat = list(self._xml.iter(self._namespace+'T'))
        try:
            self._title = flat[0]
            self._flat_contents = flat[1:]
        except IndexError:
            self._title=''
            self._flat_contents = []

    def _push(self):
        self._process.update_page_content(b'<?xml version="1.0"?>\n' + ET.tostring(self._xml))
        #refresh
        self._flatten()        

    def print(self):
        print("Title: {}".format(self._title.text))
        print("\n".join(node.text if node.text else "" for node in self._flat_contents))

    def get_lines(self, start=0, end=None):
        if end is None:
            end = len(self._flat_contents)
            
        return [node.text if node.text else "" for node in self._flat_contents[start:end]]

    def update_title(self, newtitle):
        self._title.text = newtitle
        self._push()

    def update_lines(self, lines, start=0):
        """Modify the content of lines[start:end]""" 
        for xml_line, newline in zip(self._flat_contents[start:], lines):
            xml_line.text = newline
        
        self._push()
    
    def format_lines(self, linenumbers, key, value):
        if not isinstance(linenumbers, list):
            linenumbers=[linenumbers]
        for n in linenumbers:
            self._flat_contents[n].set(key, value)
            
        self._push()
        
    
class Hierarchy():

    def __init__(self, xml=None):
        self._children = []
        if (xml != None): 
            self.__deserialize_from_xml(xml)

    def __deserialize_from_xml(self, xml):
        self._children = [Notebook(n) for n in xml]
                
    def __iter__(self):
        yield from self._children
            
    def __getitem__(self, key):
        return self._children[key]
        
    def __len__(self):
        return len(self._children)
   
class Node():
    def __init__(self):
        self.name = ""
        self._children = []

    def __str__(self):
        if self.name:
            return self.name 
        else:
            return "NO_NAME"

    def __repr__(self):
        return object.__repr__(self).rstrip(">") + " " + str(self.name) + ">"

    def __getitem__(self, key):
        return self._children[key]            

    def __iter__(self):
        yield from self._children
     
    def __len__(self):
        return len(self._children)
    

class HierarchyNode(Node):

    def __init__(self, parent=None):
        super().__init__()
        self.path = ""
        self.id = ""
        self.last_modified_time = ""
        self.synchronized = ""

    def deserialize_from_xml(self, xml):
        self._xml = xml
        self.name = xml.get("name")
        self.path = xml.get("path")
        self.id = xml.get("ID")
        self.last_modified_time = xml.get("lastModifiedTime")
              

class Notebook(HierarchyNode):

    def __init__ (self, xml=None):
        super().__init__()
        self.nickname = ""
        self.color = ""
        self.is_currently_viewed = ""
        self.recycleBin = None
        self._children = []
        if (xml != None):
            self.__deserialize_from_xml(xml)

    def __deserialize_from_xml(self, xml):
        HierarchyNode.deserialize_from_xml(self, xml)
        self.nickname = xml.get("nickname")
        self.color = xml.get("color")
        self.is_currently_viewed = xml.get("isCurrentlyViewed")
        self.recycleBin = None
        for node in xml:
            if (node.tag == namespace + "Section"):
                self._children.append(Section(node, self)) 

            elif (node.tag == namespace + "SectionGroup"):
                if(node.get("isRecycleBin")):
                    self.recycleBin = SectionGroup(node, self)
                else:
                    self._children.append(SectionGroup(node, self))


class SectionGroup(HierarchyNode):

    def __init__ (self, xml=None, parent_node=None):
        super().__init__()
        self.is_recycle_Bin = False
        self._children = []
        self.parent = parent_node
        if (xml != None):
            self.__deserialize_from_xml(xml)

    def __deserialize_from_xml(self, xml):
        HierarchyNode.deserialize_from_xml(self, xml)
        self.is_recycle_Bin = xml.get("isRecycleBin")
        for node in xml:
            if (node.tag == namespace + "SectionGroup"):
                self._children.append(SectionGroup(node, self))
            if (node.tag == namespace + "Section"):
                self._children.append(Section(node, self))


class Section(HierarchyNode):
       
    def __init__ (self, xml=None, parent_node=None):
        super().__init__()
        self.color = ""
        self.read_only = False
        self.is_currently_viewed = False      
        self._children = []
        self.parent = parent_node
        if (xml != None):
            self.__deserialize_from_xml(xml)

    def __deserialize_from_xml(self, xml):
        HierarchyNode.deserialize_from_xml(self, xml)
        self.color = xml.get("color")
        try:
            self.read_only = xml.get("readOnly")
        except Exception as e:
            self.read_only = False
        try:
            self.is_currently_viewed = xml.get("isCurrentlyViewed")      
        except Exception as e:
            self.is_currently_viewed = False

        self._children = [Page(xml=node, parent_node=self) for node in xml]


class Page(Node):
    
    def __init__ (self, xml=None, parent_node=None):
        super().__init__()
        self.id = ""
        self.date_time = ""
        self.last_modified_time = ""
        self.page_level = ""
        self.is_currently_viewed = ""
        self.parent = parent_node
        if (xml != None):                         # != None is required here, since this can return false
            self.__deserialize_from_xml(xml)


    # Get / Set Meta

    def __deserialize_from_xml (self, xml):
        self._xml = xml
        self.name = xml.get("name")
        self.id = xml.get("ID")
        self.date_time = xml.get("dateTime")
        self.last_modified_time = xml.get("lastModifiedTime")
        self.page_level = xml.get("pageLevel")
        self.is_currently_viewed = xml.get("isCurrentlyViewed")
        self._children = [Meta(xml=node) for node in xml]


class Meta():
    
    def __init__ (self, xml = None):
        self.name = ""
        self.content = ""
        if (xml!=None):
            self.__deserialize_from_xml(xml)

    def __str__(self):
        return self.name 

    def __deserialize_from_xml (self, xml):
        self._xml = xml
        self.name = xml.get("name")
        self.id = xml.get("content")


class PageContent(Node):

    def __init__ (self, xml=None):
        super().__init__()
        self.id = ""
        self.date_time = ""
        self.last_modified_time = ""
        self.page_level = ""
        self.lang = ""
        self.is_currently_viewed = ""
        self.files = []
        if (xml != None):
            self.__deserialize_from_xml(xml)
            self._xml = xml

    def __deserialize_from_xml(self, xml):
            self.name = xml.get("name")
            self.id = xml.get("ID")
            self.date_time = xml.get("dateTime")
            self.last_modified_time = xml.get("lastModifiedTime")
            self.page_level = xml.get("pageLevel")
            self.lang = xml.get("lang")
            self.is_currently_viewed = xml.get("isCurrentlyViewed")
            for node in xml:
                if (node.tag == namespace + "Outline"):
                   self._children.append(Outline(node))
                elif (node.tag == namespace + "Ink"):
                    self.files.append(Ink(node))
                elif (node.tag == namespace + "Image"):
                    self.files.append(Image(node))
                elif (node.tag == namespace + "InsertedFile"):
                    self.files.append(InsertedFile(node))       
                elif (node.tag == namespace + "MediaFile"):
                    self.files.append(MediaFile(node, self))  
                elif (node.tag == namespace + "Title"):
                    self._children.append(Title(node))    
                elif (node.tag == namespace + "MediaPlaylist"):
                    self.media_playlist = MediaPlaylist(node, self)       
    

class Title(Node):

    def __init__ (self, xml=None):
        super().__init__()
        self.style = ""
        self.lang = ""
        if (xml != None):
            self.__deserialize_from_xml(xml)

    def __str__ (self):
        return "Page Title"

    def __deserialize_from_xml(self, xml):
        self.style = xml.get("style")
        self.lang = xml.get("lang")
        for node in xml:
            if (node.tag == namespace + "OE"):
                self._children.append(OE(node, self))


class Outline(Node):

    def __init__ (self, xml=None):
        super().__init__()
        self.author = ""
        self.author_initials = ""
        self.last_modified_by = ""
        self.last_modified_by_initials = ""
        self.last_modified_time = ""
        self.id = ""
        if (xml != None):
            self.__deserialize_from_xml(xml)
            self._xml = xml

    def __str__(self):
        return "Outline"

    def __deserialize_from_xml (self, xml):     
        self.author = xml.get("author")
        self.author_initials = xml.get("authorInitials")
        self.last_modified_by = xml.get("lastModifiedBy")
        self.last_modified_by_initials = xml.get("lastModifiedByInitials")
        self.last_modified_time = xml.get("lastModifiedTime")
        self.id = xml.get("objectID")
        append = self._children.append
        for node in xml:
            if (node.tag == namespace + "OEChildren"):
                for childNode in node:
                    if (childNode.tag == namespace + "OE"):
                        append(OE(childNode, self))     


class Position():

    def __init__ (self, xml=None, parent_node=None):
        self.x = ""
        self.y = ""
        self.z = ""
        self.parent = parent_node
        if (xml!=None):
            self.__deserialize_from_xml(xml)

    def __deserialize_from_xml(self, xml):
        self.x = xml.get("x")
        self.y = xml.get("y")
        self.z = xml.get("z")


class Size():

    def __init__ (self, xml=None, parent_node=None):
        self.width = ""
        self.height = ""
        self.parent = parent_node
        if (xml!=None):
            self.__deserialize_from_xml(xml)

    def __deserialize_from_xml(self, xml):
        self.width = xml.get("width")
        self.height = xml.get("height")


class OE(Node):

    def __init__ (self, xml=None, parent_node=None):
        super().__init__()
        self.creation_time = ""
        self.last_modified_time = ""
        self.last_modified_by = ""
        self.id = ""
        self.alignment = ""
        self.quick_style_index = ""
        self.style = ""
        self.text = ""
        self.parent = parent_node
        self.files = []
        self.media_indices = []
        if (xml != None):
            self.__deserialize_from_xml(xml)
            self._xml = xml

    def __str__(self):
        try:
            return self.text
        except AttributeError:
            return "Empty OE"

    def __deserialize_from_xml(self, xml):
        self.creation_time = xml.get("creationTime")
        self.last_modified_time = xml.get("lastModifiedTime")
        self.last_modified_by = xml.get("lastModifiedBy")
        self.id = xml.get("objectID")
        self.alignment = xml.get("alignment")
        self.quick_style_index = xml.get("quickStyleIndex")
        self.style = xml.get("style")

        for node in xml:
            if (node.tag == namespace + "T"):
                if (node.text != None):
                    self.text = node.text
                else:
                    self.text = ""

            elif (node.tag == namespace + "OEChildren"):
                for childNode in node:
                    if (childNode.tag == namespace + "OE"):
                        self._children.append(OE(childNode, self))

            elif (node.tag == namespace + "Image"):
                self.files.append(Image(node, self))

            elif (node.tag == namespace + "InkWord"):
                self.files.append(Ink(node, self))

            elif (node.tag == namespace + "InsertedFile"):
                self.files.append(InsertedFile(node, self))

            elif (node.tag == namespace + "MediaFile"):
                self.files.append(MediaFile(node, self))
                
            elif (node.tag == namespace + "MediaIndex"):
                self.media_indices.append(MediaIndex(node, self))


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
            self.__deserialize_from_xml(xml)

    def __iter__ (self):
        yield None
    
    def __str__(self):
        try:
            return self.preferredName
        except AttributeError:
            return "Unnamed File"

    def __deserialize_from_xml(self, xml):
        self.path_cache = xml.get("pathCache")
        self.path_source = xml.get("pathSource")
        self.preferred_name = xml.get("preferredName")
        self.last_modified_time = xml.get("lastModifiedTime")
        self.last_modified_by = xml.get("lastModifiedBy")
        self.id = xml.get("objectID")   

  
class MediaReference():
    def __init__ (self, xml=None, parent_node=None):
        self.media_id = ""
        
    def __iter__ (self):
        yield None
    
    def __str__(self):
        return "Media Reference"

    def __deserialize_from_xml(self, xml):
        self.media_id = xml.get("mediaID")
        

class MediaPlaylist():
    def __init__ (self, xml=None, parent_node=None):
        self.media_references = []
        
    def __iter__(self):
        for c in self.media_references:
            yield c
    
    def __str__(self):
        return "Media Index"

    def __deserialize_from_xml(self, xml):
        for node in xml:
            if (node.tag == namespace + "MediaReference"):
                self.media_references.append(MediaReference(node, self))
        
        
class MediaIndex():
    def __init__ (self, xml=None, parent_node=None):
        self.media_reference = None
        self.time_index = 0
        
    def __iter__(self):
        yield None
    
    def __str__(self):
        return "Media Index"

    def __deserialize_from_xml(self, xml):
        self.time_index = xml.get("timeIndex")
        for node in xml:
            if (node.tag == namespace + "MediaReference"):
                self.media_reference = MediaReference(node, self)
                
  
class MediaFile(InsertedFile):
    def __init__ (self, xml=None, parent_node=None):
        self.media_reference = None
        super().__init__(xml, parent_node)
        
    def __iter__(self):
        yield None

    def __str__(self):
        try:
            return self.preferredName
        except AttributeError:
            return "Unnamed Media File"
            
    def __deserialize_from_xml(self, xml):
        super().__deserialize_from_xml(xml)
        for node in xml:
            if (node.tag == namespace + "MediaReference"):
                self.media_reference = MediaReference(node, self)
    
    
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
            self.__deserialize_from_xml(xml)

    def __iter__ (self):
        yield None
    
    def __str__(self):
        try:
            return self.recognizedText
        except AttributeError:
            return "Unrecognized Ink"

    def __deserialize_from_xml(self, xml):
        self.recognized_text = xml.get("recognizedText")
        self.x = xml.get("x")
        self.y = xml.get("y")
        self.ink_origin_x = xml.get("inkOriginX")
        self.ink_origin_y = xml.get("inkOriginY")
        self.width = xml.get("width")
        self.height = xml.get("height")
            
        for node in xml:
            if (node.tag == namespace + "CallbackID"):
                self.callback_id = node.get("callbackID")
            elif (node.tag == namespace + "Data"):
                self.data = node.text


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
            self.__deserialize_from_xml(xml)

    def __iter__ (self):
        yield None
    
    def __str__(self):
        return self.format + " Image"

    def __deserialize_from_xml(self, xml):
        self.format = xml.get("format")
        self.original_page_number = xml.get("originalPageNumber")
        self.last_modified_time = xml.get("lastModifiedTime")
        self.id = xml.get("objectID")
        for node in xml:
            if (node.tag == namespace + "CallbackID"):
                self.callback_id = node.get("callbackID")
            elif (node.tag == namespace + "Data"):
                if (node.text != None):
                    self.data = node.text
                

from io import BytesIO
from zipfile import ZipFile
import xml.etree.ElementTree as elemtree
import re

name_space_map = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


class DocFile:
    """Class for doc file manipulation"""

    def __init__(self, doc=None, url=None):
        self.doc = doc
        self.text = ""
        if not (doc or url):
            raise ValueError("Missing both doc and url parameter. "
                             "One of them is required.")
        if url:
            try:
                import requests
                self.doc = BytesIO(requests.get(url).content)
            except ModuleNotFoundError as e:
                print("Install requests module: pip install requests")
                raise e

        self.zip_doc = ZipFile(self.doc)
        self.file_list = self.zip_doc.namelist()

    @staticmethod
    def _qualified_name(tag):
        """
        Turns a namespace prefixed tag name into a Clark-notation
        qualified tag name for lxml.
        Source: https://github.com/python-openxml/python-docx/
        """
        prefix, root_tag = tag.split(':')
        return '{{{}}}{}'.format(name_space_map[prefix], root_tag)

    def _get_xml_document_type(self):
        return [i for i in self.file_list if re.match('word/document\d?.xml', i)]

    def _get_header_footer_text(self, type_name):
        """Gets the header text from header files found."""
        xmls = 'word/{}[0-9]*.xml'.format(type_name)
        for file_name in self.file_list:
            if re.match(xmls, file_name):
                self.text += self._xml_text(self.zip_doc.read(file_name))

    def _xml_text(self, xml):
        """Converts xml content to text."""
        text = ''
        root_xml = elemtree.fromstring(xml)
        for child_elem in root_xml.iter():
            if child_elem.tag == self._qualified_name('w:t'):
                text += child_elem.text if child_elem.text else ''
            elif child_elem.tag == self._qualified_name('w:tab'):
                text += '\t'
            elif child_elem.tag in (self._qualified_name('w:br'), self._qualified_name('w:cr')):
                text += '\n'
            elif child_elem.tag == self._qualified_name("w:p"):
                text += '\n\n'
        return text

    def get_text(self):
        # Header text
        self._get_header_footer_text(type_name="header")
        # Main text
        xml_doc_type = self._get_xml_document_type()
        if xml_doc_type:
            self.text += self._xml_text(self.zip_doc.read(xml_doc_type[0]))
        # Footer text
        self._get_header_footer_text(type_name="footer")
        self.zip_doc.close()
        return self.text

"""Zipped Xml"""
import zipfile

from lxml import etree


class ZipDocument(zipfile.ZipFile):
    """Base class for zipped-xml, like .docx and .odt"""

    def load_xml(self, path_in_zip: str) -> etree._Entity | None:
        """Parse an XML file inside the zipped doc, return root node."""
        try:
            with self.open(path_in_zip) as fobj:
                return etree.parse(fobj).getroot()
        except KeyError:
            return None

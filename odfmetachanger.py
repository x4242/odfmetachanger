"""odfmetachanger

__author__ = "0x4242"
__copright_ = "tbd"
__license__ = "tbd"
__date__ = "tbd"
__version = "tbd"
"""

import xml.etree.ElementTree as ET
import os
import shutil
import zipfile
import frontmatter

NAMESPACE = {"office" : "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
             "dc"     : "http://purl.org/dc/elements/1.1/",
             "meta"   : "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"}

ET.register_namespace("office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
ET.register_namespace("xlink", "http://www.w3.org/1999/xlink")
ET.register_namespace("dc", "http://purl.org/dc/elements/1.1/")
ET.register_namespace("meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")
ET.register_namespace("ooo", "http://openoffice.org/2004/office")
ET.register_namespace("grddl", "http://www.w3.org/2003/g/data-view#")


def alter_odf_meta_title(meta_xml_tree, new_title):
    odf_dc_title = meta_xml_tree.find(".//dc:title", NAMESPACE)
    odf_dc_title.text = new_title


def alter_odf_meta_user(meta_xml_tree, meta_name, meta_text, meta_type="string"):
    """
    <meta:user-defined meta:name="NAME" meta:value-type="TYPE">TEXT</meta:user-defined>
    """

    user_meta_elements = meta_xml_tree.findall(".//meta:user-defined", NAMESPACE)
    for element in user_meta_elements:
        # get("meta:name") doens't work, no idea why
        if element.get("{urn:oasis:names:tc:opendocument:xmlns:meta:1.0}name") == meta_name:
            element.text = meta_text
            return

    new_meta_element_attribs = {"meta:name" : meta_name, "meta:value-tape" : meta_type}
    new_meta_element = ET.SubElement(meta_xml_tree.getroot().find(".//office:meta", NAMESPACE), "meta:user-defined", new_meta_element_attribs)
    new_meta_element.text = meta_text


def read_odf_meta_data(filename):
    """Read and parse 'meta.xml' from an ODF/ODT file.

    Opens the ODF/ODT file as zip archive and opens the content of the file
    'meta.xml' which gets parsed and returned.

    Args:
        filename: the filename of the ODF/ODT file as string

    Returns:
        An ElementTree containg the parsed meta data.
    """

    odf_file = zipfile.ZipFile(filename, "r")
    meta_xml_tree = ET.parse(odf_file.open("meta.xml", "r"))
    odf_file.close()
    return meta_xml_tree


def create_new_odf_file(filename, new_meta_data):
    """Create a new ODF/ODT file with changed meta data

    Copies all content of the original file, except of 'meta.xml', to a new
    file. Adds a new 'meta.xml'. Backups/Deletes the original file and renames the new
    file.

    Args:
        filename: the filename
        new_meta_data: as string
    """

    odf_file_in = zipfile.ZipFile(filename, "r")

    tmp_filename = filename + ".tmp"
    odf_file_out = zipfile.ZipFile(tmp_filename, "w")

    for file in odf_file_in.infolist():
        # instead of copying 'meta.xml', write a new one with new meta data
        if file.filename == "meta.xml":
            new_meta_data.write("meta.xml.tmp", encoding="UTF-8", xml_declaration=True)
            tmp_xml_file = open("meta.xml.tmp")
            odf_file_out.writestr("meta.xml", tmp_xml_file.read())
            tmp_xml_file.close()
            os.remove("meta.xml.tmp")
        else:
            odf_file_out.writestr(file, odf_file_in.read(file.filename))

    odf_file_in.close()
    odf_file_out.close()

    # backup original file and rename new file
    os.rename(filename, filename + ".bak")
    os.rename(tmp_filename, filename)


def load_yaml_frontmatter(filename):
    """Load YAML frontmatter from markdown file.

    Args:
        filename: markdown filename

    Returns:
        A dict containg the frontmatter.
    """

    with open(filename, "r") as file:
        file_content = frontmatter.load(file)
        return file_content.metadata

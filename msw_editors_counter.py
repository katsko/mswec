from os import listdir
from pathlib import Path
from xml.dom import minidom
from zipfile import ZipFile

base_dir = Path('/tmp/docx_dir')


def read_xml(xml_file):
    xmldoc = minidom.parse(xml_file)
    tags = xmldoc.getElementsByTagName('cp:lastModifiedBy')
    tag = tags[0]
    last_modified_by = tag.firstChild.nodeValue
    return last_modified_by


for filename in listdir(base_dir):
    if filename.endswith('.docx'):
        with ZipFile(base_dir / filename) as zip_file:
            with zip_file.open('docProps/core.xml') as xml_file:
                last_modified_by = read_xml(xml_file)
                print(last_modified_by)

from os import listdir
from xml.dom import minidom
from zipfile import ZipFile


for filename in listdir('/tmp/docx_dir'):
    if filename.endswith('.docx'):
        with ZipFile(f'/tmp/docx_dir/{filename}') as zip_file:
            with zip_file.open('docProps/core.xml') as xml_file:
                xmldoc = minidom.parse(xml_file)
                tags = xmldoc.getElementsByTagName('cp:lastModifiedBy')
                tag = tags[0]
                last_modified_by = tag.firstChild.nodeValue
                print(last_modified_by)

from xml.dom import minidom
from zipfile import ZipFile


filename = '/tmp/docx_dir/file.docx'
with ZipFile(filename) as zip_file:
    with zip_file.open('docProps/core.xml') as xml_file:
        xmldoc = minidom.parse(xml_file)
        tags = xmldoc.getElementsByTagName('cp:lastModifiedBy')
        tag = tags[0]
        last_modified_by = tag.firstChild.nodeValue
        print(last_modified_by)

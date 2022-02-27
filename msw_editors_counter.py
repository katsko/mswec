from pathlib import Path
from xml.dom import minidom
from zipfile import ZipFile

base_dir = Path('/tmp/docx_dir')


def run():
    files = get_files(base_dir)
    editors = get_editors(files)
    print(editors)


def get_files(base_dir):
    result = []
    for filename in base_dir.iterdir():
        if filename.suffix == '.docx':
            result.append(base_dir / filename)
    return result


def get_editors(files):
    result = []
    for filename in files:
        with ZipFile(filename) as zip_file:
            with zip_file.open('docProps/core.xml') as xml_file:
                last_modified_by = read_xml(xml_file)
                result.append(last_modified_by)
    return result



def read_xml(xml_file):
    xmldoc = minidom.parse(xml_file)
    tags = xmldoc.getElementsByTagName('cp:lastModifiedBy')
    tag = tags[0]
    last_modified_by = tag.firstChild.nodeValue
    return last_modified_by


run()

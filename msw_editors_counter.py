from datetime import date
from pathlib import Path
from xml.dom import minidom
from zipfile import ZipFile

base_dir = Path('/tmp/docx_dir')


def run():
    files = get_files(base_dir)
    editors = get_editors(files)
    for dir_date, editor in editors:
        print(f'{dir_date} {editor}')


def get_files(base_dir):
    result = []
    for dir_path in base_dir.iterdir():
        if not dir_path.is_dir():
            continue
        dir_date = get_dir_date(dir_path.name)
        for path in dir_path.iterdir():
            # files on first level
            if path.is_file() and path.suffix == '.docx':
                result.append((dir_date, base_dir/path))
            elif path.is_dir():
                for inner_path in path.iterdir():
                    # files on second level
                    if inner_path.is_file() and inner_path.suffix == '.docx':
                        result.append((dir_date, base_dir/inner_path))
    return result


def get_dir_date(dir_name):
    yy = dir_name[:2]
    yyyy = int(f'20{yy}')
    mm = int(dir_name[2:4])
    dd = int(dir_name[4:6])
    return date(yyyy, mm, dd)


def get_editors(files):
    result = []
    for dir_date, filename in files:
        with ZipFile(filename) as zip_file:
            with zip_file.open('docProps/core.xml') as xml_file:
                last_modified_by = read_xml(xml_file)
                result.append((dir_date, last_modified_by))
    return result


def read_xml(xml_file):
    xmldoc = minidom.parse(xml_file)
    tags = xmldoc.getElementsByTagName('cp:lastModifiedBy')
    tag = tags[0]
    last_modified_by = tag.firstChild.nodeValue
    return last_modified_by


run()

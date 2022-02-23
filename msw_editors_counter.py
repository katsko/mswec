from xml.dom import minidom


xmldoc = minidom.parse('/tmp/docx_dir/core.xml')  # core.xml from file.docx (zip)
tags = xmldoc.getElementsByTagName('cp:lastModifiedBy')
tag = tags[0]
last_modified_by = tag.firstChild.nodeValue
print(last_modified_by)

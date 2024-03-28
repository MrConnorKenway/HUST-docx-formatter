import zipfile
import re

from lxml import etree

archive = zipfile.ZipFile('test.docx', 'r')

with archive.open('word/document.xml') as document:
    xml = document.read()

root = etree.fromstring(xml)

w = root.nsmap['w']

w_r = '{%s}r' % w

ns = {'w': w}
document = root.xpath('/w:document[1]', namespaces=ns)[0]
body = document.xpath('./w:body[1]', namespaces=ns)[0]

parser = etree.XMLParser(remove_blank_text=True)

# Find toc entries of all level 1 headings
last_main_text_title = None
toc1 = body.xpath('./w:p/w:pPr/w:pStyle[@w:val="TOC1"]', namespaces=ns)
for toc_node in toc1:
    txts = toc_node.xpath('../following-sibling::w:r/w:t', namespaces=ns)
    title_name = ''
    for t in txts:
        print(t.text, end=' ')
        title_name += t.text
    print()

    pp = toc_node.getparent().getparent()

    if re.match(r'目[ ]*录[IV0-9]+', title_name):
        # print('Ignore TOC itself in TOC')
        body.remove(toc_node.getparent().getparent())
    elif re.match(r'^[0-5]+', title_name):
        last_main_text_title = title_name
        # print('Delete page number for main text chapter')
        # there should be two tabs
        tabs = toc_node.xpath('../../w:r/w:tab', namespaces=ns)
        if len(tabs) != 2:
            print('Unknown format')
            continue
        to_be_delete = pp.getchildren()[pp.index(tabs[-1].getparent()):]
        for child in to_be_delete:
            pp.remove(child)
    elif last_main_text_title is not None:
        # print('Disable bold')
        # there should be one tabs
        tabs = toc_node.xpath('../../w:r/w:tab', namespaces=ns)
        if len(tabs) != 1:
            print('Unknown format')
            continue

        # Add parentheses around page number
        if re.match(r'^[0-9]+$', txts[-1].text):
            left_xml_str = \
            """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                    <w:r>
                        <w:t>(</w:t>
                    </w:r>
                </w:document>
            """
            right_xml_str = \
            """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                    <w:r>
                        <w:t>)</w:t>
                    </w:r>
                </w:document>
            """
            begin_nodes = txts[-1].xpath('../preceding-sibling::w:r/w:fldChar[@w:fldCharType="begin"]', namespaces=ns)
            nearest_begin_node = begin_nodes[-1]
            end_nodes = txts[-1].xpath('../following-sibling::w:r/w:fldChar[@w:fldCharType="end"]', namespaces=ns)
            nearest_end_node = end_nodes[0]
            element = etree.fromstring(left_xml_str, parser=parser)
            pp.insert(pp.index(nearest_begin_node.getparent()), element.xpath('./w:r', namespaces=ns)[0])
            element = etree.fromstring(right_xml_str, parser=parser)
            pp.insert(pp.index(nearest_end_node.getparent())+1, element.xpath('./w:r', namespaces=ns)[0])

        # Disable bold
        xml_str = \
        """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:rPr>
                    <w:b w:val="0" />
                    <w:bCs w:val="0" />
                </w:rPr>
            </w:document>
        """
        to_be_insert = pp.getchildren()[pp.index(tabs[-1].getparent()):]
        for child in to_be_insert:
            element = etree.fromstring(xml_str, parser=parser) # use parser so that lxml can pretty print
            if not child.xpath('./w:rPr/w:bCs', namespaces=ns):
                child.insert(0, element.xpath('./w:rPr', namespaces=ns)[0])
    else:
        # Disable bold
        # there should be one tabs
        tabs = toc_node.xpath('../../w:r/w:tab', namespaces=ns)
        if len(tabs) != 1:
            print('Unknown format')
            continue

        xml_str = \
        """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:rPr>
                    <w:b w:val="0" />
                    <w:bCs w:val="0" />
                </w:rPr>
            </w:document>
        """
        to_be_insert = pp.getchildren()[pp.index(tabs[-1].getparent()):]
        for child in to_be_insert:
            element = etree.fromstring(xml_str, parser=parser) # use parser so that lxml can pretty print
            if not child.xpath('./w:rPr/w:bCs', namespaces=ns):
                child.insert(0, element.xpath('./w:rPr', namespaces=ns)[0])




# Find toc entries of all level 2 headings
toc2 = body.xpath('./w:p/w:pPr/w:pStyle[@w:val="TOC2"]', namespaces=ns)
for toc_node in toc2:
    txts = toc_node.xpath('../following-sibling::w:r/w:t', namespaces=ns)
    for t in txts:
        print(t.text, end=' ')
    print()

    # Add parentheses around page number
    if re.match(r'^[0-9]+$', txts[-1].text):
        pp = toc_node.getparent().getparent()
        left_xml_str = \
        """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:r>
                    <w:t>(</w:t>
                </w:r>
            </w:document>
        """
        right_xml_str = \
        """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:r>
                    <w:t>)</w:t>
                </w:r>
            </w:document>
        """
        begin_nodes = txts[-1].xpath('../preceding-sibling::w:r/w:fldChar[@w:fldCharType="begin"]', namespaces=ns)
        nearest_begin_node = begin_nodes[-1]
        end_nodes = txts[-1].xpath('../following-sibling::w:r/w:fldChar[@w:fldCharType="end"]', namespaces=ns)
        nearest_end_node = end_nodes[0]
        element = etree.fromstring(left_xml_str, parser=parser)
        pp.insert(pp.index(nearest_begin_node.getparent()), element.xpath('./w:r', namespaces=ns)[0])
        element = etree.fromstring(right_xml_str, parser=parser)
        pp.insert(pp.index(nearest_end_node.getparent())+1, element.xpath('./w:r', namespaces=ns)[0])


# Find all level 1 headings
title1 = body.xpath('./w:p/w:pPr/w:pStyle[@w:val="1"]', namespaces=ns)
references_node = None
for toc_node in title1:
    for t in toc_node.xpath('../following-sibling::w:r/w:t', namespaces=ns):
        if t.text == "参考文献":
            references_node = toc_node
            break

if (references_node is None):
    print('Error: Cannot found chapter named 参考文献')
    exit()

ref_texts = []
references = []
ref_pos = {}

# Parse every reference entry in Chapter References
for i, paragraph in enumerate(references_node.xpath('../../following-sibling::w:p', namespaces=ns)):
    if paragraph.xpath('./w:pPr/w:pStyle[@w:val="1"]', namespaces=ns):
        break
    if paragraph.xpath('./w:bookmarkStart', namespaces=ns):
        ref_texts.append(paragraph)
        for j, bookmark in enumerate(paragraph.xpath('./w:bookmarkStart/@w:name', namespaces=ns)):
            ref_pos[bookmark] = i
            references.append(bookmark)
            if (j > 0):
                print(f'WARN: ref entry [{i+1}] has multiple bookmarks:', bookmark)

print('Found', len(ref_texts), 'references')

# [VERBOSE OUTPUT] Print text of all reference entries
# for i, node in enumerate(ref_texts):
#     print(f'[{i+1}]', end=' ')
#     for txt in node.xpath('./w:r/w:t', namespaces=ns):
#         print(txt.text, end='')
#     print()

checked = set()
sorted_ref = []


# Sort reference entries based on the occurrence order of cross references in main text
for cross_ref in body.xpath('//w:instrText[contains(text(), "REF _Ref")]', namespaces=ns):
    ref_id = cross_ref.text.split()[1]
    if ref_id and ref_id in references and ref_id not in checked:
        checked.add(ref_id)
        try:
            sorted_ref.append(ref_texts[ref_pos[ref_id]])
        except Exception as e:
            print(e)
            exit()


# Change all cross references to superscript
for cross_ref in body.xpath('//w:instrText[contains(text(), "REF _Ref")]', namespaces=ns):
    ref_id = cross_ref.text.split()[1]
    if ref_id and ref_id in references:
        """
            <w:p>                                                          ==============> pp
                <w:r>                                                      ==============> nearest_begin_node's parent
                    <w:rPr>
                        <w:vertAlign w:val="superscript" />
                    </w:rPr>
                    <w:fldChar w:fldCharType="begin" />                    ==============> nearest_begin_node
                </w:r>
                <w:r>                                                      ==============> p
                    <w:instrText>REF _Ref161329064 \r \h</w:instrText>     ==============> cross_ref
                </w:r>
                <w:r>
                    <w:rPr>
                        <w:vertAlign w:val="superscript" />
                    </w:rPr>
                    <w:fldChar w:fldCharType="end" />                      ==============> nearest_end_node
                </w:r>
            </w:p>
        """
        p = cross_ref.getparent()
        if p.tag != w_r:
            print(etree.tostring(p, encoding='UTF-8', pretty_print=True))
            assert(False)

        pp = p.getparent()
        end_nodes = cross_ref.xpath('../following-sibling::w:r/w:fldChar[@w:fldCharType="end"]', namespaces=ns)
        if len(end_nodes) == 0:
            print('Error: Empty end')
            exit()
        nearest_end_node = end_nodes[0]
        begin_nodes = cross_ref.xpath('../preceding-sibling::w:r/w:fldChar[@w:fldCharType="begin"]', namespaces=ns)
        if len(begin_nodes) == 0:
            print('Error: Empty begin')
            exit()
        nearest_begin_node = begin_nodes[-1]
        all_run_siblings = cross_ref.xpath('../../w:r', namespaces=ns)
        begin_idx = pp.index(nearest_begin_node.getparent())
        end_idx = pp.index(nearest_end_node.getparent())
        i = begin_idx
        if begin_idx >= end_idx or len(all_run_siblings) < end_idx:
            print(begin_idx, end_idx, len(all_run_siblings))
            assert(False)
        while i <= end_idx:
            run_node = all_run_siblings[i-1]
            if len(run_node.xpath('./w:rPr', namespaces=ns)) == 0:
                # no rPr
                xml_str = \
                """
                    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                        <w:rPr>
                            <w:vertAlign w:val="superscript" />
                        </w:rPr>
                    </w:document>
                """
                element = etree.fromstring(xml_str, parser=parser)
                run_node.insert(0, element.xpath('./w:rPr', namespaces=ns)[0])
            elif len(run_node.xpath('./w:rPr/w:vertAlign[@w:val="superscript"]', namespaces=ns)) == 0:
                # no superscript
                rPr_nodes = run_node.xpath('./w:rPr', namespaces=ns)
                assert(len(rPr_nodes) == 1)
                xml_str = \
                """
                    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                        <w:vertAlign w:val="superscript" />
                    </w:document>
                """
                element = etree.fromstring(xml_str, parser=parser)
                rPr_nodes[0].insert(0, element.xpath('./w:vertAlign', namespaces=ns)[0])
            i += 1


reference_start = references_node.xpath('../..', namespaces=ns)[0]

# Remove the original reference text, and replace it with sorted one
for i, r in enumerate(sorted_ref):
    body.remove(r)
    body.insert(body.index(reference_start)+1+i, r)


# [DEBUG] output test.xml for debugging purpose
with open('test.xml', 'w') as f:
    f.write(etree.tostring(root, encoding='UTF-8', pretty_print=True).decode('UTF-8'))


# Save output to another .docx file
with zipfile.ZipFile('output.docx', 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=archive.compresslevel) as dest:
    for file in archive.filelist:
        if file.filename == "word/document.xml":
            data = etree.tostring(root, encoding='UTF-8', standalone=True)
        else:
            data = archive.read(file.filename)
        dest.writestr(file.filename, data)

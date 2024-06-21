
# remove paragraphs using docx package
def remove_empty_paragraphs(doc):
    for para in doc.paragraphs:
        if not para.text.strip():
            p_element = para._element
            p_element.getparent().remove(p_element)

def extract_footnote_pos_string(doc, footnote_string):
    pos = 0
    for sentence in doc.body[0][0][0]:
        string = sentence.strip()
        p = string.find(footnote_string)
        if p != -1:
            pos = p
            break
    return string[pos-30:pos]

def extract_footnote_para_strings(doc, index_list):
    footnote_para_strings = []
    for index in index_list:
        footnote_para_strings.append(doc.body[0][0][0][index].strip())
    return footnote_para_strings

# extract_footnote using docx2python
def extract_footnote(doc):
    footnote_list = []
    for footnote in doc.footnotes_runs[0][0]:
        for specific in footnote:
            for line in specific:
                split_lines = line.split("\t")
                if "footnote" not in split_lines[0]:                    
                    footnote_list.append(split_lines[0].strip())
    return footnote_list

# handle paragraph using Document from docx-python package
def remove_string_from_paragraph(doc, target_string):
    for paragraph in doc.paragraphs:            
        if target_string in paragraph.text:
            print(paragraph.text)
            paragraph.text = paragraph.text.replace(target_string, '')

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

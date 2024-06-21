from spire.doc import *
from spire.doc.common import *

def add_footnote(file_path, para_index, refer_string, footnote_text):
    # Create a Document instance
    document = Document()
    # Load a sample Word document
    document.LoadFromFile(file_path)
    section = document.Sections[0]
    
    if section.Paragraphs.get_Item(0).Text == "Evaluation Warning: The document was created with Spire.Doc for Python.":
        section.Paragraphs.RemoveAt(0)
    
    paragraph = section.Paragraphs.get_Item(para_index)
    print(paragraph.ChildObjects.Count)
    
    selection = document.FindString(refer_string, False, True)
    
    # if selection:
    # Get the found text as a single text range
    textRange = selection.GetAsOneRange()    
    
    # Get the index position of the text range in the paragraph
    index = paragraph.ChildObjects.IndexOf(textRange)
    print(index)
    
    # # Insert the ChildObject into the paragraph
    # paragraph.ChildObjects.Insert(index, textRange)


    # Add a footnote to the paragraph
    footnote = paragraph.AppendFootnote(FootnoteType.Footnote)
    

    # Insert the footnote after the text range
    paragraph.ChildObjects.Insert(index+1, footnote)

    # Set the text content of the footnote
    text = footnote.TextBody.AddParagraph().AppendText(footnote_text)

   
    # Save the result document
    document.SaveToFile(file_path, FileFormat.Docx)
    document.Close()

def find_paragraphs_for_footnote(file_path):
    # Create a Document instance
    document = Document()
    # Load a Word document    
    document.LoadFromFile(file_path)
    # Get the first section of the document
    section = document.Sections[0]   
    
    para_footnote_indexes = []
    # Loop through the paragraphs in the section
    for y in range(section.Paragraphs.Count):    
        para = section.Paragraphs.get_Item(y)
        index = -1
        i = 0
        cnt = para.ChildObjects.Count
        while i < cnt:        
            pBase = para.ChildObjects[i] if isinstance(para.ChildObjects[i], ParagraphBase) else None
            if isinstance(pBase, Footnote):
                index = i 
                if index > -1:
                    para_footnote_indexes.append(y)
            i += 1
    document.Close
    return para_footnote_indexes



# add_footnote("translated_test_llama3_8b.docx", 1, "----footnote1----", "rrr")
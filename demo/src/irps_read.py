from docx import Document 
import os


class ParsedDocx:
    textbox = ['']


def irps_read_parse_shapes(doc):
    shapes = doc.shape.InlineShapes()
    print(shapes)    

'''
    @brief 
        parse textbox int docx file
    @return
        the text which parse in the docx file's textbox 
'''
def irps_read_parse_textbox(doc):
    text = ['']
    children = doc.element.body.iter()
    childIters = []

    for child in children:
        if (child.tag.endswith('textbox')):
            for it in child.iter():
                if (it.tag.endswith(('main}r', 'main}pPr'))):
                    childIters.append(it)
        else:
            print(child.tag)
    for it in childIters:
        if it.tag.endswith('main}pPr'):
            text.append('')
        else:
            text[-1] += it.text
        it.text = ''
    # use '###' to distingunish differ paragraph
    return text

'''
    @brief 
        read docx file
'''
def irps_read_docx(filePath):
    parsedDocx = ParsedDocx()
    doc = Document(filePath)
    parsedDocx.textbox = irps_read_parse_textbox(doc)  
    irps_read_parse_shapes(doc)
    return parsedDocx
    
# test
filePath = '/Users/haoranyan/git_rep/irps_based_on_cloud_ai/demo/test/test.docx'
tmp = irps_read_docx(filePath)
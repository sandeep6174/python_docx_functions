from docx_functions import *

document = new_document()
add_paragraph_to_doc(document, 'hi', 'number', font_size=15)
add_paragraph_to_doc(document, 'hi', 'number', font_size=20)
paragraph1 = add_paragraph_to_doc(document, 'hi', font_name='Arial')
continue_para_in_doc(paragraph1, 'hold', True, underline=True)
add_hyperlink_to_doc(paragraph1, "fds", 'google.com')
tab=add_table_to_doc(document,3,3)
edit_table_in_doc(tab,'dfgh',1,1)
document.save('test.docx')
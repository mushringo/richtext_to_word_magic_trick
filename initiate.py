from html_to_word import richtext_convertor
from docxtpl import DocxTemplate

tagged_text = "<u>KEK</u><p>Hello</p>It's my string.<em>So i can do</em><u>Whit it.</u>"
tpl = DocxTemplate('richtext_tpl.docx')
context = {
    'context_variable': richtext_convertor(tagged_text),
}

tpl.render(context)
tpl.save('richtext.docx')

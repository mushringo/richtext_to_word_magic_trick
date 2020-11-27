from html_to_word import richtext_convertor
from docxtpl import DocxTemplate

tagged_text = '<p>pur text,</p><p><strong>zhirnyj,&nbsp; <em>bold+italic, <sub>subtext,' \
              ' </sub></em></strong></p><p>' \
              '<sup>supertext, <span style="text-decoration: underline;">super-underline,' \
              ' </span></sup></p><h4><sup><span style="text-decoration: underline;">capital' \
              '</span></sup></h4><ul><li><strong>numbered-bold 1<br /></strong></li><li>' \
              '<strong>numbered-bold 2</strong></li></ul><ol><li><sup>super-number 1<br />' \
              '</sup></li><li><sup>super-number 2</sup></li></ol><p>&nbsp;</p>'
tpl = DocxTemplate('richtext_tpl.docx')
context = {
    'context_variable': richtext_convertor(tagged_text),
}

tpl.render(context)
tpl.save('richtext.docx')

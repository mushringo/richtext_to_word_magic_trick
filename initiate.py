from html_to_word import richtext_convertor
from docxtpl import DocxTemplate

tagged_text = '<p>pur text,</p><p><strong>zhirnyj,&nbsp; <em>bold+italic, <sub>subtext,' \
              ' </sub></em></strong></p><p>' \
              '<sup>supertext, <span style="text-decoration: underline;">super-underline,' \
              ' </span></sup></p><h4><sup><span style="text-decoration: underline;">capital' \
              '</span></sup></h4><ul><li><strong>numbered-bold 1<br /></strong></li><li>' \
              '<strong>numbered-bold 2</strong></li></ul><ol><li><sup>super-number 1<br />' \
              '</sup></li><li><sup>super-number 2</sup></li></ol><p>&nbsp;</p>'
h1_h6 = '<ul><li><h1>word</h1></li><li><h2>word</h2></li><li><h3>word</h3></li><li><h4>word</h4></li><li><h5>word</h5>' \
        '</li><li><h6>word</h6></li></ul>'
tpl = DocxTemplate('richtext_tpl.docx')
context = {
    'context_variable': richtext_convertor(h1_h6),
}

tpl.render(context)
tpl.save('richtext.docx')

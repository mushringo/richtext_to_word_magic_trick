from html_to_word import richtext_convertor
from docxtpl import DocxTemplate

tagged_text = '<p><span style="text-decoration: underline;"><strong>Heute definiert:</strong><strong><sub><br />' \
              '</sub></strong></span></p><ol><li><strong>Ziel 1: </strong>besser werden</li><li>Ziel 2: noch besser' \
              ' sein</li></ol><h4>Ja, das ist unser Ziel.</h4><p>&nbsp;</p>'
h1_h6 = '<ul><li><h1>word</h1></li><li><h2>word</h2></li><li><h3>word</h3></li><li><h4>word</h4></li><li><h5>word</h5>' \
        '</li><li><h6>word</h6></li></ul>'
umlaut = "Hi &Auml; er &Ouml; unsere &szlig; Aktivit&auml;ten"
tpl = DocxTemplate('richtext_tpl.docx')
context = {
    'context_variable': richtext_convertor(tagged_text),
}

tpl.render(context)
tpl.save('richtext.docx')

from html_to_word import richtext_convertor
from docxtpl import DocxTemplate

tagged_text = '<p><span style="text-decoration: underline;"><strong>Heute definiert:</strong><strong><sub><br />' \
              '</sub></strong></span></p><ol><li><strong>Ziel 1: </strong>besser werden</li><li>Ziel 2: noch besser' \
              ' sein</li></ol><h4>Ja, das ist unser Ziel.</h4><p>&nbsp;</p>'
h1_h6 = '<ul><li><h1>word</h1></li><li><h2>word</h2></li><li><h3>word</h3></li><li><h4>word</h4></li><li><h5>word</h5>' \
        '</li><li><h6>word</h6></li></ul>'
umlaut = "Hi &Auml; er &Ouml; unsere &szlig; Aktivit&auml;ten"
list2 = "<ul><li>asd</li><ul><li>dsa</li><li>dsa</li></ul><li>zxc</li><li>zxc</li></ul>"
tagged_text2 = '<p>Copy&Paste-Text</p>' \
               '<p>Ich würde gerne einen längeren Text hier verfassen, aber wenn man nichts zu sagen hat, ist es so ' \
               'schwer, etwas zu schreiben. Ein Überschrift könnte vielleicht helfen.</p> \
               <p>Oder auch nicht. Außerdem dies & jenes. Und: "solches".</p> \
<p><strong>Oder doch?</strong></p> \
<p>Es gibt natürlich noch andere Möglichkeiten, z.B.:</p> \
<ol start="1" type="1"> \
<li>Texte</li> \
<li>Zahlen</li> \
<li>Worte</li> \
<ol start="1" type="1"> \
<li>Sonstiges</li> \
<li>Anderes</li> \
<li>Diverses</li> \
</ol> \
<li>und nicht zu vergessen ganze Sätze</li> \
</ol> \
<p>Aber hilft das <strong>irgendwem</strong>? Vielleicht, <em>vielleicht</em> auch nicht. Man weiß es nicht. Aber' \
               ' solche Texte sind schon nicht ohne. <u>Vor allem</u> nicht ohne Fehler.</p> \
<ul type="disc"> \
<li>Hier zur Auswahl</li> \
<ul type="circle"> \
<li>Fehler 1</li> \
<li>Fehler 2</li> \
<li>Fehler 3</li> \
</ul> \
<li>Und sonstige Korrekturen</li> \
<li>Notwendigkeiten</li> \
</ul> \
<p>CO<sup>2</sup>-Emissionen. Man könnte daraus natürlich auch Fußnoten machen.</p> \
<p>Oder andere CO<sub>x</sub>-Sachen.</p>'
tpl = DocxTemplate('richtext_tpl.docx')
context = {
    'context_variable': richtext_convertor(tagged_text2),
}

tpl.render(context)
tpl.save('richtext.docx')

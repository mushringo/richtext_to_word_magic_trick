from docxtpl import DocxTemplate, RichText
import re

tpl = DocxTemplate('richtext_tpl.docx')
tagged_text = "<u>KEK</u><p>Hello</p>It's my string.<em>So i can do</em><b>Whatever I want</b>" \
              " <u>Whit it.</u> <sup>But it has to be correct displayed</sup><sub> And Im doing my best.</sub>KEKW" \
              "<ol><li>1</li><li>4</li><li>2</li></ol><ul><li>1</li><li>4</li><li>2</li></ul><!DOCTYPE html>"
"""
Tag legend:

p - paragraph
em - italic
b - bold
u - underline
li - lst item
ul - unordered lst
ol - ordered lst
sup - superscript
sub - subscript 
"""


# ____________________________________________________________________________________________________________________
# Main body
# --------------------------------------------------------------------------------------------------------------------


def richtext_convertor(text):
    rt = RichText()
    order = 1
    clean_sliced_array = []
    sup = False
    sub = False
    bold = False
    under = False
    italic = False
    parag = False
    bi_order = False
    bi_unorder = False

    first_cut = re.split('<', text)
    for i in first_cut:
        if i == '':
            continue
        else:
            final_cut = re.split('>', i)
            for k in final_cut:
                if k == '':
                    continue
                else:
                    clean_sliced_array.append(k)

    for i in clean_sliced_array:
        if i == '/u':
            under = False
        elif i == '/sup':
            sup = False
        elif i == '/sub':
            sub = False
        elif i == '/b':
            bold = False
        elif i == '/em':
            italic = False
        elif i == 'u':
            under = True
        elif i == 'sup':
            sup = True
        elif i == 'sub':
            sub = True
        elif i == 'b':
            bold = True
        elif i == 'em':
            italic = True
        elif i == 'p':
            parag = True
        elif i == '/p':
            rt.add('\n')
        elif i == 'ul':
            rt.add('\n')
            bi_unorder = True
        elif i == '/ul':
            bi_unorder = False
        elif i == 'ol':
            rt.add('\n')
            bi_order = True
        elif i == '/ol':
            bi_order = False
            order = 1
        elif i == '/li':
            rt.add('\n')
        elif i == 'li':
            if bi_unorder is True:
                rt.add('   ' + chr(183) + ' ')
            elif bi_order is True:
                rt.add('   ' + str(order) + '. ')
                order += 1
        else:
            if parag is True:
                rt.add('\n')
                rt.add(i, underline=under, italic=italic, bold=bold, subscript=sub, superscript=sup)
                parag = False
            else:
                rt.add(i, underline=under, italic=italic, bold=bold, subscript=sub, superscript=sup)
    return rt


# ___________________________________________________________________________________________________________________
# Output
# -------------------------------------------------------------------------------------------------------------------
context = {
    'context_variable': richtext_convertor(tagged_text),
}

tpl.render(context)
tpl.save('richtext.docx')

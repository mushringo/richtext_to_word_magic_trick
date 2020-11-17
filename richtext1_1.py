from docxtpl import DocxTemplate, RichText
import re

tpl = DocxTemplate('richtext_tpl.docx')
tagged_text = "<ol><b><li>text1</li></b>" \
              "<em><li>text2</li></em>" \
              "<u><li>text3</li></u></ol>" \
              "<ul><u><sub><li>text4</li></sub>" \
              "<li>text5</li>" \
              "<sup><li>text6</li></sup></u></ul>"

tagged_text_2 = "<u>KEK</u><p>Hello</p>It's my string.<em>So i can do</em><b>Whatever I want</b>" \
              " <u>Whit it.</u> <sup>But it has to be correct displayed</sup><sub> And Im doing my best.</sub>KEKW" \
              "<ol><li>1</li><li>4</li><li>2</li></ol><ul><li>1</li><li>4</li><li>2</li></ul>"
tag = 'p', 'em', 'b', 'u', 'sup', 'sub', 'li', 'ul', 'ol', 'color'
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

clean_sliced_array = []
rt = RichText()


# ____________________________________________________________________________________________________________________
# Main body
# --------------------------------------------------------------------------------------------------------------------


def slicer(text):
    a = re.split('<', text)
    for i in a:
        if i == '':
            continue
        else:
            b = re.split('>', i)
            for k in b:
                if k == '':
                    continue
                else:
                    clean_sliced_array.append(k)
    return clean_sliced_array


def main_un_ital(lst):
    order = 1
    sup = False
    sub = False
    bold = False
    under = False
    italic = False
    parag = False
    bi_order = False
    bi_unorder = False
    for i in lst:
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
                executor(i, under, italic, bold, sub, sup)
                parag = False
            else:
                executor(i, under, italic, bold, sub, sup)


def executor(text, u=False, em=False, bld=False, sub=False, sup=False):
    rt.add(text, underline=u, italic=em, bold=bld, subscript=sub, superscript=sup)


main_un_ital(slicer(tagged_text))

# ___________________________________________________________________________________________________________________
# Output
# -------------------------------------------------------------------------------------------------------------------
context = {
    'context_variable': rt,
}

tpl.render(context)
tpl.save('richtext.docx')

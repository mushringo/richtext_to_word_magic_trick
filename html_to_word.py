from docxtpl import RichText
import re

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
    sz = 22
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
    print(clean_sliced_array)

    for i in clean_sliced_array:
        if i == 'br /':
            rt.add('\n')
        elif i == '/u':
            under = False
        elif i == '/sup':
            sup = False
        elif i == '/h1':
            sz = 22
            bold = False
        elif i == '/h2':
            sz = 22
            bold = False
        elif i == '/h3':
            sz = 22
            bold = False
        elif i == '/h4':
            sz = 22
            bold = False
        elif i == '/h5':
            sz = 22
            bold = False
        elif i == '/h6':
            sz = 22
            bold = False
        elif i == '/sub':
            sub = False
        elif i == '/strong' or i == '/b':
            bold = False
        elif i == '/em':
            italic = False
        elif i == 'u':
            under = True
        elif i == 'h1':
            sz = 34
            bold = True
        elif i == 'h2':
            sz = 30
            bold = True
        elif i == 'h3':
            sz = 26
            bold = True
        elif i == 'h4':
            sz = 22
            bold = True
        elif i == 'h5':
            sz = 18
            bold = True
        elif i == 'h6':
            sz = 14
            bold = True
        elif i == 'sup':
            sup = True
        elif i == 'sub':
            sub = True
        elif i == 'strong' or i == 'b':
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
            if parag is True and i == '&nbsp;':
                continue
            elif parag is True and '&nbsp;' in i:
                b = re.split('&nbsp;', i)
                b = ''.join(b)
                rt.add('\n')
                rt.add(b, underline=under, italic=italic, bold=bold, subscript=sub, superscript=sup, size=sz)
                parag = False
            else:
                rt.add(i, underline=under, italic=italic, bold=bold, subscript=sub, superscript=sup, size=sz)
    return rt

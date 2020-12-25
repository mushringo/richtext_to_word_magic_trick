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
    o = 0
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
    dict = {
        '&Auml;': chr(196),
        '&auml;': chr(228),
        '&Ouml;': chr(214),
        '&ouml;': chr(246),
        '&Uuml;': chr(220),
        '&uuml;': chr(252),
        '&szlig;': chr(223)
    }

    for i in dict.keys():
        if i in text:
            splt = []
            spl = re.split(i, text)
            for n in range(0, len(spl)):
                if n == len(spl) - 1:
                    splt.append(spl[n])
                else:
                    splt.append(spl[n])
                    splt.append(dict.get(i))
            text = ''.join(splt)
        else:
            continue

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
        if i == 'br /':
            rt.add('\n')
            o += 1
        elif i == '/span':
            o += 1
            under = False
        elif i == '/sup':
            o += 1
            sup = False
        elif i == '/h1':
            o += 1
            sz = 22
            bold = False
        elif i == '/h2':
            o += 1
            sz = 22
            bold = False
        elif i == '/h3':
            o += 1
            sz = 22
            bold = False
        elif i == '/h4':
            o += 1
            sz = 22
            bold = False
        elif i == '/h5':
            o += 1
            sz = 22
            bold = False
        elif i == '/h6':
            o += 1
            sz = 22
            bold = False
        elif i == '/sub':
            o += 1
            sub = False
        elif i == '/strong' or i == '/b':
            o += 1
            bold = False
        elif i == '/em':
            o += 1
            italic = False
        elif i == 'span style="text-decoration: underline;"':
            o += 1
            under = True
        elif i == 'h1':
            o += 1
            sz = 34
            bold = True
        elif i == 'h2':
            o += 1
            sz = 30
            bold = True
        elif i == 'h3':
            o += 1
            sz = 26
            bold = True
        elif i == 'h4':
            o += 1
            sz = 22
            bold = True
        elif i == 'h5':
            o += 1
            sz = 18
            bold = True
        elif i == 'h6':
            o += 1
            sz = 14
            bold = True
        elif i == 'sup':
            o += 1
            sup = True
        elif i == 'sub':
            o += 1
            sub = True
        elif i == 'strong' or i == 'b':
            o += 1
            bold = True
        elif i == 'em':
            o += 1
            italic = True
        elif i == 'p':
            o += 1
            parag = True
        elif i == '/p':
            o += 1
            if o != len(clean_sliced_array):
                rt.add('\n')
                print(o)
            else:
                continue
        elif i == 'ul':
            o += 1
            rt.add('\n')
            bi_unorder = True
        elif i == '/ul':
            o += 1
            bi_unorder = False
        elif i == 'ol':
            o += 1
            if o != len(clean_sliced_array):
                rt.add('\n')
            else:
                continue
            bi_order = True
        elif i == '/ol':
            o += 1
            bi_order = False
            order = 1
        elif i == '/li':
            o += 1
            if o != len(clean_sliced_array):
                rt.add('\n')
            continue
        elif i == 'li':
            o += 1
            if bi_unorder is True:
                rt.add('   ' + chr(183) + ' ')
            elif bi_order is True:
                rt.add('   ' + str(order) + '. ')
                order += 1
        else:
            o += 1
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

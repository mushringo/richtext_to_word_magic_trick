from docxtpl import DocxTemplate, RichText
import re


# Test text "It's my string. <p>dddd</p> It has to be correct displayed"

tpl = DocxTemplate('richtext_tpl.docx')
tagged_text = "<p>drow</p>It's my string.<p><em><b>port</b></em> It has to be correct displayed </p>"
tag = 'p', 'em', 'b', '/p', '/em', '/b'
clean_sliced_array = []
text_fragments = []
rt = RichText()


# Creating closed tag for open one

def close_tag(tg):
    gag = tg[0] + "/" + tg[1:len(tg)]
    return gag


# Function for <em> Italic-tag

def add_italic(text):
    rt.add(text, italic=True)


# Function for <p> tag

def add_paragraph(text):
    rt.add('\n' + text + '\n')


# Slicing texts by tags and word fragments

def slicer(text):
    sliced_array = []
    splitted = re.split('<', text)
    for a in splitted:
        if a == '':
            continue
        elif a[0] == "/":
            t = []
            w = []
            for i in range(0, len(a)):
                if a[i] == '>':
                    break
                else:
                    t.append(a[i])
            t = ''.join(t)
            sliced_array.append(t)
            for i in range(len(a) - 1, 0, -1):
                if a[i] == '>':
                    for l in range(i+1, len(a)):
                        w.append(a[l])
                else:
                    continue
            w = ''.join(w)
            sliced_array.append(w)
        else:
            t = []
            w = []
            for i in range(0, len(a)):
                if a[i] == '>':
                    break
                else:
                    t.append(a[i])
            t = ''.join(t)
            sliced_array.append(t)
            for i in range(len(a) - 1, 0, -1):
                if a[i] == '>':
                    for l in range(i + 1, len(a)):
                        w.append(a[l])
                else:
                    continue

            w = ''.join(w)
            sliced_array.append(w)
    for k in range(0, len(sliced_array)):
        if sliced_array[k] == '':
            continue
        else:
            clean_sliced_array.append(sliced_array[k])
    print(clean_sliced_array)


slicer(tagged_text)


# Appending untagged and tagged text. Creating an order

"""def append_text_fragments():
    # Slicing a string

    first_slice = re.split(tag, tagged_text)
    last_slice = re.split(close_tag(tag), re.split(tag, tagged_text)[1])
    clean_text = last_slice[0]

    if first_slice[0] != '' and last_slice[1] != '':
        text_fragments.append(first_slice[0])
        text_fragments.append(clean_text)
        text_fragments.append(last_slice[1])
    elif first_slice[0] != '' and last_slice[1] == '':
        text_fragments.append(first_slice[0])
        text_fragments.append(clean_text)
    elif first_slice[0] == '' and last_slice[1] != '':
        text_fragments.append(clean_text)
        text_fragments.append(last_slice[1])

    for a in text_fragments:
        if a == clean_text:
            add_paragraph(clean_text)
        else:
            rt.add(a)"""

#
# Searching for tag

"""def if_tag(text):
    if not re.findall('<*>', text):
        rt.add(text)
    else:
        append_text_fragments()"""


# Multiple tags

def tag_next_to_tag(list):
    i = 0
    l = 1
    extratag = []
    multitag = []
    print(len(list))
    while i <= len(list):
        if list[i] in tag:
            multitag.append(list[i])
            while i + l <= len(list):
                if list[i+l] in tag:
                    multitag.append(list[i+l])
                    print(i, l)
                    l += 1
                else:
                    break
            print(i, l)
            i = i + l
            l = 1

        else:
            print(i, l)
            i += 1
        print(multitag)
        extratag.append(multitag)
        multitag = []
    for l in extratag:
        if l != '':
            multitag.append(l)
    print(extratag)




tag_next_to_tag(clean_sliced_array)

context = {
    'context_variable': rt,
}

tpl.render(context)
tpl.save('richtext.docx')

"""# -*- coding: utf-8 -*-
'''
Created : 2015-03-26
@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate, RichText
import re
tag = "<em>"


def close_tag(tag):
    tag = tag[0] + "/" + tag[1:len(tag)]
    return tag



source = '<em>dddd</em>'
tpl = DocxTemplate('templates/richtext_tpl.docx')

rt = RichText()
rt.add('some more italic', italic=True, color='#ff00ff')

""""""def add_bold():

def add_italic():

def add_underline():

def add_paragraph():

def add_bullet_list():

def add_colour():""""""


rt.add('a rich text', style='myrichtextstyle')
rt.add(' with ')
rt.add('some italic', italic=True)
rt.add(' and ')
rt.add('some violet', color='#ff00ff')
rt.add(' and ')
rt.add('some striked', strike=True)
rt.add(' and ')
rt.add('some small', size=14)
rt.add(' or ')
rt.add('big', size=60)
rt.add(' text.')
rt.add('\nYou can add an hyperlink, here to ')
rt.add('google', url_id=tpl.build_url_id('http://google.com'))
rt.add('\nEt voilà ! ')
rt.add('\n1st line')
rt.add('\n2nd line')
rt.add('\n3rd line')
rt.add('\n\n<cool>')
for ul in ['single', 'double', 'thick', 'dotted', 'dash', 'dotDash', 'dotDotDash', 'wave']:
    rt.add('\nUnderline : ' + ul + ' \n', underline=ul)
rt.add('\nFonts :\n', underline=True)
rt.add('Arial\n', font='Arial')
rt.add('Courier New\n', font='Courier New')
rt.add('Times New Roman\n', font='Times New Roman')
rt.add('\n\nHere some')
rt.add('superscript', superscript=True)
rt.add(' and some')
rt.add('subscript', subscript=True)


rt_embedded = RichText('an example of ')
rt_embedded.add(rt)

context = {
    'context_variable': rt,
    'example': rt_embedded,
}
print(context)
tpl.render(context)
tpl.save('output/richtext.docx')"""

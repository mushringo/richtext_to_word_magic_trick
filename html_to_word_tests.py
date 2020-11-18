from docxtpl import DocxTemplate, RichText
import re

tpl = DocxTemplate('richtext_tpl.docx')
tagged_text = "<u>KEK</u><p>Hello</p>It's my string.<em>So i can do</em><b>Whatever I want</b>" \
              " <u>Whit it.</u> <sup>But it has to be correct displayed</sup><sub> And Im doing my best.</sub>KEKW" \
              "<ol><li>1</li><li>4</li><li>2</li></ol><ul><li>1</li><li>4</li><li>2</li></ul>"

tagged_text_2 = "<u><em>lol</em></u>"
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
text_fragments = []
rt = RichText()
ul_list = []
ol_list = []
# _____________________________________________________________________________________________________________________
# basic functions
# ---------------------------------------------------------------------------------------------------------------------

# unordered lst


def unordered_list(text):
    rt.add('   ' + chr(183) + ' ' + text + '\n')


# ordered lst


def ordered_list(text, order):
    rt.add('   ' + str(order) + '. ' + text + '\n')


# Creating closed tag for open one


def close_tag(tg):
    gag = tg[0] + "/" + tg[1:len(tg)]
    return gag


def unclose_tag(tg):
    gag = tg[1:len(tg)]
    return gag


# Function for <p> tag

def add_paragraph(text):
    rt.add('\n' + text + '\n')


# Function for <em> Italic-tag

def add_italic(text):
    rt.add(text, italic=True)
    rt.add(' ')


# Function for <b> Bold tag

def add_bold(text):
    rt.add(text, bold=True)
    rt.add(' ')


# Function for <u> Underline tag

def add_underline(text):
    rt.add(text, underline=True, bold=True)
    rt.add(' ')


# Function for <sub> Subscript tag

def add_subscript(text):
    rt.add(text, subscript=True)
    rt.add(' ')


# Function for <sup> Superscript tag

def add_superscript(text):
    rt.add(text, superscript=True)
    rt.add(' ')


# Function for <color> Subscript tag

def add_color(text, clr):
    rt.add(text, color=str(clr))
    rt.add(' ')

# ____________________________________________________________________________________________________________________
# Complicated functions
# --------------------------------------------------------------------------------------------------------------------


# Executing tags

def executor(lst):
    for i in range(0, len(lst)):
        if lst[i][0:5] == 'color':
            add_color(lst[(i + 1)], lst[i][7:14])
        elif lst[i] in tag:
            if lst[i] == 'li':
                continue
            elif lst[i] == 'p':
                add_paragraph(lst[i+1])
            elif lst[i] == 'em':
                add_italic(lst[i+1])
            elif lst[i] == 'b':
                add_bold(lst[i+1])
            elif lst[i] == 'u':
                add_underline(lst[i+1])
            elif lst[i] == 'sub':
                add_subscript(lst[i+1])
            elif lst[i] == 'sup':
                add_superscript(lst[i+1])
            elif lst[i] == 'ul':
                for x in range(i + 1, len(lst)):
                    if lst[x] == '/ul':
                        break
                    elif lst[x] == 'li' or lst[x] == '/li':
                        continue
                    else:
                        ul_list.append(lst[x])
                rt.add('\n')
                for n in range(0, len(ul_list)):
                    unordered_list(ul_list[n])
            elif lst[i] == 'ol':
                for x in range(i + 1, len(lst)):
                    if lst[x] == '/ol':
                        break
                    elif lst[x] == 'li' or lst[x] == '/li':
                        continue
                    else:
                        ol_list.append(lst[x])
                rt.add('\n')
                for n in range(0, len(ol_list)):
                    ordered_list(ol_list[n], n + 1)
            elif lst[i] == '/ol' or lst[i] == '/ul':
                continue
            else:
                rt.add(lst[i])
        elif unclose_tag(lst[i]) in tag:
            if i+1 < len(lst) and lst[i+1] not in tag:
                if unclose_tag(lst[i+1]) not in tag:
                    rt.add(lst[i+1])
        else:
            continue



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

def executor(text, u=False, em=False, bld=False, sub=False, sup=False):
    rt.add(text, underline=u, italic=em, bold=bld, subscript=sub, superscript=sup)
# Slicing texts by tags and word fragments

def slicer(text):
    sliced_array = []
    splitted = re.split('<', text)
    print(splitted)
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
            # If there is no tag which affects text (to refactor)
            for i in range(len(a) - 1, 0, -1):
                if a[i] == '>':
                    for x in range(i+1, len(a)):
                        w.append(a[x])
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
                    for x in range(i + 1, len(a)):
                        w.append(a[x])
                else:
                    continue

            w = ''.join(w)
            sliced_array.append(w)
    for k in range(0, len(sliced_array)):
        if sliced_array[k] == '':
            continue
        else:
            clean_sliced_array.append(sliced_array[k])
    return clean_sliced_array


def slicer_2(text):
    z = []
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


print(slicer_2(tagged_text))


# Multiple tags

def tag_next_to_tag(lst):
    i = 0
    x = 1
    m = 1
    taggedarray = []
    multitag = []
    multiclosetag = []
    print(len(lst))
    while i < len(lst):
        if lst[i] in tag:
            multitag.append(lst[i])
            while i + x < len(lst):
                if lst[i+x] in tag:
                    multitag.append(lst[i+x])
                    print(i, x)
                    x += 1
                else:
                    break
            print(i, x)
            i = i + x
            x = 1
            taggedarray.append(multitag)
        elif unclose_tag(lst[i]) in tag:
            multiclosetag.append(lst[i])
            while i + m < len(lst):
                if lst[i + m] in tag:
                    multiclosetag.append(lst[i + m])
                    print(i, m)
                    x += 1
                else:
                    break
            print(i, m)
            i = i + m
            m = 1
            taggedarray.append(multiclosetag)
        else:
            print(i, x)
            i += 1
        print(multitag)
        taggedarray.append(multitag)
        multitag = []
    for x in taggedarray:
        if x != '':
            multitag.append(x)
    print(taggedarray)


def main2(lst):
    under = False
    italic = False
    for i in lst:
        if i == 'u':

            under = True
        elif i == 'em':

            italic = True
        elif i[0] == '/':
            continue
        else:

            lol(i, under, italic)







def lol(text, u=False, em=False):
    rt.add(text, underline=u, italic=em)



lol(main2(clean_sliced_array))




# ___________________________________________________________________________________________________________________
# Output
# -------------------------------------------------------------------------------------------------------------------
context = {
    'context_variable': rt,
}

tpl.render(context)
tpl.save('richtext.docx')
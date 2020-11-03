from docxtpl import DocxTemplate, RichText
import re

tpl = DocxTemplate('richtext_tpl.docx')
tagged_text = "<u>KEK</u><p>Hello</p>It's my string.<em>So i can do</em><b>Whatever I want</b>" \
              " <u>Whit it.</u> <sup>But it has to be correct displayed</sup><sub> And Im doing my best.</sub>KEKW" \
              "<ol><li>1</li><li>4</li><li>2</li></ol><ul><li>1</li><li>4</li><li>2</li></ul>"
tagged_text_2 = "<u>KEK</u><p>Hello</p>It's my string.<em>So i can do</em><b>Whatever I want</b>" \
              " <u>Whit it.</u> <sup>But it has to be correct displayed</sup><sub> And Im doing my best.</sub>KEKW" \
              "<ul><li>1</li><li>2</li></ul>"
tag = 'p', 'em', 'b', 'u', 'sup', 'sub', 'li', 'ul', 'ol', 'color'
"""
Tag legend:

p - paragraph
em - italic
b - bold
u - underline
li - list item
ul - unordered list
ol - ordered list
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

# unordered list


def unordered_list(text):
    rt.add('   ' + chr(183) + ' ' + text + '\n')


# ordered list


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

def executor(list):
    for i in range(0, len(list)):
        if list[i][0:5] == 'color':
            add_color(list[(i + 1)], list[i][7:14])
        elif list[i] in tag:
            if list[i] == 'li':
                continue
            elif list[i] == 'p':
                add_paragraph(list[i+1])
            elif list[i] == 'em':
                add_italic(list[i+1])
            elif list[i] == 'b':
                add_bold(list[i+1])
            elif list[i] == 'u':
                add_underline(list[i+1])
            elif list[i] == 'sub':
                add_subscript(list[i+1])
            elif list[i] == 'sup':
                add_superscript(list[i+1])
            elif list[i] == 'ul':
                for l in range(i + 1, len(list)):
                    if list[l] == '/ul':
                        break
                    elif list[l] == 'li' or list[l] == '/li':
                        continue
                    else:
                        ul_list.append(list[l])
                rt.add('\n')
                for n in range(0, len(ul_list)):
                    unordered_list(ul_list[n])
            elif list[i] == 'ol':
                for l in range(i + 1, len(list)):
                    if list[l] == '/ol':
                        break
                    elif list[l] == 'li' or list[l] == '/li':
                        continue
                    else:
                        ol_list.append(list[l])
                rt.add('\n')
                for n in range(0, len(ol_list)):
                    ordered_list(ol_list[n], n + 1)
            elif list[i] == '/ol' or list[i] == '/ul':
                continue
            else:
                rt.add(list[i])
        elif unclose_tag(list[i]) in tag:
            if i+1 < len(list) and list[i+1] not in tag:
                if unclose_tag(list[i+1]) not in tag:
                    rt.add(list[i+1])
        else:
            continue


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
    return clean_sliced_array


slicer(tagged_text)
executor(clean_sliced_array)


# Multiple tags

"""def tag_next_to_tag(list):
    i = 0
    l = 1
    extratag = []
    multitag = []
    print(len(list))
    while i <= len(list):
        if list[i] in tag:
            multitag.append(list[i])
            while i + l < len(list):
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

tag_next_to_tag(clean_sliced_array)"""

# ___________________________________________________________________________________________________________________
# Output
# -------------------------------------------------------------------------------------------------------------------
context = {
    'context_variable': rt,
}

tpl.render(context)
tpl.save('richtext.docx')


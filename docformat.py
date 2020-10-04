# encoding: utf-8
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.dml import MSO_THEME_COLOR_INDEX
from docx.shared import Inches
import io
import re
import docx
import requests
tstd = Document()
tstp = tstd.add_paragraph()
tstr = tstp.add_run()
bold = False
ital = False
atag = False
head = False
def main():
    f = open("ultest.html","r")
    soup = BeautifulSoup(f, 'lxml')
    print("encoding is {}".format(soup.original_encoding))
    for l in soup.descendants:
        pass
        #print(str(l))
    title = BeautifulSoup('<h1>This is a test</h1>', 'lxml')
    #document = buildDoc(soup, title)
    #document.save('C:\\Users\\etdeh\\Desktop\\test.docx')
    data = {'source':'https://www.zerohedge.com/markets/how-jpms-precious-metals-trading-desk-manipulated-markets-crime-ring-within-bank','title':title,'body':soup,'comments':'!'}
    document = buildZHdoc(**data)
    document.save('C:\\Users\\etdeh\\Desktop\\test.docx')
def buildZHdoc(**kwargs):
    source = kwargs['source']
    title = kwargs['title']
    body = kwargs['body']
    comments = kwargs['comments']
    document = Document()
    if source:
        para = document.add_paragraph()
        para.style = document.styles['Title']
        ttl = title.get_text()
        add_hyperlink(para, ttl, source, False, False, True)
    else:
        document.add_heading(title.get_text(), 0)
    main = body.find_all('div',class_="clearfix")[0].children
    for child in main:
        testChild(child,document)
    return document
def testChild(child,doc):
    print(child.name)
    if child.name == 'p':
        if type(doc) == type(tstd):
            paragraph = doc.add_paragraph()
            addP(child,paragraph)
        elif type(doc) == type(tstp):
            addP(child,doc)
        elif type(doc) == type(tstr):
            pass
    elif child.name == 'blockquote':
        if type(doc) == type(tstd):
            addBQ(child,doc)
        elif type(doc) == type(tstp):
            pass
        elif type(doc) == type(tstr):
            pass
    elif child.name == 'u':
        if type(doc) == type(tstd):
            pass
        elif type(doc) == type(tstp):
            pass
        elif type(doc) == type(tstr):
            doc.underline = True
    elif child.name == 'em':
        if type(doc) == type(tstd):
            pass
        elif type(doc) == type(tstp):
            pass
        elif type(doc) == type(tstr):
            doc.italic = True
    elif child.name == 'img':
        if type(doc) == type(tstd):
            pass
        elif type(doc) == type(tstp):
            pass
        elif type(doc) == type(tstr):
            pass
    elif child.name == 'a':
        if type(doc) == type(tstd):
            paragraph = doc.add_paragraph()
            addA(child,paragraph)
        elif type(doc) == type(tstp):
            addA(child,doc)
        elif type(doc) == type(tstr):
            pass
    elif child.name == 'picture':
        tmp = child.find_all('img')
        w = float(tmp[0]['width'])/96
        imgdata = requests.get(tmp[0]['src']).content
        imgfile = io.BytesIO(imgdata)
        if type(doc) == type(tstd):
            doc.add_picture(imgfile, width=Inches(w))
            print("got doc")
        elif type(doc) == type(tstp):
            run = doc.add_run()
            run.add_picture(imgfile, width=Inches(w))
            print("got par")
        elif type(doc) == type(tstr):
            doc.add_picture(imgfile, width=Inches(w))
            print("got run")
        else:
            print("got nothing")
    elif child.name == 'ul':
        if type(doc) == type(tstd):
            addUL(child,doc)
        elif type(doc) == type(tstp):
            pass
        elif type(doc) == type(tstr):
            pass
    elif child.name == 'ol':
        if type(doc) == type(tstd):
            pass
        elif type(doc) == type(tstp):
            pass
        elif type(doc) == type(tstr):
            pass
    elif child.name == 'iframe':
        if type(doc) == type(tstd):
            pass
        elif type(doc) == type(tstp):
            pass
        elif type(doc) == type(tstr):
            pass
def addUL(c,doc):
    for li in c.children:
        if li.name == None:
            pass
            #p.add_run(li)
        else:
            style = doc.styles['List Bullet']
            p = doc.add_paragraph("",style)
            ulr = p.add_run(li.get_text())
            if has_children(li):
                if li.contents[0].name != None:
                    testChild(li.contents[0],ulr)
def addBQ(c,doc):
    style = doc.styles['List 3']
    paragraph = doc.add_paragraph(style)
    for ch in c.children:
        testChild(ch, paragraph)
def addP(c,p):
    for child in c.children:
        if child.name == None:
            p.add_run(child)
        else:
            testChild(child,p)
def addA(c,p):
    bold = False
    ital = False
    url = c['href']
    for child in c.children:
        if child.name == None:
            pass
            add_hyperlink(p, c.get_text(), url, bold, ital, False)
        else:
            testChild(child,p)
def addPa(c,doc,style=None):
    global bold
    global ital
    global atag
    tmp = c.find_all('img')
    if len(tmp)>0:
        #print('picture here')
        w = float(tmp[0]['width'])/96
        imgdata = requests.get(tmp[0]['src']).content
        imgfile = io.BytesIO(imgdata)
        doc.add_picture(imgfile, width=Inches(w))
    else:
        if style:
            para = doc.add_paragraph(style=style)
        else:
            para = doc.add_paragraph()
        inc = 1
        for ch in c.children:
            print("ch name is {}".format(ch.name))
            if ch.name == None:
                para.add_run(ch)
            elif ch.name == 'br':
                para.add_run('\n')
            else:
                tst = ch
                testName(str(tst.name))
                while has_children(tst):
                    print("tst name is {}".format(tst.name))
                    if tst.name == 'a':
                        url = tst['href']
                    testName(tst.name)
                    try:
                        tst = tst.contents[0]
                    except:
                        tst = None
                if atag:
                    #print("test is {}".format(tst))
                    add_hyperlink(para, tst, url, bold, ital, False)
                else:
                    run = para.add_run(tst)
                    run.bold = bold
                    run.italic = ital
                bold = ital = atag = False
        print("next paragragh")
def testName(name):
    global bold
    global ital
    global atag
    if name == 'strong':
        bold = True
    elif name == 'em':
        ital = True
    elif name == 'a':
        atag = True
    #print("bold is {} and ital is {} and atag is {}".format(bold, ital, atag))
def has_children(element):
    try:
        t=element.contents
    except:
        return False
    else:
        return True
def add_hyperlink(paragraph, text, url, bflag, iflag, head):
    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )

    # Create a w:r element and a new w:rPr element
    new_run = docx.oxml.shared.OxmlElement('w:r')
    rPr = docx.oxml.shared.OxmlElement('w:rPr')

    # Join all the xml elements together add add the required text to the w:r element
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)

    # Create a new Run object and add the hyperlink into it
    r = paragraph.add_run ()
    r._r.append (hyperlink)
    r.bold = bflag
    r.italic = iflag
    
    # A workaround for the lack of a hyperlink style (doesn't go purple after using the link)
    # Delete this if using a template that has the hyperlink style in it
    r.font.color.theme_color = MSO_THEME_COLOR_INDEX.HYPERLINK
    r.font.underline = True
    if head:
       r.font.underline = False 
    return hyperlink




if __name__ == "__main__":
    main()

import docx
import textract
#specific to extracting information from word documents
import os
import shutil
import zipfile
#other tools useful in extracting the information from our document
import re
#to pretty print our xml:
import xml.dom.minidom
import re

def removeStartAndEnd(text):
    text = text[2:-1]
    return text

def removeXMLHeader(text):
    toRemove = re.search('<(.*)>',text)
    print(toRemove)
    #textNew = re.sub(toRemove,'',text)
    #print(textNew)

def quoteCheck(text):
    text = text.replace("\\xe2\\x80\\x9c","\"") #open "
    text = text.replace("\\xe2\\x80\\x9d","\"") #closed "
    return text

def extractContent(filename):
    with zipfile.ZipFile(filename, 'r') as zipObj:
    # Extract all the contents of zip file in different directory
        zipObj.extractall('temp')

def main():
    filename = 'MGMT 2150 Exam.docx'
    newText = docx.Document(filename)

    for t in newText.paragraphs:
        textT = t.text.encode('utf-8',"replace")
        textT = quoteCheck(str(textT))
        textT = removeStartAndEnd(str(textT))
        print(textT)

    for s in newText.inline_shapes:
        print(s._inline.graphic.graphicData.pic.nvPicPr.cNvPr.name)

    extractContent(filename)
    
    
    dom = xml.dom.minidom.parse("temp/word/document.xml") # or xml.dom.minidom.parseString(xml_string)
    
    #for i in range(len(dom.getElementsByTagName('w:p'))):
    #    style = dom.getElementsByTagName('w:pStyle')
    #    xmlText = dom.getElementsByTagName('w:t')
    #    stringOfText = xmlText[i].firstChild.nodeValue
    #    print(str(style[i].getAttribute('w:val')) + ":" + str(stringOfText.encode('utf-8')))

    x = dom.getElementsByTagName('w:p')[0]
    
    headingType = x.getElementsByTagName('w:pPr')[0].getElementsByTagName('w:pStyle')[0]
    textXML = x.getElementsByTagName('w:t')[0].firstChild.nodeValue.encode('utf-8')
    print(str(headingType.getAttribute('w.val')) + ":" + str(textXML))
        
    
    #checker = dom.getElementsByTagName('w:p')
    #print(checker[0].items())
    # pretty_xml_as_string = dom.toprettyxml()
    
    #print(re.search('<(.*)>',str(pretty_xml_as_string.encode("utf-8"))))
    input("Done?")
    #shutil.rmtree("temp")


main()


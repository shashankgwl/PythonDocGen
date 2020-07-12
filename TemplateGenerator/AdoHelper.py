from docx import Document
import bs4
import base64
import urllib.request
import binascii
import unicodedata
from bs4 import BeautifulSoup, NavigableString
from docxtpl import DocxTemplate, InlineImage
from azure.devops.v5_1.work_item_tracking.models import Wiql
from azure.devops.connection import Connection
from docxtpl import DocxTemplate, RichText
from msrest.authentication import BasicAuthentication
from HTMLStripper import HTMLStripper

def GetADOClient():
    credentials = BasicAuthentication('', personal_access_token)
    connection = Connection(base_url=organization_url, creds=credentials)
# Get a client (the "core" client provides access to projects, teams, etc)
    core_client = connection.clients.get_core_client()
    wit_client = connection.clients.get_work_item_tracking_client()
    return wit_client

def GetRichText(richElements):
        rt = RichText()
        for element in richElements:
            if(element.name == 'b'):
                rt.add(unicodedata.normalize('NFKD', element.text).encode('ascii', 'ignore'), bold=True)
                rt.add('\n')
            elif(element.name == 'i'):
                rt.add(unicodedata.normalize('NFKD', element.text).encode('ascii', 'ignore'), italic=True)
                rt.add('\n')
            elif(element.name == 'u'):
                rt.add(unicodedata.normalize('NFKD', element.text).encode('ascii', 'ignore'), underline=True)
                rt.add('\n')
            elif(isinstance(element,NavigableString)):
                rt.add(element)
            elif(element.name == 'br'):
                rt.add('\n')
            elif(element.name == 'ul'):
                rt.add(element.text)
            elif(element.name == 'img'):
                filename = element.attrs['src']
                rt.add('Embedded Image', url_id=doc.build_url_id(filename),bold=True,color='0000EE', underline=True)
                #requ = urllib.request.urlopen(filename)
                #imgBytes = requ.read();
                #pict =
                #"""{\pict\picscalex125\picscaley125\piccropl0\piccropr0\piccropt0\piccropb0\picw7789\pich9102\picwgoal4416\pichgoal5160\wmetafile8\bliptag-198017951"""
                #rt.add(pict);
            elif(element.name == 'a'):
                if 'href' in element.attrs:
                    href = element.attrs['href']
                    rt.add(element.text, url_id=doc.build_url_id(href),bold=True,color='0000EE',underline=True)
            elif(element.name == 'font' or element.name == 'p' or element.name == 'span' or element.name == 'div'):
                if element.attrs:
                    for keys in element.attrs:
                        if(keys == 'color'):
                            rt.add(unicodedata.normalize('NFKD', element.text + ' ').encode('ascii', 'ignore'),color=element.attrs[keys])
                        if(keys == 'style'):
                            rt.add(unicodedata.normalize('NFKD', element.text + ' ').encode('ascii', 'ignore'),style=element.attrs[keys])
                else:
                    rt.add(unicodedata.normalize('NFKD', element.text + ' ').encode('ascii', 'ignore'))
            elif (element.name == 'li'):
                continue
            else:
                rt.add(element.text)
        return rt

def EnrichNew(rawText):
        rt = RichText()
        richThings = []
        soup = BeautifulSoup(rawText, 'html.parser')
        soup.encode('utf-8')
        for tag in soup.descendants:
            if tag.name == 'br':
                richThings.append(tag)
            elif tag.name == 'img' or tag.name == 'a':
                richThings.append(tag)
            elif isinstance(tag,NavigableString):
                prev = tag.previous
                richThings.append(prev)
            else:
                continue
        
        return GetRichText(richThings)


def StripHtml(work_items):
    workItemArray = []
    stripper = HTMLStripper()
    for item in work_items:
        tempFields = {}
        desc = item.fields['System.Description']
        #kk = Enrich(desc)
        kk = EnrichNew(desc)
        tempFields['Description'] = kk
        #stripper.feed(item.fields['Microsoft.VSTS.Common.AcceptanceCriteria'])
                                                                             #tempFields['AcceptanceCriteria']
                                                                             #=
                                                                             #stripper.get_data()
        stripper.feed(item.fields['System.Title'])
        tempFields['Title'] = item.fields['System.Title']

        tempFields['ADOId'] = item.id
        #tempFields['Description'] = tempFields['System.Description']
        #tempFields['AcceptanceCr
        #iteria'] = tempFields['Microsoft.VSTS.Common.AcceptanceCriteria']
        #tempFields['Title'] = tempFields['System.Title']
        workItemArray.append(tempFields)
    
    return workItemArray



personal_access_token = 'jvoddawauhkxwzag3w2vgj64xfm54tlu6prnjxsk2k7q2buexkua'
organization_url = 'https://dev.azure.com/TELPSAADO'
wiqlQuery = Wiql(query= """select [System.Id],[System.WorkItemType],[System.Title], [System.Description] from WorkItems 
where [System.Id] in (158,160,162,167,168,169,170,173,175,178,181,202,204,208,211,212,214)""")
doc = DocxTemplate("C:\\Users\\SBHIDE\\OneDrive - Microsoft\\SHASHANK\\learning\\python\\tpl.docx")
#(8,32,33,35,36,55,56,57)
#(158,160,162,167,168,169,170,173,175,178,181,202,204,208,211,212,214)
wit_client = GetADOClient()
workItems = wit_client.query_by_wiql(wiqlQuery).work_items

work_items = (wit_client.get_work_item(int(res.id)) for res in workItems)
workItemArray = StripHtml(work_items)
finalDict = {}
finalDict['workItemArray'] = workItemArray

doc.render(finalDict)
doc.save("C:\\Users\\SBHIDE\\OneDrive - Microsoft\\SHASHANK\\learning\\python\\generated_doc.docx")

#workItemArray.append(item.fields.update({"Id":item.id}))
file = ''
from docx import Document
from docxtpl import DocxTemplate, RichText
import os
import json

def GenerateMainPage(document):
    doc = DocxTemplate("C:\\Users\\SBHIDE\\OneDrive - Microsoft\\SHASHANK\\learning\\python\\tpl.docx")
    #rt = RichText()

    contextNew = {
    "id":"34",
    "AreaPath": "Galicia Dynamics 365",
    "TeamProject": "Galicia Dynamics 365",
    "IterationPath": "Galicia Dynamics 365\\Iteration 31",
    "WorkItemType": "Requirement",
    "State": "Active",
    "Reason": "Moved to state Active",
    "AssignedTo": {
      "displayName": "Carlos Moncada",
      "url": "https://spsprodcus1.vssps.visualstudio.com/Ae2282259-e852-4507-9208-222832f13949/_apis/Identities/915fc108-4bca-648a-a237-8614e9ad296c",
      "_links": {
        "avatar": {
          "href": "https://dev.azure.com/BA-GaliciaD365/_apis/GraphProfile/MemberAvatars/aad.OTE1ZmMxMDgtNGJjYS03NDhhLWEyMzctODYxNGU5YWQyOTZj"
        }
      },
      "id": "915fc108-4bca-648a-a237-8614e9ad296c",
      "uniqueName": "v-camonc@microsoft.com",
      "imageUrl": "https://dev.azure.com/BA-GaliciaD365/_apis/GraphProfile/MemberAvatars/aad.OTE1ZmMxMDgtNGJjYS03NDhhLWEyMzctODYxNGU5YWQyOTZj",
      "descriptor": "aad.OTE1ZmMxMDgtNGJjYS03NDhhLWEyMzctODYxNGU5YWQyOTZj"
    },
    "CreatedDate": "2020-06-18T11:56:44.043Z",
    "CreatedBy": {
      "displayName": "Luciano Spiguel",
      "url": "https://spsprodcus1.vssps.visualstudio.com/Ae2282259-e852-4507-9208-222832f13949/_apis/Identities/e0b6973b-5c7a-6aaf-b86b-f585df08b472",
      "_links": {
        "avatar": {
          "href": "https://dev.azure.com/BA-GaliciaD365/_apis/GraphProfile/MemberAvatars/aad.ZTBiNjk3M2ItNWM3YS03YWFmLWI4NmItZjU4NWRmMDhiNDcy"
        }
      },
      "id": "e0b6973b-5c7a-6aaf-b86b-f585df08b472",
      "uniqueName": "v-luspig@microsoft.com",
      "imageUrl": "https://dev.azure.com/BA-GaliciaD365/_apis/GraphProfile/MemberAvatars/aad.ZTBiNjk3M2ItNWM3YS03YWFmLWI4NmItZjU4NWRmMDhiNDcy",
      "descriptor": "aad.ZTBiNjk3M2ItNWM3YS03YWFmLWI4NmItZjU4NWRmMDhiNDcy"
    },
    "ChangedDate": "2020-07-06T13:10:31.057Z",
    "ChangedBy": {
      "displayName": "Shashank Bhide",
      "url": "https://spsprodcus1.vssps.visualstudio.com/Ae2282259-e852-4507-9208-222832f13949/_apis/Identities/c4b537b1-f356-64a4-942f-556466746b31",
      "_links": {
        "avatar": {
          "href": "https://dev.azure.com/BA-GaliciaD365/_apis/GraphProfile/MemberAvatars/aad.YzRiNTM3YjEtZjM1Ni03NGE0LTk0MmYtNTU2NDY2NzQ2YjMx"
        }
      },
      "id": "c4b537b1-f356-64a4-942f-556466746b31",
      "uniqueName": "sbhide@microsoft.com",
      "imageUrl": "https://dev.azure.com/BA-GaliciaD365/_apis/GraphProfile/MemberAvatars/aad.YzRiNTM3YjEtZjM1Ni03NGE0LTk0MmYtNTU2NDY2NzQ2YjMx",
      "descriptor": "aad.YzRiNTM3YjEtZjM1Ni03NGE0LTk0MmYtNTU2NDY2NzQ2YjMx"
    },
    "CommentCount": 1,
    "Title": "Build services for \"Query Offer by IdHost\"",
    "BoardColumn": "Active",
    "BoardColumnDone": "False",
    "Microsoft.VSTS.Scheduling.StoryPoints": 5.0,
    "Microsoft.VSTS.Common.StateChangeDate": "2020-06-25T12:15:03.927Z",
    "Microsoft.VSTS.Common.Priority": 4,
    "Microsoft.VSTS.Common.StackRank": 1666666660.0,
    "Microsoft.VSTS.Common.Triage": "Pending",
    "Microsoft.VSTS.CMMI.Blocked": "No",
    "Microsoft.VSTS.CMMI.Committed": "No",
    "Microsoft.VSTS.CMMI.UserAcceptanceTest": "Not Ready",
    "Custom.Domain_MicrosoftServices": "Business Applications",
    "Custom.ProcessSequenceID": "SPOM-66",
    "Custom.PotentialISV_MicrosoftServices": "None",
    "Custom.RequirementCategory_MicrosoftServices": "Functional",
    "Custom.ReqReason_MicrosoftServices": "Accepted",
    "WEF_3DA1D7E0F139471D98B99FAD624A1FBB_Kanban.Column": "Active",
    "WEF_3DA1D7E0F139471D98B99FAD624A1FBB_Kanban.Column.Done": "False",
    "Description" : "this is the description",
    "ADescription": "*******************English******************************</span></div><div><span><span>It is required to migrate the current code that implements the Offer x IdHost Query endpoints from the .NET Framework in the Sales Architecture Front End project &quot;as is&quot; to .NET Core within the OpenShift platform and applying the POM archetype.<br></span><div><br></div><div>The endpoints are understood to be the following:<br></div><div><br></div><div>GET / People / {IdHost} / Offers<br></div><div>GET / People / {IdHost} / Offers / {Depth Level}<br></div><div>To replicate the response from the Query x IdHost Endpoints, use the original source code with the minimum dependencies necessary to have the Model classes and factories that process and build the responses.<br></div><span>In principle, it is interpreted that these queries should be exposed to the caller using at least the GET / v1 / sale-architecture / summary / simulate endpoint. Evaluate and verify if so, and extend as required.</span><br></span></div>",
    "AcceptanceCriteria": "<span><span>Responder desde el o los endpoints expuestos de \nsale-architecture de popr-loans, al menos a dos ejemplos de peticiones \nde IdHost del entorno de Integración de ADV. Puede utilizarse el cliente\n de Dynamics de Integración para identificar los ejemplos y respuestas \nválidas.</span></span>"
  }

    context = { 
             'Title' : "Contoso Corporation, Unified Service Desk Implementation" ,
             'preparedFor' : "Contoso Corp",
             'project' : 'Contoso CRM implementation',
             'Description' : 'This document describes the business requirements related to Unified Service Desk for the Ford CRM project. This document is the source of requirements for the design and subsequent implementation of USD  To obtain a complete understanding of the business requirements this document is recommended to be read along with the other documents mentioned in Appendix sections. The purpose of this document can be summarized as follows:',
             'purpose1' : 'It identifies and documents the design components of Ford USD implementation',
             'purpose2' : 'It forms the basis of the Design Document and ADO configuration/customization',
             'requirements' : [
                 {'reqId' : "1.0",'reqTitle':'Requirement title 1', 'reqDescription':'This is an amazing requirement'},
                 {'reqId' : "2.0",'reqTitle':'Requirement title 2', 'reqDescription':'This is an amazing requirement'},
                 {'reqId' : "3.0",'reqTitle':'Requirement title 3', 'reqDescription':'This is an amazing requirement'},
                 {'reqId' : "4.0",'reqTitle':'Requirement title 4', 'reqDescription':'This is an amazing requirement'}
             ],
             }

    #doc.render(contextNew,jinja_env=None, autoescape=True)
    #doc.save("generated_doc.docx")

os.chdir('C:\\Users\\SBHIDE\\OneDrive - Microsoft\\SHASHANK\\learning\\python');
print(os.getcwd())
document = Document()
#GenerateMainPage(document)
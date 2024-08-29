import os
import time
from configparser import ConfigParser, ExtendedInterpolation
import xlrd
import uuid
import csv
from bson.objectid import ObjectId
import json
from datetime import datetime
import requests
from difflib import get_close_matches
from requests import post, get, delete
import sys
import time
import shutil
from xlutils.copy import copy
import shutil
import re
from xlrd import open_workbook
from xlutils.copy import copy as xl_copy
import logging.handlers
import time
from logging.handlers import TimedRotatingFileHandler
import xlsxwriter
import argparse
import sys
from os import path
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.cell import Cell
from common_config import *
import threading
import wget
import gdown

# Global variable declaration
criteriaLookUp = dict()
millisecond = None
programNameInp = None
environment = None
observationId = None
solutionName = None
pointBasedValue = None
entityType = None
allow_multiple_submissions = None
scopeEntityType = ""
programName = None
userEntity = None
roles = ""
mainRole = ""
dictCritLookUp = {}
isProgramnamePresent = None
solutionLanguage = None
keyWords = None
entityTypeId = None
solutionDescription = None
creator = None
dikshaLoginId = None
criteriaName = None
solutionId = None
API_log = None
listOfFoundRoles = []
entityToUpload = None
programID = None
programExternalId = None
programDescription = None
criteriaLookUp = dict()
themesSheetList = []
themeRubricFileObj = dict()
criteriaLevelsReport = False
ecm_sections = dict()
criteriaLevelsCount = 0
numberOfResponses = 0
criteriaIdNameDict = dict()
criteriaLevels = list()
matchedShikshalokamLoginId = None
scopeEntities = []
scopeRoles = []
countImps = 0
ecmToSection = dict()
entitiesPGM = []
entitiesPGMID = []
solutionRolesArr = []
startDateOfResource = None
endDateOfResource = None
startDateOfProgram = None
endDateOfProgram = None
rolesPGM =None
solutionRolesArray = []
solutionStartDate = ""
solutionEndDate = ""
projectCreator = ""
orgIds = []
OrgName = []
ccRootOrgName = None
ccRootOrgId  = None
certificatetemplateid = None
question_sequence_arr = []
regex = "\"?([-a-zA-Z0-9.`?{}]+@\w+\.\w+)\"?"
class Helpers:
    def __init__(self):
        self.millisecond = None
        self.scopeEntityType = ""

    def programCreation(accessToken, parentFolder, externalId, pName, pDescription, keywords, entities, roles, orgIds,creatorKeyCloakId, creatorName,entitiesPGM,mainRole,rolesPGM):
        messageArr = []
        messageArr.append("++++++++++++ Program Creation ++++++++++++")
        # program creation url 
        programCreationurl =  internal_kong_ip + programcreationurl
        messageArr.append("Pogram Creation URL : " + programCreationurl)
        # program creation payload
        payload = json.dumps({
            "externalId": externalId,
            "name": pName,
            "description": pDescription,
            "resourceType": [
                "program"
            ],
            "language": [
                "English"
            ],
            "keywords": keywords,
            "concepts": [],
            "createdFor": orgIds,
            "rootOrganisations": orgIds,
            "startDate": startDateOfProgram,
            "endDate": endDateOfProgram,
            "imageCompression": {
                "quality": 10
            },
            "creator": creatorName,
            "owner": creatorKeyCloakId,
            "author": creatorKeyCloakId,
            "scope": {
                "entityType": scopeEntityType,
                "entities": entitiesPGMID,
                "roles": roles
            },
            "metaInformation": {
                "state":entitiesPGM.split(","),
                "roles": mainRole.split(",")
                },
                "requestForPIIConsent":True
                })
        messageArr.append("Body : " + str(payload))
        headers = {'X-authenticated-user-token': accessToken,
                   'internal-access-token': internal_access_token,
                   'Content-Type': 'application/json',
                   'Authorization':authorization}

        # program creation 
        responsePgmCreate = requests.request("POST", programCreationurl, headers=headers, data=(payload))
        messageArr.append("Program Creation Status Code : " + str(responsePgmCreate.status_code))
        messageArr.append("Program Creation Response : " + str(responsePgmCreate.text))
        messageArr.append("Program body : " + str(payload))

        # save logs 
        # createAPILog(parentFolder, messageArr)
        # # check status 
        # fileheader = [pName, ('Program Sheet Validation'), ('Passed')]
        # createAPILog(parentFolder, messageArr)
        # apicheckslog(parentFolder, fileheader)
        print(responsePgmCreate.text,responsePgmCreate)
        if responsePgmCreate.status_code == 200:
            responsePgmCreateResp = responsePgmCreate.json()
        else:
            # terminate execution
            print("Program creation API failed. Please check logs.")


    def programmappingpdpmsheetcreation(MainFilePath,accessToken, program_file,programexternalId,parentFolder):
        pdpmsheet = MainFilePath+ "/pdpmmapping/"
        if not os.path.exists(pdpmsheet):
            os.mkdir(pdpmsheet)

        wbproject = xlrd.open_workbook(program_file, on_demand=True)
        projectSheetNames = wbproject.sheet_names()

        mappingsheet = wbproject.sheet_by_name('Program Details')
        keysProject = [mappingsheet.cell(1, col_index_env).value for col_index_env in
                       range(mappingsheet.ncols)]

        pdpmcolo1 = ["user","role","entity","entityOperation","keycloak-userId","acl_school","acl_cluster","programOperation",
                    "platform_role","programs","_arrayFields"]
        with open(pdpmsheet + 'mapping.csv', 'w',encoding='utf-8') as file:
             writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
             writer.writerows([pdpmcolo1])

        wbPgm = xlrd.open_workbook(program_file, on_demand=True)
        global programNameInp
        sheetNames = wbPgm.sheet_names()
        for sheetEnv in sheetNames:
            if sheetEnv == "Instructions":
                pass
            elif sheetEnv.strip().lower() == 'program details':
                print("--->Checking Program details sheet...")
                detailsEnvSheet = wbPgm.sheet_by_name(sheetEnv)
                keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                           range(detailsEnvSheet.ncols)]
                for row_index_env in range(2, detailsEnvSheet.nrows):
                    dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                      for
                                      col_index_env in range(detailsEnvSheet.ncols)}
                    programNameInp = dictDetailsEnv['Title of the Program'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Title of the Program'] else terminatingMessage("\"Title of the Program\" must not be Empty in \"Program details\" sheet")

                extIdPGM = dictDetailsEnv['Program ID'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Program ID'] else terminatingMessage("\"Program ID\" must not be Empty in \"Program details\" sheet")

                programdesigner = dictDetailsEnv['Diksha username/user id/email id/phone no. of Program Designer'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Program ID'] else terminatingMessage("\"Diksha username/user id/email id/phone no. of Program Designer\" must not be Empty in \"Program details\" sheet")
                userDetails = Helpers.fetchUserDetails(accessToken, programdesigner)

                creatorKeyCloakId = userDetails[0]
                creatorName = userDetails[1]
                if "PROGRAM_DESIGNER" in userDetails[3]:
                    creatorKeyCloakId = userDetails[0]
                    creatorName = userDetails[1]
                else :
                    print("user does't have program designer role")

                pdpmcolo1 = [creatorName, " ", " ", " ", creatorKeyCloakId, " ", " ","ADD","PROGRAM_DESIGNER", extIdPGM, "programs"]
                with open(pdpmsheet + 'mapping.csv', 'a',encoding='utf-8') as file:
                    writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                    writer.writerows([pdpmcolo1])
                    fileheader = [creatorName,"program designer mapped successfully","Passed"]
                    # apicheckslog(parentFolder,fileheader)


            elif sheetEnv.strip().lower() == 'program manager details':
                print("--->Program Manager Details...")
                detailsEnvSheet = wbPgm.sheet_by_name(sheetEnv)
                keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                           range(detailsEnvSheet.ncols)]
                for row_index_env in range(2, detailsEnvSheet.nrows):
                    dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                      for
                                      col_index_env in range(detailsEnvSheet.ncols)}

                    if str(dictDetailsEnv['Is a SSO user?']).strip() == "YES":
                        programmanagername2 = dictDetailsEnv['Diksha user id ( profile ID)'] if dictDetailsEnv['Diksha user id ( profile ID)'] else terminatingMessage("\"Diksha user id ( profile ID)\" must not be Empty in \"Program details\" sheet")
                    else:
                        try :
                            programmanagername2 = dictDetailsEnv['Login ID on DIKSHA'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Login ID on DIKSHA'] else terminatingMessage("\"Login ID on DIKSHA\" must not be Empty in \"Program details\" sheet")
                            userDetails = Helpers.fetchUserDetails(accessToken, programmanagername2)
                        except :
                            programmanagername2 = dictDetailsEnv['Diksha user id ( profile ID)'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Diksha user id ( profile ID)'] else terminatingMessage("\"Diksha user id ( profile ID)\" must not be Empty in \"Program details\" sheet")
                            userDetails = Helpers.fetchUserDetails(accessToken, programmanagername2)

                    userDetails = Helpers.fetchUserDetails(accessToken, programmanagername2)
                    creatorKeyCloakId = userDetails[0]
                    creatorName = userDetails[1]
                    if "PROGRAM_MANAGER" in userDetails[3]:
                        creatorKeyCloakId = userDetails[0]
                        creatorName = userDetails[1]
                    else:
                        print("user does't have program manager role")

                    pdpmcolo1 = [creatorName, " ", " ", " ", creatorKeyCloakId, " ", " ","ADD","PROGRAM_MANAGER", extIdPGM, "programs"]

                    with open(pdpmsheet + 'mapping.csv', 'a',encoding='utf-8') as file:
                        writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                        writer.writerows([pdpmcolo1])
                    # messageArr.append("Response : " + str(pdpmcolo1))
                    # createAPILog(parentFolder, messageArr)

                    fileheader = [creatorName,"program manager mapped succesfully","Passed"]
                    # apicheckslog(parentFolder,fileheader)


    # this function is used for call the api and map the pdpm roles which we created
    def Programmappingapicall(MainFilePath,accessToken, program_file,parentFolder):
        urlpdpmapi = internal_kong_ip + pdpmurl
        headerpdpmApi = {
            'Authorization':authorization,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token
        }
        payload = {}
        filesProject = {
            'userRoles': open(MainFilePath + '/pdpmmapping/mapping.csv', 'rb')
        }

        responseProgrammappingApi = requests.post(url=urlpdpmapi, headers=headerpdpmApi,
                                                 data=payload,
                                                 files=filesProject)
        messageArr = ["program mapping sheet.",
                      "File path : " + MainFilePath + '/pdpmmapping/mapping.csv']
        messageArr.append("Upload status code : " + str(responseProgrammappingApi.status_code))
        # createAPILog(parentFolder, messageArr)

        if responseProgrammappingApi.status_code == 200:
            print('--->program manager and designer mapping is Success')
            with open(MainFilePath + '/pdpmmapping/mappinginternal.csv', 'w+',encoding='utf-8') as projectRes:
                projectRes.write(responseProgrammappingApi.text)
                messageArr.append("Response : " + str(responseProgrammappingApi.text))
                # createAPILog(parentFolder, messageArr)
        else:
            messageArr.append("Response : " + str(responseProgrammappingApi.text))
            # createAPILog(parentFolder, messageArr)
            # fileheader = ["PDPM mapping","PDPM mapping is failed","Failed","check PDPM sheet"]
            # apicheckslog(parentFolder,fileheader)
            sys.exit()


    
    def createFileStructForProgram(programFile):
        print("programFile:-------------",programFile)
        if not os.path.isdir('programFiles'):
            os.mkdir('programFiles')
        if "/" in str(programFile):
            fileNameSplit = str(programFile).split('/')[-1:]
            print("newfileNameSplit ;",fileNameSplit)
        else :

            fileNameSplit = os.path.basename(programFile)
        print("updatedfileNameSplit : ",fileNameSplit)

        if isinstance(fileNameSplit, list):
            fileNameSplit = fileNameSplit[0]
        # fileNameSplit = str(programFile)
        if fileNameSplit.endswith(".xlsx"):
            print("latest fileNameSplit",fileNameSplit)
            ts = str(time.time()).replace(".", "_")
            
            folderName = fileNameSplit.replace(".xlsx", "-" + str(ts))
            print("folderName :",folderName)
            os.mkdir('programFiles/' + str(folderName))
            path = os.path.join('programFiles', str(folderName))
            print("done",path)
        else:
            print("something")
        returnPathStr = os.path.join('programFiles', str(folderName))
        print("returnPathStr",returnPathStr)

        return returnPathStr
    
    def fetchScopeRole(solutionName_for_folder_path, accessToken, roleNameList):
        urlFetchRolesListApi = internal_kong_ip + listofrolesapi
        headerFetchRolesListApi = {
            'Content-Type':  content_type,
            'Authorization': authorization,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id': x_channel_id,
        }
        responseFetchRolesListApi = requests.post(url=urlFetchRolesListApi, headers=headerFetchRolesListApi)
        rolesLookup = dict()
        rolesReturn = list()
        messageArr = ["Roles list fetch API called.", "URL  : " + str(urlFetchRolesListApi),
                      "Status Code : " + str(responseFetchRolesListApi.status_code)]
        # createAPILog(solutionName_for_folder_path, messageArr)
        if responseFetchRolesListApi.status_code == 200:
            responseFetchRolesListApi = responseFetchRolesListApi.json()
            for listRoles in responseFetchRolesListApi['result']:
                eachDict = dict()
                eachDict['id'] = listRoles['_id'].lstrip().rstrip()
                eachDict['code'] = listRoles['code'].lstrip().rstrip()
                rolesLookup[listRoles['code']] = eachDict['id']
                rolesReturn.append(listRoles['code'].lstrip().rstrip())
        else:
            print("---> error in subroles API.")

        userRolesFromInp = roleNameList
        listOfFoundRoles = list()
        if len(userRolesFromInp) == 0:
            print("Roles fields must not be empty.")
        for ur in userRolesFromInp:
            rolesFlag = True
            try:
                roleDetails = rolesLookup[ur.lstrip().rstrip()]
                rolesFlag = True
            except:
                rolesFlag = False

            if rolesFlag:
                print("Role Found... : " + ur)
                listOfFoundRoles.append(ur)
            else:
                if "all" in userRolesFromInp:
                    listOfFoundRoles = ["all"]
                else:
                    print("Role error...")
                    print("Role : " + ur)
                    messageArr = ["Roles Error", "URL  : ", "Role : " + ur]
                    # createAPILog(solutionName_for_folder_path, messageArr)

        messageArr = ["Accepted Roles : " + str(listOfFoundRoles)]
        # createAPILog(solutionName_for_folder_path, messageArr)
        if len(listOfFoundRoles) == 0:
            messageArr = ["No roles matched our DB "]
            # createAPILog(solutionName_for_folder_path, messageArr)
            print("No Roles matched our DB.")
        return listOfFoundRoles


    
    def getProgramInfo(accessTokenUser, solutionName_for_folder_path, programNameInp):
        global programID, programExternalId, programDescription, isProgramnamePresent, programName
        programName = programNameInp
        programUrl = internal_kong_ip + fetchprograminfoapiurl
        # print(programUrl)
        payload = json.dumps({
            "query": {
                "name": programNameInp.lstrip().rstrip(),
                "isAPrivateProgram": False,
                "status": "active"
                },
                "mongoIdKeys": []
                })
        
        # print(programNameInp,"programNameInp")

        headersProgramSearch = {'Authorization': authorization,
                                'Content-Type': 'application/json', 'X-authenticated-user-token': accessTokenUser,
                                'internal-access-token': internal_access_token}
        responseProgramSearch = requests.post(url=programUrl, headers=headersProgramSearch,data=payload)
        # print(responseProgramSearch.text)
        messageArr = []

        if responseProgramSearch.status_code == 200:
            print('--->Program fetch API Success')
            messageArr.append("--->Program fetch API Success")
            responseProgramSearch = responseProgramSearch.json()
            countOfPrograms = len(responseProgramSearch['result'])
            messageArr.append("--->Program Count : " + str(countOfPrograms))
            if countOfPrograms == 0:
                messageArr.append("No program found with the name : " + str(programName.lstrip().rstrip()))
                messageArr.append("******************** Preparing for program Upload **********************")
                print("No program found with the name : " + str(programName.lstrip().rstrip()))
                print("******************** Preparing for program Upload **********************")
                
                return False
            else:
                getProgramDetails = []
                for eachPgm in responseProgramSearch['result']:
                    if eachPgm['isAPrivateProgram'] == False:
                        programID = eachPgm['_id']
                        programExternalId = eachPgm['externalId']
                        programDescription = eachPgm['description']
                        isAPrivateProgram = eachPgm['isAPrivateProgram']
                        getProgramDetails.append([programID, programExternalId, programDescription, isAPrivateProgram])
                        if len(getProgramDetails) == 0:
                            print("Total " + str(len(getProgramDetails)) + " backend programs found with the name : " + programName.lstrip().rstrip())
                            messageArr.append("Total " + str(len(getProgramDetails)) + " backend programs found with the name : " + programName.lstrip().rstrip())
                            
                            fileheader = ["program find api is running","found"+str(len(
                                getProgramDetails))+"programs in backend","Failed","found"+str(len(
                                getProgramDetails))+"programs ,check logs"]
                           
                        elif len(getProgramDetails) > 1:
                            print("Total " + str(len(getProgramDetails)) + " backend programs found with the name : " + programName.lstrip().rstrip())
                            messageArr.append("Total " + str(len(getProgramDetails)) + " backend programs found with the name : " + programName.lstrip().rstrip())
                           

                        else:
                            programID = getProgramDetails[0][0]
                            programExternalId = getProgramDetails[0][1]
                            programDescription = getProgramDetails[0][2]
                            isAPrivateProgram = getProgramDetails[0][3]
                            isProgramnamePresent = True
                            messageArr.append("programID : " + str(programID))
                            messageArr.append("programExternalId : " + str(programExternalId))
                            messageArr.append("programDescription : " + str(programDescription))
                            messageArr.append("isAPrivateProgram : " + str(isAPrivateProgram))
                        # createAPILog(solutionName_for_folder_path, messageArr)
        else:
            print("Program search API failed...")
            messageArr.append("Program search API failed...")
            # createAPILog(solutionName_for_folder_path, messageArr)
            # terminatingMessage("Response Code : " + str(responseProgramSearch.status_code))
        return True
    
    def fetchEntityId(solutionName_for_folder_path, accessToken, entitiesNameList, scopeEntityType):
        print(scopeEntityType,"scopeEntityType--------------")
        urlFetchEntityListApi = host+searchforlocation
        headerFetchEntityListApi = {
            'Content-Type': content_type,
            'Authorization': authorizationforhost,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id': x_channel_id,
        }
        payload = {
            "request": {
                "filters": {
                    "type": scopeEntityType
                },
                "limit": 1000
            }
        }
        responseFetchEntityListApi = requests.post(url=urlFetchEntityListApi, headers=headerFetchEntityListApi,data=json.dumps(payload))
        print(responseFetchEntityListApi,"responseFetchEntityListApi")

        messageArr = ["Entities List Fetch API executed.", "URL  : " + str(urlFetchEntityListApi),
                      "Status : " + str(responseFetchEntityListApi.status_code)]
       
        if responseFetchEntityListApi.status_code == 200:
            responseFetchEntityListApi = responseFetchEntityListApi.json()
            entitiesLookup = dict()
            entityToUpload = list()
            for listEntities in responseFetchEntityListApi['result']['response']:
                entitiesLookup[listEntities['name'].lower().lstrip().rstrip()] = listEntities['id'].lstrip().rstrip()
            entitiesFlag = False
            for eachUserEntity in entitiesNameList:
                try:
                    entityId = entitiesLookup[eachUserEntity.lower().lstrip().rstrip()]
                    entitiesFlag = True
                except:
                    entitiesFlag = False
                if entitiesFlag:
                    entityToUpload.append(entityId)
                else:
                    print("Entity Not found in DB...")
                    print("Entity name : " + str(eachUserEntity))
                    messageArr = ["Entity Not found : ", "URL  : " + str(eachUserEntity)]
                    

            messageArr = ["Entities to upload : " + str(entityToUpload)]
            
            if len(entityToUpload) == 0:
                print("--->Scope Entity error.")
            return entityToUpload
        else:
            messageArr = ["Error in Location search",str(responseFetchEntityListApi.status_code)]
            createAPILog(solutionName_for_folder_path, messageArr)
            terminatingMessage("---> Error in location search.")

    
    def programsFileCheck(filePathAddPgm, accessToken, parentFolder, MainFilePath):
        program_file = filePathAddPgm
        # open excel file 
        wbPgm = xlrd.open_workbook(filePathAddPgm, on_demand=True)
        global programNameInp
        sheetNames = wbPgm.sheet_names()
        # list of sheets in the program sheet 
        pgmSheets = ["Instructions", "Program Details", "Resource Details","Program Manager Details"]

        # checking the sheets in the program sheet 
        if (len(sheetNames) == len(pgmSheets)) and ((set(sheetNames) == set(pgmSheets))):
            print("--->Program Template detected.<---")
            # iterate through the sheets 
            for sheetEnv in sheetNames:

                if sheetEnv == "Instructions":
                    # skip Instructions sheet 
                    pass
                elif sheetEnv.strip().lower() == 'program details':
                    print("--->Checking Program details sheet...")
                    detailsEnvSheet = wbPgm.sheet_by_name(sheetEnv)
                    keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                               range(detailsEnvSheet.ncols)]
                    for row_index_env in range(2, detailsEnvSheet.nrows):
                        dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                          for
                                          col_index_env in range(detailsEnvSheet.ncols)}
                        programNameInp = dictDetailsEnv['Title of the Program'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Title of the Program'] else terminatingMessage("\"Title of the Program\" must not be Empty in \"Program details\" sheet")
                        extIdPGM = dictDetailsEnv['Program ID'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Program ID'] else terminatingMessage("\"Program ID\" must not be Empty in \"Program details\" sheet")
                        returnvalues = []
                        global entitiesPGM
                        entitiesPGM = dictDetailsEnv['Targeted state at program level'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Targeted state at program level'] else terminatingMessage("\"Targeted state at program level\" must not be Empty in \"Program details\" sheet")
                        districtentitiesPGM = dictDetailsEnv['Targeted district at program level'].encode('utf-8').decode('utf-8')
                        global startDateOfProgram, endDateOfProgram
                        startDateOfProgram = dictDetailsEnv['Start date of program']
                        endDateOfProgram = dictDetailsEnv['End date of program']
                        # taking the start date of program from program template and converting YYYY-MM-DD 00:00:00 format

                        startDateArr = str(startDateOfProgram).split("-")
                        startDateOfProgram = startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"

                        # taking the end date of program from program template and converting YYYY-MM-DD 00:00:00 format

                        endDateArr = str(endDateOfProgram).split("-")
                        endDateOfProgram = endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"

                        global scopeEntityType
                        scopeEntityType = "state"


                        if districtentitiesPGM:
                            entitiesPGM = districtentitiesPGM
                            EntityType = "district"
                        else:
                            entitiesPGM = entitiesPGM
                            EntityType = "state"

                        scopeEntityType = EntityType

                        global entitiesPGMID
                        print(entitiesPGMID,"entitiesPGMID")
                        entitiesPGMID = Helpers.fetchEntityId(parentFolder, accessToken,
                                                      entitiesPGM.lstrip().rstrip().split(","), scopeEntityType)
                        print(entitiesPGMID)
                        global orgIds



                        if not Helpers.getProgramInfo(accessToken, parentFolder, programNameInp.encode('utf-8').decode('utf-8')):
                            print("reached till here")
                            extIdPGM = dictDetailsEnv['Program ID'].encode('utf-8').decode('utf-8')
                            if str(dictDetailsEnv['Program ID']).strip() == "Do not fill this field":
                                print("change the program id")
                            descriptionPGM = dictDetailsEnv['Description of the Program'].encode('utf-8').decode('utf-8')
                            keywordsPGM = dictDetailsEnv['Keywords'].encode('utf-8').decode('utf-8')
                            entitiesPGM = dictDetailsEnv['Targeted state at program level'].encode('utf-8').decode('utf-8') 
                            districtentitiesPGM = dictDetailsEnv['Targeted district at program level'].encode('utf-8').decode('utf-8')
                            # selecting entity type based on the users input 
                            if districtentitiesPGM:
                                entitiesPGM = districtentitiesPGM
                                EntityType = "district"
                            else:
                                entitiesPGM = entitiesPGM
                                EntityType = "state"

                            scopeEntityType = EntityType

                            mainRole = dictDetailsEnv['Targeted role at program level'] 
                            # print(mainRole,"mainRole")
                            global rolesPGM
                            rolesPGM = dictDetailsEnv['Targeted subrole at program level']
                            # print(rolesPGM,rolesPGM)

                            if "teacher" in mainRole.strip().lower():
                                rolesPGM = str(rolesPGM).strip() + ",TEACHER"
                            userDetails = Helpers.fetchUserDetails(accessToken, dictDetailsEnv['Diksha username/user id/email id/phone no. of Program Designer'])
                            OrgName=userDetails[4]
                            # print(OrgName,"OrgName")
                            orgIds=Helpers.fetchOrgId(accessToken, parentFolder, OrgName)
                            print(orgIds,"orgIds")
                            creatorKeyCloakId = userDetails[0]
                            creatorName = userDetails[2]

                            messageArr = []

                            scopeEntityType = EntityType
                            # fetch entity details 
                            entitiesPGMID = Helpers.fetchEntityId(parentFolder, accessToken,entitiesPGM.lstrip().rstrip().split(","), scopeEntityType)
                            print(entitiesPGMID,"entitiesPGMID")

                            # sys.exit()
                            # fetch sub-role details 
                            rolesPGMID = Helpers.fetchScopeRole(parentFolder, accessToken, rolesPGM.lstrip().rstrip().split(","))
                            print(rolesPGMID,"rolesPGMID")

                            # sys.exit()

                            # call function to create program 
                            Helpers.programCreation(accessToken, parentFolder, extIdPGM, programNameInp, descriptionPGM,keywordsPGM.lstrip().rstrip().split(","), entitiesPGMID, rolesPGMID, orgIds,creatorKeyCloakId, creatorName,entitiesPGM,mainRole,rolesPGM)
                            # sys.exit()
                            Helpers.programmappingpdpmsheetcreation(MainFilePath, accessToken, program_file, extIdPGM,parentFolder)

                            # map PM / PD to the program 
                            Helpers.Programmappingapicall(MainFilePath, accessToken, program_file,parentFolder)

                            # check if program is created or not 
                            if Helpers.getProgramInfo(accessToken, parentFolder, programNameInp):
                                print("Program Created SuccessFully.")
                            else :
                                print("Program creation failed! Please check logs.")

                elif sheetEnv.strip().lower() == 'resource details':
                    # checking Resource details sheet 
                    print("--->Checking Resource Details sheet...")
                    detailsEnvSheet = wbPgm.sheet_by_name(sheetEnv)
                    keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                               range(detailsEnvSheet.ncols)]
                    # iterate through each row in Resource Details sheet and validate 
                    for row_index_env in range(2, detailsEnvSheet.nrows):
                        dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                          for
                                          col_index_env in range(detailsEnvSheet.ncols)}
                        resourceNamePGM = dictDetailsEnv['Name of resources in program'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Name of resources in program'] else terminatingMessage("\"Name of resources in program\" must not be Empty in \"Resource Details\" sheet")
                        resourceTypePGM = dictDetailsEnv['Type of resources'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Type of resources'] else terminatingMessage("\"Type of resources\" must not be Empty in \"Resource Details\" sheet")
                        resourceLinkOrExtPGM = dictDetailsEnv['Resource Link']
                        resourceStatusOrExtPGM = dictDetailsEnv['Resource Status'] if dictDetailsEnv['Resource Status'] else terminatingMessage("\"Resource Status\" must not be Empty in \"Resource Details\" sheet")
                        # setting start and end dates globally. 
                        global startDateOfResource, endDateOfResource
                        startDateOfResource = dictDetailsEnv['Start date of resource']
                        endDateOfResource = dictDetailsEnv['End date of resource']
                        # checking resource types and calling relevant functions 
                        if resourceTypePGM.lstrip().rstrip().lower() == "course":
                            coursemapping = courseMapToProgram(accessToken, resourceLinkOrExtPGM, parentFolder)
                            if startDateOfResource:
                                startDateArr = str(startDateOfResource).split("-")
                                bodySolutionUpdate = {"startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                                solutionUpdate(parentFolder, accessToken, coursemapping, bodySolutionUpdate)
                            if endDateOfResource:
                                endDateArr = str(endDateOfResource).split("-")
                                bodySolutionUpdate = {
                                    "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                                solutionUpdate(parentFolder, accessToken, coursemapping, bodySolutionUpdate)
                        


# Function create File structure for Solutions
    def createFileStruct(MainFilePath, addSolutionFile):
        if not os.path.isdir(MainFilePath + '/SolutionFiles'):
            os.mkdir(MainFilePath + '/SolutionFiles')
        if "\\" in str(addSolutionFile):
            fileNameSplit = str(addSolutionFile).split('\\')[-1:]
        elif "/" in str(addSolutionFile):
            fileNameSplit = str(addSolutionFile).split('/')[-1:]
        else:
            fileNameSplit = str(addSolutionFile)
        if ".xlsx" in str(fileNameSplit[0]):
            ts = str(time.time()).replace(".", "_")
            folderName = fileNameSplit[0].replace(".xlsx", "-" + str(ts))
            os.mkdir(MainFilePath + '/SolutionFiles/' + str(folderName))
            path = os.path.join(MainFilePath + '/SolutionFiles', str(folderName))
            path = os.path.join(path, str('apiHitLogs'))
            os.mkdir(path)
        else:
            print("File Error.offff")
        returnPathStr = os.path.join(MainFilePath + '/SolutionFiles', str(folderName))

        if not os.path.isdir(returnPathStr + "/user_input_file"):
            os.mkdir(returnPathStr + "/user_input_file")

        shutil.copy(addSolutionFile, os.path.join(returnPathStr + "user_input_file.xlsx"))
        # shutil.copy(programFile, os.path.join(returnPathStr + "user_input_file"))
        return returnPathStr
    

    # Generate access token for the APIs. 
    def generateAccessToken(solutionName_for_folder_path):
    # production search user api - start
        headerKeyClockUser = {'Content-Type': keyclockapicontent_type}
    
        responseKeyClockUser = requests.post(url=host + (keyclockapiurl), headers=headerKeyClockUser,
                                         data=(keyclockapibody))
        print(responseKeyClockUser)
    
        if responseKeyClockUser.status_code == 200:
            responseKeyClockUser = responseKeyClockUser.json()
            accessTokenUser = responseKeyClockUser['access_token']
            print("--->Access Token Generated!")
        return accessTokenUser


    def checkEmailValidation(email):
        if (re.search(regex, email)):
            return True
        else:
            return False
    def fetchUserDetails(accessToken, dikshaId):
        global OrgName,creatorId
        url =  host + userinfoapiurl
        headers = {'Content-Type': 'application/json',
               'Authorization': authorizationforhost,
               'X-authenticated-user-token': accessToken}
        isEmail = Helpers.checkEmailValidation(dikshaId.lstrip().rstrip())
        
        if isEmail:
            body = "{\n  \"request\": {\n    \"filters\": {\n    \t\"email\": \"" + dikshaId.lstrip().rstrip() + "\"\n    },\n      \"fields\" :[],\n    \"limit\": 1000,\n    \"sort_by\": {\"createdDate\": \"desc\"}\n  }\n}"
        else:
            body = "{\n  \"request\": {\n    \"filters\": {\n    \t\"userName\": \"" + dikshaId.lstrip().rstrip() + "\"\n    },\n      \"fields\" :[],\n    \"limit\": 1000,\n    \"sort_by\": {\"createdDate\": \"desc\"}\n  }\n}"

        responseUserSearch = requests.request("POST", url, headers=headers, data=body)
        response_json = responseUserSearch.json()
        # print(responseUserSearch.text)
        # print(json.dumps(response_json, indent=4))
        # sys.exit()
        # print(responseUserSearch)
        if responseUserSearch.status_code == 200:
            responseUserSearch = responseUserSearch.json()
            if responseUserSearch['result']['response']['content']:
                userKeycloak = responseUserSearch['result']['response']['content'][0]['userId']
                creatorId = userKeycloak
                userName = responseUserSearch['result']['response']['content'][0]['userName']
                firstName = responseUserSearch['result']['response']['content'][0]['firstName']
                rootOrgId = responseUserSearch['result']['response']['content'][0]['rootOrgId']
                for index in responseUserSearch['result']['response']['content'][0]['organisations']:
                    if rootOrgId == index['organisationId']:
                        roledetails = index['roles']
                        rootOrgName = index['orgName']
                        OrgName.append(rootOrgName)
                print(roledetails)
            else:
                print("-->Given username/email is not present in KB platform<--.")
        else:
            print(responseUserSearch.text)

        return [userKeycloak, userName, firstName,roledetails,rootOrgName,rootOrgId]
    
    def SolutionFileCheck(filePathAddPgm, accessToken, parentFolder, MainFilePath):
        global creatorId,solutionNameForSuccess
        wbPgm = xlrd.open_workbook(filePathAddPgm, on_demand=True)
        global solutionNameInp
        sheetNames = wbPgm.sheet_names()
        for sheetEnv in sheetNames:
            if sheetEnv.strip().lower() == 'details':
                print("--->Checking resource details sheet...")
                detailsEnvSheet = wbPgm.sheet_by_name(sheetEnv)
                keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                            range(detailsEnvSheet.ncols)]
                for row_index_env in range(2, detailsEnvSheet.nrows):
                    dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                        for
                                        col_index_env in range(detailsEnvSheet.ncols)}
                    solutionNameInp = dictDetailsEnv['solution_name'].encode('utf-8').decode('utf-8')
                    solutionNameForSuccess = solutionNameInp
                    global entitiesPGM

                    global startDateOfProgram, endDateOfProgram
                    startDateOfProgram = dictDetailsEnv['start_date']
                    endDateOfProgram = dictDetailsEnv['end_date']

                    # taking the start date of program from program template and converting YYYY-MM-DD 00:00:00 format

                    startDateArr = str(startDateOfProgram).split("-")
                    startDateOfProgram = startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"

                    # taking the end date of program from program template and converting YYYY-MM-DD 00:00:00 format

                    endDateArr = str(endDateOfProgram).split("-")
                    endDateOfProgram = endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"
                    if not Helpers.getProgramInfo(accessToken, parentFolder, solutionNameInp.encode('utf-8').decode('utf-8')):
                        extIdPGM = dictDetailsEnv['solution_name'].encode('utf-8').decode('utf-8')
                        programName = extIdPGM = dictDetailsEnv['solution_name'].encode('utf-8').decode('utf-8')
                        userDetails = Helpers.fetchUserDetails(accessToken, dictDetailsEnv['creator_username'])
                        OrgName=userDetails[4]
                        print(OrgName,"OrgName")
                        orgIds=Helpers.fetchOrgId(accessToken, parentFolder, OrgName)
                        creatorKeyCloakId = userDetails[0]
                        creatorName = userDetails[2]
                        if Helpers.getProgramInfo(accessToken, parentFolder, extIdPGM):
                            print("Program Created SuccessFully.")
                        else :
                            print("program creation API called")
                            Helpers.programCreation(accessToken, parentFolder, extIdPGM, programName,orgIds,creatorKeyCloakId, creatorName)

    def prepareProjectAndTasksSheets(project_inputFile, projectName_for_folder_path, accessToken):
        millisecond = int(time.time() * 1000)
        PreviousTaskname = None
        PreviousTaskid = None
        projectFilePath = projectName_for_folder_path + '/projectUpload/'
        taskFilePath = projectName_for_folder_path + '/taskUpload/'
        file_exists = os.path.isfile(projectName_for_folder_path + '/projectUpload/projectUpload.csv')
        if not os.path.exists(projectFilePath):
            os.mkdir(projectFilePath)
        if not os.path.exists(taskFilePath):
            os.mkdir(taskFilePath)

        wbproject = xlrd.open_workbook(project_inputFile, on_demand=True)
        projectSheetNames = wbproject.sheet_names()

        projectDetailsSheet = wbproject.sheet_by_name('Project upload')
        keysProject = [projectDetailsSheet.cell(1, col_index_env).value for col_index_env in
                       range(projectDetailsSheet.ncols)]
        projectColnames1 = ["title", "externalId", "categories","recommendedFor", "description", "entityType", "goal"]
        learningResource_count = 0
        for projectHeader in keysProject:
            if str(projectHeader).startswith('learningResources'):

                learningResource_count += 1
        learningResource_count = int(learningResource_count) / 2

        lr_count = 1
        for lr in range(0, int(learningResource_count)):
            projectColnames1.append("learningResources" + str(lr_count) + "-name")
            projectColnames1.append("learningResources" + str(lr_count) + "-link")
            projectColnames1.append("learningResources" + str(lr_count) + "-app")
            projectColnames1.append("learningResources" + str(lr_count) + "-id")
            lr_count += 1
        projectColnames2 = ["rationale", "primaryAudience", "taskCreationForm", "duration", "concepts", "keywords","successIndicators", "risks", "approaches", "_arrayFields"]
        for columns in projectColnames2:
            projectColnames1.append(columns)
        with open(projectFilePath + 'projectUpload.csv', 'w',encoding='utf-8') as file:
            writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
            writer.writerows([projectColnames1])

        for row_index_env in range(2, projectDetailsSheet.nrows):
            dictProjectDetails = {keysProject[col_index_env]: projectDetailsSheet.cell(row_index_env, col_index_env).value
                                  for col_index_env in range(projectDetailsSheet.ncols)}
            title = str(dictProjectDetails["title"]).encode('utf-8').decode('utf-8').strip()
            externalId = str(dictProjectDetails["projectId"]).strip() + "-" + str(millisecond)
            categories_list = ["teachers", "students", "infrastructure", "community", "educationLeader", "schoolProcess"]
            categories = str(dictProjectDetails["categories"]).encode('utf-8').decode('utf-8').split(",")
            categories_final = ""
            projectGoal = "TEMP"
            for cat in categories:
                if categories_final == "":
                    categories_final = categories_final + str(
                        (get_close_matches(cat.strip().lower().replace(" ", ""), categories_list)[0]))
                else:
                    categories_final = categories_final + "," + str(
                        (get_close_matches(cat.strip().lower().replace(" ", ""), categories_list)[0]))
            global projectCreator, projectAuthor

            projectAuthor = str(dictProjectDetails["Diksha_loginId"]).encode('utf-8').decode('utf-8').strip()
            recommendedFor = str(dictProjectDetails["recommendedFor"]).encode('utf-8').decode('utf-8').strip()
            objective = str(dictProjectDetails["objective"]).encode('utf-8').decode('utf-8').strip()
            entityType = None
            project_values = [title, externalId, categories_final,recommendedFor, objective, entityType,projectGoal]
            lr_value_count = 1
            for lr in range(0, int(learningResource_count)):
                lr_name = str(dictProjectDetails["learningResources" + str(lr_value_count) + "-name"]).strip()
                lr_link = str(dictProjectDetails["learningResources" + str(lr_value_count) + "-link"]).strip()
                if lr_name == "" and lr_link == "":
                    project_values.append("")
                    project_values.append("")
                    project_values.append("")
                    project_values.append("")
                    lr_value_count += 1
                else:
                    project_values.append(lr_name)
                    lr_link_id = lr_link.split("/")[-1]
                    project_values.append(lr_link)
                    project_values.append("Diksha")
                    project_values.append(lr_link_id)
                    lr_value_count += 1
            remaining_project_values = ["rationale", "primaryAudience", "taskCreationForm", "duration", "concepts",
                                        "keywords", "successIndicators", "risks", "approaches", "_arrayFields"] 
            for values in remaining_project_values:
                try:
                    project_values.append(str(dictProjectDetails[values]).strip())
                except:
                    if values == "_arrayFields":
                        project_values.append(
                            "categories,recommendedFor,primaryAudience,successIndicators,risks,approaches")
                    else:
                        project_values.append("")

            with open(projectFilePath + 'projectUpload.csv','a',encoding='utf-8') as file:
                writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                writer.writerows([project_values])

        tasksDetailsSheet = wbproject.sheet_by_name('Tasks upload')
        keysTasks = [tasksDetailsSheet.cell(1, col_index_env).value for col_index_env in
                     range(tasksDetailsSheet.ncols)]
        taskColumns1 = ["name", "externalId", "description", "type", "hasAParentTask", "parentTaskOperator",
                        "parentTaskValue",
                        "parentTaskId", "solutionType", "solutionSubType", "solutionId", "isDeletable"]
        taskLearningResource_count = 0

        for tasksHeader in keysTasks:
            if str(tasksHeader).startswith('learningResources'):
                taskLearningResource_count += 1
        taskLearningResource_count = int(taskLearningResource_count) / 2
        taskslr_count = 1
        for lr in range(0, int(taskLearningResource_count)):
            taskColumns1.append("learningResources" + str(taskslr_count) + "-name")
            taskColumns1.append("learningResources" + str(taskslr_count) + "-link")
            taskColumns1.append("learningResources" + str(taskslr_count) + "-app")
            taskColumns1.append("learningResources" + str(taskslr_count) + "-id")
            taskslr_count += 1
        taskColumns1.append("minNoOfSubmissionsRequired")
        taskColumns1.append("sequenceNumber")

        with open(taskFilePath + 'taskUpload.csv', 'w',encoding='utf-8') as file:
            writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
            writer.writerows([taskColumns1])
        sequenceNumber = 0
        for row_index_env in range(2, tasksDetailsSheet.nrows):
            dictTasksDetails = {keysTasks[col_index_env]: tasksDetailsSheet.cell(row_index_env, col_index_env).value
                                for col_index_env in range(tasksDetailsSheet.ncols)}
            taskName = str(dictTasksDetails["TaskTitle"]).encode('utf-8').decode('utf-8').strip()
            subtaskname = str(dictTasksDetails["Subtask"]).encode('utf-8').decode('utf-8').strip()

            if dictTasksDetails['TaskId'] :
               taskId = str(dictTasksDetails["TaskId"]).encode('utf-8').decode('utf-8').strip() + "-" + str(millisecond)
               taskminNoOfSubmissionsRequired = str(dictTasksDetails["Number of submissions for observation"]).strip()
               sequenceNumber = sequenceNumber + 1
               taskSolutionType = ""
               try:
                   taskDescription = str(dictTasksDetails["description"]).strip()
               except:
                   taskDescription = ""
               if dictTasksDetails["observation Name"] != "":
                   taskType = "observation"
               elif dictTasksDetails["learningResources1-name"] != "" and dictTasksDetails["learningResources1-link"] != "":
                   taskType = "content"
               else:
                   taskType = "simple"

               hasAParentTask = "NO"
               parentTaskOperator = ""
               parentTaskValue = ""
               parentTaskId = ""

               if dictTasksDetails["observation Name"] != "":
                   solutionNameOrId = dictTasksDetails["observation Name"].encode('utf-8').decode('utf-8')
                   taskSolutionType = "observation"
                   solutionDetailsInTask = checkEntityOfSolution(projectName_for_folder_path, solutionNameOrId, accessToken)
                   solutionSubType = solutionDetailsInTask[0]
                   solutionId = solutionDetailsInTask[1]

                   projectUpload = pd.read_csv(projectFilePath + "projectUpload.csv")
                   # updating the column value/data
                   projectUpload.loc[0, 'entityType'] = solutionDetailsInTask[0]

                   # writing into the file
                   projectUpload.to_csv(projectFilePath + "projectUpload.csv", index=False)
               else:
                   solutionId = ""
                   taskSolutionType = ""
                   solutionSubType = ""

               if str(dictTasksDetails["Mandatory task(Yes or No)"]).strip().strip().lower() == "no":
                   isDeletable = "TRUE"
               else:
                   isDeletable = "FALSE"
               task_values = [taskName, taskId, taskDescription, taskType, hasAParentTask, parentTaskOperator, parentTaskValue,
                              parentTaskId, taskSolutionType, solutionSubType, solutionId, isDeletable]
               task_lr_value_count = 1
               for task_lr in range(0, int(taskLearningResource_count)):
                   task_lr_name = str(dictTasksDetails["learningResources" + str(task_lr_value_count) + "-name"]).strip()
                   task_lr_link = str(dictTasksDetails["learningResources" + str(task_lr_value_count) + "-link"]).strip()
                   if task_lr_name == "" and task_lr_link == "":
                       task_values.append("")
                       task_values.append("")
                       task_values.append("")
                       task_values.append("")
                       task_lr_value_count += 1
                   else:
                       task_values.append(task_lr_name)
                       task_lr_link_id = task_lr_link.split("/")[-1]
                       task_values.append(task_lr_link)
                       task_values.append("Diksha")
                       task_values.append(task_lr_link_id)
                       task_lr_value_count += 1
               task_values.append(taskminNoOfSubmissionsRequired)
               task_values.append(sequenceNumber)

               # To check weather the previous-task and the curent-task Taskname & Taskid is same 
               if str(taskName) == str(PreviousTaskname) and str(taskId) == str(PreviousTaskid):
                    print("true")
               else:
                   print("false")
                   with open(taskFilePath + 'taskUpload.csv','a',encoding='utf-8') as file:
                    writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                    writer.writerows([task_values])
               subtaskname2 = str(dictTasksDetails["Subtask"]).encode('utf-8').decode('utf-8').strip()
               PreviousTaskname = taskName
               PreviousTaskid = taskId

        c = 0
        for row_index_env in range(2, tasksDetailsSheet.nrows):
            dictTasksDetails = {keysTasks[col_index_env]: tasksDetailsSheet.cell(row_index_env, col_index_env).value
                                for col_index_env in range(tasksDetailsSheet.ncols)}
            if dictTasksDetails['TaskId'] and dictTasksDetails["Subtask"]:
                if dictTasksDetails["Subtask"] != "":
                    taskHasAParentTask = "YES"
                    taskparentTaskOperator = "EQUALS"
                    taskparentTaskValue = "started"
                    c = c + 1
                    cn = "Task"+str(c)
                    parentTaskIdofsubtask = str(dictTasksDetails["TaskId"]).strip() + "-" + str(millisecond)
                    taskminNoOfSubmissionsRequired = str(dictTasksDetails["Number of submissions for observation"]).strip()
                    sequenceNumber = sequenceNumber + 1
                    try:
                        proejcttaskDescription = str(dictTasksDetails["description"]).strip()
                    except:
                        proejcttaskDescription = ""
                    if dictTasksDetails["observation Name"] != "":
                        projecttaskType = "observation"
                    elif dictTasksDetails["learningResources1-name"] != "" and dictTasksDetails[
                        "learningResources1-link"] != "":
                        projecttaskType = "content"
                    else:
                        projecttaskType = "simple"


                subtaskId = str(dictTasksDetails["TaskId"]).encode('utf-8').decode('utf-8').strip() + "-" + str(millisecond) + cn

                subtaskName1 = str(dictTasksDetails["Subtask"]).strip()
                if str(dictTasksDetails["Mandatory task(Yes or No)"]).strip().strip().lower() == "no":
                    isDeletable = "TRUE"
                else:
                    isDeletable = "FALSE"
                subtaskvalues = [subtaskName1, subtaskId,proejcttaskDescription,projecttaskType,taskHasAParentTask,taskparentTaskOperator,taskparentTaskValue,
                                 parentTaskIdofsubtask, taskSolutionType, solutionSubType, solutionId, isDeletable]
                task_lr_value_count = 1
                for task_lr in range(0, int(taskLearningResource_count)):
                    task_lr_name = str(dictTasksDetails["learningResources" + str(task_lr_value_count) + "-name"]).strip()
                    task_lr_link = str(dictTasksDetails["learningResources" + str(task_lr_value_count) + "-link"]).strip()
                    if task_lr_name == "" and task_lr_link == "":
                        task_values.append("")
                        task_values.append("")
                        task_values.append("")
                        task_values.append("")
                        task_lr_value_count += 1
                    else:
                        task_values.append(task_lr_name)
                        task_lr_link_id = task_lr_link.split("/")[-1]
                        task_values.append(task_lr_link)
                        task_values.append("Diksha")
                        task_values.append(task_lr_link_id)
                        task_lr_value_count += 1
                task_values.append(taskminNoOfSubmissionsRequired)
                task_values.append(sequenceNumber)

                with open(taskFilePath + 'taskUpload.csv', 'a',encoding='utf-8') as file:
                    writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                    writer.writerows([subtaskvalues])


            if dictTasksDetails["Subtask"] and not dictTasksDetails['TaskTitle']:
                if dictTasksDetails["Subtask"] != "":
                    taskHasAParentTask = "YES"
                    taskparentTaskOperator = "EQUALS"
                    taskparentTaskValue = "started"
                    # c = c + 1
                    # cn = "Task"+str(c)
                    parentTaskId = str(dictTasksDetails["TaskId"]).encode('utf-8').decode('utf-8').strip() + "-" + str(millisecond)
                    try:
                        proejcttaskDescription = str(dictTasksDetails["description"]).strip()
                    except:
                        proejcttaskDescription = ""
                    if dictTasksDetails["observation Name"] != "":
                        projecttaskType = "observation"
                    elif dictTasksDetails["learningResources1-name"] != "" and dictTasksDetails[
                        "learningResources1-link"] != "":
                        projecttaskType = "content"
                    else:
                        projecttaskType = "simple"

                subtaskId = str(dictTasksDetails["TaskId"]).encode('utf-8').decode('utf-8').strip() + "-" + str(millisecond) + cn

                subtaskName1 = str(dictTasksDetails["Subtask"]).encode('utf-8').decode('utf-8').strip()
                subtaskvalues = [subtaskName1, subtaskId,proejcttaskDescription,projecttaskType,taskHasAParentTask,taskparentTaskOperator,taskparentTaskValue,
                                 parentTaskId, taskSolutionType, solutionSubType, solutionId, isDeletable]

                with open(taskFilePath + 'taskUpload.csv','a',encoding='utf-8') as file:
                    writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                    writer.writerows([subtaskvalues])


    def projectUpload(projectFile, projectName_for_folder_path, accessToken):
        urlProjectUploadApi = internal_kong_ip + projectuploadapi
        headerProjectUploadApi = {
            'Authorization': authorization,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id':  x_channel_id,
            'internal-access-token': internal_access_token
        }
        project_payload = {}
        filesProject = {
            'projectTemplates': open(projectName_for_folder_path + '/projectUpload/projectUpload.csv', 'rb')
        }

        responseProjectUploadApi = requests.post(url=urlProjectUploadApi, headers=headerProjectUploadApi,data=project_payload,files=filesProject)
        messageArr = ["program mapping is success.","File path : " + projectName_for_folder_path + '/projectUpload/projectUpload.csv']
        messageArr.append("Upload status code : " + str(responseProjectUploadApi.status_code))
        # createAPILog(projectName_for_folder_path, messageArr)

        if responseProjectUploadApi.status_code == 200:
            print('ProjectUploadApi Success')
            with open(projectName_for_folder_path + '/projectUpload/projectInternal.csv','w+',encoding='utf-8') as projectRes:
                projectRes.write(responseProjectUploadApi.text)
        else:
            print("Project Upload failed.")
            messageArr.append("Response : " + str(responseProjectUploadApi.text))
            # createAPILog(projectName_for_folder_path, messageArr)
            sys.exit()

    def taskUpload(projectFile, projectName_for_folder_path, accessToken):
        projectInternalfile = open(projectName_for_folder_path + '/projectUpload/projectInternal.csv', mode='r',encoding='utf-8')
        projectInternalfile = csv.DictReader(projectInternalfile)
        for projectInternal in projectInternalfile:
            projectExternalId = projectInternal["externalId"]
            project_id = projectInternal["_SYSTEM_ID"]
            if str(project_id).strip() == "Could not pushed to kafka":
                fetchProjectIdApi = internal_kong_ip + fetchprojectlist
                headerfetchProjectIdApi = {
                    'Authorization': authorization,
                    'X-authenticated-user-token': accessToken,
                    'X-Channel-id': x_channel_id,
                    'internal-access-token': internal_access_token
                }
                fetchProjectIdPayload = {}

                responseProjectListApi = requests.get(url=fetchProjectIdApi, headers=headerfetchProjectIdApi,
                                                      data=fetchProjectIdPayload)
                messageArr = ["Tasks Upload Sheet Prepared.",
                              "File path : " + projectName_for_folder_path + '/taskUpload/taskUpload.csv']
                messageArr.append("URL : " + str(fetchProjectIdApi))
                messageArr.append("Upload status code : " + str(responseProjectListApi.status_code))
                # createAPILog(projectName_for_folder_path, messageArr)

                if responseProjectListApi.status_code == 200:
                    print('project fetch api Success')
                    responsejson = responseProjectListApi.json()
                    projectList = responsejson['result']['data']
                    for project in projectList:
                        if project['externalId'] == projectExternalId:
                            project_id = project['_id']
                else:
                    messageArr.append("Response : " + str(responseProjectListApi.text))
                    # createAPILog(projectName_for_folder_path, messageArr)
                    # terminatingMessage("project fetch api failed.")

            urlTasksUploadApi = internal_kong_ip + taskuploadapi + project_id
            headerTasksUploadApi = {
                'Authorization': authorization,
                'X-authenticated-user-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token
            }
            task_payload = {}
            filesTasks = {
                'projectTemplateTasks': open(projectName_for_folder_path + '/taskUpload/taskUpload.csv',
                                             'rb')
            }

            responseTasksUploadApi = requests.post(url=urlTasksUploadApi, headers=headerTasksUploadApi,
                                                   data=task_payload,
                                                   files=filesTasks)
            messageArr = ["Tasks Upload Sheet Prepared.",
                          "File path : " + projectName_for_folder_path + '/taskUpload/taskUpload.csv']
            messageArr.append("URL : " + str(urlTasksUploadApi))
            messageArr.append("Upload status code : " + str(responseTasksUploadApi.status_code))
            # createAPILog(projectName_for_folder_path, messageArr)

            if responseTasksUploadApi.status_code == 200:
                print('TaskUploadApi Success')
                with open(projectName_for_folder_path + '/taskUpload/taskInternal.csv','w+',encoding='utf-8') as tasksRes:
                    tasksRes.write(responseTasksUploadApi.text)
            else:
                messageArr.append("Response : " + str(responseTasksUploadApi.text))
                # createAPILog(projectName_for_folder_path, messageArr)
                # terminatingMessage("--->Tasks Upload failed.")

    def prepareaddingcertificatetemp(filePathAddProject, projectName_for_folder_path, accessToken, solutionId, programID,baseTemplate_id):
        wbproject = xlrd.open_workbook(filePathAddProject, on_demand=True)
        projectsheetforcertificate = wbproject.sheet_names()
        tasksLevelEvidance = []
        projectMinNooEvide = None
        projectLevelEvidance = []
        taskMinNooEvide =[]


        for prosheet in projectsheetforcertificate:
            if prosheet.strip().lower() == 'Project upload'.lower():
                detailsColCheck = wbproject.sheet_by_name(prosheet)
                keysColCheckDetai = [detailsColCheck.cell(0, col_index_check).value for col_index_check in
                                     range(detailsColCheck.ncols)]

                detailsEnvSheet = wbproject.sheet_by_name(prosheet)
                keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                           range(detailsEnvSheet.ncols)]
                for row_index_env in range(2, detailsEnvSheet.nrows):
                    dictDetailsEnv = {
                        keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                        for col_index_env in range(detailsEnvSheet.ncols)}

                    projectLevelMinNooEvidence = dictDetailsEnv["Minimum No. of Evidence"]
                    print(projectLevelMinNooEvidence)
                    projectLevelEvidance = dictDetailsEnv["Project Level Evidence"].lower()
                    if projectLevelMinNooEvidence == "":
                        projectLevelMinNooEvidence = 1  # Set default value to 1
                        projectMinNooEvide = int(projectLevelMinNooEvidence)
                    else:
                        projectMinNooEvide = int(projectLevelMinNooEvidence)


        for prosheet in projectsheetforcertificate:
            if prosheet.strip().lower() == 'Tasks upload'.lower():
                detailsColCheck = wbproject.sheet_by_name(prosheet)
                keysColCheckDetai = [detailsColCheck.cell(0, col_index_check).value for col_index_check in
                                     range(detailsColCheck.ncols)]

                detailsEnvSheet = wbproject.sheet_by_name(prosheet)
                keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                           range(detailsEnvSheet.ncols)]

                for row_index_env in range(2, detailsEnvSheet.nrows):
                    dictDetailsEnv = {
                        keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                        for col_index_env in range(detailsEnvSheet.ncols)}


                    taskLevelEvidence = dictDetailsEnv["Task Level Evidence"].lower()
                    minNoOfEvidence = dictDetailsEnv["Minimum No. of Evidence"]

                    if taskLevelEvidence == "yes":
                        tasksLevelEvidance.append(dictDetailsEnv["TaskTitle"])
                        if minNoOfEvidence == "":
                            minNoOfEvidence = 1  # Set default value to 1
                            taskMinNooEvide.append(minNoOfEvidence)
                        else:
                            taskMinNooEvide.append(minNoOfEvidence)



        addcetificateFilePath = projectName_for_folder_path + '/addCertificate/'
        if not os.path.exists(addcetificateFilePath):
            os.mkdir(addcetificateFilePath)

        urladdcertificate = internal_kong_ip + addcertificatetemplate
        headeraddcertificateApi = {
            'Authorization': authorization,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token,
            'Content-Type': 'application/json'
        }

        if str(projectLevelEvidance).strip().lower() == "yes":
            payload = {}
            payload['criteria'] = {}
            payload['criteria']['validationText'] = "Complete validation message"
            payload['criteria']['expression'] = ""
            payload['criteria']['conditions'] = {}
            payload['criteria']['conditions']['C1'] = {}
            payload['criteria']['conditions']['C1']['validationText'] = "Submit your project."
            payload['criteria']['conditions']['C1']['expression'] = "C1"
            payload['criteria']['conditions']['C1']['conditions'] = {}
            payload['criteria']['conditions']['C1']['conditions']['C1'] = {}
            payload['criteria']['conditions']['C1']['conditions']['C1']['scope'] = "project"
            payload['criteria']['conditions']['C1']['conditions']['C1']['key'] = "status"
            payload['criteria']['conditions']['C1']['conditions']['C1']['operator'] = "=="
            payload['criteria']['conditions']['C1']['conditions']['C1']['value'] = "submitted"
            payload['criteria']['conditions']['C2'] = {}
            payload['criteria']['conditions']['C2']['validationText'] = f"Add {int(projectMinNooEvide)} evidence at the project level",
            payload['criteria']['conditions']['C2']['expression'] = "C1"
            payload['criteria']['conditions']['C2']['conditions'] = {}
            payload['criteria']['conditions']['C2']['conditions']['C1'] = {}
            payload['criteria']['conditions']['C2']['conditions']['C1']['scope'] = "project"
            payload['criteria']['conditions']['C2']['conditions']['C1']['key'] = "attachments"
            payload['criteria']['conditions']['C2']['conditions']['C1']['function'] = "count"
            payload['criteria']['conditions']['C2']['conditions']['C1']['filter'] = {}
            payload['criteria']['conditions']['C2']['conditions']['C1']['filter']['key'] = "type"
            payload['criteria']['conditions']['C2']['conditions']['C1']['filter']['value'] = "all"  
            payload['criteria']['conditions']['C2']['conditions']['C1']['operator'] = ">="
            payload['criteria']['conditions']['C2']['conditions']['C1']['value'] = int(projectMinNooEvide)
            payload['issuer'] ={}
            payload['issuer']['name']=""
            payload['status'] = "active"
            payload['solutionId'] = solutionId
            payload['programId'] = programID
            payload['baseTemplateId'] = ""

        else:
            # str(projectLevelEvidance).strip().lower() == "no":
            payload = {}
            payload['criteria'] = {}
            payload['criteria']['validationText'] = "Complete validation message"
            payload['criteria']['expression'] = ""
            payload['criteria']['conditions'] = {}
            payload['criteria']['conditions']['C1'] = {}
            payload['criteria']['conditions']['C1']['validationText'] = "Submit your project."
            payload['criteria']['conditions']['C1']['expression'] = "C1"
            payload['criteria']['conditions']['C1']['conditions'] = {}
            payload['criteria']['conditions']['C1']['conditions']['C1'] = {}
            payload['criteria']['conditions']['C1']['conditions']['C1']['scope'] = "project"
            payload['criteria']['conditions']['C1']['conditions']['C1']['key'] = "status"
            payload['criteria']['conditions']['C1']['conditions']['C1']['operator'] = "=="
            payload['criteria']['conditions']['C1']['conditions']['C1']['value'] = "submitted"
            payload['issuer'] ={}
            payload['issuer']['name']=""
            payload['status'] = "active"
            payload['solutionId'] = solutionId
            payload['programId'] = programID
            payload['baseTemplateId'] = ""


        if prosheet.strip().lower() == 'Certificate details'.lower():
            print("--->Checking Certificate details  sheet...")
            detailsColCheck = wbproject.sheet_by_name(prosheet)
            keysColCheckDetai = [detailsColCheck.cell(0, col_index_check).value for col_index_check in
                    range(detailsColCheck.ncols)]

            detailsEnvSheet = wbproject.sheet_by_name(prosheet)
            keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                           range(detailsEnvSheet.ncols)]
            for row_index_env in range(2, detailsEnvSheet.nrows):

                dictDetailsEnv = {
                        keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                        for
                        col_index_env in range(detailsEnvSheet.ncols)}
                certificateissuer = dictDetailsEnv['Certificate issuer'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Certificate issuer'] else terminatingMessage("\"Certificate issuer\" must not be Empty in \"Certificate details\" sheet")
                payload["issuer"]["name"] = certificateissuer

                Typeofcertificate = dictDetailsEnv['Type of certificate'] if dictDetailsEnv['Type of certificate'] in ["One Logo - One Signature", "One Logo - Two Signature", "Two Logo - One Signature","Two Logo - Two Signature"] else terminatingMessage("\"Type of certificate\" must not be Empty in \"Certificate details\" sheet")

                payload["baseTemplateId"]=baseTemplate_id

        projectInternalfile = open(projectName_for_folder_path + '/projectUpload/projectInternal.csv', mode='r',encoding='utf-8')
        projectInternalfile = csv.DictReader(projectInternalfile)
        for projectInternal in projectInternalfile:
            projectExternalId = projectInternal["externalId"]
            project_id = projectInternal["_SYSTEM_ID"]

        taskinternalfile = open(projectName_for_folder_path + '/taskUpload/taskInternal.csv', mode='r',encoding='utf-8')
        taskinternalfile = csv.DictReader(taskinternalfile)
        projectTemplatefile = open(projectName_for_folder_path + '/solutionDetails/solutionDetails.csv', mode='r',encoding='utf-8')
        projectTemplatefile = csv.DictReader(projectTemplatefile)
        for Projecttemp in projectTemplatefile:
            projectTemplateId = Projecttemp["duplicateTemplate_id"]
        c = 2
        for task in taskinternalfile:
            if task['name'] in tasksLevelEvidance:
                hasAparent = task["hasAParentTask"]
                if task["hasAParentTask"].lower() == "no":

                    task_id = task["_SYSTEM_ID"]

                    c = c + 1
                    cn = "C" + str(c)
                    taskconditions = {
                        cn: {
                            "validationText": f"Add {int(taskMinNooEvide[c-3])} evidence for the task {tasksLevelEvidance[c-3]}",
                            "expression": "C1",
                            "conditions": {
                                "C1": {
                                    "scope": "task",
                                    "key": "attachments",
                                    "function": "count",
                                    "filter": {
                                        "key": "type",
                                        "value": "all"
                                    },
                                    "operator": ">=",
                                    "value": int(taskMinNooEvide[c-3]),
                                    "taskDetails": [
                                        task_id
                                    ]
                                }
                            }
                        }
                    }
                    payload["criteria"]["conditions"].update(taskconditions)
            else:
                pass


        condition = ""
        for a, i in enumerate(payload["criteria"]["conditions"]):
            if a == 0:
                condition = condition + str(i)
            else:
                condition = condition + "&&" + str(i)
        payload["criteria"]["expression"] = condition


        print(json.dumps(payload, indent=1))
        # sys.exit()

        responseaddcertificateUploadApi = requests.request("POST",url=urladdcertificate, headers=headeraddcertificateApi,
                                               data=json.dumps(payload))
        messageArr = ["Add certificate json is prepared",
                      "File path : " + projectName_for_folder_path + '/addCertificate/Addcertificate.text']
        messageArr.append("URL : " + str(responseaddcertificateUploadApi))
        messageArr.append("Upload status code : " + str(responseaddcertificateUploadApi.status_code))
        # createAPILog(projectName_for_folder_path, messageArr)
        with open(projectName_for_folder_path + '/addCertificate/Addcertificatejson.json',
                  'w+',encoding='utf-8') as tasksRes:
            tasksRes.write(json.dumps(payload))

        if responseaddcertificateUploadApi.status_code == 200:
            responseaddcetificate = responseaddcertificateUploadApi.json()
            certificatetemplateid = responseaddcetificate['result']['id']
            print("-->Certificate template id generated <--", certificatetemplateid)


            with open(projectName_for_folder_path + '/addCertificate/Addcertificate.text',
                      'w+',encoding='utf-8') as tasksRes:
                tasksRes.write(responseaddcertificateUploadApi.text)

        else:
            print("Add certificate mission failed please check logs")
            messageArr.append("Response : " + str(responseaddcertificateUploadApi.text))
            # createAPILog(projectName_for_folder_path, messageArr)
            sys.exit()

        urluploadcertificatepi =internal_kong_ip + uploadcertificatetosvg + certificatetemplateid

        headeruploadcertificateApi = {
            'Authorization': authorization,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token
        }
        task_payload = {}
        task_file = []
        certificateaddtotemplate = ('file', ( 'Dowloaded.svg',open(projectName_for_folder_path + '/Dowloadedsvg/Dowloaded.svg', 'rb'), 'image/svg+xml'))
        task_file.append(certificateaddtotemplate)


        responseDownloadsvgApi = requests.request("POST",url=urluploadcertificatepi, headers=headeruploadcertificateApi,
                                               data=task_payload,
                                               files=task_file)
        if responseDownloadsvgApi.status_code == 200:
            responseeditsvg = responseDownloadsvgApi.json()
            svgid = responseeditsvg['result']['data']['templateId']

            urlsolutionupdateapi = internal_kong_ip + updatecertificatesolu + solutionId

            headersolutionupdateApi = {
                'Authorization': authorization,
                'X-authenticated-user-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': 'application/json'
            }

            certificate_payload = json.dumps({
                'certificateTemplateId':certificatetemplateid
            })
            responseupdatecertificateApi = requests.request("POST", url=urlsolutionupdateapi,
                                                      headers=headersolutionupdateApi,
                                                      data=certificate_payload)


            if responseupdatecertificateApi.status_code == 200:
                print("--->certificate added to the solution<---")

            else:
                print("error in updating solution")
                sys.exit()

            urlprojecttemplateapi = internal_kong_ip + updateprojecttemplate + projectTemplateId
            headerprojectrtemplateupdateApi = {
                'Authorization': authorization,
                'X-authenticated-user-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': 'application/json'
            }

            certificate_payload = json.dumps({
                'certificateTemplateId': certificatetemplateid
            })
            responseupdatecertificateApi = requests.request("POST", url=urlprojecttemplateapi,
                                                            headers=headerprojectrtemplateupdateApi,
                                                            data=certificate_payload)
            if responseupdatecertificateApi.status_code == 200:
                print("--->Certificate added to project<---")

            else:
                print("error in updating certificate with project")
                sys.exit()
# Th    is function is used to add SVG to the certificate based on type of certificate


    def editsvg(accessToken,filePathAddProject,projectName_for_folder_path,baseTemplate_id):
        wbproject = xlrd.open_workbook(filePathAddProject, on_demand=True)
        projectsheetforcertificate = wbproject.sheet_names()
        for prosheet in projectsheetforcertificate:
            if prosheet.strip().lower() == 'Certificate details'.lower():
                print("--->Checking Certificate details  sheet...")
                detailsColCheck = wbproject.sheet_by_name(prosheet)
                keysColCheckDetai = [detailsColCheck.cell(0, col_index_check).value for col_index_check in
                                     range(detailsColCheck.ncols)]

                detailsEnvSheet = wbproject.sheet_by_name(prosheet)
                keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                           range(detailsEnvSheet.ncols)]
                for row_index_env in range(2, detailsEnvSheet.nrows):

                    dictDetailsEnv = {
                        keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                        for
                        col_index_env in range(detailsEnvSheet.ncols)}
                    certificateissuer = dictDetailsEnv['Certificate issuer'].encode('utf-8').decode('utf-8')
                    Typeofcertificate = dictDetailsEnv['Type of certificate']
                    Certificateisuuer = dictDetailsEnv['Certificate issuer'].encode('utf-8').decode('utf-8')
                    Logo1 = dictDetailsEnv['Logo - 1']
                    authsignaturelogo1 = dictDetailsEnv['Authorised Signature Image - 1']
                    authrigedsignaturename1 = dictDetailsEnv['Authorised Signature Name - 1'].encode('utf-8').decode('utf-8')
                    authrigeddesignation1 = dictDetailsEnv['Authorised Designation - 1'].encode('utf-8').decode('utf-8')
                    authrigedlogo2 = dictDetailsEnv['Authorised Signature Image - 2']
                    authrigedsignaturename2 = dictDetailsEnv['Authorised Signature Name - 2'].encode('utf-8').decode('utf-8')
                    authrigeddesignation2 = dictDetailsEnv['Authorised Designation - 2'].encode('utf-8').decode('utf-8')

                    payload = {}
                    downloadedfiles = []
                    baseTemplateId = ''
                    if Typeofcertificate == 'One Logo - One Signature':
                        print("-->This is One Logo - One Signature<--")

                        stateLogo1 = ('stateLogo1',('logo1.jpg',open(projectName_for_folder_path +'/Logofile/logo1.jpg' ,'rb'),'image/jpeg'))
                        downloadedfiles.append(stateLogo1)
                        payload['stateTitle'] = Certificateisuuer
                        signatureImg1 = ('signatureImg1',('signature1.jpg',open(projectName_for_folder_path +'/Logofile/signature1.jpg','rb'),'image/jpeg'))
                        downloadedfiles.append(signatureImg1)
                        payload['signatureTitleName1'] = authrigedsignaturename1
                        payload['signatureTitleDesignation1'] = authrigeddesignation1
                        baseTemplateId=baseTemplate_id


                    elif Typeofcertificate == 'One Logo - Two Signature':
                        print("-->This is One Logo - Two Signature<--")

                        stateLogo1 = ('stateLogo1', (
                        'logo1.jpg', open(projectName_for_folder_path + '/Logofile/logo1.jpg', 'rb'), 'image/jpeg'))
                        downloadedfiles.append(stateLogo1)
                        payload['stateTitle'] = Certificateisuuer
                        signatureImg1 = ('signatureImg1', (
                        'signature1.jpg', open(projectName_for_folder_path + '/Logofile/signature1.jpg', 'rb'),
                        'image/jpeg'))
                        downloadedfiles.append(signatureImg1)
                        signatureImg2 = ('signatureImg2', ('signature2.jpg', open(projectName_for_folder_path + '/Logofile/signature2.jpg', 'rb'),'image/jpeg'))
                        downloadedfiles.append(signatureImg2)
                        payload['signatureTitleName1'] = authrigedsignaturename1
                        payload['signatureTitleDesignation1'] = authrigeddesignation1
                        payload['signatureTitleName2'] = authrigedsignaturename2
                        payload['signatureTitleDesignation2'] = authrigeddesignation2
                        baseTemplateId=baseTemplate_id

                    elif Typeofcertificate == 'Two Logo - One Signature':
                        print("-->This is Two Logo - One Signature<--")
                        stateLogo1 = ('stateLogo1', (
                            'logo1.jpg', open(projectName_for_folder_path + '/Logofile/logo1.jpg', 'rb'), 'image/jpeg'))
                        downloadedfiles.append(stateLogo1)
                        payload['stateTitle'] = Certificateisuuer
                        signatureImg1 = ('signatureImg1', ('signature1.jpg', open(projectName_for_folder_path + '/Logofile/signature1.jpg', 'rb'),'image/jpeg'))
                        downloadedfiles.append(signatureImg1)
                        stateLogo2 = ('stateLogo2', ('logo2.jpg', open(projectName_for_folder_path + '/Logofile/logo2.jpg', 'rb'), 'image/jpeg'))
                        downloadedfiles.append(stateLogo2)
                        payload['signatureTitleName1'] = authrigedsignaturename1
                        payload['signatureTitleDesignation1'] = authrigeddesignation1
                        baseTemplateId=baseTemplate_id

                    elif Typeofcertificate == 'Two Logo - Two Signature':
                        print("-->This is Two Logo - Two Signature<--")
                        stateLogo1 = ('stateLogo1', ('logo1.jpg', open(projectName_for_folder_path + '/Logofile/logo1.jpg', 'rb'), 'image/jpeg'))
                        downloadedfiles.append(stateLogo1)
                        payload['stateTitle'] = Certificateisuuer
                        signatureImg1 = ('signatureImg1', ('signature1.jpg', open(projectName_for_folder_path + '/Logofile/signature1.jpg', 'rb'),'image/jpeg'))
                        downloadedfiles.append(signatureImg1)
                        stateLogo2 = ('stateLogo2', ('logo2.jpg', open(projectName_for_folder_path + '/Logofile/logo2.jpg', 'rb'), 'image/jpeg'))
                        downloadedfiles.append(stateLogo2)
                        signatureImg2 = ('signatureImg2', ('signature2.jpg', open(projectName_for_folder_path + '/Logofile/signature2.jpg', 'rb'),'image/jpeg'))
                        downloadedfiles.append(signatureImg2)
                        payload['signatureTitleName1'] = authrigedsignaturename1
                        payload['signatureTitleDesignation1'] = authrigeddesignation1
                        payload['signatureTitleName2'] = authrigedsignaturename2
                        payload['signatureTitleDesignation2'] = authrigeddesignation2
                        baseTemplateId=baseTemplate_id

                    urleditnigsvgApi =  internal_kong_ip + editsvgtemp + baseTemplateId
                    headereditingsvgApi = {
                        'Authorization': authorization,
                        'X-authenticated-user-token': accessToken,
                        'X-Channel-id': x_channel_id,
                        'internal-access-token': internal_access_token

                    }
                    responseeditsvg = requests.request("POST",url=urleditnigsvgApi, headers=headereditingsvgApi,data=payload, files=downloadedfiles)

                    if responseeditsvg.status_code == 200:
                        responseeditsvg = responseeditsvg.json()
                        svgid = responseeditsvg['result']['url']
                        filesvg = svgid
                        Logofilepath = projectName_for_folder_path + '/Dowloadedsvg/'
                        if not os.path.exists(Logofilepath):
                            os.mkdir(Logofilepath)
                        dest_file = Logofilepath + 'Dowloaded.svg'
                        Logofile1 = gdown.download(filesvg, dest_file, quiet=False)

                    else:
                        print("-->Error in downloading SVG file please check logs<--")

    


    def fetchCertificateBaseTemplate(filePathAddProject,accessToken,projectName_for_folder_path):
        wbproject = xlrd.open_workbook(filePathAddProject, on_demand=True)
        projectsheetforcertificate = wbproject.sheet_names()
        for prosheet in projectsheetforcertificate:
            if prosheet.strip().lower() == 'Certificate details'.lower():
                detailsColCheck = wbproject.sheet_by_name(prosheet)
                keysColCheckDetai = [detailsColCheck.cell(0, col_index_check).value for col_index_check in
                                     range(detailsColCheck.ncols)]

                detailsEnvSheet = wbproject.sheet_by_name(prosheet)
                keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                           range(detailsEnvSheet.ncols)]
                for row_index_env in range(2, detailsEnvSheet.nrows):
                    dictDetailsEnv = {
                        keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                        for col_index_env in range(detailsEnvSheet.ncols)}

                    typeOfCertificate = dictDetailsEnv["Type of certificate"]
                    print(typeOfCertificate)

        urldbFind = internal_kong_ip + dbfindapi
        headerdbFindApi = {
            'Authorization':  authorization,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token,
            'Content-Type': 'application/json'
        }
        payload = json.dumps({
            "query": {},
            "mongoIdKeys": []
        })

        responsedbFindApi = requests.request("POST", url=urldbFind, headers=headerdbFindApi,
                                             data=payload)
        if responsedbFindApi.status_code == 200:
            responseaddcetificate = responsedbFindApi.json()
            result_list = responseaddcetificate['result']
            baseTemplateLookup = {}
            for i in result_list:
                baseTemplateLookup[i['code']] = i['_id']
            typeOfCertificate=typeOfCertificate.lower()
            typeOfCertificate=typeOfCertificate.replace("-","_")
            typeOfCertificate = typeOfCertificate.replace(" ","")
            baseTemplateCode= certificatetypeof[typeOfCertificate]
            print(baseTemplateCode,"baseTemplateCode")
            print(baseTemplateLookup,"baseTemplateLookup")

            return baseTemplateLookup[baseTemplateCode]

        else:
            print("--->Error in fetching DBfind data please give proper code value<---")
            #messageArr.append("Response : " + str(responseaddcetificate.text))
            #createAPILog(projectName_for_folder_path, messageArr)
            sys.exit()


    def downloadlogosign(filePathAddProject,projectName_for_folder_path):
        wbproject = xlrd.open_workbook(filePathAddProject, on_demand=True)
        projectsheetforcertificate = wbproject.sheet_names()
        for prosheet in projectsheetforcertificate:
            if prosheet.strip().lower() == 'Certificate details'.lower():
                print("--->Checking Certificate details  sheet...")
                detailsColCheck = wbproject.sheet_by_name(prosheet)
                keysColCheckDetai = [detailsColCheck.cell(0, col_index_check).value for col_index_check in range(detailsColCheck.ncols)]

                detailsEnvSheet = wbproject.sheet_by_name(prosheet)
                keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                           range(detailsEnvSheet.ncols)]
                for row_index_env in range(2, detailsEnvSheet.nrows):

                    dictDetailsEnv = {
                        keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                        for
                        col_index_env in range(detailsEnvSheet.ncols)}
                    certificateissuer = dictDetailsEnv['Certificate issuer'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Certificate issuer'] else terminatingMessage("\"Certificate issuer\" must not be Empty in \"Certificate details\" sheet")

                    typeOfCertificate = dictDetailsEnv['Type of certificate'] if dictDetailsEnv['Type of certificate'] else terminatingMessage("\"Type of certificate\" must not be Empty in \"Certificate details\" sheet")

                    if typeOfCertificate == 'One Logo - One Signature':
                       Logo1 = dictDetailsEnv['Logo - 1']
                       logo_split = str(Logo1).split('/')[5]

                       file_url = 'https://drive.google.com/uc?export=download&id='+logo_split

                       Logofilepath = projectName_for_folder_path + '/Logofile/'
                       if not os.path.exists(Logofilepath):
                           os.mkdir(Logofilepath)
                       dest_file = Logofilepath + '/logo1.jpg'
                       Logofile1 = gdown.download(file_url, dest_file,quiet=False)


                       Authsign1 = dictDetailsEnv['Authorised Signature Image - 1']
                       logo_split = str(Authsign1).split('/')[5]

                       file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split


                       dest_file = Logofilepath + '/signature1.jpg'
                       signature1 = gdown.download(file_url, dest_file, quiet=False)

                    elif typeOfCertificate == 'One Logo - Two Signature':

                        Logo1 = dictDetailsEnv['Logo - 1']
                        logo_split = str(Logo1).split('/')[5]

                        file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split

                        Logofilepath = projectName_for_folder_path + '/Logofile/'
                        if not os.path.exists(Logofilepath):
                            os.mkdir(Logofilepath)
                        dest_file = Logofilepath + '/logo1.jpg'
                        Logofile1 = gdown.download(file_url, dest_file, quiet=False)


                        Authsign1 = dictDetailsEnv['Authorised Signature Image - 1']
                        logo_split = str(Authsign1).split('/')[5]

                        file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split

                        dest_file = Logofilepath + '/signature1.jpg'
                        signature1 = gdown.download(file_url, dest_file, quiet=False)

                        Authsign2 = dictDetailsEnv['Authorised Signature Image - 2']
                        logo_split = str(Authsign2).split('/')[5]

                        file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split


                        dest_file = Logofilepath + '/signature2.jpg'
                        signature2 = gdown.download(file_url, dest_file, quiet=False)

                    elif typeOfCertificate == 'Two Logo - One Signature':

                        Logo1 = dictDetailsEnv['Logo - 1']
                        logo_split = str(Logo1).split('/')[5]

                        file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split

                        Logofilepath = projectName_for_folder_path + '/Logofile/'
                        if not os.path.exists(Logofilepath):
                            os.mkdir(Logofilepath)
                        dest_file = Logofilepath + '/logo1.jpg'
                        Logofile1 = gdown.download(file_url, dest_file, quiet=False)


                        Logo2 = dictDetailsEnv['Logo - 2']
                        logo_split = str(Logo2).split('/')[5]

                        file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split


                        dest_file = Logofilepath + '/logo2.jpg'
                        Logofile2 = gdown.download(file_url, dest_file, quiet=False)


                        Authsign1 = dictDetailsEnv['Authorised Signature Image - 1']
                        logo_split = str(Authsign1).split('/')[5]


                        file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split


                        dest_file = Logofilepath + '/signature1.jpg'
                        signature1 = gdown.download(file_url, dest_file, quiet=False)


                    elif typeOfCertificate == 'Two Logo - Two Signature':

                        Logo1 = dictDetailsEnv['Logo - 1']
                        logo_split = str(Logo1).split('/')[5]


                        file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split

                        Logofilepath = projectName_for_folder_path + '/Logofile/'
                        if not os.path.exists(Logofilepath):
                            os.mkdir(Logofilepath)
                        dest_file = Logofilepath + '/logo1.jpg'
                        Logofile1 = gdown.download(file_url, dest_file, quiet=False)


                        Logo2 = dictDetailsEnv['Logo - 2']
                        logo_split = str(Logo2).split('/')[5]


                        file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split


                        dest_file = Logofilepath + '/logo2.jpg'
                        Logofile2 = gdown.download(file_url, dest_file, quiet=False)


                        Authsign1 = dictDetailsEnv['Authorised Signature Image - 1']
                        logo_split = str(Authsign1).split('/')[5]


                        file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split


                        dest_file = Logofilepath + '/signature1.jpg'
                        signature1 = gdown.download(file_url, dest_file, quiet=False)


                        Authsign2 = dictDetailsEnv['Authorised Signature Image - 2']
                        logo_split = str(Authsign2).split('/')[5]


                        file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split


                        dest_file = Logofilepath + '/signature2.jpg'
                        signature2 = gdown.download(file_url, dest_file, quiet=False)

                    else:
                        print("--->Logos and signature downlading are failed(check if drive link are  Anyone with the link or not)<---")




    def fetchSolutionDetailsFromProgramSheet(solutionName_for_folder_path, programFile, solutionId, accessToken):
        global solutionRolesArray, solutionStartDate, solutionEndDate
        urlFetchSolutionApi = internal_kong_ip + fetchsolutiondoc + solutionId

        headerFetchSolutionApi = {
            'Content-Type': 'application/json',
            'Authorization': authorization,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token
        }
        payloadFetchSolutionApi = {}

        responseFetchSolutionApiUrl = requests.post(url=urlFetchSolutionApi, headers=headerFetchSolutionApi,
                                                 data=payloadFetchSolutionApi)
        responseFetchSolutionJson = responseFetchSolutionApiUrl.json()
        messageArr = ["Solution Fetch Link.",
                      "solution name : " + responseFetchSolutionJson["result"]["name"],
                      "solution ExternalId : " + responseFetchSolutionJson["result"]["externalId"]]
        messageArr.append("Upload status code : " + str(responseFetchSolutionApiUrl.status_code))
        # createAPILog(solutionName_for_folder_path, messageArr)

        if responseFetchSolutionApiUrl.status_code == 200:
            print('Fetch solution Api Success')

            solutionName = responseFetchSolutionJson["result"]["name"]

            xfile = openpyxl.load_workbook(programFile)

            resourceDetailsSheet = xfile.get_sheet_by_name('Resource Details')
            rowCountRD = resourceDetailsSheet.max_row
            columnCountRD = resourceDetailsSheet.max_column
            for row in range(3, rowCountRD + 1):
                if resourceDetailsSheet["A" + str(row)].value == solutionName:
                    solutionMainRole = str(resourceDetailsSheet["E" + str(row)].value).strip()
                    solutionRolesArray = str(resourceDetailsSheet["F" + str(row)].value).split(",") if str(
                        resourceDetailsSheet["E" + str(row)].value).split(",") else []
                    if "teacher" in solutionMainRole.strip().lower():
                        solutionRolesArray.append("TEACHER")
                    solutionStartDate = resourceDetailsSheet["G" + str(row)].value
                    solutionEndDate = resourceDetailsSheet["H" + str(row)].value
        return [solutionRolesArray, solutionStartDate, solutionEndDate]





    def solutionCreationAndMapping(projectName_for_folder_path, entityToUpload, listOfFoundRoles, accessToken,programFile):
        SolutionFilePath = projectName_for_folder_path + '/solutionDetails/'
        if not os.path.exists(SolutionFilePath):
            os.mkdir(SolutionFilePath)
        with open(projectName_for_folder_path + '/solutionDetails/solutionDetails.csv', 'w',encoding='utf-8') as file:
            writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
            writer.writerows(
                [["solutionExtId", "solutionName", "solutionDescription", "solution_id", "programExternalId", "entityType",
                  "scopeEntityType", "entityNames", "roles", "duplicateTemplateExtId", "duplicateTemplate_id"]])

        projectInternalfile = open(projectName_for_folder_path + '/projectUpload/projectInternal.csv', mode='r',encoding='utf-8')
        projectInternalfile = csv.DictReader(projectInternalfile)
        for projectInternal in projectInternalfile:
            projectExternalId = projectInternal["externalId"]
            project_id = projectInternal["_SYSTEM_ID"]
            project_name = projectInternal["title"]
            project_description = projectInternal["description"]
            if projectInternal["entityType"]:
                projectEntityType = projectInternal["entityType"]
            else:
                projectEntityType = "school"
            solutionExternalId = projectExternalId + "-PROJECT-SOLUTION"

            urlCreateProjectSolutionApi = internal_kong_ip + projectsolutioncreationapi
            headerCreateSolutionApi = {
                'Content-Type': content_type,
                'Authorization': authorization,
                'X-authenticated-user-token': accessToken,
                'X-Channel-id': x_channel_id
            }
            sol_payload = {
                "createdFor": orgIds,
                "rootOrganisations": orgIds,
                "programExternalId": programExternalId,
                "entityType": projectEntityType,
                "externalId": solutionExternalId,
                "name": project_name,
                "description": project_description
            }
            responseCreateSolutionApi = requests.post(url=urlCreateProjectSolutionApi,headers=headerCreateSolutionApi, data=json.dumps(sol_payload))

            messageArr = ["Project Solution Created.","URL : " + str(urlCreateProjectSolutionApi),"Status Code : " + str(responseCreateSolutionApi.status_code),"Response : " + str(responseCreateSolutionApi.text)]
            if responseCreateSolutionApi.status_code == 200:
                responseCreateSolutionApi = responseCreateSolutionApi.json()
                solutionId = responseCreateSolutionApi['result']['_id']
                messageArr.append("Solution Generated : " + str(solutionId))
                # createAPILog(projectName_for_folder_path, messageArr)
                print("ProjectSolutionCreationApi Success")
                duplicateTemplateExtId = projectExternalId + '_IMPORTED'
                queryparamsMapProjectSolutionApi = projectExternalId + '?solutionId=' + solutionExternalId
                urlMapProjectSolutionApi = internal_kong_ip + mapsolutiontoproject
                headerMapSolutionProject = {
                    'Content-Type': content_type,
                    'Authorization': authorization,
                    'X-authenticated-user-token': accessToken,
                    'X-Channel-id': x_channel_id
                }
                payloadMapSolutionProject = {
                    "externalId": duplicateTemplateExtId,
                    "rating": 5
                }
                responseMapProjectSolutionApi = requests.post(
                    url=urlMapProjectSolutionApi + queryparamsMapProjectSolutionApi,
                    headers=headerMapSolutionProject, data=json.dumps(payloadMapSolutionProject))

                messageArr = ["Successfully mapped the project to Solution",
                              "URL : " + str(urlMapProjectSolutionApi + queryparamsMapProjectSolutionApi),
                              "Status Code : " + str(responseMapProjectSolutionApi.status_code),
                              "Response : " + str(responseMapProjectSolutionApi.text)]
                if responseMapProjectSolutionApi.status_code == 200:
                    responseMapProjectSolutionApi = responseMapProjectSolutionApi.json()
                    duplicateTemplateId = responseMapProjectSolutionApi['result']['_id']
                    messageArr.append("duplicate TemplateId successfully created: " + str(duplicateTemplateId))
                    # createAPILog(projectName_for_folder_path, messageArr)
                    print("MapSolutionToProjectApi Sucsess")
                    with open(projectName_for_folder_path + '/solutionDetails/solutionDetails.csv', 'a',encoding='utf-8') as file:
                        writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                        writer.writerows([[solutionExternalId, project_name, project_description, solutionId,
                                           programExternalId, projectEntityType,
                                           scopeEntityType, entityToUpload, listOfFoundRoles, duplicateTemplateExtId,
                                           duplicateTemplateId]])
                    solutionDetails = Helpers.fetchSolutionDetailsFromProgramSheet(projectName_for_folder_path, programFile,
                                                                           solutionId, accessToken)
                    scopeEntities = entitiesPGMID
                    scopeRoles = solutionDetails[0]
                    bodySolutionUpdate = {
                        "scope": {"entityType": scopeEntityType, "entities": scopeEntities, "roles": scopeRoles}}
                    Helpers.solutionUpdate(projectName_for_folder_path, accessToken, solutionId, bodySolutionUpdate)

                    userDetails = Helpers.fetchUserDetails(accessToken, projectAuthor)
                    matchedShikshalokamLoginId = userDetails[0]
                    projectCreator = userDetails[2]

                    bodySolutionUpdate = {
                        "creator": projectCreator, "author": matchedShikshalokamLoginId}
                    Helpers.solutionUpdate(projectName_for_folder_path, accessToken, solutionId, bodySolutionUpdate)
                    # Below script will convert date DD-MM-YYYY TO YYYY-MM-DD 00:00:00 to match the code syntax

                    if solutionDetails[1]:
                        startDateArr = str(solutionDetails[1]).split("-")
                        bodySolutionUpdate = {
                            "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                        Helpers.solutionUpdate(projectName_for_folder_path, accessToken, solutionId, bodySolutionUpdate)
                    if solutionDetails[2]:
                        endDateArr = str(solutionDetails[2]).split("-")
                        bodySolutionUpdate = {
                            "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                        Helpers.solutionUpdate(projectName_for_folder_path, accessToken, solutionId, bodySolutionUpdate)
                else:
                    print("Map project to solution api failed.")
                return [solutionExternalId, solutionId]
            else:
                print("Project solution creation api failed.")
                sys.exit()



    def prepareProgramSuccessSheet(MainFilePath, solutionName_for_folder_path, programFile, solutionExternalId, solutionId,accessToken):
        urlFetchSolutionApi = internal_kong_ip + fetchsolutiondoc + solutionId
        headerFetchSolutionApi = {
            'Authorization': authorization,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token
        }
        payloadFetchSolutionApi = {}

        responseFetchSolutionApi = requests.post(url=urlFetchSolutionApi, headers=headerFetchSolutionApi,
                                                 data=payloadFetchSolutionApi)
        responseFetchSolutionJson = responseFetchSolutionApi.json()
        messageArr = ["Solution Fetch Link.",
                      "solution name : " + responseFetchSolutionJson["result"]["name"],
                      "solution ExternalId : " + responseFetchSolutionJson["result"]["externalId"]]
        messageArr.append("Upload status code : " + str(responseFetchSolutionApi.status_code))
        # createAPILog(solutionName_for_folder_path, messageArr)

        if responseFetchSolutionApi.status_code == 200:
            print('Fetch solution Api Success')
            solutionName = responseFetchSolutionJson["result"]["name"]
        urlFetchSolutionLinkApi = internal_kong_ip + fetchlink + solutionId
        headerFetchSolutionLinkApi = {
            'Authorization': authorization,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token
        }
        payloadFetchSolutionLinkApi = {}

        responseFetchSolutionLinkApi = requests.post(url=urlFetchSolutionLinkApi, headers=headerFetchSolutionLinkApi,
                                                     data=payloadFetchSolutionLinkApi)
        messageArr = ["Solution Fetch Link.","solution id : " + solutionId,"solution ExternalId : " + solutionExternalId]
        messageArr.append("Upload status code : " + str(responseFetchSolutionLinkApi.status_code))
        # createAPILog(solutionName_for_folder_path, messageArr)

        if responseFetchSolutionLinkApi.status_code == 200:
            print('Fetch solution Link Api Success')
            responseProjectUploadJson = responseFetchSolutionLinkApi.json()
            solutionLink = responseProjectUploadJson["result"]
        #     messageArr.append("Response : " + str(responseFetchSolutionLinkApi.text))
        #     # createAPILog(solutionName_for_folder_path, messageArr)

        #     if os.path.exists(MainFilePath + "/" + str(programFile).replace(".xlsx", "") + '-SuccessSheet.xlsx'):
        #         xfile = openpyxl.load_workbook(
        #             MainFilePath + "/" + str(programFile).replace(".xlsx", "") + '-SuccessSheet.xlsx')
        #     else:
        #         xfile = openpyxl.load_workbook(programFile)

        #     resourceDetailsSheet = xfile.get_sheet_by_name('Resource Details')

        #     greenFill = PatternFill(start_color='0000FF00',
        #                             end_color='0000FF00',
        #                             fill_type='solid')
        #     rowCountRD = resourceDetailsSheet.max_row
        #     columnCountRD = resourceDetailsSheet.max_column
        #     for row in range(3, rowCountRD + 1):
        #         if str(resourceDetailsSheet["B" + str(row)].value).rstrip().lstrip().lower() == "course":
        #             resourceDetailsSheet["D1"] = ""
        #             resourceDetailsSheet["E1"] = ""
        #             resourceDetailsSheet['I2'] = "External id of the resource"
        #             resourceDetailsSheet['J2'] = "link to access the resource/Response"
        #             resourceDetailsSheet['I2'].fill = greenFill
        #             resourceDetailsSheet['J2'].fill = greenFill
        #             resourceDetailsSheet['I' + str(row)] = solutionExternalId
        #             resourceDetailsSheet['J' + str(row)] = "The course has been successfully mapped to the program"
        #             resourceDetailsSheet['I' + str(row)].fill = greenFill
        #             resourceDetailsSheet['J' + str(row)].fill = greenFill
        #         elif str(resourceDetailsSheet["A" + str(row)].value).strip() == solutionName:
        #             resourceDetailsSheet["D1"] = ""
        #             resourceDetailsSheet["E1"] = ""
        #             resourceDetailsSheet['I2'] = "External id of the resource"
        #             resourceDetailsSheet['J2'] = "link to access the resource/Response"
        #             resourceDetailsSheet['I2'].fill = greenFill
        #             resourceDetailsSheet['J2'].fill = greenFill
        #             resourceDetailsSheet['I' + str(row)] = solutionExternalId
        #             resourceDetailsSheet['J' + str(row)] = solutionLink
        #             resourceDetailsSheet['I' + str(row)].fill = greenFill
        #             resourceDetailsSheet['J' + str(row)].fill = greenFill

        #     programFile = str(programFile).replace(".xlsx", "")
        #     xfile.save(MainFilePath + "/" + programFile + '-SuccessSheet.xlsx')
        #     print("Program success sheet is created")

        else:
            print("Fetch solution link API Failed")
            messageArr.append("Response : " + str(responseFetchSolutionLinkApi.text))
            # createAPILog(solutionName_for_folder_path, messageArr)
            sys.exit()
        return solutionLink

# fetch org Ids 
    def fetchOrgId(accessToken, parentFolder, OrgName):
        url = host + fetchorgdetails
        headers = {'Content-Type': 'application/json',
                   'Authorization': authorization,
                   'x-authenticated-user-token': accessToken}
        orgIds = []
        organisations = str(OrgName).split(",")
        for org in organisations:
            orgBody = {"id": "",
                       "ts": "",
                       "params": {
                           "msgid": "",
                           "resmsgid": "",
                           "status": "success"
                       },
                       "request": {
                           "filters": {
                               "orgName": str(org).strip()
                           }
                       }}

            responseOrgSearch = requests.request("POST", url, headers=headers, data=json.dumps(orgBody))
            if responseOrgSearch.status_code == 200:
                responseOrgSearch = responseOrgSearch.json()
                if responseOrgSearch['result']['response']['content']:
                    orgId = responseOrgSearch['result']['response']['content'][0]['id']
                    orgIds.append(orgId)
                else:
                    print("Email is not present in KB")
            else:
               print(responseOrgSearch.text)
                
        return orgIds

    def solutionUpdate(solutionName_for_folder_path, accessToken, solutionId, bodySolutionUpdate):
        solutionUpdateApi = internal_kong_ip + solutionupdateapi + str(solutionId)
        print("solutionUpdateApi:",solutionUpdateApi)
        headerUpdateSolutionApi = {
            'Content-Type': 'application/json',
            'Authorization': authorization,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id': x_channel_id,
            "internal-access-token": internal_access_token
            }
        responseUpdateSolutionApi = requests.post(url=solutionUpdateApi, headers=headerUpdateSolutionApi,data=json.dumps(bodySolutionUpdate))
        
        if responseUpdateSolutionApi.status_code == 200:
            print("Solution Update Success.")
            return True
        else:
            print("Solution Update Failed.")
            return False

    def createSurveySolution(parentFolder, wbSurvey, accessToken):
        # print(accessToken)
        print(wbSurvey,"wbSurveywbSurvey")
    
        sheetNames1 = wbSurvey.sheet_names()
        for sheetEnv in sheetNames1:
            if sheetEnv.strip().lower() == 'details':
                surveySolutionCreationReqBody = {}
                detailsEnvSheet = wbSurvey.sheet_by_name(sheetEnv)
                keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                       range(detailsEnvSheet.ncols)]

                for row_index_env in range(2, detailsEnvSheet.nrows):
                    dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                  for
                                  col_index_env in range(detailsEnvSheet.ncols)}
                    # print(dictDetailsEnv,"dictDetailsEnv")
                    surveySolutionCreationReqBody['name'] = dictDetailsEnv['solution_name'].encode('utf-8').decode('utf-8')
                    surveySolutionCreationReqBody["description"] = "survey Solution"
                    surveySolutionExternalId = str(uuid.uuid1())
                    surveySolutionCreationReqBody["externalId"] = surveySolutionExternalId
                    # if dictDetailsEnv['creator_username'].encode('utf-8').decode('utf-8') == "":
                    #     exceptionHandlingFlag = True
                    #     print('survey_creator_username column should not be empty in the details sheet')
                    #     sys.exit()
                    # else:
                    #     surveySolutionCreationReqBody['creator'] = dictDetailsEnv['Name_of_the_creator'].encode('utf-8').decode('utf-8')

                    
                    userDetails = Helpers.fetchUserDetails(accessToken, dictDetailsEnv['creator_username'])
                    # print(userDetails)
                    surveySolutionCreationReqBody['author'] = userDetails[0]
                    # print("surveySolutionCreationReqBody",surveySolutionCreationReqBody)

                    # Below script will convert date DD-MM-YYYY TO YYYY-MM-DD 00:00:00 to match the code syntax 

                    if dictDetailsEnv["start_date"]:
                        if type(dictDetailsEnv["start_date"]) == str:
                            startDateArr = None
                            startDateArr = (dictDetailsEnv["start_date"]).split("-")
                            surveySolutionCreationReqBody["startDate"] = startDateArr[2] + "-" + startDateArr[1] + "-" + \
                                                                     startDateArr[0] + " 00:00:00"
                        elif type(dictDetailsEnv["start_date"]) == float:
                            surveySolutionCreationReqBody["startDate"] = (
                            xlrd.xldate.xldate_as_datetime(dictDetailsEnv["start_date"],
                                                           wbSurvey.datemode)).strftime("%Y/%m/%d")
                        else:
                            surveySolutionCreationReqBody["startDate"] = ""
                        if dictDetailsEnv["end_date"]:
                            if type(dictDetailsEnv["end_date"]) == str:
                                print("enter 1")

                                endDateArr = None
                                endDateArr = (dictDetailsEnv["end_date"]).split("-")
                                surveySolutionCreationReqBody["endDate"] = endDateArr[2] + "-" + endDateArr[1] + "-" + \
                                                                       endDateArr[0] + " 23:59:59"
                            elif type(dictDetailsEnv["end_date"]) == float:
                                print("enter 2")
                                surveySolutionCreationReqBody["endDate"] = (
                                    xlrd.xldate.xldate_as_datetime(dictDetailsEnv["end_date"],
                                                               wbSurvey.datemode)).strftime("%Y/%m/%d")
                            else:
                                print("enter 3")
                                surveySolutionCreationReqBody["endDate"] = ""
                            enDt = surveySolutionCreationReqBody["endDate"]
                        
                            urlCreateSolutionApi =internal_kong_ip_survey+ surveysolutioncreationapiurl
                            print(urlCreateSolutionApi)
                            headerCreateSolutionApi = {
                            'Content-Type': 'application/json',
                            'Authorization': authorization,
                            'X-authenticated-user-token': accessToken,
                            'X-Channel-id': x_channel_id,
                            'appName': appname
                        }
                            # print(surveySolutionCreationReqBody)
                            print(headerCreateSolutionApi)
                            # sys.exit()
                            responseCreateSolutionApi = requests.post(url=urlCreateSolutionApi,
                                                                  headers=headerCreateSolutionApi,
                                                                  data=json.dumps(surveySolutionCreationReqBody))
                            print(responseCreateSolutionApi.text)
                            responseInText = responseCreateSolutionApi.text
                        
                            if responseCreateSolutionApi.status_code == 200:
                                responseCreateSolutionApi = responseCreateSolutionApi.json()
                                urlSearchSolution = internal_kong_ip_core + fetchsolutiondetails + "survey&page=1&limit=10&search=" + str(surveySolutionExternalId)
                                print(urlSearchSolution)
                                responseSearchSolution = requests.request("POST", urlSearchSolution,
                                                                      headers=headerCreateSolutionApi)
                            
                                if responseSearchSolution.status_code == 200:
                                    responseSearchSolutionApi = responseSearchSolution.json()
                                    surveySolutionExternalId = None
                                    surveySolutionExternalId = responseSearchSolutionApi['result']['data'][0]['externalId']
                                else:
                                    print("Solution fetch API failed")
                                    print("URL : " + urlSearchSolution)

                                solutionId = None
                                solutionId = responseCreateSolutionApi["result"]["solutionId"]
                                bodySolutionUpdate = {"creator": userDetails[2]}

                                return [solutionId, surveySolutionExternalId]
                            
                            else:
                                print("somethinghere i found")

    def schedule_deletion(returnPathStr):
        def delete_file():
            try:
                time.sleep(15)
                if os.path.exists(returnPathStr):
                    if os.path.isfile(returnPathStr):
                        os.remove(returnPathStr)
                        print(f"File {returnPathStr} deleted successfully.")

                    elif os.path.isdir(returnPathStr):
                        shutil.rmtree(returnPathStr)
                        print(f"Directory {returnPathStr} deleted successfully.")
                else:
                    print(f"File {returnPathStr} not found.")
            except Exception as e:
                print(f"Error deleting file: {e}")

        threading.Thread(target=delete_file, daemon=True).start()


        
    def mainFunc(MainFilePath, programFile, addObservationSolution, millisecond, isProgramnamePresent, isCourse,scopeEntityType=scopeEntityType):
        scopeEntityType = scopeEntityType

        parentFolder = Helpers.createFileStruct(MainFilePath, addObservationSolution)
        

        accessToken = Helpers.generateAccessToken(parentFolder)
        wbObservation = xlrd.open_workbook(addObservationSolution, on_demand=True)
        print("wbObservation",wbObservation)
        Helpers.programsFileCheck(programFile, accessToken, parentFolder, MainFilePath)
        wbprogram = xlrd.open_workbook(programFile, on_demand=True)
        programSheetNames = wbprogram.sheet_names()
        wbproject = xlrd.open_workbook(addObservationSolution, on_demand=True)
        projectSheetNames = wbproject.sheet_names()
        for programSheets in programSheetNames:
            if programSheets.strip().lower() == 'program details':
                print("Checking program details sheet...")
                programDetailsSheet = wbprogram.sheet_by_name(programSheets)
                keysEnv = [programDetailsSheet.cell(1, col_index_env).value for col_index_env in
                           range(programDetailsSheet.ncols)]
                for row_index_env in range(2, programDetailsSheet.nrows):
                    dictProgramDetails = {
                        keysEnv[col_index_env]: programDetailsSheet.cell(row_index_env, col_index_env).value
                        for col_index_env in range(programDetailsSheet.ncols)}
                    programName = dictProgramDetails['Title of the Program'].encode('utf-8').decode('utf-8')
                    isProgramnamePresent = False
                    if programName == "":
                        isProgramnamePresent = False
                    else:
                        isProgramnamePresent = True
                    scopeEntityType = scopeEntityType
                    userEntity = dictProgramDetails['Targeted state at program level'].encode('utf-8').decode('utf-8').lstrip().rstrip().split(",") if dictProgramDetails['Targeted state at program level'] else terminatingMessage("\"scope_entity\" must not be Empty in \"details\" sheet")
                    
        for sheets in projectSheetNames:
            if sheets.strip().lower() == 'Project upload'.lower():
                print("Checking project upload sheet...")
                projectsheet = wbproject.sheet_by_name(sheets)
                keysEnv = [projectsheet.cell(1, col_index_env).value for col_index_env in
                           range(projectsheet.ncols)]
                for row_index_env in range(1, projectsheet.nrows):
                    projectDetails = {keysEnv[col_index_env]: projectsheet.cell(row_index_env, col_index_env).value
                                      for col_index_env in range(projectsheet.ncols)}
                    ProjectName = projectDetails["title"].encode('utf-8').decode('utf-8')
                    print(ProjectName)
                    entityType = "school"
        try:
            def addProjectFunc(filePathAddProject, projectName_for_folder_path, millisAddObs):
                print('Add Project Function Called')
                
                projectName_for_folder = None
                
                if not path.exists(projectName_for_folder_path):
                    os.mkdir(projectName_for_folder_path)
                # copy input file to drive file
                if not path.exists(projectName_for_folder_path + "/user_input_file"):
                    os.mkdir(projectName_for_folder_path + "/user_input_file")
                shutil.copy(filePathAddProject, projectName_for_folder_path + "/user_input_file")
                shutil.copy(programFile, projectName_for_folder_path + "/user_input_file")
                messageArr = ["Access token generated.", "Access token : " + accessToken, "Solution file created.",
                              "Path : " + projectName_for_folder_path]
                # createAPILog(projectName_for_folder_path, messageArr)
                wbproject = xlrd.open_workbook(filePathAddProject, on_demand=True)
                projectsheetforcertificate = wbproject.sheet_names()
                for prosheet in projectsheetforcertificate:
                    if prosheet.strip().lower() == 'Project upload'.lower():
                        detailsColCheck = wbproject.sheet_by_name(prosheet)
                        keysColCheckDetai = [detailsColCheck.cell(0, col_index_check).value for col_index_check in
                                             range(detailsColCheck.ncols)]
                        detailsEnvSheet = wbproject.sheet_by_name(prosheet)
                        keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                                   range(detailsEnvSheet.ncols)]
                        for row_index_env in range(2, detailsEnvSheet.nrows):
                            # print(dictDetailsEnv)
                            # sys.exit()
                            dictDetailsEnv = {
                                keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                for
                                col_index_env in range(detailsEnvSheet.ncols)}
                            if str(dictDetailsEnv['has certificate']).lower() == 'No'.lower():
                                Helpers.prepareProjectAndTasksSheets(addObservationSolution, projectName_for_folder_path,
                                                             accessToken)
                                # sys.exit()
                                Helpers.projectUpload(addObservationSolution, projectName_for_folder_path, accessToken)
                                Helpers.taskUpload(addObservationSolution, projectName_for_folder_path, accessToken)
                                ProjectSolutionResp = Helpers.solutionCreationAndMapping(projectName_for_folder_path,
                                                                                 entityToUpload,
                                                                                 listOfFoundRoles, accessToken,programFile)
                                ProjectSolutionExternalId = ProjectSolutionResp[0]
                                ProjectSolutionId = ProjectSolutionResp[1]
                                solutionlink = Helpers.prepareProgramSuccessSheet(MainFilePath, projectName_for_folder_path, programFile,
                                                           ProjectSolutionExternalId,
                                                           ProjectSolutionId, accessToken)
                                print(solutionlink)
                            elif str(dictDetailsEnv['has certificate']).lower()== 'Yes'.lower():
                                print("---->this is certificate with project<---")
                                baseTemplate_id=Helpers.fetchCertificateBaseTemplate(filePathAddProject,accessToken,projectName_for_folder_path)
                                # sys.exit()
                                Helpers.downloadlogosign(filePathAddProject,projectName_for_folder_path)
                                Helpers.editsvg(accessToken,filePathAddProject,projectName_for_folder_path,baseTemplate_id)
                                Helpers.prepareProjectAndTasksSheets(addObservationSolution, projectName_for_folder_path,accessToken)
                                Helpers.projectUpload(addObservationSolution, projectName_for_folder_path, accessToken)
                                Helpers.taskUpload(addObservationSolution, projectName_for_folder_path, accessToken)
                                ProjectSolutionResp = Helpers.solutionCreationAndMapping(projectName_for_folder_path,entityToUpload,listOfFoundRoles, accessToken,programFile)
                                ProjectSolutionExternalId = ProjectSolutionResp[0]
                                ProjectSolutionId = ProjectSolutionResp[1]
                                print("---------------------------------------------")
                                certificatetemplateid= Helpers.prepareaddingcertificatetemp(filePathAddProject,projectName_for_folder_path, accessToken,ProjectSolutionId,programID,baseTemplate_id)
                                solutionlink = Helpers.prepareProgramSuccessSheet(MainFilePath, projectName_for_folder_path, programFile,
                                                           ProjectSolutionExternalId,
                                                           ProjectSolutionId, accessToken)
                                print(solutionlink)
                return solutionlink

        # return solutionlink
                                
        except:
            print("Terminated")
        millisecond = int(time.time() * 1000)
        projectSolutionLink=addProjectFunc(addObservationSolution, parentFolder, millisecond)
        return projectSolutionLink
        print(projectSolutionLink,"projectSolutionLink")
        print("Done.")

        # Helpers.SolutionFileCheck(addObservationSolution, accessToken, parentFolder, MainFilePath)
        # print(wbObservation,"wbObservation")
        # surveyResp = Helpers.createSurveySolution(parentFolder, wbObservation, accessToken)
        # surTempExtID = surveyResp[1]
        # # surveyChildId = surveyResp[0]
        # bodySolutionUpdate = {"status": "active", "isDeleted": False}
        # Helpers.solutionUpdate(parentFolder, accessToken, surveyResp[0], bodySolutionUpdate)
        # surveyChildId = Helpers.uploadSurveyQuestions(parentFolder, wbObservation, addObservationSolution, accessToken, surTempExtID,
        #                         surveyResp[0], millisecond)
        
        # local = os.getcwd()
        # sucessSheetName = Helpers.preparesolutionUploadSheet(MainFilePath,parentFolder,surveyChildId)
        # print("surveychild id",surveyChildId)
        # clickheretodownload = Helpers.uploadSuccessSheetToBucket(surveyChildId,sucessSheetName,accessToken)
        # Helpers.schedule_deletion(MainFilePath)
        # return [surveyChildId,local+'/'+sucessSheetName,clickheretodownload]
    
    
    def loadSurveyFile(programFile):
        MainFilePath = Helpers.createFileStructForProgram(programFile)
        wbPgm = xlrd.open_workbook(programFile, on_demand=True)
        sheetNames = wbPgm.sheet_names()
        pgmSheets = ["Instructions", "Program Details", "Resource Details","Program Manager Details"]
        print(sheetNames)
        print(pgmSheets)
        if len(sheetNames) == len(pgmSheets) and sheetNames == pgmSheets:
            print("--->Program Template detected.<---")

            for sheetEnv in sheetNames:
                if sheetEnv.strip().lower() == 'program details':
                    print("Checking program details sheet...")
                    programDetailsSheet = wbPgm.sheet_by_name(sheetEnv)
                    keysEnv = [programDetailsSheet.cell(1, col_index_env).value for col_index_env in
                               range(programDetailsSheet.ncols)]
                    for row_index_env in range(2, programDetailsSheet.nrows):
                        dictProgramDetails = {
                            keysEnv[col_index_env]: programDetailsSheet.cell(row_index_env, col_index_env).value
                            for col_index_env in range(programDetailsSheet.ncols)}
                        programName = dictProgramDetails['Title of the Program'].encode('utf-8').decode('utf-8')
                        print(programName,"programName")
                        isProgramnamePresent = False
                        if programName == "":
                            isProgramnamePresent = False
                        else:
                            isProgramnamePresent = True
                        print(isProgramnamePresent,"isProgramnamePresent")
                        scopeEntityType = "scopeEntityType"
                        print(scopeEntityType,"scopeEntityType")
                        userEntity = dictProgramDetails['Targeted state at program level'].encode('utf-8').decode('utf-8').lstrip().rstrip().split(
                            ",") 
                        print(userEntity,"userEntity")
                if sheetEnv.strip().lower() == 'resource details':
                    print("--->Checking Resource Details sheet...")
                    messageArr = []
                    messageArr.append("--->Checking Resource Details sheet...")
                    detailsEnvSheet = wbPgm.sheet_by_name(sheetEnv)
                    keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                               range(detailsEnvSheet.ncols)]
                    for row_index_env in range(2, detailsEnvSheet.nrows):
                        millisecond = int(time.time() * 1000)
                        dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                          for
                                          col_index_env in range(detailsEnvSheet.ncols)}
                        resourceNamePGM = dictDetailsEnv['Name of resources in program'].encode('utf-8').decode('utf-8')
                        resourceTypePGM = dictDetailsEnv['Type of resources'].encode('utf-8').decode('utf-8')
                        resourceLinkOrExtPGM = dictDetailsEnv['Resource Link'] 
                        if str(dictDetailsEnv['Type of resources']).lower().strip() == "course":
                            isCourse = False
                        else:
                            isCourse = False
                            resourceStatus = dictDetailsEnv['Resource Status']
                            if resourceStatus.strip()=="New Upload":
                                print("--->Resource Name : "+str(resourceNamePGM))
                                resourceLinkOrExtPGM = str(resourceLinkOrExtPGM).split('/')[5]
                                file_url = 'https://docs.google.com/spreadsheets/d/' + resourceLinkOrExtPGM + '/export?format=xlsx'
                                if not os.path.isdir('InputFiles'):
                                    os.mkdir('InputFiles')
                                dest_file = 'InputFiles'
                                addObservationSolution = wget.download(file_url, dest_file)
                                print("--->solution input file successfully downloaded " + str(addObservationSolution))


        solutionSL = Helpers.mainFunc(MainFilePath,programFile,addObservationSolution, millisecond, isProgramnamePresent,isCourse)
        print(solutionSL)
        return solutionSL

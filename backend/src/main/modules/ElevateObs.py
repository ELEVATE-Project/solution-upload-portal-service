import imaplib
import base64
import os
import time
from configparser import ConfigParser, ExtendedInterpolation
import wget
import urllib
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
import xlwt
import xlutils
from xlutils.copy import copy
import shutil
import re
from xlrd import open_workbook
from xlutils.copy import copy as xl_copy
import logging
import logging.handlers
import time
from logging.handlers import TimedRotatingFileHandler
import xlsxwriter
import argparse
import sys
from os import path
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.cell import Cell
import gdown
import jwt
# get current working directory
currentDirectory = os.getcwd()

# Read config file 
config = ConfigParser()
config.read('common_config/config.ini')


# email regex
regex = "\"?([-a-zA-Z0-9.`?{}]+@\w+\.\w+)\"?"

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
rolesPGMID = None
solutionDescription = None
creator = None
dikshaLoginId = None
criteriaName = None
solutionId = None
API_log = None
listOfFoundRoles = []
stateEntitiesPGM = []
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
entitiesType = []
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

class ElevateObservation:

    def terminatingMessage(msg):
        print(msg)
        sys.exit()

    def valid_file(param):
        base, ext = os.path.splitext(param)
        if ext.lower() not in ('.xlsx'):
            raise argparse.ArgumentTypeError('File must have a csv extension')
        return param

    def createFileStructForProgram(programFile):
        if not os.path.isdir('programFiles'):
            os.mkdir('programFiles')
        if "\\" in str(programFile):
            fileNameSplit = str(programFile).split('\\')[-1:]
        elif "/" in str(programFile):
            fileNameSplit = str(programFile).split('/')[-1:]
        else:
            fileNameSplit = str(programFile)
        if ".xlsx" in fileNameSplit:
            ts = str(time.time()).replace(".", "_")
            folderName = fileNameSplit.replace(".xlsx", "-" + str(ts))
            os.mkdir('programFiles/' + str(folderName))
            path = os.path.join('programFiles', str(folderName))
        else:
            ElevateObservation.terminatingMessage("File Error.")
        returnPathStr = os.path.join('programFiles', str(folderName))
        return returnPathStr

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
            ElevateObservation.terminatingMessage("File Error.offff")
        returnPathStr = os.path.join(MainFilePath + '/SolutionFiles', str(folderName))

        if not os.path.isdir(returnPathStr + "/user_input_file"):
            os.mkdir(returnPathStr + "/user_input_file")

        shutil.copy(addSolutionFile, os.path.join(returnPathStr + "user_input_file.xlsx"))
        return returnPathStr

    def apicheckslog(solutionName_for_folder_path, messageArr):
        file_exists = solutionName_for_folder_path + '/apiHitLogs/apiLogs.csv'
        # global fileheader
        fileheader = ["Resource","Process","Status","Remark"]

        if not path.exists(file_exists):
            with open(file_exists, 'w', newline='',encoding='utf-8') as file:
                writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                writer.writerows([fileheader])
        with open(file_exists, 'a', newline='',encoding='utf-8') as file:
            writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
            writer.writerows([messageArr])

    def createAPILog(solutionName_for_folder_path, messageArr):
        file_exists = solutionName_for_folder_path + '/apiHitLogs/apiLogs.txt'
        # check if the file existis or not and create a file 
        if not path.exists(file_exists):
            API_log = open(file_exists, "w", encoding='utf-8')
            API_log.write("===============================================================================")
            API_log.write("\n")
            API_log.write("ENVIRONMENT : " + str(environment))
            API_log.write("\n")
            API_log.write("===============================================================================")
            API_log.write("\n")
            API_log.close()

        API_log = open(file_exists, "a", encoding='utf-8')
        API_log.write("\n")
        for msg in messageArr:
            API_log.write(msg)
            API_log.write("\n")
        API_log.close()

    def generateAccessToken(solutionName_for_folder_path):
        # production search user api - start
        headerKeyClockUser = {'Content-Type': config.get(environment, 'content-type')}
        
        # responseKeyClockUser = requests.post(url=config.get(environment, 'host') + config.get(environment, 'keyclockAPIUrl'), headers=headerKeyClockUser,
        #                                      data=str(config.get(environment, 'keyclockAPIBody')))
        loginBody = {
            'email' : config.get(environment, 'email'),
            'password' : config.get(environment, 'password')
        }
        responseKeyClockUser = requests.request("POST", config.get(environment, 'userLoginHost') + config.get(environment, 'keyclockapiurl'), headers=headerKeyClockUser, data=json.dumps(loginBody))
        messageArr = []
        messageArr.append("URL : " + str(config.get(environment, 'keyclockAPIUrl')))
        messageArr.append("Body : " + str(config.get(environment, 'keyclockAPIBody')))
        messageArr.append("Status Code : " + str(responseKeyClockUser))
        if responseKeyClockUser.status_code == 200:
            responseKeyClockUser = responseKeyClockUser.json()
            accessTokenUser = responseKeyClockUser['result']['access_token']
            messageArr.append("Acccess Token : " + str(accessTokenUser))
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            fileheader = ["Access Token","Access Token succesfully genarated","Passed"]
            ElevateObservation.apicheckslog(solutionName_for_folder_path,fileheader)
            print("--->Access Token Generated!")
            return accessTokenUser
        
        print("Error in generating Access token")
        print("Status code : " + str(responseKeyClockUser.status_code))
        ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
        fileheader = ["Access Token", "Error in generating Access token", "Failed",responseKeyClockUser.status_code+"Check access token api"]
        ElevateObservation.apicheckslog(solutionName_for_folder_path, fileheader)
        fileheader = ["Access Token", "Error in generating Access token", "Failed","Check Headers of api"]
        ElevateObservation.apicheckslog(solutionName_for_folder_path, fileheader)
        ElevateObservation.terminatingMessage("Please check API logs.")
    
    def fetchEntityType(solutionName_for_folder_path, accessToken, entitiesPGM, scopeEntityType):
        urlFetchEntityListApi = config.get(environment, 'elevateentityhost') + config.get(environment, 'searchForLocation')
        headerFetchEntityListApi = {
            'Content-Type': config.get(environment, 'Content-Type'),
            'internal-access-token': config.get(environment, 'internal-access-token'),
        }

        # Initialize a dictionary to store entity types for each entity
        entityTypes = []

        # Loop through each entity name in the entitiesPGM list
        for entityName in entitiesPGM:
            entityName = entityName.strip()  # Remove any extra spaces

            # Prepare the payload for the API request
            payload = {
                "query": {
                    "metaInformation.name": entityName  # Use the current entity name
                },
                "projection": [
                    "entityType"
                ]
            }
            data = json.dumps(payload)

            # Make the API call inside the loop to send one request per entity
            responseFetchEntityListApi = requests.post(url=urlFetchEntityListApi, headers=headerFetchEntityListApi, data=data)
            # Log API call details
            messageArr = ["Entities List Fetch API executed for entity: " + entityName, 
                        "URL  : " + str(urlFetchEntityListApi),
                        "Status : " + str(responseFetchEntityListApi.status_code)]
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)

            # Check if the API call was successful
            if responseFetchEntityListApi.status_code == 200:
                responseFetchEntityListApi = responseFetchEntityListApi.json()

                # Loop through the result to find the entityType
                entityToUpload = None  # Initialize for each entity
                for listEntities in responseFetchEntityListApi['result']:
                    entityToUpload = listEntities['entityType']
                    # entityToUpload = listEntities.get('entityType', '').lower().strip()

                    # If a valid entityType is found, store it in the dictionary and break out of the loop
                    if entityToUpload:
                        entityTypes.append(entityToUpload)
                        break

                # If no entityType is found for this entity, raise an error for that specific entity
                if not entityToUpload:
                    raise ValueError(f"Entity type not found for entity '{entityName}'.")
            else:
                # Handle cases where the API call fails for a specific entity
                raise RuntimeError(f"Failed to fetch entity type for '{entityName}'. Status code: {responseFetchEntityListApi.status_code}")
        # Return all found entity types
        return entityTypes

    def fetchEntityId(solutionName_for_folder_path, accessToken, entitiesNameList, scopeEntityType):
        urlFetchEntityListApi = config.get(environment, 'elevateentityhost')+config.get(environment, 'searchForLocation')
        headerFetchEntityListApi = {
            'Content-Type': config.get(environment, 'Content-Type'),
            'internal-access-token': config.get(environment, 'internal-access-token'),
        }
        payload = {

        "query" : {
            "entityType": {
                "$in": scopeEntityType
            }
        },

        "projection": [
            "_id","metaInformation.name"
        ]
        }
        data=json.dumps(payload)
        responseFetchEntityListApi = requests.post(url=urlFetchEntityListApi, headers=headerFetchEntityListApi,data=json.dumps(payload))
        messageArr = ["Entities List Fetch API executed.", "URL  : " + str(urlFetchEntityListApi),
                    "Status : " + str(responseFetchEntityListApi.status_code)]
        ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
        if responseFetchEntityListApi.status_code == 200:
            responseFetchEntityListApi = responseFetchEntityListApi.json()
            entitiesLookup = dict()
            entityToUpload = list()
            for listEntities in responseFetchEntityListApi['result']:
                entitiesLookup[listEntities['metaInformation']['name'].lower().lstrip().rstrip()] = listEntities['_id'].lstrip().rstrip()
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
                    ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)

            messageArr = ["Entities to upload : " + str(entityToUpload)]
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            if len(entityToUpload) == 0:
                ElevateObservation.terminatingMessage("--->Scope Entity error.")
            return entityToUpload
        else:
            messageArr = ["Error in Location search",str(responseFetchEntityListApi.status_code)]
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            ElevateObservation.terminatingMessage("---> Error in location search.")

    def getProgramInfo(accessTokenUser, solutionName_for_folder_path, programNameInp):
        if programNameInp:
            global programID, programExternalId, programDescription, isProgramnamePresent, programName
            programName = programNameInp
            programUrl = config.get(environment, 'internal_kong_ip') + config.get(environment, 'fetchProgramInfoApiUrl')
            payload = json.dumps({
                "query": {
                    "name": programNameInp.lstrip().rstrip(),
                    "isAPrivateProgram": False,
                    "status": "active"
                    },
                    "mongoIdKeys": []
                    })
            
            headersProgramSearch =  {'Content-Type': 'application/json', 'X-auth-token': accessTokenUser}
            responseProgramSearch = requests.post(url=programUrl, headers=headersProgramSearch,data=payload)
            messageArr = []

            messageArr.append("Program Search API")
            messageArr.append("URL : " + programUrl)
            messageArr.append("Status Code : " + str(responseProgramSearch.status_code))
            messageArr.append("Response : " + str(responseProgramSearch.text))
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            messageArr = []
            if responseProgramSearch.status_code == 200:
                print('--->Program fetch API Success')
                messageArr.append("--->Program fetch API Success")
                responseProgramSearch = responseProgramSearch.json()
                print(responseProgramSearch,"this is the progrma")
                countOfPrograms = len(responseProgramSearch['result'])
                messageArr.append("--->Program Count : " + str(countOfPrograms))
                if countOfPrograms == 0:
                    messageArr.append("No program found with the name : " + str(programName.lstrip().rstrip()))
                    messageArr.append("******************** Preparing for program Upload **********************")
                    print("No program found with the name : " + str(programName.lstrip().rstrip()))
                    print("******************** Preparing for program Upload **********************")
                    ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                    fileheader = ["Program name fetch","Successfully fetched program name","Passed"]
                    ElevateObservation.apicheckslog(solutionName_for_folder_path,fileheader)
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
                                ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                                fileheader = ["program find api is running","found"+str(len(
                                    getProgramDetails))+"programs in backend","Failed","found"+str(len(
                                    getProgramDetails))+"programs ,check logs"]
                                ElevateObservation.apicheckslog(solutionName_for_folder_path,fileheader)
                                ElevateObservation.terminatingMessage("Aborting...")
                            elif len(getProgramDetails) > 1:
                                print("Total " + str(len(getProgramDetails)) + " backend programs found with the name : " + programName.lstrip().rstrip())
                                messageArr.append("Total " + str(len(getProgramDetails)) + " backend programs found with the name : " + programName.lstrip().rstrip())
                                ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)

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
                            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            else:
                print("Program search API failed...")
                messageArr.append("Program search API failed...")
                ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                ElevateObservation.terminatingMessage("Response Code : " + str(responseProgramSearch.status_code))
            return True

    def fetchUserDetails(environment, accessToken, dikshaId):
        global OrgName
        decoded_token = jwt.decode(accessToken, options={"verify_signature": False}, algorithms=["HS256"])
        data=decoded_token.get('data')
        user_id = data.get('id')  
        url = config.get(environment, 'userLoginHost') + config.get(environment, 'userInfoApiUrl')+"/"+str(user_id)
        messageArr = ["User search API called."]
        headers = {'Content-Type': 'application/json',
                'internal_access_token': config.get(environment, 'internal-access-token')}
    
        # isEmail = checkEmailValidation(dikshaId.lstrip().rstrip())
        # if isEmail:
        #     body = "{\n  \"request\": {\n    \"filters\": {\n    \t\"email\": \"" + dikshaId.lstrip().rstrip() + "\"\n    },\n      \"fields\" :[],\n    \"limit\": 1000,\n    \"sort_by\": {\"createdDate\": \"desc\"}\n  }\n}"
        # else:
        #     body = "{\n  \"request\": {\n    \"filters\": {\n    \t\"userName\": \"" + dikshaId.lstrip().rstrip() + "\"\n    },\n      \"fields\" :[],\n    \"limit\": 1000,\n    \"sort_by\": {\"createdDate\": \"desc\"}\n  }\n}"
        
        responseUserSearch = requests.request("GET", url, headers=headers)
        if responseUserSearch.status_code == 200:
            responseUserSearch = responseUserSearch.json()
            if responseUserSearch['result']:
                userKeycloak = responseUserSearch['result']['id']
                userName = responseUserSearch['result']['name']
                firstName = responseUserSearch['result']['name']
                rootOrgId = responseUserSearch['result']['organization']['id']
                for index in responseUserSearch['result']['user_roles']:
                    if rootOrgId == index['organization_id']:
                        roledetails = index['title']
                        # rootOrgName = index['orgName']
                        # OrgName.append(index['orgName'])
                print(roledetails)
            else:
                ElevateObservation.terminatingMessage("-->Given username/email is not present in Elevate platform<--.")
        else:
            print(responseUserSearch.text)
            ElevateObservation.terminatingMessage("User fetch API failed. Check logs.")
        return [userKeycloak, userName, firstName,roledetails,rootOrgId]

    def programCreation(accessToken, parentFolder, externalId, pName, pDescription, keywords, entities, roles, orgIds,creatorKeyCloakId, creatorName,entitiesPGM,mainRole,rolesPGM):
        messageArr = []
        messageArr.append("++++++++++++ Program Creation ++++++++++++")
        # program creation url 
        programCreationurl = config.get(environment, 'INTERNAL_KONG_IP') + config.get(environment, 'programCreationurl')
        messageArr.append("Program Creation URL : " + programCreationurl)
        # program creation payload
        scope={}
        for i in range(len(scopeEntityType)):
            entity_type = scopeEntityType[i]
            entity_value = entities[i]
            if entity_type in scope:
                scope[entity_type].append(entity_value)
            else:
                scope[entity_type] = [entity_value]
                        # bodySolutionUpdate = {
                        #     "scope": {"entityType": scopeEntityType, "entities": scopeEntities, "roles": scopeRoles}}
        scope["roles"] = roles
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
            "scope": scope,
            "metaInformation": {
                "state":stateEntitiesPGM.split(","),
                "roles": mainRole.split(",")
                },
                "requestForPIIConsent":True
                })
        messageArr.append("Body : " + str(payload))
        headers = {'X-auth-token': accessToken,
                'internal-access-token': config.get(environment, 'internal-access-token'),
                'Content-Type': 'application/json',
                'Authorization':config.get(environment, 'Authorization')}
        
        # program creation 
        responsePgmCreate = requests.request("POST", programCreationurl, headers=headers, data=(payload))
        messageArr.append("Program Creation Status Code : " + str(responsePgmCreate.status_code))
        messageArr.append("Program Creation Response : " + str(responsePgmCreate.text))
        messageArr.append("Program body : " + str(payload))

        # save logs 
        ElevateObservation.createAPILog(parentFolder, messageArr)
        # check status 
        fileheader = [pName, ('Program Sheet Validation'), ('Passed')]
        ElevateObservation.createAPILog(parentFolder, messageArr)
        ElevateObservation.apicheckslog(parentFolder, fileheader)
        if responsePgmCreate.status_code == 200:
            responsePgmCreateResp = responsePgmCreate.json()
        else:
            # terminate execution
            ElevateObservation.terminatingMessage("Program creation API failed. Please check logs.")

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
                        programNameInp = dictDetailsEnv['Title of the Program'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Title of the Program'] else ElevateObservation.terminatingMessage("\"Title of the Program\" must not be Empty in \"Program details\" sheet")
                        extIdPGM = dictDetailsEnv['Program ID'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Program ID'] else ElevateObservation.terminatingMessage("\"Program ID\" must not be Empty in \"Program details\" sheet")
                        descriptionPGM = dictDetailsEnv['Description of the Program'].encode('utf-8').decode('utf-8') if dictDetailsEnv[
                                'Description of the Program'] else ElevateObservation.terminatingMessage(
                                "\"Description of the Program\" must not be Empty in \"Program details\" sheet")
                        keywordsPGM = dictDetailsEnv['Keywords'].encode('utf-8').decode('utf-8')
                        returnvalues = []
                        global entitiesPGM
                        entitiesPGM = dictDetailsEnv['Targeted entities at program level'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Targeted entities at program level'] else ElevateObservation.terminatingMessage("\"Targeted entities at program level\" must not be Empty in \"Program details\" sheet")
                        global stateEntitiesPGM
                        stateEntitiesPGM = dictDetailsEnv['Targeted state at program level'].encode('utf-8').decode('utf-8')
                        global mainRole
                        mainRole = dictDetailsEnv['Targeted role at program level'] if dictDetailsEnv['Targeted role at program level'] else ElevateObservation.terminatingMessage("\"Targeted role at program level\" must not be Empty in \"Program details\" sheet")
                        global rolesPGM
                        rolesPGM = dictDetailsEnv['Targeted subrole at program level'] if dictDetailsEnv['Targeted subrole at program level'] else ElevateObservation.terminatingMessage("\"Targeted subrole at program level\" must not be Empty in \"Program details\" sheet")  
                        global rolesPGMID
                        rolesPGMID=rolesPGM.lstrip().rstrip().split(",")
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
                        # scopeEntityType = "state"
                        global entitiesType
                        entitiesType = ElevateObservation.fetchEntityType(parentFolder, accessToken,
                                                    entitiesPGM.lstrip().rstrip().split(","), scopeEntityType)
                        if entitiesPGM:
                            entitiesPGM = entitiesPGM
                            scopeEntityType = entitiesType

                        global entitiesPGMID
                        entitiesPGMID = ElevateObservation.fetchEntityId(parentFolder, accessToken,
                                                    entitiesPGM.lstrip().rstrip().split(","), scopeEntityType)
                        global orgIds
                        


                        if not ElevateObservation.getProgramInfo(accessToken, parentFolder, programNameInp.encode('utf-8').decode('utf-8')):
                            extIdPGM = dictDetailsEnv['Program ID'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Program ID'] else ElevateObservation.terminatingMessage("\"Program ID\" must not be Empty in \"Program details\" sheet")
                            if str(dictDetailsEnv['Program ID']).strip() == "Do not fill this field":
                                ElevateObservation.terminatingMessage("change the program id")
                            descriptionPGM = dictDetailsEnv['Description of the Program'].encode('utf-8').decode('utf-8') if dictDetailsEnv[
                                'Description of the Program'] else ElevateObservation.terminatingMessage(
                                "\"Description of the Program\" must not be Empty in \"Program details\" sheet")
                            keywordsPGM = dictDetailsEnv['Keywords'].encode('utf-8').decode('utf-8')
                            entitiesPGM = dictDetailsEnv['Targeted entities at program level'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Targeted entities at program level'] else ElevateObservation.terminatingMessage("\"Targeted entities at program level\" must not be Empty in \"Program details\" sheet")
                            stateEntitiesPGM = dictDetailsEnv['Targeted state at program level'].encode('utf-8').decode('utf-8')
                            # selecting entity type based on the users input 
                            if entitiesPGM:
                                entitiesPGM = entitiesPGM
                                scopeEntityType = entitiesType


                            mainRole = dictDetailsEnv['Targeted role at program level'] if dictDetailsEnv['Targeted role at program level'] else ElevateObservation.terminatingMessage("\"Targeted role at program level\" must not be Empty in \"Program details\" sheet")
                            # global rolesPGM
                            rolesPGM = dictDetailsEnv['Targeted subrole at program level'] if dictDetailsEnv['Targeted subrole at program level'] else ElevateObservation.terminatingMessage("\"Targeted subrole at program level\" must not be Empty in \"Program details\" sheet")
                            
                            if "teacher" in mainRole.strip().lower():
                                rolesPGM = str(rolesPGM).strip() + ",TEACHER"
                            userDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dictDetailsEnv['Elevate username/user id/email id/phone no. of Program Designer'])
                            OrgName=userDetails[4]
                            # orgIds=fetchOrgId(environment, accessToken, parentFolder, OrgName)
                            creatorKeyCloakId = userDetails[0]
                            creatorName = userDetails[2]
                            
                            messageArr = []

                            scopeEntityType = entitiesType
                            # fetch entity details 
                            entitiesPGMID = ElevateObservation.fetchEntityId(parentFolder, accessToken,entitiesPGM.lstrip().rstrip().split(","), scopeEntityType)
                            
                            # sys.exit()
                            # fetch sub-role details 
                            # rolesPGMID = fetchScopeRole(parentFolder, accessToken, rolesPGM.lstrip().rstrip().split(","))
                            # global rolesPGMID
                            rolesPGMID=rolesPGM.lstrip().rstrip().split(",")
                            # sys.exit()
                            # call function to create program 
                            ElevateObservation.programCreation(accessToken, parentFolder, extIdPGM, programNameInp, descriptionPGM,keywordsPGM.lstrip().rstrip().split(","), entitiesPGMID, rolesPGMID, orgIds,creatorKeyCloakId, creatorName,entitiesPGM,mainRole,rolesPGM)
                            # sys.exit()
                            # programmappingpdpmsheetcreation(MainFilePath, accessToken, program_file, extIdPGM,parentFolder)

                            # map PM / PD to the program 
                            # Programmappingapicall(MainFilePath, accessToken, program_file,parentFolder)

                            # check if program is created or not 
                            if ElevateObservation.getProgramInfo(accessToken, parentFolder, programNameInp):
                                print("Program Created SuccessFully.")
                            else :
                                ElevateObservation.terminatingMessage("Program creation failed! Please check logs.")
                        else :
                            userDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dictDetailsEnv['Elevate username/user id/email id/phone no. of Program Designer'])
                            OrgName=userDetails[4]
                            # orgIds=fetchOrgId(environment, accessToken, parentFolder, OrgName)
                            creatorKeyCloakId = userDetails[0]
                            creatorName = userDetails[2]
                            ElevateObservation.programCreation(accessToken, parentFolder, extIdPGM, programNameInp, descriptionPGM,keywordsPGM.lstrip().rstrip().split(","), entitiesPGMID, rolesPGMID, orgIds,creatorKeyCloakId, creatorName,entitiesPGM,mainRole,rolesPGM)
                            ElevateObservation.getProgramInfo(accessToken, parentFolder, extIdPGM)

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
                        resourceNamePGM = dictDetailsEnv['Name of resources in program'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Name of resources in program'] else ElevateObservation.terminatingMessage("\"Name of resources in program\" must not be Empty in \"Resource Details\" sheet")
                        resourceTypePGM = dictDetailsEnv['Type of resources'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Type of resources'] else ElevateObservation.terminatingMessage("\"Type of resources\" must not be Empty in \"Resource Details\" sheet")
                        resourceLinkOrExtPGM = dictDetailsEnv['Resource Link']
                        resourceStatusOrExtPGM = dictDetailsEnv['Resource Status'] if dictDetailsEnv['Resource Status'] else ElevateObservation.terminatingMessage("\"Resource Status\" must not be Empty in \"Resource Details\" sheet")
                        # setting start and end dates globally. 
                        global startDateOfResource, endDateOfResource
                        startDateOfResource = dictDetailsEnv['Start date of resource']
                        endDateOfResource = dictDetailsEnv['End date of resource']
                        # checking resource types and calling relevant functions 
                        # if resourceTypePGM.lstrip().rstrip().lower() == "course":
                        #     coursemapping = courseMapToProgram(accessToken, resourceLinkOrExtPGM, parentFolder)
                        #     if startDateOfResource:
                        #         startDateArr = str(startDateOfResource).split("-")
                        #         bodySolutionUpdate = {"startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                        #         solutionUpdate(parentFolder, accessToken, coursemapping, bodySolutionUpdate)
                        #     if endDateOfResource:
                        #         endDateArr = str(endDateOfResource).split("-")
                        #         bodySolutionUpdate = {
                        #             "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                        #         solutionUpdate(parentFolder, accessToken, coursemapping, bodySolutionUpdate)

    def typeofresource(filePathAddObs, accessToken, parentFolder):
        global criteriaLevelsReport, scopeRoles, criteriaLevels, scopeEntityType , ccRootOrgName , ccRootOrgId, errorVar
        wbObservation1 = xlrd.open_workbook(filePathAddObs, on_demand=True)
        sheetNames1 = wbObservation1.sheet_names()
        ecmIds = list()
        criteriaLevels = list()
        criteriaExternalIds = list()
        rubrics_sheet_names = ['Instructions', 'details', 'framework', 'ECMs or Domains', 'questions','Criteria_Rubric-Scoring', 'Domain(theme)_rubric_scoring']
        rubrics_sheet_IMP_names = ['Instructions', 'details', 'framework', 'ECMs or Domains', 'questions','Criteria_Rubric-Scoring', 'Domain(theme)_rubric_scoring', 'Imp mapping']
        observation_sheet_names = ['Instructions', 'details', 'criteria', 'questions']
        survey_sheet_names = ['Instructions', 'details', 'questions']
        project_sheet_names = ['Instructions', 'Project upload', 'Tasks upload','Certificate details']

        # 1-with rubrics , 2 - with out rubrics , 3 - survey , 4 - Project 5 - With rubric and IMP

        global environment, observationId, solutionName, pointBasedValue, entityType, allow_multiple_submissions, programName, userEntity, roles, isProgramnamePresent, solutionLanguage, keyWords, entityTypeId, solutionDescription, creator, dikshaLoginId
        if (len(rubrics_sheet_names) == len(sheetNames1)) and ((set(rubrics_sheet_names) == set(sheetNames1))):
            print("--->Observation with rubrics file detected.<---")
            typeofSolution = 1
        elif (len(observation_sheet_names) == len(sheetNames1)) and ((set(observation_sheet_names) == set(sheetNames1))):
            print("--->Observation without rubrics file detected.<---")
            typeofSolution = 2
        elif (len(survey_sheet_names) == len(sheetNames1)) and ((set(survey_sheet_names) == set(sheetNames1))):
            print("--->Survey file detected.<---")
            typeofSolution = 3
        elif (len(project_sheet_names) == len(sheetNames1)) and ((set(project_sheet_names) == set(sheetNames1))):
            print("--->Project file detected.<---")
            typeofSolution = 4
        elif (len(rubrics_sheet_IMP_names) == len(sheetNames1)) and ((set(rubrics_sheet_IMP_names) == set(sheetNames1))):
            print("--->Observation with rubrics and IMP file detected.<---")
            typeofSolution = 5
        else:
            typeofSolution = 0
            print(typeofSolution)
            errorVar = ("Please check the Input sheet.")
        return typeofSolution
    
    def criteriaUpload(solutionName_for_folder_path, wbObservation, millisAddObs, accessToken, tabName, projectDrivenFlag):
        criteriaColNames = ["criteriaId", "criteria_name"]
        criteriaSheet = wbObservation.sheet_by_name(tabName)
        keys = [criteriaSheet.cell(1, col_index).value for col_index in range(criteriaSheet.ncols)]
        criteriaUploadFieldnames = ['criteriaID', 'criteriaName']
        dictCriteriaToCsv = dict()
        criteriaLevelsFromFramework = dict()
        global criteriaLevelsCount
        if tabName == "framework":
            fetchLevelsFromFramework = wbObservation.sheet_by_name('framework')
            if projectDrivenFlag:
                criteriaImpDict = {}
                impsToCriteria = wbObservation.sheet_by_name('Imp mapping')
                keysFromImpSheet = [impsToCriteria.cell(1, col_index).value for col_index in range(impsToCriteria.ncols)]
                for row_indexImp in range(2, impsToCriteria.nrows):
                    dictImp = {keysFromImpSheet[col_index]: impsToCriteria.cell(row_indexImp, col_index).value for col_index in range(impsToCriteria.ncols)}
                    criteriaImpDict[dictImp['criteriaId'].strip()] = {}
                    for levls in range(1, countImps + 1):
                        criteriaImpDict[dictImp['criteriaId'].strip()].update({'L' + str(levls) + '-improvement-projects': dictImp['L' + str(levls) + '-improvement-projects'].strip()})

            keysFromFrameWork = [fetchLevelsFromFramework.cell(1, col_index).value for col_index in
                                range(fetchLevelsFromFramework.ncols)]
            levelCount = 1

            for eachHeaders in keysFromFrameWork:
                if eachHeaders == "L" + str(levelCount) + " description":
                    levelCount += 1
            levelCount = levelCount - 1

            for row_indexFrameWork in range(2, fetchLevelsFromFramework.nrows):
                dictFramework = {
                    keysFromFrameWork[col_index]: fetchLevelsFromFramework.cell(row_indexFrameWork, col_index).value for
                    col_index in range(fetchLevelsFromFramework.ncols)}
                criteriaLevelsFromFramework[dictFramework["Criteria ID"]] = {}

                for levlsNo in range(1, levelCount + 1):
                    criteriaLevelsFromFramework[dictFramework["Criteria ID"]].update(
                        {"L" + str(levlsNo): dictFramework["L" + str(levlsNo) + " description"]})
                    if not "L" + str(levlsNo) in criteriaColNames:
                        criteriaColNames.append("L" + str(levlsNo))

            for row_index in range(2, criteriaSheet.nrows):
                dictCriteria = {keys[col_index]: criteriaSheet.cell(row_index, col_index).value for col_index in
                                range(criteriaSheet.ncols)}
                dictCriteriaToCsv = {}

                dictCriteriaToCsv['criteriaID'] = dictCriteria['Criteria ID'].strip() + '_' + str(millisAddObs)
                criteriaLookUp[dictCriteriaToCsv['criteriaID'].strip()] = dictCriteria['Criteria Name'].encode('utf-8').decode('utf-8')
                dictCriteriaToCsv['criteriaName'] = dictCriteria['Criteria Name'].encode('utf-8').decode('utf-8')
                criteriaName = dictCriteria['Criteria Name'].encode('utf-8').decode('utf-8')
                dictCriteriaToCsv['type'] = 'auto'
                for levlsNo in range(1, levelCount + 1):
                    dictCriteriaToCsv['L' + str(levlsNo)] = dictCriteria["L" + str(levlsNo) + " description"]
                if projectDrivenFlag:
                    for eachImps in criteriaImpDict[dictCriteria['Criteria ID'].strip()]:
                        dictCriteriaToCsv[eachImps] = criteriaImpDict[dictCriteria['Criteria ID'].strip()][eachImps]

                if not 'type' in criteriaUploadFieldnames:
                    criteriaUploadFieldnames.append('type')
                for eachCols in criteriaColNames:
                    if not eachCols in ['criteria_id', 'criteria_name', 'type', "criteriaId"]:
                        if not eachCols in criteriaUploadFieldnames:
                            criteriaUploadFieldnames.append(eachCols)
                if projectDrivenFlag:
                    for levls in range(1, countImps + 1):
                        if not (str('L' + str(levls) + '-improvement-projects') in criteriaUploadFieldnames):
                            criteriaUploadFieldnames.append('L' + str(levls) + '-improvement-projects')
                criteriaFilePath = solutionName_for_folder_path + '/criteriaUpload/'
                file_exists = os.path.isfile(solutionName_for_folder_path + '/criteriaUpload/uploadSheet.csv')
                criteriaLevelsCount = levelCount
                if not os.path.exists(criteriaFilePath):
                    os.mkdir(criteriaFilePath)
                with open(solutionName_for_folder_path + '/criteriaUpload/uploadSheet.csv', 'a',encoding='utf-8') as criteriaUploadFile:
                    writerCriteriaUpload = csv.DictWriter(criteriaUploadFile, fieldnames=list(criteriaUploadFieldnames),
                                                        lineterminator='\n')
                    if not file_exists:
                        writerCriteriaUpload.writeheader()
                    writerCriteriaUpload.writerow(dictCriteriaToCsv)
                    
        elif tabName == "criteria":
            criteriaSheet = wbObservation.sheet_by_name(tabName)
            keys = [criteriaSheet.cell(1, col_index).value for col_index in range(criteriaSheet.ncols)]
            for row_index in range(2, criteriaSheet.nrows):
                dictCriteria = {keys[col_index]: criteriaSheet.cell(row_index, col_index).value for col_index in
                                range(criteriaSheet.ncols)}
                dictCriteria['criteriaID'] = dictCriteria['criteria_id'].encode('utf-8').decode('utf-8').strip() + '_' + str(millisAddObs)
                criteriaLookUp[dictCriteria['criteriaID']] = dictCriteria['criteria_name'].encode('utf-8').decode('utf-8')
                del dictCriteria['criteria_id']
                dictCriteria['criteriaName'] = dictCriteria['criteria_name'].encode('utf-8').decode('utf-8')
                criteriaName = dictCriteria['criteria_name']
                del dictCriteria['criteria_name']
                dictCriteria['L1'] = 'NA'
                dictCriteria['type'] = 'auto'
                criteriaFilePath = solutionName_for_folder_path + '/criteriaUpload/'
                file_exists = os.path.isfile(solutionName_for_folder_path + '/criteriaUpload/uploadSheet.csv')
                if not os.path.exists(criteriaFilePath):
                    os.mkdir(criteriaFilePath)
                criteriaUploadFieldnames = []
                criteriaUploadFieldnames = ['criteriaID', 'criteriaName', 'L1', 'type']
                with open(solutionName_for_folder_path + '/criteriaUpload/uploadSheet.csv', 'a',encoding='utf-8') as criteriaUploadFile:
                    writerCriteriaUpload = csv.DictWriter(criteriaUploadFile, fieldnames=criteriaUploadFieldnames,
                                                        lineterminator='\r')
                    if not file_exists:
                        writerCriteriaUpload.writeheader()
                    writerCriteriaUpload.writerow(dictCriteria)

        urlCriteriaUploadApi = config.get(environment, 'INTERNAL_KONG_IP')+config.get(environment, 'criteriaUploadApiUrl')
        headerCriteriaUploadApi = {
            'Authorization': config.get(environment, 'Authorization'),
            'X-auth-token': accessToken,
            'X-Channel-id': config.get(environment, 'X-Channel-id')
        }
        filesCriteria = {
            'criteria': open(solutionName_for_folder_path + '/criteriaUpload/uploadSheet.csv', 'rb')
        }

        responseCriteriaUploadApi = requests.post(url=urlCriteriaUploadApi, headers=headerCriteriaUploadApi,
                                                files=filesCriteria)
        messageArr = ["Criteria Upload Sheet Prepared.",
                    "File path : " + solutionName_for_folder_path + '/criteriaUpload/uploadSheet.csv']
        messageArr.append("Upload status code : " + str(responseCriteriaUploadApi.status_code))
        ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)

        if responseCriteriaUploadApi.status_code == 200:
            print('CriteriaUploadApi Success')
            with open(solutionName_for_folder_path + '/criteriaUpload/uploadInternalIdsSheet.csv', 'w+',encoding='utf-8') as criteriaRes:
                criteriaRes.write(responseCriteriaUploadApi.text)
        else:

            messageArr.append("Response : " + str(responseCriteriaUploadApi.text))
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
        ElevateObservation.terminatingMessage("Criteria Upload failed.")
    
    def frameWorkUpload(solutionName_for_folder_path, wbObservation, millisAddObs, accessToken):
        global criteriaLevelsReport
        dateTime = datetime.now()
        frameworkDocInsertObj = {}
        frameworkExternalId = None
        frameworkExternalId = uuid.uuid1()
        frameworkExternalId = str(frameworkExternalId)
        frameworkDocInsertObj['externalId'] = frameworkExternalId
        frameworkDocInsertObj['name'] = solutionName.strip()
        frameworkDocInsertObj['description'] = solutionDescription
        frameworkDocInsertObj['parentId'] = None
        frameworkDocInsertObj['resourceType'] = ['Observations Framework']
        frameworkDocInsertObj['language'] = solutionLanguage
        frameworkDocInsertObj['levelToScoreMapping'] = dict()
        if keyWords and (keyWords != 'Framework' or keyWords != 'Frameworks' or keyWords != 'Observation' or keyWords != 'Observations'):
            keywordsFinalArr = ['Framework', 'Observation']
            keywordsArr = keyWords.encode('utf-8').decode('utf-8').split(',')
            for keyw in keywordsArr:
                keywordsFinalArr.append(keyw)
            frameworkDocInsertObj['keywords'] = keywordsFinalArr
        else:
            frameworkDocInsertObj['keywords'] = ['Framework', 'Observation']
        frameworkDocInsertObj['concepts'] = []
        frameworkDocInsertObj['createdFor'] = [ccRootOrgId]  # createdForArr
        frameworkDocInsertObj['rootOrg'] = [ccRootOrgId]  # rootOrgArr
        criteriaFrameworkArr = []
        with open(solutionName_for_folder_path + '/criteriaUpload/uploadInternalIdsSheet.csv', 'r',encoding='utf-8') as criteriaInternalFile:
            criteriaInternalReader = csv.DictReader(criteriaInternalFile)
            criteriaWeightage = 100 / (len(list(criteriaInternalReader)))
            criteriaInternalFile.seek(0, 0)
            next(criteriaInternalReader, None)
            for crit in criteriaInternalReader:
                dictCritInter = {}
                dictCritInter = dict(crit)
                criteriaFrameworkObj = {
                    'criteriaId': str(ObjectId(dictCritInter['Criteria Internal Id'])),
                    'weightage': criteriaWeightage
                }
                criteriaFrameworkArr.append(criteriaFrameworkObj)
        criteriaInternalFile.close()
        frameworkDocInsertObj['themes'] = [{
            'type': 'theme',
            'label': 'theme',
            'name': 'Observation Theme',
            'externalId': 'OB',
            'weightage': 100,
            'criteria': criteriaFrameworkArr
        }]

        if not pointBasedValue.lower() == "null":
            frameworkDocInsertObj['flattenedThemes'] = {
                "type": "theme",
                "label": "theme",
                "name": "Observation Theme",
                "externalId": "OB",
                "weightage": 1,
                "criteria": criteriaFrameworkArr,
                "rubric": {
                    "expressionVariables": {
                        "SCORE": "OB.sumOfPointsOfAllChildren()"
                    },
                    "levels": {
                        "L1": {
                            "expression": "(0<=SCORE<=100000)"
                        }
                    }
                },
                "hierarchyLevel": 0,
                "hierarchyTrack": []
            }
            frameworkDocInsertObj['scoringSystem'] = pointBasedValue
            frameworkDocInsertObj['isRubricDriven'] = True
            criteriaLevelsReport = True
            frameworkDocInsertObj['themes'] = [{
                'type': 'theme',
                'label': 'theme',
                'name': 'Observation Theme',
                'externalId': 'OB',
                'weightage': 100,
                'criteria': criteriaFrameworkArr,
                "rubric": {
                    "expressionVariables": {
                        "SCORE": "OB.sumOfPointsOfAllChildren()"
                    },
                    "levels": {
                        "L1": {
                            "expression": "(0<=SCORE<=100000)"
                        }
                    }
                }
            }]
            for levs in range(1, criteriaLevelsCount + 1):
                levelToScore = {"L" + str(levs): {'points': levs * 10, 'label': 'Level ' + str(levs)}}
                frameworkDocInsertObj['levelToScoreMapping'].update(levelToScore)
            frameworkDocInsertObj['noOfRatingLevels'] = criteriaLevelsCount
            
        else:
            frameworkDocInsertObj['scoringSystem'] = None
            frameworkDocInsertObj['isRubricDriven'] = False

        frameworkDocInsertObj['entityTypeId'] = entityTypeId
        frameworkDocInsertObj['entityType'] = entityType
        frameworkDocInsertObj['type'] = 'observation'
        frameworkDocInsertObj['subType'] = entityType
        frameworkDocInsertObj['status'] = "active"
        frameworkDocInsertObj['updatedBy'] = 'INITIALIZE'
        frameworkDocInsertObj['createdBy'] = 'INITIALIZE'
        frameworkDocInsertObj['createdAt'] = str(dateTime)
        frameworkDocInsertObj['updatedAt'] = str(dateTime)
        frameworkDocInsertObj['author'] = matchedShikshalokamLoginId
        frameworkDocInsertObj['isTempObTest'] = 'observationAutomation'

        # Adding Credits and license into Frameworks
        frameworkDocInsertObj['creator'] = str(creator)
        frameworkDocInsertObj['license'] = {}
        frameworkDocInsertObj['license']['author'] = str(creator)
        frameworkDocInsertObj['license']['creator'] = str(creator)
        frameworkDocInsertObj['license']['copyright'] = str(ccRootOrgName)
        frameworkDocInsertObj['license']['copyrightYear'] = int(dateTime.strftime("%Y"))
        frameworkDocInsertObj['license']['contentType'] = "Observation"
        frameworkDocInsertObj['license']['organisation'] = [ccRootOrgName]
        frameworkDocInsertObj['license']['orgDetails'] = {}
        frameworkDocInsertObj['license']['orgDetails']['email'] = None
        frameworkDocInsertObj['license']['orgDetails']['orgName'] = ccRootOrgName
        frameworkDocInsertObj['license']['licenseDetails'] = {}
        frameworkDocInsertObj['license']['licenseDetails']['name'] = "CC BY 4.0"
        frameworkDocInsertObj['license']['licenseDetails']['url'] = "https://creativecommons.org/licenses/by/4.0/legalcode"
        frameworkDocInsertObj['license']['licenseDetails']['description'] = "For details see below:"
    
        urlCreateFrameworkApi = config.get(environment, 'INTERNAL_KONG_IP') + config.get(environment, 'frameworkCreationApi')
        frameworkFilePath = solutionName_for_folder_path + '/framework/'
        file_exists_framework = os.path.isfile(solutionName_for_folder_path + '/framework/uploadFile.json')
        if not os.path.exists(frameworkFilePath):
            os.mkdir(frameworkFilePath)

        with open(frameworkFilePath + "uploadFile.json", "w",encoding='utf-8') as outfile:
            json.dump(frameworkDocInsertObj, outfile)
        headerFrameworkUploadApi = {'Authorization': config.get(environment, 'Authorization'),
                                    'X-auth-token': accessToken,
                                    'X-Channel-id': config.get(environment, 'X-Channel-id')}
        filesFramework = {'framework': open(solutionName_for_folder_path + '/framework/uploadFile.json', 'rb')}

        responseFrameworkUploadApi = requests.post(url=urlCreateFrameworkApi, headers=headerFrameworkUploadApi,
                                                files=filesFramework)
        messageArr = ["Framwork json file created.",
                    "File loc : " + solutionName_for_folder_path + '/framework/uploadFile.json',
                    "Framework upload API called,", "Status code : " + str(responseFrameworkUploadApi.status_code)]
        ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
        if responseFrameworkUploadApi.status_code == 200:
            print('Framework upload Success')
            return frameworkExternalId

        else:
            messageArr = ["Framwork upload Failed.", "Response : " + responseFrameworkUploadApi.text]
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            print('Framework upload api failed in ' + environment,
                'status_code response from api is ' + str(responseFrameworkUploadApi.status_code))
            sys.exit()
            
    def themesUpload(solutionName_for_folder_path, wbObservation, millisAddObs, accessToken, frameworkExternalId,obsWORubWS):
        global dictCritLookUp
        with open(solutionName_for_folder_path + '/criteriaUpload/uploadInternalIdsSheet.csv', 'r',encoding='utf-8') as criteriaInternalFile:
            criteriaInternalReader = csv.DictReader(criteriaInternalFile)
            for crit in criteriaInternalReader:
                dictCritLookUp[crit['Criteria External Id']] = crit['Criteria Internal Id']
        if obsWORubWS:
            print("Themes Observation without rubrics with scores")
            themeUploadFieldnames = ["theme", "aoi", "indicators", "criteriaInternalId"]
            themesUploadCsv = dict()
            for dictCritLookUpKey, dictCritLookUpValue in dictCritLookUp.items():
                themesUploadCsv['theme'] = "Observation Theme" + "###" + "OB" + "###40"
                themesUploadCsv['aoi'] = ""
                themesUploadCsv['indicators'] = ""
                themesUploadCsv['criteriaInternalId'] = dictCritLookUpValue + "###40"
                themeFilePath = solutionName_for_folder_path + '/themeUpload/'
                file_exists = os.path.isfile(solutionName_for_folder_path + '/themeUpload/uploadSheet.csv')

                if not os.path.exists(themeFilePath):
                    os.mkdir(themeFilePath)
                with open(solutionName_for_folder_path + '/themeUpload/uploadSheet.csv', 'a',encoding='utf-8') as themeUploadFile:
                    writerthemeUpload = csv.DictWriter(themeUploadFile, fieldnames=list(themeUploadFieldnames),
                                                    lineterminator='\n')
                    if not file_exists:
                        writerthemeUpload.writeheader()
                    writerthemeUpload.writerow(themesUploadCsv)

        else:
            frameWorkSheet = wbObservation.sheet_by_name('framework')
            keys = [frameWorkSheet.cell(1, col_index).value for col_index in range(frameWorkSheet.ncols)]
            themeUploadFieldnames = ["theme", "aoi", "indicators", "criteriaInternalId"]
            themesUploadCsv = dict()
            for row_index in range(2, frameWorkSheet.nrows):
                dictCriteria = {keys[col_index]: frameWorkSheet.cell(row_index, col_index).value for col_index in
                                range(frameWorkSheet.ncols)}
                themesUploadCsv['theme'] = dictCriteria['Domain Name'].encode('utf-8').decode('utf-8') + "###" + dictCriteria['Domain ID'] + "###40"
                themesUploadCsv['aoi'] = ""
                themesUploadCsv['indicators'] = ""
                themesUploadCsv['criteriaInternalId'] = dictCritLookUp[dictCriteria['Criteria ID'].strip() + '_' + str(
                    millisAddObs)] + "###40"  # if dictCriteria['Criteria ID'] else  ""
                themeFilePath = solutionName_for_folder_path + '/themeUpload/'
                file_exists = os.path.isfile(solutionName_for_folder_path + '/themeUpload/uploadSheet.csv')

                if not os.path.exists(themeFilePath):
                    os.mkdir(themeFilePath)
                with open(solutionName_for_folder_path + '/themeUpload/uploadSheet.csv', 'a',encoding='utf-8') as themeUploadFile:
                    writerthemeUpload = csv.DictWriter(themeUploadFile, fieldnames=list(themeUploadFieldnames),
                                                    lineterminator='\n')
                    if not file_exists:
                        writerthemeUpload.writeheader()
                    writerthemeUpload.writerow(themesUploadCsv)

        urlThemesUploadApi = config.get(environment, 'INTERNAL_KONG_IP')+config.get(environment, 'themeUploadApiUrl') + frameworkExternalId
        headerThemesUploadApi = {'Authorization': config.get(environment, 'Authorization'),
                                'X-auth-token': accessToken,
                                'X-Channel-id': config.get(environment, 'X-Channel-id')}
        filesThemes = {'themes': open(solutionName_for_folder_path + '/themeUpload/uploadSheet.csv', 'rb')}
        responseThemeUploadApi = requests.post(url=urlThemesUploadApi, headers=headerThemesUploadApi, files=filesThemes)
        messageArr = ["Themes upload sheet prepared.",
                    "File path : " + solutionName_for_folder_path + '/themeUpload/uploadSheet.csv',
                    "Theme upload to framework API called.", "URL : " + urlThemesUploadApi,
                    "Status code : " + str(responseThemeUploadApi.status_code)]
        ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
        if responseThemeUploadApi.status_code == 200:
            print('Theme UploadApi Success')
            with open(solutionName_for_folder_path + '/themeUpload/uploadInternalIdsSheet.csv', 'w+',encoding='utf-8') as criteriaRes:
                criteriaRes.write(responseThemeUploadApi.text)
        else:
            messageArr = ["Themes upload failed.", "Response : " + str(responseThemeUploadApi.text)]
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            ElevateObservation.terminatingMessage("Theme upload failed.")

    def createSolutionFromFramework(solutionName_for_folder_path, accessToken, frameworkExternalId):
        urlCreateSolutionApi = config.get(environment, 'INTERNAL_KONG_IP')+ config.get(environment, 'solutionCreationApiUrl')
        headerCreateSolutionApi = {
            'Content-Type': config.get(environment, 'Content-Type'),
            'Authorization': config.get(environment, 'Authorization'),
            'X-auth-token': accessToken,
            'X-Channel-id': config.get(environment, 'X-Channel-id')
        }
        queryparamsCreateSolutionApi = '?frameworkId=' + str(frameworkExternalId) + '&entityType=' + entityType
        responseCreateSolutionApi = requests.post(url=urlCreateSolutionApi + queryparamsCreateSolutionApi,
                                                headers=headerCreateSolutionApi)

        messageArr = ["Solution Created from Framework.",
                    "URL : " + str(urlCreateSolutionApi + queryparamsCreateSolutionApi),
                    "Status Code : " + str(responseCreateSolutionApi.status_code),
                    "Response : " + str(responseCreateSolutionApi.text)]
        ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
        messageArr = []
        if responseCreateSolutionApi.status_code == 200:
            responseCreateSolutionApi = responseCreateSolutionApi.json()
            solutionId = responseCreateSolutionApi['result']['templateId']
            messageArr.append("Parent Solution Generated : " + str(solutionId))
            print("Parent Solution Generated : " + str(solutionId))
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
        else:
            messageArr.append("Solution from framework api failed.")
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            ElevateObservation.terminatingMessage("Solution from framework api failed.")
        return solutionId

    def solutionUpdate(solutionName_for_folder_path, accessToken, solutionId, bodySolutionUpdate):
        solutionUpdateApi = config.get(environment, 'INTERNAL_KONG_IP') + config.get(environment, 'solutionUpdateApi') + str(solutionId)
        headerUpdateSolutionApi = {
            'Content-Type': 'application/json',
            'Authorization': config.get(environment, 'Authorization'),
            'X-auth-token': accessToken,
            'X-Channel-id': config.get(environment, 'X-Channel-id'),
            "internal-access-token": config.get(environment, 'internal-access-token')
            }
        print(bodySolutionUpdate,"this is a solution update 2216")
        responseUpdateSolutionApi = requests.post(url=solutionUpdateApi, headers=headerUpdateSolutionApi,data=json.dumps(bodySolutionUpdate))
        messageArr = ["Solution Update API called.", "URL : " + str(solutionUpdateApi), "Body : " + str(bodySolutionUpdate),"Response : " + str(responseUpdateSolutionApi.text),"Status Code : " + str(responseUpdateSolutionApi.status_code)]
        ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
        if responseUpdateSolutionApi.status_code == 200:
            print("Solution Update Success.")
            return True
        else:
            print("Solution Update Failed.")
            return False
    
    def questionUpload(filePathAddObs, solutionName_for_folder_path, frameworkExternalId, millisAddObs, accessToken,
                   solutionId, typeofSolution):
        wbObservation = xlrd.open_workbook(filePathAddObs, on_demand=True)
        excelBook = open_workbook(filePathAddObs)
        sheetNam = excelBook.sheet_names()
        shCnt = 0
        countColSeq = 0
        questShee = wbObservation.sheet_by_name('questions')
        Qukeys = [questShee.cell(1, col_index).value for col_index in range(questShee.ncols)]
        countColSeq = Qukeys.index('question_sequence')
        questionsResponseDict = dict()

        for i in sheetNam:
            if i.strip().lower() == 'questions':
                sheetNam1 = excelBook.sheets()[shCnt]
            shCnt = shCnt + 1
        dataSort = [sheetNam1.row_values(i) for i in range(sheetNam1.nrows)]
        labels = dataSort[1]
        dataSort = dataSort[2:]
        dataSort.sort(key=lambda x: int(x[countColSeq]))
        openWorkBookSort = xlrd.open_workbook(filePathAddObs)
        openWorkBookSort1 = xl_copy(openWorkBookSort)
        sheet1 = openWorkBookSort1.add_sheet('questions_sequence_sorted')
        print("Question Sorted.")
        for idx, label in enumerate(labels):
            sheet1.write(0, idx, label)

        for idx_r, row in enumerate(dataSort):
            for idx_c, value in enumerate(row):
                sheet1.write(idx_r + 1, idx_c, value)

        openWorkBookSort1.save(filePathAddObs)
        wbObservation = xlrd.open_workbook(filePathAddObs, on_demand=True)
        questionsSheet = wbObservation.sheet_by_name('questions_sequence_sorted')
        keys2 = [questionsSheet.cell(0, col_index2).value for col_index2 in range(questionsSheet.ncols)]
        questionsList = list()
        for row_index2 in range(1, questionsSheet.nrows):
            d2 = {keys2[col_index2]: questionsSheet.cell(row_index2, col_index2).value for col_index2 in
                range(questionsSheet.ncols)}
            questionsList.append(d2)
        questionSeqByEcmDict = dict()
        questionSeqByEcmSectionDict = dict()
        questionSeqByEcmArr = []
        quesSeqCnt = 1.0
        questionUploadFieldnames = []
        questionUploadExceptSliderFieldnames = []
        questionUploadSliderFieldNames = []
        if typeofSolution == 1:
            for ques00 in questionsList:
                questionSeqByEcmDict[ecmToSection[ques00['section_id']] + "_" + str(millisAddObs)] = {
                    ecm_sections[ecmToSection[ques00['section_id']] + "_" + str(millisAddObs)]: []}
        elif typeofSolution == 2:
            questionSeqByEcmDict["OB"] = {
                "S1": []
            }

        for ques1 in questionsList:
            if not pointBasedValue.lower() == "null":
                questionUploadExceptSliderFieldnames = ['solutionId', 'criteriaExternalId', 'name', 'evidenceMethod',
                                                        'section', 'instanceParentQuestionId', 'hasAParentQuestion',
                                                        'parentQuestionOperator', 'parentQuestionValue', 'parentQuestionId',
                                                        'externalId', 'question0', 'question1', 'tip', 'hint',
                                                        'instanceIdentifier', 'responseType', 'dateFormat', 'autoCapture',
                                                        'validation', 'validationIsNumber', 'validationRegex',
                                                        'validationMax', 'validationMin', 'file', 'fileIsRequired',
                                                        'fileUploadType', 'allowAudioRecording', 'minFileCount',
                                                        'maxFileCount', 'caption', 'questionGroup', 'modeOfCollection',
                                                        'accessibility', 'showRemarks', 'rubricLevel', 'isAGeneralQuestion',
                                                        'R1', 'R1-hint', 'R2', 'R2-hint', 'R3', 'R3-hint', 'R4', 'R4-hint',
                                                        'R5', 'R5-hint', 'R6', 'R6-hint', 'R7', 'R7-hint', 'R8', 'R8-hint',
                                                        'R9', 'R9-hint', 'R10', 'R10-hint', 'R11', 'R11-hint', 'R12',
                                                        'R12-hint', 'R13', 'R13-hint', 'R14', 'R14-hint', 'R15', 'R15-hint',
                                                        'R16', 'R16-hint', 'R17', 'R17-hint', 'R18', 'R18-hint', 'R19',
                                                        'R19-hint', 'R20', 'R20-hint', 'R1-score', 'R2-score', 'R3-score',
                                                        'R4-score', 'R5-score', 'R6-score', 'R7-score', 'R8-score',
                                                        'R9-score', 'R10-score', 'R11-score', 'R12-score', 'R13-score',
                                                        'R14-score', 'R15-score', 'R16-score', 'R17-score', 'R18-score',
                                                        'R19-score', 'R20-score', 'weightage', 'sectionHeader', 'page',
                                                        'questionNumber', '_arrayFields', 'prefillFromEntityProfile',
                                                        'isEditable', 'entityFieldName']
                if ques1['question_response_type'].strip().lower() == 'slider' and ques1['slider_value_with_score'].strip():
                    noOfSliderColumn = ques1['slider_value_with_score'].strip().split(',')
                    possibleSliderColumn = (int(ques1['max_number_value']) + 1) - (int(ques1['min_number_value']))
                    sliderCnt = int(ques1['min_number_value'])
                    if len(noOfSliderColumn) == possibleSliderColumn:
                        for sliderIndex, sliCn in enumerate(noOfSliderColumn):
                            questionUploadSliderFieldNames.append('slider-value-' + str(sliderIndex + 1))
                            questionUploadSliderFieldNames.append('slider-value-' + str(sliderIndex + 1) + '-score')
            else:
                questionUploadFieldnames = ['solutionId', 'criteriaExternalId', 'name', 'evidenceMethod', 'section',
                                            'instanceParentQuestionId', 'hasAParentQuestion', 'parentQuestionOperator',
                                            'parentQuestionValue', 'parentQuestionId', 'externalId', 'question0',
                                            'question1', 'tip', 'hint', 'instanceIdentifier', 'responseType', 'dateFormat',
                                            'autoCapture', 'validation', 'validationIsNumber', 'validationRegex',
                                            'validationMax', 'validationMin', 'file', 'fileIsRequired', 'fileUploadType',
                                            'allowAudioRecording', 'minFileCount', 'maxFileCount', 'caption',
                                            'questionGroup', 'modeOfCollection', 'accessibility', 'showRemarks',
                                            'rubricLevel', 'isAGeneralQuestion', 'R1', 'R1-hint', 'R2', 'R2-hint', 'R3',
                                            'R3-hint', 'R4', 'R4-hint', 'R5', 'R5-hint', 'R6', 'R6-hint', 'R7', 'R7-hint',
                                            'R8', 'R8-hint', 'R9', 'R9-hint', 'R10', 'R10-hint', 'R11', 'R11-hint', 'R12',
                                            'R12-hint', 'R13', 'R13-hint', 'R14', 'R14-hint', 'R15', 'R15-hint', 'R16',
                                            'R16-hint', 'R17', 'R17-hint', 'R18', 'R18-hint', 'R19', 'R19-hint', 'R20',
                                            'R20-hint', 'sectionHeader', 'page', 'questionNumber', '_arrayFields',
                                            'prefillFromEntityProfile', 'isEditable', 'entityFieldName']
        if len(questionUploadExceptSliderFieldnames) > 0:
            if len(questionUploadSliderFieldNames) > 0:
                questionUploadFieldnames = questionUploadExceptSliderFieldnames + questionUploadSliderFieldNames
            else:
                questionUploadFieldnames = questionUploadExceptSliderFieldnames
        for ques in questionsList:
            questionFilePath = solutionName_for_folder_path + '/questionUpload/'
            file_exists_ques = os.path.isfile(solutionName_for_folder_path + '/questionUpload/uploadSheet.csv')
            if not os.path.exists(questionFilePath):
                os.mkdir(questionFilePath)
            with open(solutionName_for_folder_path + '/questionUpload/uploadSheet.csv', 'a',
                    encoding='utf-8') as questionUploadFile:
                writerQuestionUpload = csv.DictWriter(questionUploadFile, fieldnames=questionUploadFieldnames,
                                                    lineterminator='\n')
                if not file_exists_ques:
                    writerQuestionUpload.writeheader()
                questionFileObj = {}
                observationExternalId = None
                observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
                questionFileObj['solutionId'] = observationExternalId
                questionFileObj['criteriaExternalId'] = ques['criteria_id'].strip() + '_' + str(millisAddObs)
                try:
                    questionFileObj['name'] = criteriaLookUp[questionFileObj['criteriaExternalId']]
                except:
                    print("criteria Id error....")
                    print(questionFileObj['criteriaExternalId'] + " not found.")
                    sys.exit()
                if typeofSolution == 1 or typeofSolution == 5:
                    questionFileObj['evidenceMethod'] = ecmToSection[ques['section_id']] + "_" + str(millisAddObs)
                    questionFileObj['section'] = ques['section_id']
                elif typeofSolution == 2:
                    questionFileObj['evidenceMethod'] = "OB"
                    questionFileObj['section'] = "S1"
                questionsResponseDict[ques['question_id'].strip() + '_' + str(millisAddObs)] = {
                    "response(R1)": ques["response(R1)".replace(" ", "")],
                    "response(R2)": ques["response(R2)".replace(" ", "")],
                    "response(R3)": ques["response(R3)".replace(" ", "")],
                    "response(R4)": ques["response(R4)".replace(" ", "")],
                    "response(R5)": ques["response(R5)".replace(" ", "")],
                    "response(R6)": ques["response(R6)".replace(" ", "")],
                    "response(R7)": ques["response(R7)".replace(" ", "")],
                    "response(R8)": ques["response(R8)".replace(" ", "")],
                    "response(R9)": ques["response(R9)".replace(" ", "")],
                    "response(R10)": ques["response(R10)".replace(" ", "")],
                    "response(R11)": ques["response(R11)".replace(" ", "")],
                    "response(R12)": ques["response(R12)".replace(" ", "")],
                    "response(R13)": ques["response(R13)".replace(" ", "")],
                    "response(R14)": ques["response(R14)".replace(" ", "")],
                    "response(R15)": ques["response(R15)".replace(" ", "")],
                    "response(R16)": ques["response(R16)".replace(" ", "")],
                    "response(R17)": ques["response(R17)".replace(" ", "")],
                    "response(R18)": ques["response(R18)".replace(" ", "")],
                    "response(R19)": ques["response(R19)".replace(" ", "")],
                    "response(R20)": ques["response(R20)".replace(" ", "")]}
                hasInstanceParentFlag = False
                if ques['instance_parent_question_id'].encode('utf-8').decode('utf-8'):
                    hasInstanceParentFlag = True
                    questionFileObj['instanceParentQuestionId'] = ques['instance_parent_question_id'].encode('utf-8').decode('utf-8').strip() + '_' + str(
                        millisAddObs)
                    questionFileObj['hasAParentQuestion'] = 'NO'
                else:
                    hasInstanceParentFlag = False
                    questionFileObj['instanceParentQuestionId'] = 'NA'
                notEqualsFlag = False
                if ques['parent_question_id'].encode('utf-8').decode('utf-8').strip():
                    questionFileObj['hasAParentQuestion'] = 'YES'
                    if ques['show_when_parent_question_value_is'].encode('utf-8').decode('utf-8').lower().lstrip().rstrip() == 'or' or ques[
                        'show_when_parent_question_value_is'].encode('utf-8').decode('utf-8').lower().lstrip().rstrip() == '||':
                        notEqualsFlag = False
                        questionFileObj['parentQuestionOperator'] = '||'
                        questionFileObj['parentQuestionValue'] = ques['parent_question_value'].encode('utf-8').decode('utf-8').lstrip().rstrip().replace(
                            " ", "")
                    elif ques['show_when_parent_question_value_is'].lower().lstrip().rstrip() == 'equals':
                        notEqualsFlag = False
                        questionFileObj['parentQuestionOperator'] = "EQUALS"
                        questionFileObj['parentQuestionValue'] = ques['parent_question_value'].encode('utf-8').decode('utf-8').lstrip().rstrip().replace(
                            " ", "")
                    elif ques['show_when_parent_question_value_is'].encode('utf-8').decode('utf-8').lstrip().rstrip() == 'NOT_EQUALS_TO' or ques[
                        'show_when_parent_question_value_is'].encode('utf-8').decode('utf-8').lower().lstrip().rstrip() == 'NOT_EQUALS_TO'.lower():
                        notEqualsFlag = True
                        questionFileObj['parentQuestionOperator'] = "||"
                    else:
                        questionFileObj['parentQuestionOperator'] = ""
                    if type(ques['parent_question_value']) != str:
                        if (ques['parent_question_value'] and ques['parent_question_value'].is_integer() == True):
                            questionFileObj['parentQuestionValue'] = int(ques['parent_question_value'])
                        elif (ques['parent_question_value'] and ques['parent_question_value'].is_integer() == False):
                            questionFileObj['parentQuestionValue'] = ques[
                                'parent_question_value'].encode('utf-8').decode('utf-8').lstrip().rstrip().replace(" ", "")
                    else:
                        questionFileObj['parentQuestionId'] = ques['parent_question_id'].encode('utf-8').decode('utf-8').strip() + '_' + str(millisAddObs)
                        if notEqualsFlag:
                            Qkeys = ques.keys()
                            final_parent_question_value = str()
                            avoidResponses = ques['parent_question_value'].lstrip().rstrip().split(",")
                            for i in Qkeys:
                                searchResponse = re.search("^response\(R[0-9]\)$|^response\(R[0-2][0-9]\)$", i)
                                if searchResponse:
                                    try:
                                        responseCheck = questionsResponseDict[questionFileObj['parentQuestionId']][
                                            searchResponse.string]
                                    except:
                                        print(questionFileObj[
                                                'parentQuestionId'] + " Referenced before intialising in questions sheet.")
                                        print("Please check question sequesnce...")
                                        print("Aborting...")
                                        messageArr = [questionFileObj[
                                                        'parentQuestionId'] + " Referenced before intialising in questions sheet.",
                                                    "Please check question sequesnce...", ]
                                        ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                                        sys.exit()
                                    if responseCheck:
                                        if not searchResponse.string.replace("response(", "").replace(")",
                                                                                                    "") in avoidResponses:
                                            final_parent_question_value += searchResponse.string.replace("response(",
                                                                                                        "").replace(")",
                                                                                                                    "") + ","
                            questionFileObj['parentQuestionValue'] = final_parent_question_value.encode('utf-8').decode('utf-8').rstrip(",").lstrip(",")
                        else:
                            pass
                else:
                    questionFileObj['parentQuestionOperator'] = None
                    questionFileObj['parentQuestionValue'] = None
                    questionFileObj['parentQuestionId'] = None
                questionFileObj['externalId'] = ques['question_id'].strip() + '_' + str(millisAddObs)
                if typeofSolution == 1:
                    questionSeqByEcmDict[questionFileObj['evidenceMethod']][
                        ecm_sections[questionFileObj['evidenceMethod']]].append(
                        ques['question_id'].strip() + '_' + str(millisAddObs))
                elif typeofSolution == 2:
                    questionSeqByEcmDict["OB"]["S1"].append(ques['question_id'].strip() + '_' + str(millisAddObs))

                questionFileObj['question0'] = ques['question_primary_language'].encode('utf-8').decode('utf-8')
                if not questionFileObj['question0']:
                    questionFileObj['question0'] = None
                if ques['question_secondory_language']:
                    questionFileObj['question1'] = ques['question_secondory_language'].encode('utf-8').decode('utf-8')
                else:
                    questionFileObj['question1'] = None
                if ques['question_tip']:
                    questionFileObj['tip'] = ques['question_tip'].encode('utf-8').decode('utf-8')
                else:
                    questionFileObj['tip'] = None
                if ques['question_hint']:
                    questionFileObj['hint'] = ques['question_hint'].encode('utf-8').decode('utf-8')
                else:
                    questionFileObj['hint'] = None
                if ques['instance_identifier']:
                    questionFileObj['instanceIdentifier'] = ques['instance_identifier'].encode('utf-8').decode('utf-8')
                else:
                    questionFileObj['instanceIdentifier'] = None
                if ques['question_response_type'].strip().lower():
                    questionFileObj['responseType'] = ques['question_response_type'].strip().lower()
                if questionFileObj['responseType'] == "date":
                    questionFileObj['dateFormat'] = "DD-MM-YYYY"
                    if ques['date_auto_capture'] and ques['date_auto_capture'] == 1 or str(
                            ques['date_auto_capture']).lower() == "true":
                        questionFileObj['autoCapture'] = 'TRUE'
                    elif ques['date_auto_capture'] and ques['date_auto_capture'] == 0 or str(
                            ques['date_auto_capture']).lower() == "false":
                        questionFileObj['autoCapture'] = 'FALSE'
                    else:
                        questionFileObj['autoCapture'] = 'FALSE'

                else:
                    questionFileObj['dateFormat'] = ""
                    questionFileObj['autoCapture'] = None
                if ques['response_required']:
                    if ques['response_required'] == 1 or str(ques['response_required']).lower() == "true":
                        questionFileObj['validation'] = 'TRUE'
                    else:
                        questionFileObj['validation'] = 'FALSE'
                else:
                    questionFileObj['validation'] = 'FALSE'
                if ques['question_response_type'].strip().lower() == 'number':
                    questionFileObj['validationIsNumber'] = 'TRUE'
                    questionFileObj['validationRegex'] = 'isNumber'
                    if (ques['max_number_value'] and ques['max_number_value'].is_integer() == True):
                        questionFileObj['validationMax'] = int(ques['max_number_value'])
                    elif (ques['max_number_value'] and ques['max_number_value'].is_integer() == False):
                        questionFileObj['validationMax'] = ques['max_number_value']
                    else:
                        questionFileObj['validationMax'] = 10000
                    if (ques['min_number_value'] and ques['min_number_value'].is_integer() == True):
                        questionFileObj['validationMin'] = int(ques['min_number_value'])
                    elif (ques['min_number_value'] and ques['min_number_value'].is_integer() == False):
                        questionFileObj['validationMin'] = ques['min_number_value']
                    else:
                        questionFileObj['validationMin'] = 0
                elif ques['question_response_type'].strip().lower() == 'slider':
                    questionFileObj['validationIsNumber'] = None
                    questionFileObj['validationRegex'] = 'isNumber'
                    if (ques['max_number_value'] and ques['max_number_value'].is_integer() == True):
                        questionFileObj['validationMax'] = int(ques['max_number_value'])
                    elif (ques['max_number_value'] and ques['max_number_value'].is_integer() == False):
                        questionFileObj['validationMax'] = ques['max_number_value']
                    else:
                        questionFileObj['validationMax'] = 5
                    if (ques['min_number_value'] and ques['min_number_value'].is_integer() == True):
                        questionFileObj['validationMin'] = int(ques['min_number_value'])
                    elif (ques['min_number_value'] and ques['min_number_value'].is_integer() == False):
                        questionFileObj['validationMin'] = ques['min_number_value']
                    else:
                        questionFileObj['validationMin'] = 0
                else:
                    questionFileObj['validationIsNumber'] = None
                    questionFileObj['validationRegex'] = None
                    questionFileObj['validationMax'] = None
                    questionFileObj['validationMin'] = None
                if ques['file_upload'] == 1 or ques['file_upload'] == "TRUE":
                    questionFileObj['file'] = 'Snapshot'
                    questionFileObj['fileIsRequired'] = 'TRUE'
                    questionFileObj['fileUploadType'] = 'image/jpeg,docx,pdf,ppt'
                    questionFileObj['minFileCount'] = 0
                    questionFileObj['maxFileCount'] = 10
                else:
                    questionFileObj['file'] = 'NA'
                    questionFileObj['fileIsRequired'] = "FALSE"
                    questionFileObj['fileUploadType'] = None
                    questionFileObj['minFileCount'] = None
                    questionFileObj['maxFileCount'] = None
                questionFileObj['allowAudioRecording'] = False
                questionFileObj['caption'] = 'FALSE'
                questionFileObj['questionGroup'] = 'A1'
                questionFileObj['modeOfCollection'] = 'onfield'
                questionFileObj['accessibility'] = 'No'
                if ques['show_remarks'] == 1 or ques['show_remarks'] == "TRUE":
                    questionFileObj['showRemarks'] = 'TRUE'
                else:
                    questionFileObj['showRemarks'] = 'FALSE'
                questionFileObj['rubricLevel'] = None
                questionFileObj['isAGeneralQuestion'] = None
                if not pointBasedValue.lower() == "null":
                    if ques['question_response_type'].strip().lower() == 'radio' or ques[
                        'question_response_type'].strip() == 'multiselect':
                        questionFileObj['R1-score'] = ques['Score for R1']
                        questionFileObj['R2-score'] = ques['Score for R2']
                        questionFileObj['R3-score'] = ques['Score for R3']
                        questionFileObj['R4-score'] = ques['Score for R4']
                        questionFileObj['R5-score'] = ques['Score for R5']
                        questionFileObj['R6-score'] = ques['Score for R6']
                        questionFileObj['R7-score'] = ques['Score for R7']
                        questionFileObj['R8-score'] = ques['Score for R8']
                        questionFileObj['R9-score'] = ques['Score for R9']
                        questionFileObj['R10-score'] = ques['Score for R10']
                        questionFileObj['R11-score'] = ques['Score for R11']
                        questionFileObj['R12-score'] = ques['Score for R12']
                        questionFileObj['R13-score'] = ques['Score for R13']
                        questionFileObj['R14-score'] = ques['Score for R14']
                        questionFileObj['R15-score'] = ques['Score for R15']
                        questionFileObj['R16-score'] = ques['Score for R16']
                        questionFileObj['R17-score'] = ques['Score for R17']
                        questionFileObj['R18-score'] = ques['Score for R18']
                        questionFileObj['R19-score'] = ques['Score for R19']
                        questionFileObj['R20-score'] = ques['Score for R20']
                    if ques['question_response_type'].strip().lower() == 'slider' and ques[
                        'slider_value_with_score'].strip():
                        noOfSliderColumnQuestionVal = ques['slider_value_with_score'].strip().split(',')
                        possibleSliderColumnQuesVal = (int(ques['max_number_value']) + 1) - (int(ques['min_number_value']))
                        if len(noOfSliderColumnQuestionVal) == possibleSliderColumnQuesVal:
                            for sliVal in noOfSliderColumnQuestionVal:
                                sliValArr = []
                                sliValArr = sliVal.split(':')
                                questionFileObj['slider-value-' + str(sliValArr[0])] = sliValArr[0]
                                questionFileObj['slider-value-' + str(sliValArr[0]) + '-score'] = sliValArr[1]
                    if str(ques['question_weightage']):
                        questionFileObj['weightage'] = ques['question_weightage']
                    else:
                        questionFileObj['weightage'] = 0
                if ques['question_response_type'].strip().lower() == 'radio' or ques[
                    'question_response_type'].strip() == 'multiselect':
                    if type(ques['response(R1)']) != str:
                        if (ques['response(R1)'] and ques['response(R1)'].is_integer() == True):
                            questionFileObj['R1'] = int(ques['response(R1)'])
                        elif (ques['response(R1)'] and ques['response(R1)'].is_integer() == False):
                            questionFileObj['R1'] = ques['response(R1)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R1'] = ques['response(R1)']
                    if type(ques['response(R1)_hint']) != str:
                        if (ques['response(R1)_hint'] and ques['response(R1)_hint'].is_integer() == True):
                            questionFileObj['R1-hint'] = int(ques['response(R1)_hint'])
                        elif (ques['response(R1)_hint'] and ques['response(R1)_hint'].is_integer() == False):
                            questionFileObj['R1-hint'] = ques['response(R1)_hint'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R1-hint'] = ques['response(R1)_hint'].encode('utf-8').decode('utf-8')
                    if type(ques['response(R2)']) != str:
                        if (ques['response(R2)'] and ques['response(R2)'].is_integer() == True):
                            questionFileObj['R2'] = int(ques['response(R2)'])
                        elif (ques['response(R2)'] and ques['response(R2)'].is_integer() == False):
                            questionFileObj['R2'] = ques['response(R2)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R2'] = ques['response(R2)']
                    if type(ques['response(R2)_hint']) != str:
                        if (ques['response(R2)_hint'] and ques['response(R2)_hint'].is_integer() == True):
                            questionFileObj['R2-hint'] = int(ques['response(R2)_hint'])
                        elif (ques['response(R2)_hint'] and ques['response(R2)_hint'].is_integer() == False):
                            questionFileObj['R2-hint'] = ques['response(R2)_hint'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R2-hint'] = ques['response(R2)_hint']
                    if type(ques['response(R3)']) != str:
                        if (ques['response(R3)'] and ques['response(R3)'].is_integer() == True):
                            questionFileObj['R3'] = int(ques['response(R3)'])
                        elif (ques['response(R3)'] and ques['response(R3)'].is_integer() == False):
                            questionFileObj['R3'] = ques['response(R3)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R3'] = ques['response(R3)']
                    if type(ques['response(R3)_hint']) != str:
                        if (ques['response(R3)_hint'] and ques['response(R3)_hint'].is_integer() == True):
                            questionFileObj['R3-hint'] = int(ques['response(R3)_hint'])
                        elif (ques['response(R3)_hint'] and ques['response(R3)_hint'].is_integer() == False):
                            questionFileObj['R3-hint'] = ques['response(R3)_hint'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R3-hint'] = ques['response(R3)_hint']
                    if type(ques['response(R4)']) != str:
                        if (ques['response(R4)'] and ques['response(R4)'].is_integer() == True):
                            questionFileObj['R4'] = int(ques['response(R4)'])
                        elif (ques['response(R4)'] and ques['response(R4)'].is_integer() == False):
                            questionFileObj['R4'] = ques['response(R4)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R4'] = ques['response(R4)']
                    if type(ques['response(R4)_hint']) != str:
                        if (ques['response(R4)_hint'] and ques['response(R4)_hint'].is_integer() == True):
                            questionFileObj['R4-hint'] = int(ques['response(R4)_hint'])
                        elif (ques['response(R4)_hint'] and ques['response(R4)_hint'].is_integer() == False):
                            questionFileObj['R4-hint'] = ques['response(R4)_hint'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R4-hint'] = ques['response(R4)_hint']
                    if type(ques['response(R5)']) != str:
                        if (ques['response(R5)'] and ques['response(R5)'].is_integer() == True):
                            questionFileObj['R5'] = int(ques['response(R5)'])
                        elif (ques['response(R5)'] and ques['response(R5)'].is_integer() == False):
                            questionFileObj['R5'] = ques['response(R5)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R5'] = ques['response(R5)']
                    if type(ques['response(R5)_hint']) != str:
                        if (ques['response(R5)_hint'] and ques['response(R5)_hint'].is_integer() == True):
                            questionFileObj['R5-hint'] = int(ques['response(R5)_hint'])
                        elif (ques['response(R5)_hint'] and ques['response(R5)_hint'].is_integer() == False):
                            questionFileObj['R5-hint'] = ques['response(R5)_hint'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R5-hint'] = ques['response(R5)_hint']
                    if type(ques['response(R6)']) != str:
                        if (ques['response(R6)'] and ques['response(R6)'].is_integer() == True):
                            questionFileObj['R6'] = int(ques['response(R6)'])
                        elif (ques['response(R6)'] and ques['response(R6)'].is_integer() == False):
                            questionFileObj['R6'] = ques['response(R6)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R6'] = ques['response(R6)']
                    if type(ques['response(R6)_hint']) != str:
                        if (ques['response(R6)_hint'] and ques['response(R6)_hint'].is_integer() == True):
                            questionFileObj['R6-hint'] = int(ques['response(R6)_hint'])
                        elif (ques['response(R6)_hint'] and ques['response(R6)_hint'].is_integer() == False):
                            questionFileObj['R6-hint'] = ques['response(R6)_hint'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R6-hint'] = ques['response(R6)_hint']
                    if type(ques['response(R7)']) != str:
                        if (ques['response(R7)'] and ques['response(R7)'].is_integer() == True):
                            questionFileObj['R7'] = int(ques['response(R7)'])
                        elif (ques['response(R7)'] and ques['response(R7)'].is_integer() == False):
                            questionFileObj['R7'] = ques['response(R7)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R7'] = ques['response(R7)']
                    if type(ques['response(R7)_hint']) != str:
                        if (ques['response(R7)_hint'] and ques['response(R7)_hint'].is_integer() == True):
                            questionFileObj['R7-hint'] = int(ques['response(R7)_hint'])
                        elif (ques['response(R7)_hint'] and ques['response(R7)_hint'].is_integer() == False):
                            questionFileObj['R7-hint'] = ques['response(R7)_hint'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R7-hint'] = ques['response(R7)_hint']
                    if type(ques['response(R8)']) != str:
                        if (ques['response(R8)'] and ques['response(R8)'].is_integer() == True):
                            questionFileObj['R8'] = int(ques['response(R8)'])
                        elif (ques['response(R8)'] and ques['response(R8)'].is_integer() == False):
                            questionFileObj['R8'] = ques['response(R8)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R8'] = ques['response(R8)']
                    if type(ques['response(R8)_hint']) != str:
                        if (ques['response(R8)_hint'] and ques['response(R8)_hint'].is_integer() == True):
                            questionFileObj['R8-hint'] = int(ques['response(R8)_hint'])
                        elif (ques['response(R8)_hint'] and ques['response(R8)_hint'].is_integer() == False):
                            questionFileObj['R8-hint'] = ques['response(R8)_hint'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R8-hint'] = ques['response(R8)_hint']
                    if type(ques['response(R9)']) != str:
                        if (ques['response(R9)'] and ques['response(R9)'].is_integer() == True):
                            questionFileObj['R9'] = int(ques['response(R9)'])
                        elif (ques['response(R9)'] and ques['response(R9)'].is_integer() == False):
                            questionFileObj['R9'] = ques['response(R9)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R9'] = ques['response(R9)']
                    if type(ques['response(R9)_hint']) != str:
                        if (ques['response(R9)_hint'] and ques['response(R9)_hint'].is_integer() == True):
                            questionFileObj['R9-hint'] = int(ques['response(R9)_hint'])
                        elif (ques['response(R9)_hint'] and ques['response(R9)_hint'].is_integer() == False):
                            questionFileObj['R9-hint'] = ques['response(R9)_hint'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R9-hint'] = ques['response(R9)_hint']
                    if type(ques['response(R10)']) != str:
                        if (ques['response(R10)'] and ques['response(R10)'].is_integer() == True):
                            questionFileObj['R10'] = int(ques['response(R10)'])
                        elif (ques['response(R10)'] and ques['response(R10)'].is_integer() == False):
                            questionFileObj['R10'] = ques['response(R10)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R10'] = ques['response(R10)']
                    if type(ques['response(R10)_hint']) != str:
                        if (ques['response(R10)_hint'] and ques['response(R10)_hint'].is_integer() == True):
                            questionFileObj['R10-hint'] = int(ques['response(R10)_hint'])
                        elif (ques['response(R10)_hint'] and ques['response(R10)_hint'].is_integer() == False):
                            questionFileObj['R10-hint'] = ques['response(R10)_hint'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R10-hint'] = ques['response(R10)_hint']
                    if type(ques['response(R11)']) != str:
                        if (ques['response(R11)'] and ques['response(R11)'].is_integer() == True):
                            questionFileObj['R11'] = int(ques['response(R11)'])
                        elif (ques['response(R11)'] and ques['response(R11)'].is_integer() == False):
                            questionFileObj['R11'] = ques['response(R11)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R11'] = ques['response(R11)']
                    if type(ques['response(R11)_hint']) != str:
                        if (ques['response(R11)_hint'] and ques['response(R11)_hint'].is_integer() == True):
                            questionFileObj['R11-hint'] = int(ques['response(R11)_hint'])
                        elif (ques['response(R11)_hint'] and ques['response(R11)_hint'].is_integer() == False):
                            questionFileObj['R11-hint'] = ques['response(R11)_hint'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R11-hint'] = ques['response(R11)_hint']
                    if type(ques['response(R12)']) != str:
                        if (ques['response(R12)'] and ques['response(R12)'].is_integer() == True):
                            questionFileObj['R12'] = int(ques['response(R12)'])
                        elif (ques['response(R12)'] and ques['response(R12)'].is_integer() == False):
                            questionFileObj['R12'] = ques['response(R12)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R12'] = ques['response(R12)']
                    if type(ques['response(R12)_hint']) != str:
                        if (ques['response(R12)_hint'] and ques['response(R12)_hint'].is_integer() == True):
                            questionFileObj['R12-hint'] = int(ques['response(R12)_hint'])
                        elif (ques['response(R12)_hint'] and ques['response(R12)_hint'].is_integer() == False):
                            questionFileObj['R12-hint'] = ques['response(R12)_hint'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R12-hint'] = ques['response(R12)_hint']
                    if type(ques['response(R13)']) != str:
                        if (ques['response(R13)'] and ques['response(R13)'].is_integer() == True):
                            questionFileObj['R13'] = int(ques['response(R13)'])
                        elif (ques['response(R13)'] and ques['response(R13)'].is_integer() == False):
                            questionFileObj['R13'] = ques['response(R13)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R13'] = ques['response(R13)']
                    if type(ques['response(R13)_hint']) != str:
                        if (ques['response(R13)_hint'] and ques['response(R13)_hint'].is_integer() == True):
                            questionFileObj['R13-hint'] = int(ques['response(R13)_hint'])
                        elif (ques['response(R13)_hint'] and ques['response(R13)_hint'].is_integer() == False):
                            questionFileObj['R13-hint'] = ques['response(R13)_hint'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R13-hint'] = ques['response(R13)_hint']
                    if type(ques['response(R14)']) != str:
                        if (ques['response(R14)'] and ques['response(R14)'].is_integer() == True):
                            questionFileObj['R14'] = int(ques['response(R14)'])
                        elif (ques['response(R14)'] and ques['response(R14)'].is_integer() == False):
                            questionFileObj['R14'] = ques['response(R14)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R14'] = ques['response(R14)']
                    if type(ques['response(R14)_hint']) != str:
                        if (ques['response(R14)_hint'] and ques['response(R14)_hint'].is_integer() == True):
                            questionFileObj['R14-hint'] = int(ques['response(R14)_hint'])
                        elif (ques['response(R14)_hint'] and ques['response(R14)_hint'].is_integer() == False):
                            questionFileObj['R14-hint'] = ques['response(R14)_hint'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R14-hint'] = ques['response(R14)_hint']
                    if type(ques['response(R15)']) != str:
                        if (ques['response(R15)'] and ques['response(R15)'].is_integer() == True):
                            questionFileObj['R15'] = int(ques['response(R15)'])
                        elif (ques['response(R15)'] and ques['response(R15)'].is_integer() == False):
                            questionFileObj['R15'] = ques['response(R15)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R15'] = ques['response(R15)']
                    if type(ques['response(R15)_hint']) != str:
                        if (ques['response(R15)_hint'] and ques['response(R15)_hint'].is_integer() == True):
                            questionFileObj['R15-hint'] = int(ques['response(R15)_hint'])
                        elif (ques['response(R15)_hint'] and ques['response(R15)_hint'].is_integer() == False):
                            questionFileObj['R15-hint'] = ques['response(R15)_hint'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R15-hint'] = ques['response(R15)_hint']
                    if type(ques['response(R16)']) != str:
                        if (ques['response(R16)'] and ques['response(R16)'].is_integer() == True):
                            questionFileObj['R16'] = int(ques['response(R16)'])
                        elif (ques['response(R16)'] and ques['response(R16)'].is_integer() == False):
                            questionFileObj['R16'] = ques['response(R16)'].encode('utf-8').decode('utf-8')
                    else:
                        questionFileObj['R16'] = ques['response(R16)']
                    if type(ques['response(R16)_hint']) != str:
                        if (ques['response(R16)_hint'] and ques['response(R16)_hint'].is_integer() == True):
                            questionFileObj['R16-hint'] = int(ques['response(R16)_hint'])
                        elif (ques['response(R16)_hint'] and ques['response(R16)_hint'].is_integer() == False):
                            questionFileObj['R16-hint'] = ques['response(R16)_hint']
                    else:
                        questionFileObj['R16-hint'] = ques['response(R16)_hint']
                    if type(ques['response(R17)']) != str:
                        if (ques['response(R17)'] and ques['response(R17)'].is_integer() == True):
                            questionFileObj['R17'] = int(ques['response(R17)'])
                        elif (ques['response(R17)'] and ques['response(R17)'].is_integer() == False):
                            questionFileObj['R17'] = ques['response(R17)']
                    else:
                        questionFileObj['R17'] = ques['response(R17)']
                    if type(ques['response(R17)_hint']) != str:
                        if (ques['response(R17)_hint'] and ques['response(R17)_hint'].is_integer() == True):
                            questionFileObj['R17-hint'] = int(ques['response(R17)_hint'])
                        elif (ques['response(R17)_hint'] and ques['response(R17)_hint'].is_integer() == False):
                            questionFileObj['R17-hint'] = ques['response(R17)_hint']
                    else:
                        questionFileObj['R17-hint'] = ques['response(R17)_hint']
                    if type(ques['response(R18)']) != str:
                        if (ques['response(R18)'] and ques['response(R18)'].is_integer() == True):
                            questionFileObj['R18'] = int(ques['response(R18)'])
                        elif (ques['response(R18)'] and ques['response(R18)'].is_integer() == False):
                            questionFileObj['R18'] = ques['response(R18)']
                    else:
                        questionFileObj['R18'] = ques['response(R18)']
                    if type(ques['response(R18)_hint']) != str:
                        if (ques['response(R18)_hint'] and ques['response(R18)_hint'].is_integer() == True):
                            questionFileObj['R18-hint'] = int(ques['response(R18)_hint'])
                        elif (ques['response(R18)_hint'] and ques['response(R18)_hint'].is_integer() == False):
                            questionFileObj['R18-hint'] = ques['response(R18)_hint']
                    else:
                        questionFileObj['R18-hint'] = ques['response(R18)_hint']
                    if type(ques['response(R19)']) != str:
                        if (ques['response(R19)'] and ques['response(R19)'].is_integer() == True):
                            questionFileObj['R19'] = int(ques['response(R19)'])
                        elif (ques['response(R19)'] and ques['response(R19)'].is_integer() == False):
                            questionFileObj['R19'] = ques['response(R19)']
                    else:
                        questionFileObj['R19'] = ques['response(R19)']
                    if type(ques['response(R19)_hint']) != str:
                        if (ques['response(R19)_hint'] and ques['response(R19)_hint'].is_integer() == True):
                            questionFileObj['R19-hint'] = int(ques['response(R19)_hint'])
                        elif (ques['response(R19)_hint'] and ques['response(R19)_hint'].is_integer() == False):
                            questionFileObj['R19-hint'] = ques['response(R19)_hint']
                    else:
                        questionFileObj['R19-hint'] = ques['response(R19)_hint']
                    if type(ques['response(R20)']) != str:
                        if (ques['response(R20)'] and ques['response(R20)'].is_integer() == True):
                            questionFileObj['R20'] = int(ques['response(R20)'])
                        elif (ques['response(R20)'] and ques['response(R20)'].is_integer() == False):
                            questionFileObj['R20'] = ques['response(R20)']
                    else:
                        questionFileObj['R20'] = ques['response(R20)']
                    if type(ques['response(R20)_hint']) != str:
                        if (ques['response(R20)_hint'] and ques['response(R20)_hint'].is_integer() == True):
                            questionFileObj['R20-hint'] = int(ques['response(R20)_hint'])
                        elif (ques['response(R20)_hint'] and ques['response(R20)_hint'].is_integer() == False):
                            questionFileObj['R20-hint'] = ques['response(R20)_hint']
                    else:
                        questionFileObj['R20-hint'] = ques['response(R20)_hint']
                else:
                    questionFileObj['R1'] = None
                    questionFileObj['R1-hint'] = None
                    questionFileObj['R2'] = None
                    questionFileObj['R2-hint'] = None
                    questionFileObj['R3'] = None
                    questionFileObj['R3-hint'] = None
                    questionFileObj['R4'] = None
                    questionFileObj['R4-hint'] = None
                    questionFileObj['R5'] = None
                    questionFileObj['R5-hint'] = None
                    questionFileObj['R6'] = None
                    questionFileObj['R6-hint'] = None
                    questionFileObj['R7'] = None
                    questionFileObj['R7-hint'] = None
                    questionFileObj['R8'] = None
                    questionFileObj['R8-hint'] = None
                    questionFileObj['R9'] = None
                    questionFileObj['R9-hint'] = None
                    questionFileObj['R10'] = None
                    questionFileObj['R10-hint'] = None
                    questionFileObj['R11'] = None
                    questionFileObj['R11-hint'] = None
                    questionFileObj['R12'] = None
                    questionFileObj['R12-hint'] = None
                    questionFileObj['R13'] = None
                    questionFileObj['R13-hint'] = None
                    questionFileObj['R14'] = None
                    questionFileObj['R14-hint'] = None
                    questionFileObj['R15'] = None
                    questionFileObj['R15-hint'] = None
                    questionFileObj['R16'] = None
                    questionFileObj['R16-hint'] = None
                    questionFileObj['R17'] = None
                    questionFileObj['R17-hint'] = None
                    questionFileObj['R18'] = None
                    questionFileObj['R18-hint'] = None
                    questionFileObj['R19'] = None
                    questionFileObj['R19-hint'] = None
                    questionFileObj['R20'] = None
                    questionFileObj['R20-hint'] = None
                    questionFileObj['_arrayFields'] = None
                if ques['section_header']:
                    questionFileObj['sectionHeader'] = ques['section_header'].encode('utf-8').decode('utf-8')
                else:
                    questionFileObj['sectionHeader'] = None
                questionFileObj['page'] = ques['page']
                if type(ques['question_number']) != str:
                    if ques['question_number'] and ques['question_number'].is_integer() == True:
                        questionFileObj['questionNumber'] = int(ques['question_number'])
                    elif ques['question_number']:
                        questionFileObj['questionNumber'] = ques['question_number']
                else:
                    questionFileObj['questionNumber'] = ques['question_number']
                questionFileObj['prefillFromEntityProfile'] = None
                questionFileObj['isEditable'] = 'TRUE'
                questionFileObj['entityFieldName'] = None
                questionFileObj['_arrayFields'] = 'parentQuestionValue'
                writerQuestionUpload.writerow(questionFileObj)
        bodySolutionUpdate = {"questionSequenceByEcm": questionSeqByEcmDict}
        ElevateObservation.solutionUpdate(solutionName_for_folder_path, accessToken, solutionId, bodySolutionUpdate)

        urlQuestionsUploadApi = config.get(environment, 'INTERNAL_KONG_IP') + config.get(environment, 'questionUploadApiUrl')
        headerQuestionUploadApi = {'Authorization': config.get(environment, 'Authorization'),
                                'X-auth-token': accessToken,
                                'X-Channel-id': config.get(environment, 'X-Channel-id')}
        filesQuestion = {
            'questions': open(solutionName_for_folder_path + '/questionUpload/uploadSheet.csv', 'rb')
        }
        responseQuestionUploadApi = requests.post(url=urlQuestionsUploadApi, headers=headerQuestionUploadApi,
                                                files=filesQuestion)
        messageArr = ["Question Upload sheet prepared.",
                    "File loc : " + solutionName_for_folder_path + '/questionUpload/uploadSheet.csv',
                    "Question upload API called.", "Status code : " + str(responseQuestionUploadApi.status_code)]
        ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
        if responseQuestionUploadApi.status_code == 200:
            print('QuestionUploadApi Success')
            with open(solutionName_for_folder_path + '/questionUpload/uploadInternalIdsSheet.csv','w+',
                    encoding='utf-8') as questionRes:
                questionRes.write(responseQuestionUploadApi.text)
        else:
            messageArr = ["Question Upload Failed.", "Response : " + str(responseQuestionUploadApi.text)]
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            ElevateObservation.terminatingMessage("Question Upload failed.")

    def fetchSolutionCriteria(solutionName_for_folder_path, observationId, accessToken):
        url = config.get(environment, 'INTERNAL_KONG_IP') + config.get(environment, 'ferchSolutionCriteria') + observationId

        headers = {
            'Authorization': config.get(environment, 'Authorization'),
            'X-auth-token': accessToken,
            'internal-access-token': config.get(environment, 'internal-access-token')
        }

        response = requests.request("POST", url, headers=headers)
        messageArr = ["Criteria solution fetch API called.", "Status Code  : " + str(response.status_code), "URL : " + url]
        ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)

        os.mkdir(solutionName_for_folder_path + "/solutionCriteriaFetch/")
        if response.status_code == 200:
            print("Solution criteria fetched.")
            with open(solutionName_for_folder_path + "/solutionCriteriaFetch/solutionCriteriaDetails.csv",
                    'w+',encoding='utf-8') as solutionCriteriaFetch:
                solutionCriteriaFetch.write(response.text)
        else:
            messageArr = ["Criteria solution fetch API failed.", "Response  : " + str(response.text)]
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            ElevateObservation.terminatingMessage("Solution criteria fetch failed. Status Code : " + str(response.status_code))

    def uploadCriteriaRubrics(solutionName_for_folder_path, wbObservation, millisAddObs, accessToken, frameworkExternalId,
                          withRubricsFlag):
        if withRubricsFlag:
            criteriaRubricSheet = wbObservation.sheet_by_name('Criteria_Rubric-Scoring')
            dictSolCritLookUp = dict()
            filePath = os.path.join(solutionName_for_folder_path + "/solutionCriteriaFetch/", "solutionCriteriaDetails.csv")
            with open(filePath, 'r',encoding='utf-8') as criteriaInternalFile:
                criteriaInternalReader = csv.DictReader(criteriaInternalFile)
                for crit in criteriaInternalReader:
                    dictSolCritLookUp[crit['criteriaID']] = [crit['criteriaInternalId'], crit['criteriaName']]
        else:
            criteriaRubricSheet = wbObservation.sheet_by_name('criteria')
            dictSolCritLookUp = dict()
            filePath = os.path.join(solutionName_for_folder_path + "/solutionCriteriaFetch/", "solutionCriteriaDetails.csv")
            with open(filePath, 'r',encoding='utf-8') as criteriaInternalFile:
                criteriaInternalReader = csv.DictReader(criteriaInternalFile)
                for crit in criteriaInternalReader:
                    dictSolCritLookUp[crit['criteriaID']] = [crit['criteriaInternalId'], crit['criteriaName']]

        keys = [criteriaRubricSheet.cell(1, col_index).value for col_index in range(criteriaRubricSheet.ncols)]
        criteriaRubricUploadFieldnames = ["externalId", "name", "criteriaId", "weightage", "expressionVariables"]

        if withRubricsFlag:
            for cl in criteriaLevels:
                criteriaRubricUploadFieldnames.append("L" + str(cl))
        else:
            criteriaRubricUploadFieldnames.append("L1")
        criteriaRubricUpload = dict()
        criteriaRubricsFilePath = solutionName_for_folder_path + '/criteriaRubrics/'
        file_exists_ques = os.path.isfile(solutionName_for_folder_path + '/criteriaRubrics/uploadSheet.csv')
        if not os.path.exists(criteriaRubricsFilePath):
            os.mkdir(criteriaRubricsFilePath)
        if withRubricsFlag:
            for row_index in range(2, criteriaRubricSheet.nrows):
                file_exists_ques = os.path.isfile(solutionName_for_folder_path + '/criteriaRubrics/uploadSheet.csv')
                with open(solutionName_for_folder_path + '/criteriaRubrics/uploadSheet.csv', 'a',
                        encoding='utf-8') as questionUploadFile:
                    writerQuestionUpload = csv.DictWriter(questionUploadFile, fieldnames=criteriaRubricUploadFieldnames,
                                                        lineterminator='\n')
                    if not file_exists_ques:
                        writerQuestionUpload.writeheader()
                    dictCriteriaRubric = {keys[col_index]: criteriaRubricSheet.cell(row_index, col_index).value for
                                        col_index in range(criteriaRubricSheet.ncols)}
                    criteriaRubricUpload['externalId'] = dictCriteriaRubric['criteriaId'] + "_" + str(millisAddObs)
                    print(criteriaRubricUpload['externalId'])
                    criteriaRubricUpload['name'] = dictSolCritLookUp[criteriaRubricUpload['externalId']][1]
                    criteriaRubricUpload['criteriaId'] = dictSolCritLookUp[criteriaRubricUpload['externalId']][0]
                    if dictCriteriaRubric['weightage']:
                        criteriaRubricUpload['weightage'] = dictCriteriaRubric['weightage']
                    else:
                        criteriaRubricUpload['weightage'] = 0
                    criteriaRubricUpload['expressionVariables'] = "SCORE=" + criteriaRubricUpload[
                        'criteriaId'] + ".scoreOfAllQuestionInCriteria()"
                    for cl in criteriaLevels:
                        criteriaRubricUpload['L' + str(cl)] = dictCriteriaRubric['L' + str(cl) + " SCORE"]
                    writerQuestionUpload.writerow(criteriaRubricUpload)
        else:
            for criteriaIds, criteriaDetails in dictSolCritLookUp.items():
                file_exists_ques = os.path.isfile(solutionName_for_folder_path + '/criteriaRubrics/uploadSheet.csv')
                with open(solutionName_for_folder_path + '/criteriaRubrics/uploadSheet.csv', 'a',
                        encoding='utf-8') as questionUploadFile:
                    writerQuestionUpload = csv.DictWriter(questionUploadFile, fieldnames=criteriaRubricUploadFieldnames,
                                                        lineterminator='\n')
                    if not file_exists_ques:
                        writerQuestionUpload.writeheader()
                    criteriaRubricUpload['externalId'] = criteriaIds
                    criteriaRubricUpload['name'] = criteriaDetails[1]
                    criteriaRubricUpload['weightage'] = 1
                    criteriaRubricUpload['criteriaId'] = criteriaDetails[0]
                    criteriaRubricUpload['expressionVariables'] = 'SCORE=' + str(
                        criteriaDetails[0]) + '.scoreOfAllQuestionInCriteria()'
                    criteriaRubricUpload['L1'] = '0<=SCORE<=100000'
                    writerQuestionUpload.writerow(criteriaRubricUpload)

        urlCriteriaRubricUploadApi = config.get(environment, 'INTERNAL_KONG_IP') + config.get(environment,'criteriaRubricUploadApiUrl') + frameworkExternalId + "-OBSERVATION-TEMPLATE"
        headerCriteriaRubricUploadApi = {
            'Authorization': config.get(environment, 'Authorization'),
            'X-auth-token': accessToken,
            'X-Channel-id': config.get(environment, 'X-Channel-id')
        }
        filesCriteriaRubric = {
            'criteria': open(solutionName_for_folder_path + '/criteriaRubrics/uploadSheet.csv', 'rb')
        }
        responseCriteriaRubricUploadApi = requests.post(url=urlCriteriaRubricUploadApi,
                                                        headers=headerCriteriaRubricUploadApi, files=filesCriteriaRubric)
        messageArr = ["Criteria Rubric upload sheet prepared.",
                    "File Loc : " + solutionName_for_folder_path + '/criteriaRubrics/uploadSheet.csv',
                    "Status Code : " + str(responseCriteriaRubricUploadApi.status_code)]
        ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
        if responseCriteriaRubricUploadApi.status_code == 200:
            with open(solutionName_for_folder_path + '/criteriaRubrics/uploadInternalIdsSheet.csv',
                    'w+',encoding='utf-8') as criteriaRubricRes:
                criteriaRubricRes.write(responseCriteriaRubricUploadApi.text)
        else:
            messageArr = ["Criteria Rubric upload Failed.", "Response : " + str(responseCriteriaRubricUploadApi.text)]
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            ElevateObservation.terminatingMessage("Criteria Rubric upload Failed.")

    def uploadThemeRubrics(solutionName_for_folder_path, wbObservation, accessToken, frameworkExternalId, withRubricsFlag):
        themeRubricUploadFieldnames = ["externalId", "name", "weightage"]
        themeRubricsFilePath = os.path.join(solutionName_for_folder_path, "themeRubrics/")
        if not os.path.exists(themeRubricsFilePath):
            os.mkdir(themeRubricsFilePath)
        themeRubricUpload = dict()
        if withRubricsFlag:
            themeRubricSheet = wbObservation.sheet_by_name('Domain(theme)_rubric_scoring')
            keys = [themeRubricSheet.cell(1, col_index).value for col_index in range(themeRubricSheet.ncols)]
            themeRubricUploadFieldnames = ["externalId", "name", "weightage"]
            if withRubricsFlag:
                for cl in criteriaLevels:
                    themeRubricUploadFieldnames.append("L" + str(cl))
            else:
                themeRubricUploadFieldnames.append("L1")

            for row_index in range(2, themeRubricSheet.nrows):
                file_exists_ques = os.path.isfile(solutionName_for_folder_path + '/themeRubrics/uploadSheet.csv')
                with open(solutionName_for_folder_path + '/themeRubrics/uploadSheet.csv', 'a',
                        encoding='utf-8') as themeRubricsUploadFile:
                    writerThemeRubricsUpload = csv.DictWriter(themeRubricsUploadFile,
                                                            fieldnames=themeRubricUploadFieldnames, lineterminator='\n')
                    if not file_exists_ques:
                        writerThemeRubricsUpload.writeheader()

                    dictThemeRubric = {keys[col_index]: themeRubricSheet.cell(row_index, col_index).value for col_index in
                                    range(themeRubricSheet.ncols)}
                    themeRubricUpload['externalId'] = dictThemeRubric['domain_Id']
                    themeRubricUpload['name'] = dictThemeRubric['domain_name'].encode('utf-8').decode('utf-8')
                    if dictThemeRubric['weightage']:
                        themeRubricUpload['weightage'] = dictThemeRubric['weightage']
                    else:
                        themeRubricUpload['weightage'] = 0
                    if withRubricsFlag:
                        for cl in criteriaLevels:
                            themeRubricUpload['L' + str(cl)] = dictThemeRubric['L' + str(cl)]
                    else:
                        themeRubricUpload['L1'] = '0<=SCORE<=100000'
                    writerThemeRubricsUpload.writerow(themeRubricUpload)
        else:
            themeRubricUploadFieldnames.append("L1")
            file_exists_ques = os.path.isfile(solutionName_for_folder_path + '/themeRubrics/uploadSheet.csv')
            with open(solutionName_for_folder_path + '/themeRubrics/uploadSheet.csv', 'a',
                    encoding='utf-8') as themeRubricsUploadFile:
                writerThemeRubricsUpload = csv.DictWriter(themeRubricsUploadFile, fieldnames=themeRubricUploadFieldnames,
                                                        lineterminator='\n')
                if not file_exists_ques:
                    writerThemeRubricsUpload.writeheader()
                themeRubricUpload['externalId'] = "OB"
                themeRubricUpload['name'] = "Observation Theme"
                themeRubricUpload['weightage'] = 1
                themeRubricUpload['L1'] = '0<=SCORE<=100000'
                writerThemeRubricsUpload.writerow(themeRubricUpload)
        urlThemeRubricUploadApi = config.get(environment, 'INTERNAL_KONG_IP') + config.get(environment,'themeRubricUploadApiUrl') + frameworkExternalId + "-OBSERVATION-TEMPLATE"
        headerThemeRubricUploadApi = {
            'Authorization': config.get(environment, 'Authorization'),
            'X-auth-token': accessToken,
            'X-Channel-id': config.get(environment, 'X-Channel-id')
        }
        filesThemeRubric = {
            'themes': open(solutionName_for_folder_path + '/themeRubrics/uploadSheet.csv', 'rb')
        }
        responseThemeRubricUploadApi = requests.post(url=urlThemeRubricUploadApi, headers=headerThemeRubricUploadApi,
                                                    files=filesThemeRubric)
        if responseThemeRubricUploadApi.status_code == 200:
            print('ThemeRubricUploadApi Success')
            with open(solutionName_for_folder_path + '/themeRubrics/uploadInternalIdsSheet.csv', 'w+',encoding='utf-8') as themeRubricRes:
                themeRubricRes.write(responseThemeRubricUploadApi.text)
        else:
            messageArr = ['theme rubric upload api failed in ' + environment,
                        ' status_code response from api is ' + str(responseThemeRubricUploadApi.status_code),
                        "Response : " + str(responseThemeRubricUploadApi.text)]
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            ElevateObservation.terminatingMessage(
                'theme rubric upload api failed in ' + environment + ' status_code response from api is ' + str(
                    responseThemeRubricUploadApi.status_code))
    
    def fetchSolutionDetailsFromProgramSheet(solutionName_for_folder_path, programFile, solutionId, accessToken):
        global solutionRolesArray, solutionStartDate, solutionEndDate
        urlFetchSolutionApi = config.get(environment, 'INTERNAL_KONG_IP') + config.get(environment, 'fetchSolutionDoc') + solutionId
        
        headerFetchSolutionApi = {
            'Content-Type': 'application/json',
            'Authorization': config.get(environment, 'Authorization'),
            'X-auth-token': accessToken,
            'X-Channel-id': config.get(environment, 'X-Channel-id'),
            'internal-access-token': config.get(environment, 'internal-access-token')
        }
        payloadFetchSolutionApi = {}

        responseFetchSolutionApiUrl = requests.post(url=urlFetchSolutionApi, headers=headerFetchSolutionApi,
                                                data=payloadFetchSolutionApi)
        responseFetchSolutionJson = responseFetchSolutionApiUrl.json()
        messageArr = ["Solution Fetch Link.",
                    "solution name : " + responseFetchSolutionJson["result"]["name"],
                    "solution ExternalId : " + responseFetchSolutionJson["result"]["externalId"]]
        messageArr.append("Upload status code : " + str(responseFetchSolutionApiUrl.status_code))
        ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)

        if responseFetchSolutionApiUrl.status_code == 200:
            print('Fetch solution Api Success')
            
            solutionName = responseFetchSolutionJson["result"]["name"]
            xfile = openpyxl.load_workbook(programFile)
            sheet_name = 'Resource Details'.strip()
            resourceDetailsSheet = xfile[sheet_name]
            rowCountRD = resourceDetailsSheet.max_row
            columnCountRD = resourceDetailsSheet.max_column
            for row in range(3, rowCountRD + 1):
                cell_value = resourceDetailsSheet["A" + str(row)].value
                if cell_value is not None and str(cell_value).strip() == str(solutionName).strip():
                    solutionMainRole = str(resourceDetailsSheet["E" + str(row)].value).strip()
                    solutionRolesArray = str(resourceDetailsSheet["F" + str(row)].value).split(",") if str(resourceDetailsSheet["E" + str(row)].value).split(",") else []
                    if "teacher" in solutionMainRole.strip().lower():
                        solutionRolesArray.append("TEACHER")
                    solutionStartDate = resourceDetailsSheet["G" + str(row)].value
                    solutionEndDate = resourceDetailsSheet["H" + str(row)].value
        return [solutionRolesArray, solutionStartDate, solutionEndDate]

    def fetchSolutionDetailsFromResourceSheet(solutionName_for_folder_path, programFile, solutionId, accessToken,typeofSolution):
        global solutionRolesArray, solutionStartDate, solutionEndDate
        urlFetchSolutionApi = config.get(environment, 'INTERNAL_KONG_IP') + config.get(environment, 'fetchSolutionDoc') + solutionId
        
        headerFetchSolutionApi = {
            'Content-Type': 'application/json',
            'Authorization': config.get(environment, 'Authorization'),
            'X-auth-token': accessToken,
            'X-Channel-id': config.get(environment, 'X-Channel-id'),
            'internal-access-token': config.get(environment, 'internal-access-token')
        }
        payloadFetchSolutionApi = {}

        responseFetchSolutionApiUrl = requests.post(url=urlFetchSolutionApi, headers=headerFetchSolutionApi,
                                                data=payloadFetchSolutionApi)
        responseFetchSolutionJson = responseFetchSolutionApiUrl.json()
        messageArr = ["Solution Fetch Link.",
                    "solution name : " + responseFetchSolutionJson["result"]["name"],
                    "solution ExternalId : " + responseFetchSolutionJson["result"]["externalId"]]
        messageArr.append("Upload status code : " + str(responseFetchSolutionApiUrl.status_code))
        ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)

        if responseFetchSolutionApiUrl.status_code == 200:
            print('Fetch solution Api Success')
            
            solutionName = responseFetchSolutionJson["result"]["name"]
            xfile = openpyxl.load_workbook(programFile)
            sheet_name = 'details'.strip()
            resourceDetailsSheet = xfile[sheet_name]
            rowCountRD = resourceDetailsSheet.max_row
            columnCountRD = resourceDetailsSheet.max_column
            for row in range(3, rowCountRD + 1):
                cell_value = resourceDetailsSheet["A" + str(row)].value
                if cell_value is not None and str(cell_value).strip() == str(solutionName).strip():
                    # solutionMainRole = str(resourceDetailsSheet["E" + str(row)].value).strip()
                    # solutionRolesArray = str(resourceDetailsSheet["F" + str(row)].value).split(",") if str(resourceDetailsSheet["E" + str(row)].value).split(",") else []
                    # if "teacher" in solutionMainRole.strip().lower():
                    #     solutionRolesArray.append("TEACHER")
                    if typeofSolution ==1 :
                        solutionStartDate = resourceDetailsSheet["J" + str(row)].value
                        solutionEndDate = resourceDetailsSheet["K" + str(row)].value
                    else :
                        solutionStartDate = resourceDetailsSheet["H" + str(row)].value
                        solutionEndDate = resourceDetailsSheet["I" + str(row)].value
        return [solutionRolesArray, solutionStartDate, solutionEndDate]


    def createChild(solutionName_for_folder_path, observationExternalId, accessToken):
        childObservationExternalId = str(observationExternalId + "_CHILD")
        urlSol_prog_mapping = config.get(environment, 'INTERNAL_KONG_IP') + config.get(environment,'solutionToprogramMAppingApiUrl') + "?solutionId=" + observationExternalId + "&entityType=" + entityType
        
        payloadSol_prog_mapping = {
            "externalId": childObservationExternalId,
            "name": solutionName.lstrip().rstrip(),
            "description": solutionDescription.lstrip().rstrip(),
            "programExternalId": programExternalId
        }
        headersSol_prog_mapping = {'Authorization': config.get(environment, 'Authorization'),
                                'X-auth-token': accessToken,
                                'Content-Type': config.get(environment, 'Content-Type')}
        responseSol_prog_mapping = requests.request("POST", urlSol_prog_mapping, headers=headersSol_prog_mapping,
                                                    data=json.dumps(payloadSol_prog_mapping))
        messageArr = ["Create child API called.", "URL : " + urlSol_prog_mapping,
                    "Status code : " + str(responseSol_prog_mapping.status_code),
                    "Response : " + responseSol_prog_mapping.text, "body : " + str(payloadSol_prog_mapping)]
        if responseSol_prog_mapping.status_code == 200:
            if programName :
                print("Solution mapped to program : " + programName)
            print("Child solution : " + childObservationExternalId)

            responseSol_prog_mapping = responseSol_prog_mapping.json()
            child_id = responseSol_prog_mapping['result']['_id']
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            print("child solutionId: " + child_id)
            return [child_id, childObservationExternalId]
        else:
            print("Unable to create child solution")

            messageArr.append("Unable to create child solution")
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            return False
    
    def prepareProgramSuccessSheet(MainFilePath, solutionName_for_folder_path, programFile, solutionExternalId, solutionId,accessToken):
        urlFetchSolutionApi = config.get(environment, 'INTERNAL_KONG_IP') + config.get(environment, 'fetchSolutionDoc') + solutionId
        headerFetchSolutionApi = {
            'Authorization': config.get(environment, 'Authorization'),
            'X-auth-token': accessToken,
            'X-Channel-id': config.get(environment, 'X-Channel-id'),
            'internal-access-token': config.get(environment, 'internal-access-token')
        }
        payloadFetchSolutionApi = {}

        responseFetchSolutionApi = requests.post(url=urlFetchSolutionApi, headers=headerFetchSolutionApi,
                                                data=payloadFetchSolutionApi)
        responseFetchSolutionJson = responseFetchSolutionApi.json()
        messageArr = ["Solution Fetch Link.",
                    "solution name : " + responseFetchSolutionJson["result"]["name"],
                    "solution ExternalId : " + responseFetchSolutionJson["result"]["externalId"]]
        messageArr.append("Upload status code : " + str(responseFetchSolutionApi.status_code))
        ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)

        if responseFetchSolutionApi.status_code == 200:
            print('Fetch solution Api Success')
            solutionName = responseFetchSolutionJson["result"]["name"]
        urlFetchSolutionLinkApi = config.get(environment, 'internal_kong_ip') + config.get(environment, 'fetchLink') + solutionId
        headerFetchSolutionLinkApi = {
            'Authorization': config.get(environment, 'Authorization'),
            'X-auth-token': accessToken,
            'X-Channel-id': config.get(environment, 'X-Channel-id'),
            'internal-access-token': config.get(environment, 'internal-access-token')
        }
        payloadFetchSolutionLinkApi = {}

        responseFetchSolutionLinkApi = requests.get(url=urlFetchSolutionLinkApi, headers=headerFetchSolutionLinkApi,
                                                    data=payloadFetchSolutionLinkApi)

        messageArr = ["Solution Fetch Link.","solution id : " + solutionId,"solution ExternalId : " + solutionExternalId]
        messageArr.append("Upload status code : " + str(responseFetchSolutionLinkApi.status_code))
        ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)

        if responseFetchSolutionLinkApi.status_code == 200:
            print('Fetch solution Link Api Success')
            responseProjectUploadJson = responseFetchSolutionLinkApi.json()
            solutionLink = responseProjectUploadJson["result"]
            messageArr.append("Response : " + str(responseFetchSolutionLinkApi.text))
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)

            if os.path.exists(MainFilePath + "/" + str(programFile).replace(".xlsx", "") + '-SuccessSheet.xlsx'):
                xfile = openpyxl.load_workbook(
                    MainFilePath + "/" + str(programFile).replace(".xlsx", "") + '-SuccessSheet.xlsx')
            else:
                xfile = openpyxl.load_workbook(programFile)
            print(xfile.sheetnames)

            sheet_name = 'Resource Details'.strip()

            resourceDetailsSheet = xfile[sheet_name]

            greenFill = PatternFill(start_color='0000FF00',
                                    end_color='0000FF00',
                                    fill_type='solid')
            rowCountRD = resourceDetailsSheet.max_row
            columnCountRD = resourceDetailsSheet.max_column
            for row in range(3, rowCountRD + 1):
                if str(resourceDetailsSheet["B" + str(row)].value).rstrip().lstrip().lower() == "course":
                    resourceDetailsSheet["D1"] = ""
                    resourceDetailsSheet["E1"] = ""
                    resourceDetailsSheet['I2'] = "External id of the resource"
                    resourceDetailsSheet['J2'] = "link to access the resource/Response"
                    resourceDetailsSheet['I2'].fill = greenFill
                    resourceDetailsSheet['J2'].fill = greenFill
                    resourceDetailsSheet['I' + str(row)] = solutionExternalId
                    resourceDetailsSheet['J' + str(row)] = "The course has been successfully mapped to the program"
                    resourceDetailsSheet['I' + str(row)].fill = greenFill
                    resourceDetailsSheet['J' + str(row)].fill = greenFill
                elif str(resourceDetailsSheet["A" + str(row)].value).strip() == solutionName:
                    resourceDetailsSheet["D1"] = ""
                    resourceDetailsSheet["E1"] = ""
                    resourceDetailsSheet['I2'] = "External id of the resource"
                    resourceDetailsSheet['J2'] = "link to access the resource/Response"
                    resourceDetailsSheet['I2'].fill = greenFill
                    resourceDetailsSheet['J2'].fill = greenFill
                    resourceDetailsSheet['I' + str(row)] = solutionExternalId
                    resourceDetailsSheet['J' + str(row)] = solutionLink
                    resourceDetailsSheet['I' + str(row)].fill = greenFill
                    resourceDetailsSheet['J' + str(row)].fill = greenFill

            programFile = str(programFile).replace(".xlsx", "")
            xfile.save(MainFilePath + "/" + programFile + '-SuccessSheet.xlsx')
            print("Program success sheet is created")
            return solutionLink
        else:
            print("Fetch solution link API Failed")
            messageArr.append("Response : " + str(responseFetchSolutionLinkApi.text))
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            sys.exit()

    def check_sequence(arr):
        for i in range(1, len(arr)):
            if arr[i] != arr[i - 1] + 1:
                return False
        return True
    
    def ObsWRValidate(wbObservation1, accessToken, parentFolder,typeofSolution):
        print("Validating Observation temp....")
        global errorVar, entityType, solutionName, solutionDescription, scopeEntityType, dikshaLoginId, pointBasedValue
        ObsImpFlag = False
        print(typeofSolution)
        try:
            # wbObservation1 = xlrd.open_workbook(filePathAddObs, on_demand=True)
            sheetNames1 = wbObservation1.sheet_names()
            ecmIds = list()
            criteriaLevels = list()
            criteriaExternalIds = list()
            rubrics_sheet_names = ['Instructions', 'details', 'framework', 'ECMs or Domains', 'questions','Criteria_Rubric-Scoring', 'Domain(theme)_rubric_scoring']
            rubrics_sheet_IMP_names = ['Instructions', 'details', 'framework', 'ECMs or Domains', 'questions','Criteria_Rubric-Scoring', 'Domain(theme)_rubric_scoring', 'Imp mapping']
            if (len(rubrics_sheet_names) == len(sheetNames1)) and ((set(rubrics_sheet_names) == set(sheetNames1))):
                print("--->Observation with rubrics file detected.<---")
            elif (len(rubrics_sheet_IMP_names) == len(sheetNames1)) and ((set(rubrics_sheet_IMP_names) == set(sheetNames1))):
                print("--->Observation with rubrics and IMP file detected.<---")
                typeofSolution = 5
            for sheetEnv in sheetNames1:
                questionsequenceArr =[]
                if sheetEnv == "Instructions":
                    pass
                else:
                    if sheetEnv.strip().lower() == 'details':
                        print("--->Checking details sheet...")
                        detailsCols = ["observation_solution_name", "observation_solution_description", "Elevate_loginId","Name_of_the_creator", "language", "allow_multiple_submissions", "keywords","scoring_system", "entity_type","start_date","end_date"]
                        detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                        keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                                range(detailsEnvSheet.ncols)]
                        for row_index_env in range(2, detailsEnvSheet.nrows):
                            dictDetailsEnv = {
                                keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                col_index_env in range(detailsEnvSheet.ncols)}
                            if set(detailsCols) == set(dictDetailsEnv.keys()):
                                solutionName = dictDetailsEnv['observation_solution_name'].encode('utf-8').decode('utf-8') if dictDetailsEnv['observation_solution_name'] else ElevateObservation.terminatingMessage("\"observation_solution_name\" must not be Empty in \"details\" sheet")
                                dikshaLoginId = dictDetailsEnv['Elevate_loginId'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Elevate_loginId'] else ElevateObservation.terminatingMessage("\"Elevate_loginId\" must not be Empty in \"details\" sheet")
                                ccUserDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dikshaLoginId)
                                # if not "CONTENT_CREATOR" in ccUserDetails[3]:
                                #     terminatingMessage("---> "+dikshaLoginId +" is not a CONTENT_CREATOR in Diksha " + environment)
                                # ccRootOrgName = ccUserDetails[4]
                                # ccRootOrgId = ccUserDetails[5]
                                print(str(dictDetailsEnv['scoring_system']).encode('utf-8').decode('utf-8'),"1245")
                                solutionDescription = dictDetailsEnv['observation_solution_description'].encode('utf-8').decode('utf-8')
                                pointBasedValue = str(dictDetailsEnv['scoring_system']).encode('utf-8').decode('utf-8') if dictDetailsEnv['scoring_system'] else ElevateObservation.terminatingMessage("\"scoring_system\" must not be Empty in \"details\" sheet")
                                entityType = dictDetailsEnv['entity_type'].encode('utf-8').decode('utf-8') if dictDetailsEnv['entity_type'] else ElevateObservation.terminatingMessage("\"entity_type\" must not be Empty in \"details\" sheet")

                                solutionLanguage = dictDetailsEnv['language'].split(",") if dictDetailsEnv['language'] else [""]
                                keyWords = dictDetailsEnv['keywords'].encode('utf-8').decode('utf-8')
                                creator = dictDetailsEnv['Name_of_the_creator'].encode('utf-8').decode('utf-8')  if dictDetailsEnv['Name_of_the_creator'] else ElevateObservation.terminatingMessage("\"Name_of_the_creator\" must not be Empty in \"details\" sheet")
                                allow_multiple_submissions = dictDetailsEnv['allow_multiple_submissions']
                                if allow_multiple_submissions == 1 or allow_multiple_submissions == 'TRUE':
                                    allow_multiple_submissions = True
                                else:
                                    allow_multiple_submissions = False

                                scopeEntityType = scopeEntityType

                                isProgramnamePresent = False
                                if programName == "":
                                    isProgramnamePresent = False
                                else:
                                    isProgramnamePresent = True
                                    ElevateObservation.getProgramInfo(accessToken, parentFolder, programName)
                            else:
                                ElevateObservation.terminatingMessage("--->Columns Mismatch in Details Sheet.")
                    if sheetEnv and sheetEnv.strip().lower() == 'framework':
                        frameworkCols = ["Domain ID", "Domain Name", "Criteria ID", "criteria_name", "L1 description","L2 description", "L3 description"]
                        print("--->Checking frameworks sheet...")
                        detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                        keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                                range(detailsEnvSheet.ncols)]
                        listOfThemeCriteria = list()
                        for row_index_env in range(1, detailsEnvSheet.nrows):
                            dictDetailsEnv = {
                                keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                col_index_env in range(detailsEnvSheet.ncols)}
                            countLevelUp = 1
                            for eachColNameCheck in keysEnv:
                                if "L" + str(countLevelUp) + " description" == eachColNameCheck:
                                    countLevelUp += 1
                            for i in range(1, countLevelUp):
                                if not i in criteriaLevels:
                                    criteriaLevels.append(i)

                            if dictDetailsEnv['Criteria ID'].encode('utf-8').decode('utf-8'):
                                if not [dictDetailsEnv['Domain ID'], dictDetailsEnv['Criteria ID']] in listOfThemeCriteria:
                                    listOfThemeCriteria.append([dictDetailsEnv['Domain ID'], dictDetailsEnv['Criteria ID']])
                                else:
                                    ElevateObservation.terminatingMessage("Theme , criteria combo repeating in framework sheet.")
                            if not dictDetailsEnv['Domain ID']:
                                ElevateObservation.terminatingMessage("Domain ID cannot be empty in framework sheet.")
                            if not dictDetailsEnv['Domain Name']:
                                ElevateObservation.terminatingMessage("Theme cannot be empty in framework sheet.")

                            if dictDetailsEnv['Criteria ID']:
                                criteriaExternalIds.append(dictDetailsEnv['Criteria ID'].lower())
                    if sheetEnv.strip().lower() == 'ecms or domains':
                        print("--->Checking ECMs sheet...")
                        global ecmToSection
                        detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                        keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                                range(detailsEnvSheet.ncols)]
                        for row_index_env in range(2, detailsEnvSheet.nrows):
                            dictDetailsEnv = {
                                keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                col_index_env in range(detailsEnvSheet.ncols)}
                            if dictDetailsEnv['ECM Id/Domian ID'].lower() not in ecmIds:
                                ecmIds.append(dictDetailsEnv['ECM Id/Domian ID'].lower())
                            if not dictDetailsEnv['ECM Id/Domian ID']:
                                ElevateObservation.terminatingMessage("ECM Id/Domian ID cannot be empty in ecm\'s sheet.")
                            if not dictDetailsEnv['section_id']:
                                ElevateObservation.terminatingMessage("section_id cannot be empty in ecm\'s sheet.")
                            if not dictDetailsEnv['section_name']:
                                ElevateObservation.terminatingMessage("section_name cannot be empty in ecm\'s sheet.")
                            if not dictDetailsEnv['ECM Name/Domain Name']:
                                ElevateObservation.terminatingMessage("ECM Name/Domain Name cannot be empty in ecm\'s sheet.")
                            ecmToSection[dictDetailsEnv['section_id']] = dictDetailsEnv['ECM Id/Domian ID']
                    if sheetEnv.strip().lower() == 'questions':
                        print("--->Checking questions sheet...")
                        quesExtIds = list()
                        detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                        keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                                range(detailsEnvSheet.ncols)]
                        global numberOfResponses
                        numberOfResponses = 0
                        for qKeys in keysEnv:
                            countRespo = re.search(r"response\(R[0-9]|[1-9][0-9]|100\)$", qKeys)
                            if countRespo and not "_hint" in qKeys and "response" in qKeys:
                                numberOfResponses += 1

                        for n in range(1, numberOfResponses + 1):
                            if not "Score for R" + str(n) in keysEnv or not "response(R" + str(n) + ")_hint" in keysEnv:
                                ElevateObservation.terminatingMessage("Mandatory Key: " + "Score for R" + str(n) + " or " + "response(R" + str(
                                    n) + ")_hint is missing")
                        for row_index_env in range(2, detailsEnvSheet.nrows):
                            dictDetailsEnv = {
                                keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                col_index_env in range(detailsEnvSheet.ncols)}
                            quesExtIds.append(dictDetailsEnv['question_id'].encode('utf-8').decode('utf-8').lower())

                            if not dictDetailsEnv['criteria_id']:
                                ElevateObservation.terminatingMessage("criteria_id cannot be empty in questions sheet.")
                            if not dictDetailsEnv['criteria_id'].lower() in criteriaExternalIds:
                                ElevateObservation.terminatingMessage("Criteria ID : " + dictDetailsEnv['criteria_id'] + " in question sheet not present in criteria sheet.")
                            question_sequence = dictDetailsEnv['question_sequence'] if dictDetailsEnv['question_sequence'] else terminatingMessage("\"question_sequence\" must not be Empty in \"questions\" sheet")

                            questionsequenceArr.append(question_sequence)
                            question_sequence_arr = questionsequenceArr
                            if not dictDetailsEnv['question_primary_language']:
                                ElevateObservation.terminatingMessage("question_primary_language cannot be empty in questions sheet.")
                            if not dictDetailsEnv['question_response_type']:
                                ElevateObservation.terminatingMessage("question_response_type cannot be empty in questions sheet.")
                            if not dictDetailsEnv['question_id']:
                                ElevateObservation.terminatingMessage("question_id cannot be empty in questions sheet.")
                            if not dictDetailsEnv['criteria_id']:
                                ElevateObservation.terminatingMessage("criteria_id : " + str(
                                    dictDetailsEnv['criteria_id']) + "  cannot be empty in questions sheet.")
                            if not dictDetailsEnv['criteria_id'].lower() in criteriaExternalIds:
                                ElevateObservation.terminatingMessage("criteria_id : " + str(dictDetailsEnv['criteria_id']) + " in questions sheet is not matching the criteria upload.")
                        if not len(question_sequence_arr) == len(set(question_sequence_arr)):
                                ElevateObservation.terminatingMessage("\"question_sequence\" must be Unique in \"questions\" sheet")
                        if not len(quesExtIds) == len(set(quesExtIds)):
                            ElevateObservation.terminatingMessage("Duplicate question_id detected in questions sheet.")
                        if not ElevateObservation.check_sequence(question_sequence_arr): ElevateObservation.terminatingMessage("\"question_sequence\" must be in sequence in \"questions\" sheet")
                    if typeofSolution == 5:
                        if sheetEnv.strip().lower() == 'imp mapping':
                            print("--->Checking Imp mapping sheet...")
                            global countImps
                            countImps = 1
                            detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                            keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                                    range(detailsEnvSheet.ncols)]
                            for row_index_env in range(2, detailsEnvSheet.nrows):
                                dictDetailsEnv = {
                                    keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                    col_index_env in range(detailsEnvSheet.ncols)}
                            for eachCols in dictDetailsEnv.keys():
                                if eachCols.strip() == "L" + str(countImps) + "-improvement-projects":
                                    countImps += 1
                            countImps = countImps - 1
                    if not pointBasedValue.lower() == "null":
                        if sheetEnv.strip().lower() == 'Criteria_Rubric-Scoring':
                            print("--->Checking Criteria Rubrics sheet")
                            cR_extIds = list()
                            detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                            keysEnv = [detailsEnvSheet.cell(0, col_index_env).value for col_index_env in
                                    range(detailsEnvSheet.ncols)]
                            listOfCRs = ["criteriaId", "weightage"]
                            for cl in criteriaLevels:
                                listOfCRs.append("L" + str(cl))
                            for keyys in keysEnv:
                                if not keyys in listOfCRs:
                                    print("--->" + keyys + " : unwanted column detected...")
                                    print("==>PS :  unwanted column will be ignored while uploading...")
                            for row_index_env in range(1, detailsEnvSheet.nrows):
                                dictDetailsEnv = {
                                    keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                    col_index_env in range(detailsEnvSheet.ncols)}
                                cR_extIds.append(dictDetailsEnv['criteriaId'].lower())
                                for cl in criteriaLevels:
                                    if not dictDetailsEnv["L" + str(cl)]:
                                        ElevateObservation.terminatingMessage("L" + str(cl) + " must not be empty in criteria_rubric.")
                                if dictDetailsEnv['criteriaId']:
                                    ElevateObservation.terminatingMessage("criteriaId must be empty in criteria_rubric sheet.")
                                if not dictDetailsEnv['weightage']:
                                    ElevateObservation.terminatingMessage("weightage cannot be empty in criteria_rubric sheet.")
                            if not len(cR_extIds) == len(set(cR_extIds)):
                                ElevateObservation.terminatingMessage("Duplicate externalId detected in criteria_rubric sheet.")
                        if sheetEnv.strip().lower() == 'Domain(theme)_rubric_scoring':
                            print("--->Checking Theme Rubrics sheet")
                            detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                            keysEnv = [detailsEnvSheet.cell(0, col_index_env).value for col_index_env in
                                    range(detailsEnvSheet.ncols)]
                            for row_index_env in range(1, detailsEnvSheet.nrows):
                                dictDetailsEnv = {
                                    keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                    col_index_env in range(detailsEnvSheet.ncols)}
                                if not dictDetailsEnv['domain_Id']:
                                    ElevateObservation.terminatingMessage("domain_Id cannot be empty in theme_rubric sheet.")
                                if not dictDetailsEnv['domain_name']:
                                    ElevateObservation.terminatingMessage("domain_name cannot be empty in theme_rubric sheet.")
                                if not dictDetailsEnv['weightage']:
                                    ElevateObservation.terminatingMessage("weightage cannot be empty in theme_rubric sheet.")

            if errorVar == "":
                return True
            else:
                print(errorVar,"3292")
                return False
        except Exception as e:
            print(f"Error during ECM processing: {str(e)}")
            print(errorVar,"3270")

    def mainFunc(MainFilePath, programFile, addObservationSolution, millisecond, isProgramnamePresent, isCourse,
             scopeEntityType=scopeEntityType):
        scopeEntityType = scopeEntityType
        if not isCourse:
            parentFolder = ElevateObservation.createFileStruct(MainFilePath, addObservationSolution)
            accessToken = ElevateObservation.generateAccessToken(parentFolder)
            ElevateObservation.programsFileCheck(programFile, accessToken, parentFolder, MainFilePath)
            # typeofSolution = validateSheets(addObservationSolution, accessToken, parentFolder)
            typeofSolution = ElevateObservation.typeofresource(addObservationSolution, accessToken, parentFolder)
            print(typeofSolution,"this is type of solution")
            # sys.exit()
            wbObservation = xlrd.open_workbook(addObservationSolution, on_demand=True)
            projectSheetNames = wbObservation.sheet_names()
            wbProgram = xlrd.open_workbook(programFile, on_demand=True)
            programSheetNames = wbProgram.sheet_names()
            for programSheets in programSheetNames:
                if programSheets.strip().lower() == 'program details':
                    print("Checking program details sheet...")
                    programDetailsSheet = wbProgram.sheet_by_name(programSheets)
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

                        if dictProgramDetails.get('Targeted state at program level'):
                            userEntity = dictProgramDetails['Targeted state at program level'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "\"Targeted state at program level\" must not be Empty in \"details\" sheet"
        
            if typeofSolution == 1 or typeofSolution == 5:
                if typeofSolution == 5:
                    impLedObsFlag = True
                else:
                    impLedObsFlag = False
            for sheets in projectSheetNames:
                if sheets.strip().lower() == 'details'.lower() and typeofSolution in [1, 5]:
                    ResourceSheet = wbObservation.sheet_by_name(sheets)
                    keysEnv = [ResourceSheet.cell(1, col_index_env).value for col_index_env in range(ResourceSheet.ncols)]
                    dictDetailsEnv = {keysEnv[col_index_env]: ResourceSheet.cell(row_index_env, col_index_env).value
                                    for col_index_env in range(ResourceSheet.ncols)}
                    ObsWRResourceName = dictDetailsEnv['observation_solution_name'].encode('utf-8').decode('utf-8')
                    try:
                        if not ElevateObservation.ObsWRValidate(wbObservation, accessToken, parentFolder,typeofSolution,):
                            print("Error during validation of Observation with Rubric file ....")
                            finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                            return finalObsRubricSolutionLink   
                        print("validation successful")
                        def addObsWRFunc(parentFolder, wbObservation, millisecond, accessToken):
                            ElevateObservation.criteriaUpload(parentFolder, wbObservation, millisecond, accessToken, "framework", impLedObsFlag)
                            
                            userDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dikshaLoginId)
                            matchedShikshalokamLoginId = userDetails[0]
                            
                            frameworkExternalId = ElevateObservation.frameWorkUpload(parentFolder, wbObservation, millisecond, accessToken)
                            observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
                            ElevateObservation.themesUpload(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, False)
                            solutionId = ElevateObservation.createSolutionFromFramework(parentFolder, accessToken, frameworkExternalId)

                            ecmsSheet = wbObservation.sheet_by_name('ECMs or Domains')
                            keys = [ecmsSheet.cell(1, col_index).value for col_index in range(ecmsSheet.ncols)]
                            ecm_update = dict()
                            ecm_dict = dict()
                            section = dict()
                            ecmSeqCount = 1
                            for row_index in range(2, ecmsSheet.nrows):
                                dictECMs = {keys[col_index]: ecmsSheet.cell(row_index, col_index).value for col_index in
                                            range(ecmsSheet.ncols)}
                                EMC_ID = dictECMs['ECM Id/Domian ID'].encode('utf-8').decode('utf-8').strip() + '_' + str(millisecond)
                                ECM_NAME = dictECMs['ECM Name/Domain Name'].encode('utf-8').decode('utf-8').strip()
                                section.update({dictECMs['section_id']: dictECMs['section_name']})
                                ecm_sections[EMC_ID] = dictECMs['section_id']
                                if dictECMs['Is ECM Mandatory?']:
                                    if dictECMs['Is ECM Mandatory?'] == "TRUE" or dictECMs['Is ECM Mandatory?'] == 1:
                                        dictECMs['Is ECM Mandatory?'] = False
                                    elif dictECMs['Is ECM Mandatory?'] == "FALSE" or dictECMs['Is ECM Mandatory?'] == 0:
                                        dictECMs['Is ECM Mandatory?'] = True
                                else:
                                    dictECMs['Is ECM Mandatory?'] = False
                                ecm_update[EMC_ID] = {
                                    "externalId": EMC_ID, "tip": None, "name": ECM_NAME, "description": None,
                                    "modeOfCollection": "onfield",
                                    "canBeNotApplicable": dictECMs['Is ECM Mandatory?'],
                                    "notApplicable": False, "canBeNotAllowed": dictECMs['Is ECM Mandatory?'],
                                    "remarks": None,
                                    "sequenceNo": ecmSeqCount
                                    }
                                ecmSeqCount += 1
                            ecm_dict['evidenceMethods'] = ecm_update
                            bodySolutionUpdate = ecm_dict
                            ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                            bodySolutionUpdate = {"sections": section}
                            ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                            bodySolutionUpdate = {"status": "active", "isDeleted": False, "criteriaLevelReport": criteriaLevelsReport}
                            ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                            excelBook = open_workbook(addObservationSolution)
                            ElevateObservation.questionUpload(addObservationSolution, parentFolder, frameworkExternalId, millisecond, accessToken,solutionId,typeofSolution)
                            if not pointBasedValue.lower() == "null":
                                bodySolutionUpdate = {"isRubricDriven": True}
                                ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                                ElevateObservation.fetchSolutionCriteria(parentFolder, observationExternalId, accessToken)
                                ElevateObservation.uploadCriteriaRubrics(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, True)
                                ElevateObservation.uploadThemeRubrics(parentFolder, wbObservation, accessToken, frameworkExternalId, True)
                            else:
                                print("Observation with scoring system : null.")
                            bodySolutionUpdate = {'allowMultipleAssessemts': allow_multiple_submissions, "creator": creator}
                            ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                            solutionDetails = ElevateObservation.fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, solutionId, accessToken)
                            if solutionDetails[1]:
                                startDateArr = str(solutionDetails[1]).split("-")
                                bodySolutionUpdate = {
                                    "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                                ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                            if solutionDetails[2]:
                                print(solutionDetails[2],"this is 5294")
                                endDateArr = str(solutionDetails[2]).split("-")
                                bodySolutionUpdate = {
                                    "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                                ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                            if isProgramnamePresent:
                                childId = ElevateObservation.createChild(parentFolder, observationExternalId, accessToken)
                                if childId[0]:
                                    solutionDetails = ElevateObservation.fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, childId[0],
                                                                                        accessToken)
                                    scopeEntities = entitiesPGMID
                                    scopeRoles = solutionDetails[0]
                                    scope = {}
                                    for i in range(len(entitiesType)):
                                        entity_type = entitiesType[i]
                                        entity_value = scopeEntities[i]
                                        if entity_type in scope:
                                            scope[entity_type].append(entity_value)
                                        else:
                                            scope[entity_type] = [entity_value]
                                    # bodySolutionUpdate = {
                                    #     "scope": {"entityType": scopeEntityType, "entities": scopeEntities, "roles": scopeRoles}}
                                    scope["roles"] = scopeRoles
                                    bodySolutionUpdate = {
                                        "scope": scope
                                    }
                                    
                                    ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate)
                                    if solutionDetails[1]:
                                        startDateArr = str(solutionDetails[1]).split("-")
                                        bodySolutionUpdate = {
                                            "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[
                                                0] + " 00:00:00"}
                                        ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate)
                                    if solutionDetails[2]:
                                        endDateArr = str(solutionDetails[2]).split("-")
                                        bodySolutionUpdate = {
                                            "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                                        ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate)
                                    ObsRubricSolutionLink = ElevateObservation.prepareProgramSuccessSheet(MainFilePath, parentFolder, programFile, childId[1], childId[0],
                                                            accessToken)
                                    if not ObsRubricSolutionLink:
                                        finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                        return finalObsRubricSolutionLink
                                    else:
                                        finalObsRubricSolutionLink = {ObsWRResourceName: ObsRubricSolutionLink}
                                        return finalObsRubricSolutionLink
                            else:
                                print("No program name detected.")
                        millisecond = int(time.time() * 1000)
                        ObsWRSolutionLink = addObsWRFunc(parentFolder, wbObservation, millisecond, accessToken)
                        return ObsWRSolutionLink
                    except Exception as e:
                        print(f"Error occurred during project creation: {str(e)}")
                        # raise RuntimeError("The project creation failed due to an unexpected error")
                        solutionError = str(e)
                        print(errorVar,"3266")
                        if errorVar == "":
                            finalObsRubricSolutionLink = {ObsWRResourceName: solutionError}
                        else:
                            finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                        # print(projectSolutionLink, "3247")
                        return finalObsRubricSolutionLink
                
                elif typeofSolution == 2:
                    ResourceSheet = wbObservation.sheet_by_name(sheets)
                    keysEnv = [ResourceSheet.cell(1, col_index_env).value for col_index_env in range(ResourceSheet.ncols)]
                    
                    # Collect observation solution details
                    dictDetailsEnv = {
                        keysEnv[col_index_env]: ResourceSheet.cell(row_index_env, col_index_env).value
                        for col_index_env in range(ResourceSheet.ncols)
                    }
                    ObsWORResourceName = dictDetailsEnv['observation_solution_name'].encode('utf-8').decode('utf-8')
                    pointBasedValue = "null"
                    try:
                        def addObsWORFunc(parentFolder, wbObservation, millisecond, accessToken):
                            if not ElevateObservation.ObsWORValidate(wbObservation, accessToken, parentFolder):
                                print(errorVar,"---->validation error 3931")
                                ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                return ObsWORSolutionLink
                            print("Create Observation Function called ....")
                            ElevateObservation.criteriaUpload(parentFolder, wbObservation, millisecond, accessToken, "criteria", False)
                            
                            userDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dikshaLoginId)
                            matchedShikshalokamLoginId = userDetails[0]
                            
                            frameworkExternalId = ElevateObservation.frameWorkUpload(parentFolder, wbObservation, millisecond, accessToken)
                            observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
                            ElevateObservation.themesUpload(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, True)
                            solutionId = ElevateObservation.createSolutionFromFramework(parentFolder, accessToken, frameworkExternalId)
                            sectionsObj = {"sections": {'S1': 'Observation Question'}}
                            ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, sectionsObj)
                            ecmObj = {}
                            ecmExternalId = None
                            ecmObj = {
                                "evidenceMethods": {'OB': {'externalId': 'OB', 'tip': None, 'name': 'Observation', 'description': None,
                                                        'modeOfCollection': 'onfield', 'canBeNotApplicable': False,
                                                        'notApplicable': False, 'canBeNotAllowed': False, 'remarks': None}}}
                            ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, ecmObj)
                            ElevateObservation.questionUpload(addObservationSolution, parentFolder, frameworkExternalId, millisecond, accessToken,
                                        solutionId, typeofSolution)
                            # fetchSolutionCriteria(parentFolder, observationExternalId, accessToken)
                            if not pointBasedValue.lower() == "null":
                                ElevateObservation.uploadCriteriaRubrics(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, False)
                                ElevateObservation.uploadThemeRubrics(parentFolder, wbObservation, accessToken, frameworkExternalId, False)
                            bodySolutionUpdate = {"status": "active", "isDeleted": False, "allowMultipleAssessemts": True,
                                                "creator": creator}
                            ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)

                            solutionDetails = ElevateObservation.fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, solutionId, accessToken)
                            # Below script will convert date DD-MM-YYYY TO YYYY-MM-DD 00:00:00 to match the code syntax

                            if solutionDetails[1]:
                                startDateArr = str(solutionDetails[1]).split("-")
                                bodySolutionUpdate = {
                                    "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                                ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                            if solutionDetails[2]:
                                endDateArr = str(solutionDetails[2]).split("-")
                                bodySolutionUpdate = {
                                    "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                                ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                            if isProgramnamePresent:
                                childId = ElevateObservation.createChild(parentFolder, observationExternalId, accessToken)
                                if childId[0]:
                                    solutionDetails = ElevateObservation.fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, childId[0],
                                                                                        accessToken)
                                    scopeEntities = entitiesPGMID
                                    print(entitiesType,solutionDetails,"this is 5429")
                                    scopeRoles = solutionDetails[0]
                                    scope = {}
                                    for i in range(len(entitiesType)):
                                        entity_type = entitiesType[i]
                                        entity_value = scopeEntities[i]
                                        if entity_type in scope:
                                            scope[entity_type].append(entity_value)
                                        else:
                                            scope[entity_type] = [entity_value]
                                    # bodySolutionUpdate = {
                                    #     "scope": {"entityType": scopeEntityType, "entities": scopeEntities, "roles": scopeRoles}}
                                    scope["roles"] = scopeRoles
                                    bodySolutionUpdate = {
                                        "scope": scope
                                    }
                                    ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate)
                                    if solutionDetails[1]:
                                        startDateArr = str(solutionDetails[1]).split("-")
                                        bodySolutionUpdate = {
                                            "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[
                                                0] + " 00:00:00"}
                                        ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate)
                                    if solutionDetails[2]:
                                        endDateArr = str(solutionDetails[2]).split("-")
                                        bodySolutionUpdate = {
                                            "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                                        ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate)
                                    ObsSolutionLink = ElevateObservation.prepareProgramSuccessSheet(MainFilePath, parentFolder, programFile, childId[1], childId[0],
                                                            accessToken)
                                    if not ObsSolutionLink:
                                        ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                        return ObsWORSolutionLink
                                    else:
                                        print(ObsSolutionLink)
                                        finalObsSolutionLink = {ObsWORResourceName: ObsSolutionLink}
                                        return finalObsSolutionLink
                            else:
                                print("No program name detected.")

                        millisecond = int(time.time() * 1000)
                        ObsWORSolutionLink = addObsWORFunc(parentFolder, wbObservation, millisecond, accessToken)
                        return ObsWORSolutionLink
                    except Exception as e:
                        print(f"Error occurred: {str(e)}")
                        solutionError = str(e)
                        print(errorVar,"3266")
                        if errorVar == "":
                            ObsWORSolutionLink = {ObsWORResourceName: solutionError}
                        else:
                            ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                        return ObsWORSolutionLink
        else :
            parentFolder = ElevateObservation.createFileStruct(MainFilePath, addObservationSolution)
            accessToken = ElevateObservation.generateAccessToken(parentFolder)
            typeofSolution = ElevateObservation.typeofresource(addObservationSolution, accessToken, parentFolder)            
            wbObservation = xlrd.open_workbook(addObservationSolution, on_demand=True)
            print(typeofSolution,"this is type of solution")
            if typeofSolution == 1 or typeofSolution == 5:
                if typeofSolution == 5:
                    impLedObsFlag = True
                else:
                    impLedObsFlag = False
                ElevateObservation.criteriaUpload(parentFolder, wbObservation, millisecond, accessToken, "framework", impLedObsFlag)
                
                userDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dikshaLoginId)
                matchedShikshalokamLoginId = userDetails[0]
                
                frameworkExternalId = ElevateObservation.frameWorkUpload(parentFolder, wbObservation, millisecond, accessToken)
                observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
                ElevateObservation.themesUpload(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, False)
                solutionId = ElevateObservation.createSolutionFromFramework(parentFolder, accessToken, frameworkExternalId)

                ecmsSheet = wbObservation.sheet_by_name('ECMs or Domains')
                keys = [ecmsSheet.cell(1, col_index).value for col_index in range(ecmsSheet.ncols)]
                ecm_update = dict()
                ecm_dict = dict()
                section = dict()
                ecmSeqCount = 1
                for row_index in range(2, ecmsSheet.nrows):
                    dictECMs = {keys[col_index]: ecmsSheet.cell(row_index, col_index).value for col_index in
                                range(ecmsSheet.ncols)}
                    EMC_ID = dictECMs['ECM Id/Domian ID'].encode('utf-8').decode('utf-8').strip() + '_' + str(millisecond)
                    ECM_NAME = dictECMs['ECM Name/Domain Name'].encode('utf-8').decode('utf-8').strip()
                    section.update({dictECMs['section_id']: dictECMs['section_name']})
                    ecm_sections[EMC_ID] = dictECMs['section_id']
                    if 'Is ECM Mandatory?' in dictECMs:  # Ensure the key exists
                        value = dictECMs['Is ECM Mandatory?']
                        if value == "TRUE" or value == 1:
                            dictECMs['Is ECM Mandatory?'] = False
                        elif value == "FALSE" or value == 0:
                            dictECMs['Is ECM Mandatory?'] = True
                    else:
                        dictECMs['Is ECM Mandatory?'] = False
                    ecm_update[EMC_ID] = {
                        "externalId": EMC_ID, "tip": None, "name": ECM_NAME, "description": None,
                        "modeOfCollection": "onfield",
                        "canBeNotApplicable": dictECMs['Is ECM Mandatory?'],
                        "notApplicable": False, "canBeNotAllowed": dictECMs['Is ECM Mandatory?'],
                        "remarks": None,
                        "sequenceNo": ecmSeqCount
                        }
                    ecmSeqCount += 1
                ecm_dict['evidenceMethods'] = ecm_update
                bodySolutionUpdate = ecm_dict
                ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                bodySolutionUpdate = {"sections": section}
                ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                bodySolutionUpdate = {"status": "active", "isDeleted": False, "criteriaLevelReport": criteriaLevelsReport}
                ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                excelBook = open_workbook(addObservationSolution)
                ElevateObservation.questionUpload(addObservationSolution, parentFolder, frameworkExternalId, millisecond, accessToken,solutionId,typeofSolution)
                if not pointBasedValue.lower() == "null":
                    bodySolutionUpdate = {"isRubricDriven": True}
                    ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                    ElevateObservation.fetchSolutionCriteria(parentFolder, observationExternalId, accessToken)
                    ElevateObservation.uploadCriteriaRubrics(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, True)
                    ElevateObservation.uploadThemeRubrics(parentFolder, wbObservation, accessToken, frameworkExternalId, True)
                else:
                    print("Observation with scoring system : null.")
                bodySolutionUpdate = {'allowMultipleAssessemts': allow_multiple_submissions, "creator": creator}
                ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                global solutionRolesArray, solutionStartDate, solutionEndDate
                xfile = openpyxl.load_workbook(programFile)
                sheet_name = 'details'.strip()
                resourceDetailsSheet = xfile[sheet_name]
                solutionDetails = ElevateObservation.fetchSolutionDetailsFromResourceSheet(parentFolder, programFile, solutionId, accessToken,typeofSolution)
                if solutionDetails[1]:
                    startDateArr = str(solutionDetails[1]).split("-")
                    bodySolutionUpdate = {
                        "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                    ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                if solutionDetails[2]:
                    print(solutionDetails[2],"this is 5294")
                    endDateArr = str(solutionDetails[2]).split("-")
                    bodySolutionUpdate = {
                        "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                    ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                if isProgramnamePresent:
                    childId = ElevateObservation.createChild(parentFolder, observationExternalId, accessToken)
                    if childId[0]:
                        # solutionDetails = fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, childId[0],
                        #                                                        accessToken)
                        # scopeEntities = entitiesPGMID
                        # scopeRoles = solutionDetails[0]
                        # scope = {}
                        # for i in range(len(entitiesType)):
                        #     entity_type = entitiesType[i]
                        #     entity_value = scopeEntities[i]
                        # if entity_type in scope:
                        #    scope[entity_type] = []
                        #    scope[entity_type].append(entity_value)
                        # else:
                        #     scope[entity_type] = [entity_value]
                        # bodySolutionUpdate = {
                        #     "scope": {"entityType": scopeEntityType, "entities": scopeEntities, "roles": scopeRoles}}
                        # scope["roles"] = [rolesPGM]
                        # bodySolutionUpdate = {
                        #     "scope": scope
                        #  }
                        
                        ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate)
                        if solutionDetails[1]:
                            startDateArr = str(solutionDetails[1]).split("-")
                            bodySolutionUpdate = {
                                "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[
                                    0] + " 00:00:00"}
                            ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate)
                        if solutionDetails[2]:
                            endDateArr = str(solutionDetails[2]).split("-")
                            bodySolutionUpdate = {
                                "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                            ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate)
                        ElevateObservation.prepareProgramSuccessSheet(MainFilePath, parentFolder, programFile, childId[1], childId[0],
                                                accessToken)
                else:
                    print("No program name detected.")
            elif typeofSolution == 2:
                ElevateObservation.criteriaUpload(parentFolder, wbObservation, millisecond, accessToken, "criteria", False)
                
                userDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dikshaLoginId)
                matchedShikshalokamLoginId = userDetails[0]
                
                frameworkExternalId = ElevateObservation.frameWorkUpload(parentFolder, wbObservation, millisecond, accessToken)
                observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
                ElevateObservation.themesUpload(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, True)
                solutionId = ElevateObservation.createSolutionFromFramework(parentFolder, accessToken, frameworkExternalId)
                sectionsObj = {"sections": {'S1': 'Observation Question'}}
                ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, sectionsObj)
                ecmObj = {}
                ecmExternalId = None
                ecmObj = {
                    "evidenceMethods": {'OB': {'externalId': 'OB', 'tip': None, 'name': 'Observation', 'description': None,
                                            'modeOfCollection': 'onfield', 'canBeNotApplicable': False,
                                            'notApplicable': False, 'canBeNotAllowed': False, 'remarks': None}}}
                ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, ecmObj)
                ElevateObservation.questionUpload(addObservationSolution, parentFolder, frameworkExternalId, millisecond, accessToken,
                            solutionId, typeofSolution)
                ElevateObservation.fetchSolutionCriteria(parentFolder, observationExternalId, accessToken)
                if not pointBasedValue.lower() == "null":
                    ElevateObservation.uploadCriteriaRubrics(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, False)
                    ElevateObservation.uploadThemeRubrics(parentFolder, wbObservation, accessToken, frameworkExternalId, False)
                bodySolutionUpdate = {"status": "active", "isDeleted": False, "allowMultipleAssessemts": True,
                                    "creator": creator}
                ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)

                solutionDetails = ElevateObservation.fetchSolutionDetailsFromResourceSheet(parentFolder, programFile, solutionId, accessToken,typeofSolution)
                # Below script will convert date DD-MM-YYYY TO YYYY-MM-DD 00:00:00 to match the code syntax

                if solutionDetails[1]:
                    startDateArr = str(solutionDetails[1]).split("-")
                    bodySolutionUpdate = {
                        "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                    ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                if solutionDetails[2]:
                    endDateArr = str(solutionDetails[2]).split("-")
                    bodySolutionUpdate = {
                        "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                    ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate)
                if isProgramnamePresent:
                    childId = ElevateObservation.createChild(parentFolder, observationExternalId, accessToken)
                    if childId[0]:
                        # solutionDetails = fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, childId[0],
                        #                                                        accessToken)
                        # scopeEntities = entitiesPGMID
                        # scopeRoles = solutionDetails[0]
                        # scope = {}
                        # print(entitiesType,"this is a 5755")
                        # for i in range(len(entitiesType)):
                        #    entity_type = entitiesType[i]
                        #    entity_value = scopeEntities[i]
                        # if entity_type in scope:
                        #     scope[entity_type] = []
                        #     scope[entity_type].append(entity_value)
                        # else:
                        #     scope[entity_type] = [entity_value]
                        # # bodySolutionUpdate = {
                        #  #     "scope": {"entityType": scopeEntityType, "entities": scopeEntities, "roles": scopeRoles}}
                        # scope["roles"] = scopeRoles
                        # bodySolutionUpdate = {
                        #      "scope": scope
                        # }
                        ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate)
                        if solutionDetails[1]:
                            startDateArr = str(solutionDetails[1]).split("-")
                            bodySolutionUpdate = {
                                "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[
                                    0] + " 00:00:00"}
                            ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate)
                        if solutionDetails[2]:
                            endDateArr = str(solutionDetails[2]).split("-")
                            bodySolutionUpdate = {
                                "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                            ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate)
                        ElevateObservation.prepareProgramSuccessSheet(MainFilePath, parentFolder, programFile, childId[1], childId[0],
                                                accessToken)
                else:
                    print("No program name detected.")

    def loadSurveyFile(programFile):
        solutionDict = {}
        start_time = time.time()
        parser = argparse.ArgumentParser()
        parser.add_argument('--programFile', '--resourceFile', type=ElevateObservation.valid_file)
        parser.add_argument('--env', '--env')
        argument = parser.parse_args()
        programFile = argument.programFile
        environment = argument.env
        millisecond = int(time.time() * 1000)
        MainFilePath = ElevateObservation.createFileStructForProgram(programFile)
        wbPgm = xlrd.open_workbook(programFile, on_demand=True)
        sheetNames = wbPgm.sheet_names()
        pgmSheets = ["Instructions", "Program Details", "Resource Details","Program Manager Details"]
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
                        isProgramnamePresent = False
                        if programName == "":
                            isProgramnamePresent = False
                        else:
                            isProgramnamePresent = True
                        scopeEntityType = scopeEntityType
                        userEntity = dictProgramDetails['Targeted entities at program level'].encode('utf-8').decode('utf-8').lstrip().rstrip().split(
                            ",") if \
                            dictProgramDetails['Targeted entities at program level'] else ElevateObservation.terminatingMessage("\"scope_entity\" must not be Empty in \"details\" sheet")
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
                        resourceNamePGM = dictDetailsEnv['Name of resources in program'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Name of resources in program'] else ElevateObservation.terminatingMessage("\"Name of resources in program\" must not be Empty in \"Resource Details\" sheet")
                        resourceTypePGM = dictDetailsEnv['Type of resources'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Type of resources'] else ElevateObservation.terminatingMessage("\"Type of resources\" must not be Empty in \"Resource Details\" sheet")
                        resourceLinkOrExtPGM = dictDetailsEnv['Resource Link'] if dictDetailsEnv['Resource Link'] else ElevateObservation.terminatingMessage("\"Resource Link\" must not be Empty in \"Resource Details\" sheet")
                        if str(dictDetailsEnv['Type of resources']).lower().strip() == "course":
                            isCourse = False
                        else:
                            isCourse = False
                            resourceStatus = dictDetailsEnv['Resource Status'] if dictDetailsEnv['Resource Status'] else ElevateObservation.terminatingMessage("\"Resource Status\" must not be Empty in \"Resource Details\" sheet")
                            if resourceStatus.strip()=="New Upload":
                                print("--->Resource Name : "+str(resourceNamePGM))
                                resourceLinkOrExtPGM = str(resourceLinkOrExtPGM).split('/')[5]
                                file_url = 'https://docs.google.com/spreadsheets/d/' + resourceLinkOrExtPGM + '/export?format=xlsx'
                                if not os.path.isdir('InputFiles'):
                                    os.mkdir('InputFiles')
                                dest_file = 'InputFiles'
                                download_file = wget.download(file_url, dest_file)
                                # print("--->solution input file successfully downloaded" + str(addObservationSolution))
                                # ElevateObservation.mainFunc(MainFilePath, programFile, addObservationSolution, millisecond, isProgramnamePresent,isCourse, )\
                                downloaded_file.append(download_file)

            print("--->Solution input file successfully downloaded: " + str(downloaded_file))
            for addObservationSolution in downloaded_file:
                print(f"Processing file: {addObservationSolution}")
                solutionSL = ElevateObservation.mainFunc(MainFilePath, programFile, addObservationSolution, millisecond, isProgramnamePresent, isCourse,
             scopeEntityType=scopeEntityType)
                print(solutionSL)
                print(solutionSL.items(),"3400")
                for resourceName, solutionLink in solutionSL.items():
                    solutionDict[resourceName] = solutionLink
            downloaded_file = None
        else :
            MainFilePath = ElevateObservation.createFileStructForProgram(programFile)
            print(programFile,"58222")
            addObservationSolution = programFile
            wbPgm = xlrd.open_workbook(programFile, on_demand=True)
            millisecond = int(time.time() * 1000)
            # Specify the local path of the Excel file
            local=os.getcwd()
            resourceLinkOrExtPGMcopy = local+'/'+str(programFile)
            if not os.path.isdir('InputFiles'):
                os.mkdir('InputFiles')
            shutil.copy(resourceLinkOrExtPGMcopy,'InputFiles' )
            print("--->solution input file successfully copied")
            ElevateObservation.mainFunc(MainFilePath, programFile, os.path.join('InputFiles',programFile),millisecond ,isProgramnamePresent =True,isCourse=True)
        
        result = {
            "solutionDict": solutionDict,
            "programName": programName 
        }
        return json.dumps(result)
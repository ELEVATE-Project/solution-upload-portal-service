import os
import jwt
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
import pandas as pd
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
# from common_config import *
from backend.src.main.modules.common_config import *
# from common_config import elevateuserhost, userlogin, keyclockAPIBody
import jwt
import threading
import wget
import gdown
from dotenv import load_dotenv
from pathlib import Path
# get current working directory
currentDirectory = os.getcwd()

# Read config file 
config = ConfigParser()
config.read('common_config/config.ini')


# email regex
regex = "\"?([-a-zA-Z0-9.`?{}]+@\w+\.\w+)\"?"

env_path = Path(__file__).resolve().parents[1] / "apiServices" / "src" / "main" / ".env"

# Load the .env file
load_dotenv(dotenv_path=env_path)

SECRET_KEY = os.getenv("SECRET_KEY")
ADMIN_TOKEN = os.getenv("admin-token")
internal_access_token = os.getenv("internal_access_token")
adminTokenHeaderName = os.getenv("adminTokenHeaderName")
jwtTokenSecret = os.getenv("jwtTokenSecret")
projAdminAccessToken = os.getenv("projAdminAccessToken")
adminAccessToken = os.getenv("adminAccessToken")
authorization = os.getenv("authorization")
authorizationforhost = os.getenv("authorizationforhost")
appname = os.getenv("appname")
x_channel_id = os.getenv("x_channel_id")
host = os.getenv("host")
userLoginHost = os.getenv("userLoginHost")
internal_kong_ip = os.getenv("internal_kong_ip")
elevateprojecthost = os.getenv("elevateprojecthost")
elevateentityhost = os.getenv("elevateentityhost")
identifier = os.getenv("identifier")
password = os.getenv("password")
origin = os.getenv("origin")


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
stateEntitiesPGM = []
districtEntitiesPGM = []
blockEntitiesPGM = []
clusterEntitiesPGM = []
schoolEntitiesPGM = []
orgIdForScope = []
entityToUpload = None
programID = None
programExternalId = None
programDescription = None
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
tenantID = None
orgIDFromTemplate = None
roleOfResourceCreator = 'admin'
solutionDict = {}
entityHierarchy = []
isExternalProgram = ""
mainRoles = ""

class ElevateObservation:

    # def terminatingMessage(msg):
    #     print(msg)
    #     sys.exit()

    #helper function 
    def clean_single_value(value):
        value = str(value).strip()
        try:
            f = float(value)
            if f.is_integer():
                return str(int(f))
            return str(f)
        except ValueError:
            return value  # not a number, return as-is
        
    def append_to_list(base, items_to_add):
        if not isinstance(base, list):
            base = [base]

        if isinstance(items_to_add, list):
            return base + items_to_add
        else:
            return base + [items_to_add]

    def normalize_cell_value(value):
        if isinstance(value, str) and ',' in value:
            return [ElevateObservation.clean_single_value(part) for part in value.split(',') if part.strip()]
        return ElevateObservation.clean_single_value(value)



    def valid_file(param):
        base, ext = os.path.splitext(param)
        if ext.lower() not in ('.xlsx'):
            raise argparse.ArgumentTypeError('File must have a csv extension')
        return param

    def createFileStructForProgram(programFile):
        if not os.path.isdir('programFiles'):
            os.mkdir('programFiles')

        if "/" in str(programFile):
            fileNameSplit = str(programFile).split('/')[-1]
        else:
            fileNameSplit = os.path.basename(programFile)

        print(fileNameSplit, "fileNameSplit")

        folderName = None  # Default value

        if fileNameSplit.endswith(".xlsx"):
            ts = str(time.time()).replace(".", "_")
            folderName = fileNameSplit.replace(".xlsx", "-" + str(ts))
            os.mkdir(os.path.join('programFiles', folderName))

        if folderName is None:  # Handle cases where the file is not an Excel file
            print("Error: Unsupported file type. Returning default path.")
            return None  # You can return an error message or a default path instead

        returnPathStr = os.path.join('programFiles', folderName)
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
            print("File Error.offff")
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
        global errorVar
        accessTokenUser = None  # define it upfront

        try:
            print("Generating Access Token...")
            headerKeyClockUser = {'Content-Type': "application/x-www-form-urlencoded",'origin': "default-qa.tekdinext.com"}
            # responseKeyClockUser = requests.post(url=config.get(environment, 'elevateuserhost') + config.get(environment, 'userlogin'), headers=headerKeyClockUser,
                                                #  data=json.dumps(config.get(environment, 'keyclockAPIBody')))
            # Elevateproject.terminatingMessage(type(json.loads(config.get(environment, 'keyclockAPIBody'))))\
            loginBody = {
                'email' : identifier,
                'password' : password
            }
            responseKeyClockUser = requests.post(userLoginHost + keyclockapiurl , headers=headerKeyClockUser, data=loginBody)


            messageArr = []
            messageArr.append("URL : " + userLoginHost)
            messageArr.append("Body : " + str(loginBody))
            messageArr.append("Status Code : " + str(responseKeyClockUser.status_code))
            if responseKeyClockUser.status_code == 200:
                responseKeyClockUser = responseKeyClockUser.json()
                accessTokenUser = responseKeyClockUser['result']['access_token']
                # jwtToken  = jwtTokenSecret
                # decode = jwt.decode(accessTokenUser, jwtToken , algorithms=["HS256"])
                # ElevateObservation.getRolesAndTenantIdAndOrgIdFromUserToken(decode)
                messageArr.append("Acccess Token : " + str(accessTokenUser))
                ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                fileheader = ["Access Token","Access Token succesfully genarated","Passed"]
                ElevateObservation.apicheckslog(solutionName_for_folder_path,fileheader)
                print("--->Access Token Generated!")
                print("qqqqqqqqqqqqqqqqqqqqqqqqqqq")
            else:
                print("Error in generating Access token")
                print("Status code : " + str(responseKeyClockUser.status_code))
                print(responseKeyClockUser.text)
                errorVar = str(responseKeyClockUser.text)

            return accessTokenUser

        except Exception as e:
            errorVar = f"Error occurred: {str(e)}"
            print(errorVar, "---> API-Error")
            return accessTokenUser
    
    def getRolesAndTenantIdAndOrgIdFromUserToken(decodedToken):
        rolesInToken = decodedToken['data']['organizations'][0]['roles']
        global roleOfResourceCreator
        for element in rolesInToken:
            if element.get('title') == 'org_admin':
             roleOfResourceCreator = 'org_admin'
            elif element.get('title') == 'tenant_admin':
             roleOfResourceCreator = 'tenant_admin'

        global tenantID 
        tenantID = ElevateObservation.clean_single_value(decodedToken['data']['tenant_code'])
        print(tenantID,"306")
        global orgIDFromTemplate 
        orgIDFromTemplate = ElevateObservation.clean_single_value(decodedToken['data']['organizations'][0].get('id'))
        print(orgIDFromTemplate,"312")

    def fetchEntityType(solutionName_for_folder_path, accessToken, entitiesPGM, scopeEntityType,):
        urlFetchEntityListApi = elevateentityhost + searchforlocation
        print(urlFetchEntityListApi,"urlFetchEntityListApi")
        headerFetchEntityListApi = {
            'Content-Type': content_type,
            'internal-access-token': internal_access_token,
        }
        print(headerFetchEntityListApi,"headerFetchEntityListApi")
        # Initialize a dictionary to store entity types for each entity
        entityTypes = []
        entityTypeID =[]
        # Loop through each entity name in the entitiesPGM list
        for entityName in entitiesPGM:
            entityName = entityName.strip()  # Remove any extra spaces

            # Prepare the payload for the API request
            payload = {
                "query": {
                    "metaInformation.name": entityName,
                    "tenantId":tenantID,  # Use the current entity name
                    "entityType": scopeEntityType[0]
                    # "orgIds": {"$in":ElevateObservation.append_to_list(ElevateObservation.normalize_cell_value(orgIDFromTemplate),'ALL')},
                },
                "projection": [
                    "entityType","_id"
                ]
            }
            data = json.dumps(payload)
            print(data,"data")
            # Make the API call inside the loop to send one request per entity
            responseFetchEntityListApi = requests.post(url=urlFetchEntityListApi, headers=headerFetchEntityListApi, data=data)
            # Log API call details
            print(responseFetchEntityListApi.text,"responseFetchEntityListApi")
            messageArr = ["Entities List Fetch API executed for entity: " + entityName, 
                        "URL  : " + str(urlFetchEntityListApi),
                        "Status : " + str(responseFetchEntityListApi.status_code)]
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)

            # Check if the API call was successful
            if responseFetchEntityListApi.status_code == 200:
                responseFetchEntityListApi = responseFetchEntityListApi.json()

                # Loop through the result to find the entityType
                global entityToUpload
                entityToUpload = None  # Initialize for each entity
                for listEntities in responseFetchEntityListApi['result']:
                    entityToUpload = listEntities['entityType']
                    entityId = listEntities['_id']
                    print(entityId,"entityId")
                    # entityToUpload = listEntities.get('entityType', '').lower().strip()

                    # If a valid entityType is found, store it in the dictionary and break out of the loop
                    if entityToUpload:
                        entityTypes.append(entityToUpload)
                        # break

                    if entityId:
                        entityTypeID.append(entityId)
                        # print("Entity ID found:", entityId)

                # If no entityType is found for this entity, raise an error for that specific entity
                if not entityToUpload:
                    raise ValueError(f"Entity type not found for entity '{entityName}'.")
            else:
                # Handle cases where the API call fails for a specific entity
                raise RuntimeError(f"Failed to fetch entity type for '{entityName}'. Status code: {responseFetchEntityListApi.status_code}")
        # Return all found entity types
        return entityTypes,entityTypeID

    def fetchEntityId(solutionName_for_folder_path, accessToken, entitiesNameList, scopeEntityType,entitiesPGM):
        try:
            global errorVar
            urlFetchEntityListApi = elevateentityhost + searchforlocation
            headerFetchEntityListApi = {
                'Content-Type': content_type,
                'internal-access-token': internal_access_token,
                'origin': origin
            }
            payload = {
                "query" : {
                    "entityType": {
                        "$in": scopeEntityType
                    },
                    "metaInformation.name":
                    {
                        "$in": entitiesPGM.split(",")
                    },
                    "tenantId":tenantID ,
                    # "orgIds": {"$in":ElevateObservation.append_to_list(ElevateObservation.normalize_cell_value(orgIDFromTemplate),'ALL')}
                },

                "projection": [
                    "_id","metaInformation.name"
                ]
                }
            data=json.dumps(payload)
            print(data,"payload")
            responseFetchEntityListApi = requests.post(url=urlFetchEntityListApi, headers=headerFetchEntityListApi,data=json.dumps(payload))
            messageArr = ["Entities List Fetch API executed.", "URL  : " + str(urlFetchEntityListApi),
                        "Status : " + str(responseFetchEntityListApi.status_code)]
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)

            print(responseFetchEntityListApi.text,"responseFetchEntityListApi-------")
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
                    print("--->Scope Entity error.")
                return entityToUpload
            else:
                messageArr = ["Error in Location search",str(responseFetchEntityListApi.status_code)]
                errorvar = str(responseFetchEntityListApi.text)
                ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                print("---> Error in location search.")
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")


    def fetchEntityParentChilds(solutionName_for_folder_path, accessToken, entityIds):
        try:
            global errorVar
            if not isinstance(entityIds, list):
                entityIds = [entityIds]

            hierarchy = ["state", "district", "block", "cluster", "school"]
            merged_output = {level: [] for level in hierarchy}

            for entityId in entityIds:
                urlFetchEntity = elevateentityhost + fetchDetailsEntity + entityId
                print(urlFetchEntity, "urlFetchEntity")

                headers = {
                    'Content-Type': content_type,
                    'tenantId': tenantID,
                    'Authorization': f'Bearer {accessToken}'
                }

                response = requests.get(url=urlFetchEntity, headers=headers)

                if response.status_code == 200:
                    responseJson = response.json()
                    result = responseJson.get("result", [])[0]

                    parent_info = result.get("parentInformation", {})
                    current_entity_type = result.get("entityType").lower()
                    current_entity_id = result.get("_id")

                    current_index = hierarchy.index(current_entity_type)

                    for i in range(current_index):
                        level = hierarchy[i]
                        if level in parent_info and parent_info[level]:
                            val = parent_info[level][0]["_id"]
                            if val not in merged_output[level]:
                                merged_output[level].append(val)

                    if current_entity_id not in merged_output[current_entity_type]:
                        merged_output[current_entity_type].append(current_entity_id)

                    for i in range(current_index + 1, len(hierarchy)):
                        if not merged_output[hierarchy[i]]:  # only add ALL if empty
                            merged_output[hierarchy[i]].append("ALL")

                else:
                    errorVar = response.text
                    print(f"---> Error in fetching entity details for {entityId}. "
                        f"Status {response.status_code} Response {response.text}")

            # Final check: ensure each level has at least "ALL" if empty
            for level in hierarchy:
                if not merged_output[level]:
                    merged_output[level].append("ALL")

            print("Structured Entity Hierarchy:", json.dumps(merged_output, indent=2))
            return merged_output

        except Exception as e:
            errorVar = str(e)
            print("Error occurred:", errorVar)
            return None


    def getProgramInfo(accessToken, solutionName_for_folder_path, programNameInp):
        try:
            if programNameInp:
                global programID, programExternalId, programDescription, isProgramnamePresent, programName, isExternalProgram,tenantID,orgIdForScope
                programName = programNameInp
                programUrl = elevateprojecthost + fetchprograminfoapiurl
                print(programUrl,"programUrl")
                # print(programUrl,"payload")
                payload = json.dumps({
                    "query": {
                        "name": programNameInp.lstrip().rstrip(),
                        "isAPrivateProgram": False,
                        "status": "active",
                        "tenantId": tenantID,
                        },
                        "mongoIdKeys": []
                        })
                print(payload,"payload")
                headersProgramSearch =  {'Content-Type': 'application/json', 'X-auth-token': accessToken}
                print(headersProgramSearch,"headersProgramSearch")
                responseProgramSearch = requests.post(url=programUrl, headers=headersProgramSearch,data=payload)
                print(responseProgramSearch.text,"responseProgramSearch")
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
                    countOfPrograms = len(responseProgramSearch['result'])
                    messageArr.append("--->Program Count : " + str(countOfPrograms))
                    if countOfPrograms == 0:
                        programUrl = elevateprojecthost + fetchprograminfoapiurl
                        print(programUrl,"programUrl")
                        # print(programUrl,"payload")
                        responseProgramSearch = requests.post(url=programUrl, headers=headersProgramSearch,data=payload)
                        print(responseProgramSearch.text,"responseProgramSearch")
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
                                isExternalProgram = 'false'
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
                                            # ElevateObservation.terminatingMessage("Aborting...")
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
                                return True
                    else:
                        getProgramDetails = []
                        isExternalProgram = 'true'
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
                                    # ElevateObservation.terminatingMessage("Aborting...")
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
                        return True
                else:
                    print("Program search API failed...")
                    messageArr.append("Program search API failed...")
                    ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                    # Helpers.terminatingMessage("Response Code : " + str(responseProgramSearch.status_code))
                    errorVar = str(responseProgramSearch.text)
                    return False
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")
                

    def fetchUserDetails(environment, accessToken, dikshaId):
        global OrgName, errorVar
        error_message = ""
        try:
            decoded_token = jwt.decode(accessToken, options={"verify_signature": False}, algorithms=["HS256"])
            data=decoded_token.get('data')
            user_id = data.get('id')  
            url = userLoginHost + userinfoapiurl
            messageArr = ["User search API called."]
            headers = {# 'Content-Type': 'application/json',
                    'internal-access-token': internal_access_token,
                    'X-auth-token': accessToken}
           
            responseUserSearch = requests.request("GET", url, headers=headers)
            print(responseUserSearch.text,"responseUserSearch")
            if responseUserSearch.status_code == 200:
                responseUserSearch = responseUserSearch.json()
                if responseUserSearch['result']:
                    userKeycloak = responseUserSearch['result']['id']
                    userName = responseUserSearch['result']['name']
                    firstName = responseUserSearch['result']['name']
                    rootOrgId = responseUserSearch['result']['organizations'][0]['id']
                    roledetails = None
                    roles = responseUserSearch['result']['organizations'][0]['roles']
                    for index in roles:
                        if rootOrgId == index['organization_id']:
                            roledetails = index['title']
                    # userKeycloak = responseUserSearch['result']['id']
                    # userName = responseUserSearch['result']['name']
                    # firstName = responseUserSearch['result']['name']
                    # rootOrgId = responseUserSearch['result']['organization']['id']
                    # for index in responseUserSearch['result']['user_roles']:
                    #     if rootOrgId == index['organization_id']:
                    #         roledetails = index['title']
                    #         # rootOrgName = index['orgName']
                    #         # OrgName.append(index['orgName'])
                    # print(roledetails)
                    # sys.exit()
                    return [userKeycloak, userName, firstName,roledetails,rootOrgId]
                else:
                    print("-->Given username/email is not present in the platform<--.")
                    return False
            else:
                error_message = ""
                if responseUserSearch.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"FetchSolutionApiUrl-Client Error {responseUserSearch.status_code}: {responseUserSearch.text}"
                elif responseUserSearch.status_code in [500, 502, 503, 504]:
                    error_message = f"FetchSolutionApiUrl-Server Error {responseUserSearch.status_code}: {responseUserSearch.text}"
                else:
                    error_message = f"FetchSolutionApiUrl-Unexpected Error {responseUserSearch.status_code}: {responseUserSearch.text}"
                errorVar = error_message
                return False
        except Exception as e:
            errorVar = f"Error occurred: {str(e)}"
            print(errorVar)     
            

    def programCreation(accessToken, parentFolder, externalId, pName, pDescription, keywords, entities, roles, orgIds,creatorKeyCloakId, creatorName,entitiesPGM,mainRole,rolesPGM,entityHierarchy):
        global errorVar,scopeEntityType,orgIdForScope,mainRoles,programExternalId
        print(orgIDFromTemplate,"orgIDFromTemplate")
        print(orgIdForScope,"orgIdForScope")
        try: 
            messageArr = []
            messageArr.append("++++++++++++ Program Creation ++++++++++++")
            # program creation url 
            ProgramCreationurl = elevateprojecthost + programcreationurl
            messageArr.append("Program Creation URL : " + ProgramCreationurl)
            # print(ProgramCreationurl,"ProgramCreationurl")
            # program creation payload
            scope = {
                "organizations": orgIdForScope,
                "professional_subroles": rolesPGMID,
                "professional_role": mainRole
            }

            scope.update(entityHierarchy)
            programExternalId = externalId
            payload = json.dumps({
            "externalId": programExternalId,
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
                "recommendedFor" : roles
                },
                "requestForPIIConsent":True
                })
            # messageArr.append("Body : " + str(payload))
            headers = {'X-auth-token': accessToken,
                'internal-access-token': internal_access_token,
                'Content-Type': 'application/json',
                'Authorization':authorization,
                'tenantId': tenantID ,
                'orgid': orgIDFromTemplate,
                adminTokenHeaderName: adminAccessToken
                }
            print(headers,"headers")
            # program creation 
            responsePgmCreate = requests.request("POST", ProgramCreationurl, headers=headers, data=(payload))
            print(responsePgmCreate.text,"responsePgmCreate")
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
                # print(responsePgmCreateResp,"responsePgmCreateResp")
                print("program created successful....")
                return True
            else:
                error_message = ""
                if responsePgmCreate.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"PgmCreate-Client Error {responsePgmCreate.status_code}: {responsePgmCreate.text}"
                elif responsePgmCreate.status_code in [500, 502, 503, 504]:
                    error_message = f"PgmCreate-Server Error {responsePgmCreate.status_code}: {responsePgmCreate.text}"
                else:
                    error_message = f"PgmCreate-Unexpected Error {responsePgmCreate.status_code}: {responsePgmCreate.text}"
                errorVar = error_message
                # terminate execution
                print("Program creation API failed. Please check logs.")
                return errorVar
        except Exception as e:
            errorVar = str(e)
            print(errorVar)

    def programsFileCheck(filePathAddPgm, accessToken, parentFolder, MainFilePath):
        global errorVar, entityHierarchy,tenantID,orgIDFromTemplate,orgIdForScope
        errorVar = ""
        program_file = filePathAddPgm
        # open excel file 
        wbPgm = xlrd.open_workbook(filePathAddPgm, on_demand=True)
        global programNameInp
        sheetNames = wbPgm.sheet_names()
        # list of sheets in the program sheet 
        pgmSheets = ["Instructions", "Program Details", "Resource Details","Program Manager Details","Role-Subrole Mapping"]

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
                        if dictDetailsEnv.get('Title of the Program'):
                            programNameInp = dictDetailsEnv['Title of the Program'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "\"Title of the Program\" must not be Empty in \"Program details\" sheet"
                        
                        if dictDetailsEnv.get('Program ID'):
                            extIdPGM = dictDetailsEnv['Program ID'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "\"Program ID\" must not be Empty in \"Program details\" sheet"
                        if dictDetailsEnv.get('Description of the Program'):
                            descriptionPGM = dictDetailsEnv['Description of the Program'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "\"Keywords\" must not be Empty in \"Program details\" sheet"
                        if dictDetailsEnv.get('Keywords'):
                            keywordsPGM = dictDetailsEnv['Keywords'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "\"Keywords\" must not be Empty in \"Program details\" sheet"

                        global stateEntitiesPGM,entitiesPGM,districtEntitiesPGM,blockEntitiesPGM,clusterEntitiesPGM,schoolEntitiesPGM
                        if dictDetailsEnv.get('Targeted state at program level'):
                            stateEntitiesPGM = dictDetailsEnv['Targeted state at program level'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "\"Targeted state at program level\" must not be Empty in \"Program details\" sheet"
                        if dictDetailsEnv.get('Targeted District at program level'):
                            districtEntitiesPGM = dictDetailsEnv['Targeted District at program level'].encode('utf-8').decode('utf-8')
                        # else:
                        #     errorVar = "\"Targeted District at program level\" must not be Empty in \"Program details\" sheet"

                        if dictDetailsEnv.get('Targeted Block at program level'):
                            blockEntitiesPGM = dictDetailsEnv['Targeted Block at program level'].encode('utf-8').decode('utf-8')
                        if dictDetailsEnv.get('Targeted Cluster at program level'):
                            clusterEntitiesPGM = dictDetailsEnv['Targeted Cluster at program level'].encode('utf-8').decode('utf-8')
                        if dictDetailsEnv.get('Targeted School at program level'):
                            schoolEntitiesPGM = dictDetailsEnv['Targeted School at program level'].encode('utf-8').decode('utf-8')
                       

                        global mainRole,rolesPGMID,rolesPGM,mainRoleproff
                        mainRole = dictDetailsEnv['Targeted role at program level']
                        newProgramRole = mainRole.split(",")
                        programRoleArray = list(newProgramRole)
                        if dictDetailsEnv.get('Targeted role at program level'):
                            mainRole = dictDetailsEnv['Targeted role at program level'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "\"Targeted role at program level\" must not be Empty in \"Program details\" sheet"
                        if dictDetailsEnv.get('Targeted subrole at program level'):
                            rolesPGM = dictDetailsEnv['Targeted subrole at program level'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "\"Targeted subrole at program level\" must not be Empty in \"Program details\" sheet"

                        mainRoles = str(mainRole).strip().encode('utf-8').decode('utf-8').split(",")
                        subRoles = str(rolesPGM).strip().encode('utf-8').decode('utf-8').split(",")

                        mainRoles = [r.strip() for r in mainRoles if r.strip()]
                        subRoles = [r.strip() for r in subRoles if r.strip()]

                        verifiedRoles = ElevateObservation.validate_roles_against_api(mainRoles, subRoles)
                        print(verifiedRoles,"verifiedRoles")
                        mainRoleproff = verifiedRoles[0]
                        rolesPGMID = verifiedRoles[1]

                        print("mainRole", mainRoleproff)
                        print("rolesPGMID", rolesPGMID)
                        global startDateOfProgram, endDateOfProgram, ReffstartDateOfProgram, ReffendDateOfProgram
                        if dictDetailsEnv.get('Start date of program'):
                            startDateOfProgram = dictDetailsEnv['Start date of program']
                        else:
                            errorVar = "\"Start date of program\" must not be Empty in \"Program details\" sheet"
                        # startDateOfProgram = dictDetailsEnv['Start date of program']
                        if dictDetailsEnv.get('End date of program'):
                            endDateOfProgram = dictDetailsEnv['End date of program']
                        else:
                            errorVar = "\"End date of program\" must not be Empty in \"Program details\" sheet"
                        ReffstartDateOfProgram = dictDetailsEnv['Start date of program']
                        ReffendDateOfProgram = dictDetailsEnv['End date of program']
                        
                        startDateArr = str(startDateOfProgram).split("-")
                        startDateOfProgram = startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"

                        # taking the end date of program from program template and converting YYYY-MM-DD 00:00:00 format

                        endDateArr = str(endDateOfProgram).split("-")
                        endDateOfProgram = endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"

                        global scopeEntityType

                        if schoolEntitiesPGM:
                            entitiesPGM = schoolEntitiesPGM
                            EntityType = "school"

                        elif clusterEntitiesPGM:
                            entitiesPGM = clusterEntitiesPGM
                            EntityType = "cluster"

                        elif blockEntitiesPGM:
                            entitiesPGM = blockEntitiesPGM
                            EntityType = "block"

                        elif districtEntitiesPGM:
                            entitiesPGM = districtEntitiesPGM
                            EntityType = "district"

                        else:
                            entitiesPGM = stateEntitiesPGM
                            EntityType = "state"

                        # scopeEntityType = "state"
                        scopeEntityType = [EntityType] if isinstance(EntityType, str) else EntityType
                        print("scopeEntityType --->", scopeEntityType)

                        entitiesPGMs = entitiesPGM
                        print("entitiesPGMs", entitiesPGMs)

                        print("scopeEntityType", scopeEntityType)
                        global entitiesType
                        print(districtEntitiesPGM, "entitiesPGM")
                        entitiesType = ElevateObservation.fetchEntityType(parentFolder, accessToken,
                                                    entitiesPGMs.lstrip().rstrip().split(","), scopeEntityType)

                        print("entitiesType", entitiesType)
                        if scopeEntityType:
                            entitiesPGM = entitiesPGM
                            scopeEntityType = entitiesType[0]

                        global entitiesPGMID
                        entitiesPGMID = entitiesType[1]
                        print(entitiesPGMID,"1068")
                        # entitiesPGMID = ElevateObservation.fetchEntityId(parentFolder, accessToken,
                        #                             entitiesPGMs.lstrip().rstrip().split(","), scopeEntityType,entitiesPGM)
                        
                        entityHierarchy = ElevateObservation.fetchEntityParentChilds(parentFolder, scopeEntityType, entitiesPGMID)
                        print("fetchedhirearchy", entityHierarchy)
                        
                        print("entitiesPGMID882", entitiesPGMID)                        

                        if not ElevateObservation.getProgramInfo(accessToken, parentFolder, programNameInp.encode('utf-8').decode('utf-8')):
                            extIdPGM = dictDetailsEnv['Program ID'].encode('utf-8').decode('utf-8')
                            if str(dictDetailsEnv['Program ID']).strip() == "Do not fill this field":
                                print ("change the program id")
                            descriptionPGM = dictDetailsEnv['Description of the Program'].encode('utf-8').decode('utf-8')
                            keywordsPGM = dictDetailsEnv['Keywords'].encode('utf-8').decode('utf-8')
                            if dictDetailsEnv.get('Targeted state at program level'):
                                stateEntitiesPGM = dictDetailsEnv['Targeted state at program level'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "\"Targeted state at program level\" must not be Empty in \"Program details\" sheet"
                            if dictDetailsEnv.get('Targeted District at program level'):
                                districtEntitiesPGM = dictDetailsEnv['Targeted District at program level'].encode('utf-8').decode('utf-8')
                            # else:
                            #     errorVar = "\"Targeted District at program level\" must not be Empty in \"Program details\" sheet"

                            if dictDetailsEnv.get('Targeted Block at program level'):
                                blockEntitiesPGM = dictDetailsEnv['Targeted Block at program level'].encode('utf-8').decode('utf-8')
                            if dictDetailsEnv.get('Targeted Cluster at program level'):
                                clusterEntitiesPGM = dictDetailsEnv['Targeted Cluster at program level'].encode('utf-8').decode('utf-8')
                            if dictDetailsEnv.get('Targeted School at program level'):
                                schoolEntitiesPGM = dictDetailsEnv['Targeted School at program level'].encode('utf-8').decode('utf-8')
                                                   
                            mainRole = dictDetailsEnv['Targeted role at program level']
                            # global rolesPGM
                            rolesPGM = dictDetailsEnv['Targeted subrole at program level']
                            userDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dictDetailsEnv['Username/user id/email id/phone no. of Program Designer'])
                            # userDetails=["222","1","name","1","1"]
                            print(userDetails,"userDetails")
                            OrgName=userDetails[4]
                            # orgIds=fetchOrgId(environment, accessToken, parentFolder, OrgName)
                            creatorKeyCloakId = userDetails[0]
                            creatorName = userDetails[2]
                            
                            messageArr = []

                            scopeEntityType = entitiesType[0]
                            # fetch entity details 
                            # entitiesPGMID = ElevateObservation.fetchEntityId(parentFolder, accessToken,entitiesPGMs.lstrip().rstrip().split(","), scopeEntityType,entitiesPGM)
                            entitiesPGMID = entitiesType[1]
                            print("entitiesPGMID915", entitiesPGMID)
                            entityHierarchy = ElevateObservation.fetchEntityParentChilds(parentFolder, scopeEntityType, entitiesPGMID)
                            
                            # sys.exit()
                            # fetch sub-role details 
                            # rolesPGMID = fetchScopeRole(parentFolder, accessToken, rolesPGM.lstrip().rstrip().split(","))
                            # global rolesPGMID
                            # mainRole=mainRole.lstrip().rstrip().split(",")
                            # rolesPGMID=rolesPGM.lstrip().rstrip().split(",")
                            # sys.exit()
                            # call function to create program 
                            if not ElevateObservation.programCreation(accessToken, parentFolder, extIdPGM, programNameInp,descriptionPGM, keywordsPGM.lstrip().rstrip().split(","),entitiesPGMID, programRoleArray, orgIds, creatorKeyCloakId, creatorName,entitiesPGM, mainRoleproff, rolesPGM, entityHierarchy):
                                return False
                            # sys.exit()
                            # programmappingpdpmsheetcreation(MainFilePath, accessToken, program_file, extIdPGM,parentFolder)

                            # map PM / PD to the program 
                            # Programmappingapicall(MainFilePath, accessToken, program_file,parentFolder)

                            # check if program is created or not
                            print(programNameInp,"programNameInp") 
                            if ElevateObservation.getProgramInfo(accessToken, parentFolder, programNameInp):
                                print("Program Created SuccessFully.")
                            else :
                                print("Program creation failed! Please check logs.")
                                return False
                        else :
                            userDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dictDetailsEnv['Username/user id/email id/phone no. of Program Designer'])
                            # userDetails=["222","1","name","1","1"]
                            OrgName=userDetails[4]
                            creatorKeyCloakId = userDetails[0]
                            creatorName = userDetails[2]
                            
                            if not ElevateObservation.getProgramInfo(accessToken, parentFolder, programNameInp):
                                print("Program creation failed! Please check logs.")
                                return False

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
                        if dictDetailsEnv.get('Name of resources in program'):
                            resourceNamePGM = dictDetailsEnv['Name of resources in program'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "\"Name of resources in program\" must not be Empty in \"Program details\" sheet"

                        if dictDetailsEnv.get('Type of resources'):
                            resourceTypePGM = dictDetailsEnv['Type of resources'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "\"Type of resources\" must not be Empty in \"Program details\" sheet"

                        if dictDetailsEnv.get('Resource Link'):
                            resourceLinkOrExtPGM = dictDetailsEnv['Resource Link'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "\"Resource Link\" must not be Empty in \"Program details\" sheet"

                        if dictDetailsEnv.get('Resource Status'):
                            resourceStatusOrExtPGM = dictDetailsEnv['Resource Status']
                        else:
                            errorVar = "\"Resource Status\" must not be Empty in \"Program details\" sheet"
                        
                        global startDateOfResource, endDateOfResource
                        if dictDetailsEnv.get('Start date of resource'):
                            startDateOfResource = dictDetailsEnv['Start date of resource']
                        else:
                            errorVar = "\"Start date of resource\" must not be Empty in \"Program details\" sheet"
                        # startDateOfResource = dictDetailsEnv['Start date of resource']
                        if dictDetailsEnv.get('End date of resource'):
                            endDateOfResource = dictDetailsEnv['End date of resource']
                        else:
                            errorVar = "\"End date of resource\" must not be Empty in \"Program details\" sheet"
                        
                        if errorVar == "":
                            return True
                        else:
                            return False
            
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
        # project_sheet_names = ['Instructions', 'Project upload', 'Tasks upload','Certificate details']

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
        # elif (len(project_sheet_names) == len(sheetNames1)) and ((set(project_sheet_names) == set(sheetNames1))):
        #     print("--->Project file detected.<---")
        #     typeofSolution = 4
        elif (len(rubrics_sheet_IMP_names) == len(sheetNames1)) and ((set(rubrics_sheet_IMP_names) == set(sheetNames1))):
            print("--->Observation with rubrics and IMP file detected.<---")
            typeofSolution = 5
        else:
            typeofSolution = 0
            print(typeofSolution)
            errorVar = ("Please check the Input sheet.")
        return typeofSolution
    
    def criteriaUpload(solutionName_for_folder_path, wbObservation, millisAddObs, accessToken, tabName, projectDrivenFlag):
        global errorVar,criteriaName
        error_message = ""
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
        try:
            urlCriteriaUploadApi = internal_kong_ip + criteriauploadapiurl
            headerCriteriaUploadApi = {
                "internal-access-token": internal_access_token,
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'tenantId': tenantID ,
                'orgid': orgIDFromTemplate,
                adminTokenHeaderName: adminAccessToken
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
                return True
            else:
                error_message = ""
                if responseCriteriaUploadApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"CriteriaUploadApi-Client Error {responseCriteriaUploadApi.status_code}: {responseCriteriaUploadApi.text}"
                elif responseCriteriaUploadApi.status_code in [500, 502, 503, 504]:
                    error_message = f"CriteriaUploadApi-Server Error {responseCriteriaUploadApi.status_code}: {responseCriteriaUploadApi.text}"
                else:
                    error_message = f"CriteriaUploadApi-Unexpected Error {responseCriteriaUploadApi.status_code}: {responseCriteriaUploadApi.text}"
                errorVar = error_message
                messageArr.append("Response : " + str(responseCriteriaUploadApi.text))
                # errorVar = str(responseCriteriaUploadApi.text)
                ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                print("Criteria Upload failed.")
                return False
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")
    
    def frameWorkUpload(solutionName_for_folder_path, wbObservation, millisAddObs, accessToken):
        global criteriaLevelsReport,errorVar,pointBasedValue,entityType,solutionLanguage,keyWords,entityTypeId
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
        try:
            urlCreateFrameworkApi = internal_kong_ip + frameworkcreationapi
            frameworkFilePath = solutionName_for_folder_path + '/framework/'
            file_exists_framework = os.path.isfile(solutionName_for_folder_path + '/framework/uploadFile.json')
            if not os.path.exists(frameworkFilePath):
                os.mkdir(frameworkFilePath)

            with open(frameworkFilePath + "uploadFile.json", "w",encoding='utf-8') as outfile:
                json.dump(frameworkDocInsertObj, outfile)
            headerFrameworkUploadApi = {'Authorization': authorization,
                                        "internal-access-token": internal_access_token,
                                        'X-auth-token': accessToken,
                                        'X-Channel-id': x_channel_id,
                                        'tenantId': tenantID ,
                                        'orgid': orgIDFromTemplate,
                                        adminTokenHeaderName: adminAccessToken}
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
                error_message = ""
                if responseFrameworkUploadApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"FrameworkUploadApi-Client Error {responseFrameworkUploadApi.status_code}: {responseFrameworkUploadApi.text}"
                elif responseFrameworkUploadApi.status_code in [500, 502, 503, 504]:
                    error_message = f"FrameworkUploadApi-Server Error {responseFrameworkUploadApi.status_code}: {responseFrameworkUploadApi.text}"
                else:
                    error_message = f"FrameworkUploadApi-Unexpected Error {responseFrameworkUploadApi.status_code}: {responseFrameworkUploadApi.text}"
                errorVar = error_message
                messageArr = ["Framwork upload Failed.", "Response : " + responseFrameworkUploadApi.text]
                # errorVar = str(responseFrameworkUploadApi.text)
                ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                print('Framework upload api failed in ',
                    'status_code response from api is ' + str(responseFrameworkUploadApi.status_code))
                return False
                
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")
            
    def themesUpload(solutionName_for_folder_path, wbObservation, millisAddObs, accessToken, frameworkExternalId,obsWORubWS):
        global dictCritLookUp,errorVar
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
        try:
            urlThemesUploadApi = internal_kong_ip + themeuploadapiurl + frameworkExternalId
            headerThemesUploadApi = {'Authorization': authorization,
                                    "internal-access-token": internal_access_token,
                                    'X-auth-token': accessToken,
                                    'X-Channel-id': x_channel_id,
                                    'tenantId': tenantID ,
                                    'orgid': orgIDFromTemplate,
                                    adminTokenHeaderName: adminAccessToken}
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
                return True
            else:
                error_message = ""
                if responseThemeUploadApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"ThemeUploadApi-Client Error {responseThemeUploadApi.status_code}: {responseThemeUploadApi.text}"
                elif responseThemeUploadApi.status_code in [500, 502, 503, 504]:
                    error_message = f"ThemeUploadApi-Server Error {responseThemeUploadApi.status_code}: {responseThemeUploadApi.text}"
                else:
                    error_message = f"ThemeUploadApi-Unexpected Error {responseThemeUploadApi.status_code}: {responseThemeUploadApi.text}"
                errorVar = error_message
                messageArr = ["Themes upload failed.", "Response : " + str(responseThemeUploadApi.text)]
                # errorVar = str(responseThemeUploadApi.text)
                ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                print("Theme upload failed.")
                return False
                # sys.exit()
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")

    def createSolutionFromFramework(solutionName_for_folder_path, accessToken, frameworkExternalId):
        global errorVar,entityType,solutionId,isExternalProgram
        error_message = ""
        try:
            urlCreateSolutionApi = internal_kong_ip + solutioncreationapiurl
            headerCreateSolutionApi = {
                'Content-Type': content_type,
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantID ,
                'orgid': orgIDFromTemplate,
                adminTokenHeaderName: adminAccessToken
            }
            queryparamsCreateSolutionApi = '?frameworkId=' + str(frameworkExternalId) + '&entityType=' + entityType + '&isExternalProgram=' + isExternalProgram
            # queryparamsCreateSolutionApi = '?frameworkId=' + str(frameworkExternalId) + '&entityType=' + entityType
            print(queryparamsCreateSolutionApi)
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
                return solutionId
            else:
                error_message = ""
                if responseCreateSolutionApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"CreateSolutionApi-Client Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}"
                elif responseCreateSolutionApi.status_code in [500, 502, 503, 504]:
                    error_message = f"CreateSolutionApi-Server Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}"
                else:
                    error_message = f"CreateSolutionApi-Unexpected Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}"
                errorVar = error_message
                messageArr.append("Solution from framework api failed.")
                ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                # errorVar = str(responseCreateSolutionApi.text)
                print("Solution from framework api failed.")
                return False
                # sys.exit()
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")
            

    def solutionUpdate(solutionName_for_folder_path, accessToken, solutionId, bodySolutionUpdate):
        global errorVar
        error_message = ""
        try:
            solutionUpdateApi = internal_kong_ip + solutionupdateapi + str(solutionId)
            headerUpdateSolutionApi = {
                'Content-Type': 'application/json',
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                "internal-access-token": internal_access_token,
                'tenantId': tenantID ,
                'orgid': orgIDFromTemplate,
                adminTokenHeaderName: adminAccessToken
                }
            responseUpdateSolutionApi = requests.post(url=solutionUpdateApi, headers=headerUpdateSolutionApi,data=json.dumps(bodySolutionUpdate))
            messageArr = ["Solution Update API called.", "URL : " + str(solutionUpdateApi), "Body : " + str(bodySolutionUpdate),"Response : " + str(responseUpdateSolutionApi.text),"Status Code : " + str(responseUpdateSolutionApi.status_code)]
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            if responseUpdateSolutionApi.status_code == 200:
                print("Solution Update Success.")
                return True
            else:
                error_message = ""
                if responseUpdateSolutionApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"UpdateSolutionApi-Client Error {responseUpdateSolutionApi.status_code}: {responseUpdateSolutionApi.text}"
                elif responseUpdateSolutionApi.status_code in [500, 502, 503, 504]:
                    error_message = f"UpdateSolutionApi-Server Error {responseUpdateSolutionApi.status_code}: {responseUpdateSolutionApi.text}"
                else:
                    error_message = f"UpdateSolutionApi-Unexpected Error {responseUpdateSolutionApi.status_code}: {responseUpdateSolutionApi.text}"
                errorVar = error_message
                ElevateObservation.createAPILog(solutionName_for_folder_path, errorVar)
                return False
            
        except Exception as e:
            errorVar = error_message
            ElevateObservation.createAPILog(solutionName_for_folder_path, [f"Exception: {str(e)}"])
    
    def questionUpload(filePathAddObs, solutionName_for_folder_path, frameworkExternalId, millisAddObs, accessToken,
                   solutionId, typeofSolution):
        global errorVar, pointBasedValue
        error_message = ""
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
        if not ElevateObservation.solutionUpdate(solutionName_for_folder_path, accessToken, solutionId, bodySolutionUpdate):
            return False
        try:
            urlQuestionsUploadApi = internal_kong_ip + questionuploadapiurl
            headerQuestionUploadApi = {'Authorization': authorization,
                                       "internal-access-token": internal_access_token,
                                    'X-auth-token': accessToken,
                                    'X-Channel-id': x_channel_id,
                                    'tenantId': tenantID ,
                                    'orgid': orgIDFromTemplate,
                                    adminTokenHeaderName: adminAccessToken
                                    }
            filesQuestion = {
                'questions': open(solutionName_for_folder_path + '/questionUpload/uploadSheet.csv', 'rb')
            }
            responseQuestionUploadApi = requests.post(url=urlQuestionsUploadApi, headers=headerQuestionUploadApi,
                                                    files=filesQuestion)
            print(responseQuestionUploadApi.text,"responseQuestionUploadApi")
            messageArr = ["Question Upload sheet prepared.",
                        "File loc : " + solutionName_for_folder_path + '/questionUpload/uploadSheet.csv',
                        "Question upload API called.", "Status code : " + str(responseQuestionUploadApi.status_code)]
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            if responseQuestionUploadApi.status_code == 200:
                print('QuestionUploadApi Success')
                with open(solutionName_for_folder_path + '/questionUpload/uploadInternalIdsSheet.csv','w+',
                        encoding='utf-8') as questionRes:
                    questionRes.write(responseQuestionUploadApi.text)
                return True
            else:
                    error_message = ""
                    if responseQuestionUploadApi.status_code in [400, 401, 403, 404, 422]:
                        error_message = f"QuestionUploadApi-Client Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}"
                    elif responseQuestionUploadApi.status_code in [500, 502, 503, 504]:
                        error_message = f"QuestionUploadApi-Server Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}"
                    else:
                        error_message = f"QuestionUploadApi-Unexpected Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}"
                    errorVar = error_message
                    messageArr = ["Question Upload Failed.", "Response : " + str(responseQuestionUploadApi.text)]
                    ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                    # errorVar = str(responseQuestionUploadApi.text)
                    print("Question Upload failed.")
                    return False
                    # sys.exit()
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")

    def fetchSolutionCriteria(solutionName_for_folder_path, observationId, accessToken):
        global errorVar
        error_message = ""
        try:
            url = internal_kong_ip + ferchsolutioncriteria + observationId

            headers = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'internal-access-token': internal_access_token,
                'tenantId': tenantID ,
                'orgid': orgIDFromTemplate,
                adminTokenHeaderName: adminAccessToken
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
                return True
            else:
                error_message = ""
                if response.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"QuestionUploadApi-Client Error {response.status_code}: {response.text}"
                elif response.status_code in [500, 502, 503, 504]:
                    error_message = f"QuestionUploadApi-Server Error {response.status_code}: {response.text}"
                else:
                    error_message = f"QuestionUploadApi-Unexpected Error {response.status_code}: {response.text}"
                errorVar = error_message
                messageArr = ["Criteria solution fetch API failed.", "Response  : " + str(response.text)]
                ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                # errorVar = str(response.text)
                print("Solution criteria fetch failed. Status Code : " + str(response.status_code))
                return False
                # sys.exit()
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")

    def uploadCriteriaRubrics(solutionName_for_folder_path, wbObservation, millisAddObs, accessToken, frameworkExternalId,
                          withRubricsFlag):
        global errorVar
        error_message = ""
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
        try:
            urlCriteriaRubricUploadApi = internal_kong_ip + criteriarubricuploadapiurl + frameworkExternalId + "-OBSERVATION-TEMPLATE"
            headerCriteriaRubricUploadApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                "internal-access-token": internal_access_token,
                'tenantId': tenantID ,
                'orgid': orgIDFromTemplate,
                adminTokenHeaderName: adminAccessToken
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
                return True
            else:
                error_message = ""
                if responseCriteriaRubricUploadApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"CriteriaRubricUploadApi-Client Error {responseCriteriaRubricUploadApi.status_code}: {responseCriteriaRubricUploadApi.text}"
                elif responseCriteriaRubricUploadApi.status_code in [500, 502, 503, 504]:
                    error_message = f"CriteriaRubricUploadApi-Server Error {responseCriteriaRubricUploadApi.status_code}: {responseCriteriaRubricUploadApi.text}"
                else:
                    error_message = f"CriteriaRubricUploadApi-Unexpected Error {responseCriteriaRubricUploadApi.status_code}: {responseCriteriaRubricUploadApi.text}"
                errorVar = error_message
                messageArr = ["Criteria Rubric upload Failed.", "Response : " + str(responseCriteriaRubricUploadApi.text)]
                ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                errorVar = str(responseCriteriaRubricUploadApi.text)
                print("Criteria Rubric upload Failed.")
                return False
                # sys.exit()
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")

    def uploadThemeRubrics(solutionName_for_folder_path, wbObservation, accessToken, frameworkExternalId, withRubricsFlag):
        global errorVar,criteriaLevels
        error_message = ""
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
                print(criteriaLevels,"criteriaLevels")
                for cl in criteriaLevels:
                    themeRubricUploadFieldnames.append("L" + str(cl))
            else:
                themeRubricUploadFieldnames.append("L1")
            print(themeRubricUploadFieldnames,"themeRubricUploadFieldnames")
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
        try:
            urlThemeRubricUploadApi = internal_kong_ip + themerubricuploadapiurl + frameworkExternalId + "-OBSERVATION-TEMPLATE"
            headerThemeRubricUploadApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantID ,
                'orgid': orgIDFromTemplate,
                adminTokenHeaderName: adminAccessToken
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
                return True
            else:
                error_message = ""
                if responseThemeRubricUploadApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"ThemeRubricUploadApi-Client Error {responseThemeRubricUploadApi.status_code}: {responseThemeRubricUploadApi.text}"
                elif responseThemeRubricUploadApi.status_code in [500, 502, 503, 504]:
                    error_message = f"ThemeRubricUploadApi-Server Error {responseThemeRubricUploadApi.status_code}: {responseThemeRubricUploadApi.text}"
                else:
                    error_message = f"ThemeRubricUploadApi-Unexpected Error {responseThemeRubricUploadApi.status_code}: {responseThemeRubricUploadApi.text}"
                errorVar = error_message
                messageArr = ['theme rubric upload api failed in ' + environment,
                            ' status_code response from api is ' + str(responseThemeRubricUploadApi.status_code),
                            "Response : " + str(responseThemeRubricUploadApi.text)]
                ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                # errorVar = str(responseThemeRubricUploadApi.text)
                print('theme rubric upload api failed in ' + environment + ' status_code response from api is ' + str(responseThemeRubricUploadApi.status_code))
                return False
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")
    
    def fetchSolutionDetailsFromProgramSheet(solutionName_for_folder_path, programFile, solutionId, accessToken):
        global solutionRolesArray, solutionStartDate, solutionEndDate, errorVar
        error_message = ""
        try:
            urlFetchSolutionApi = internal_kong_ip + fetchsolutiondoc + solutionId
            headerFetchSolutionApi = {
                'Content-Type': 'application/json',
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantID ,
                'orgid': orgIDFromTemplate,
                adminTokenHeaderName: adminAccessToken
            }
            payloadFetchSolutionApi = {}
            responseFetchSolutionApiUrl = requests.post(url=urlFetchSolutionApi, headers=headerFetchSolutionApi,
                                                    data=payloadFetchSolutionApi)
            print(responseFetchSolutionApiUrl.text,'responseFetchSolutionApiUrl')
            responseFetchSolutionJson = responseFetchSolutionApiUrl.json()
            messageArr = ["Solution Fetch Link.",
                        "solution name : " + responseFetchSolutionJson["result"]["name"],
                        "solution ExternalId : " + responseFetchSolutionJson["result"]["externalId"]]
            messageArr.append("Upload status code : " + str(responseFetchSolutionApiUrl.status_code))
            ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
            if responseFetchSolutionApiUrl.status_code == 200:
                solutionName = responseFetchSolutionJson["result"]["name"]
                print(solutionName,"solutionName")
                xfile = openpyxl.load_workbook(programFile)
                sheet_name = 'Resource Details'.strip()
                resourceDetailsSheet = xfile[sheet_name]
                rowCountRD = resourceDetailsSheet.max_row
                columnCountRD = resourceDetailsSheet.max_column
                for row in range(3, rowCountRD + 1):
                    cell_value = resourceDetailsSheet["A" + str(row)].value
                    if cell_value is not None and str(cell_value).strip() == str(solutionName).strip():
                        solutionNameCell = resourceDetailsSheet[f"A{row}"].value
                        if resourceDetailsSheet["A" + str(row)].value == solutionName:
                            solutionMainRole = str(resourceDetailsSheet["E" + str(row)].value).split(",")
                            solutionRolesArray = str(resourceDetailsSheet["F" + str(row)].value).split(",")
                            solutionStartDate = resourceDetailsSheet["G" + str(row)].value
                            solutionEndDate = resourceDetailsSheet["H" + str(row)].value
                            return [solutionMainRole,solutionRolesArray, solutionStartDate, solutionEndDate]
                
            else:
                error_message = ""
                if responseFetchSolutionApiUrl.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"FetchSolutionApiUrl-Client Error {responseFetchSolutionApiUrl.status_code}: {responseFetchSolutionApiUrl.text}"
                elif responseFetchSolutionApiUrl.status_code in [500, 502, 503, 504]:
                    error_message = f"FetchSolutionApiUrl-Server Error {responseFetchSolutionApiUrl.status_code}: {responseFetchSolutionApiUrl.text}"
                else:
                    error_message = f"FetchSolutionApiUrl-Unexpected Error {responseFetchSolutionApiUrl.status_code}: {responseFetchSolutionApiUrl.text}"
                errorVar = error_message
                print(error_message)
                print(errorVar)
                return False
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            print("-----> API-Error",errorVar)
            return False

    def fetchSolutionDetailsFromResourceSheet(solutionName_for_folder_path, programFile, solutionId, accessToken,typeofSolution):
        global solutionRolesArray, solutionStartDate, solutionEndDate
        urlFetchSolutionApi = internal_kong_ip + fetchsolutiondoc + solutionId
        
        headerFetchSolutionApi = {
            'Content-Type': 'application/json',
            'Authorization': authorization,
            'X-auth-token': accessToken,
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
        global errorVar,solutionName, solutionDescription,entityType,programExternalId,isExternalProgram
        error_message=""
        try:
            childObservationExternalId = str(observationExternalId + "_CHILD")
            urlSol_prog_mapping = internal_kong_ip + solutiontoprogrammappingapiurl + "?solutionId=" + observationExternalId + "&entityType=" + entityType
            print(urlSol_prog_mapping,"urlSol_prog_mapping")
            if isExternalProgram == 'true':
                payloadSol_prog_mapping = {
                    "externalId": childObservationExternalId,
                    "name": solutionName.lstrip().rstrip(),
                    "description": solutionDescription.lstrip().rstrip(),
                    "programExternalId": programID
                }
            else:
                payloadSol_prog_mapping = {
                    "externalId": childObservationExternalId,
                    "name": solutionName.lstrip().rstrip(),
                    "description": solutionDescription.lstrip().rstrip(),
                    "programExternalId": programExternalId
                }
            print(payloadSol_prog_mapping,"payloadSol_prog_mapping")
            headersSol_prog_mapping = {'Authorization': authorization,
                                    'X-auth-token': accessToken,
                                    'Content-Type': content_type,
                                    'internal-access-token': internal_access_token,
                                    'tenantId': tenantID ,
                                    'orgid': orgIDFromTemplate,
                                    adminTokenHeaderName: adminAccessToken
                                    }
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
                error_message = ""
                if responseSol_prog_mapping.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"Sol_prog_mapping-Client Error {responseSol_prog_mapping.status_code}: {responseSol_prog_mapping.text}"
                elif responseSol_prog_mapping.status_code in [500, 502, 503, 504]:
                    error_message = f"Sol_prog_mapping-Server Error {responseSol_prog_mapping.status_code}: {responseSol_prog_mapping.text}"
                else:
                    error_message = f"Sol_prog_mapping-Unexpected Error {responseSol_prog_mapping.status_code}: {responseSol_prog_mapping.text}"
                errorVar = error_message
                print("Unable to create child solution")
                # errorVar = str(responseSol_prog_mapping.text)
                messageArr.append("Unable to create child solution")
                ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                return False
                # return False
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")
    
    def prepareProgramSuccessSheet(MainFilePath, solutionName_for_folder_path, programFile, solutionExternalId, solutionId,accessToken):
        global errorVar
        error_message = ""
        try: 
            urlFetchSolutionApi = internal_kong_ip + fetchsolutiondoc + solutionId
            headerFetchSolutionApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantID ,
                'orgid': orgIDFromTemplate,
                adminTokenHeaderName: adminAccessToken
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
                solutionLink = "https: Deeplink Solution created successfully."
                return solutionLink
                # urlFetchSolutionLinkApi = internal_kong_ip + fetchlink + solutionId
                # headerFetchSolutionLinkApi = {
                #     'Authorization': authorization,
                #     'X-auth-token': accessToken,
                #     'X-Channel-id': x_channel_id,
                #     'internal-access-token': internal_access_token
                # }
                # payloadFetchSolutionLinkApi = {}

                # responseFetchSolutionLinkApi = requests.get(url=urlFetchSolutionLinkApi, headers=headerFetchSolutionLinkApi,
                #                                             data=payloadFetchSolutionLinkApi)

                # messageArr = ["Solution Fetch Link.","solution id : " + solutionId,"solution ExternalId : " + solutionExternalId]
                # messageArr.append("Upload status code : " + str(responseFetchSolutionLinkApi.status_code))
                # ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                # print(responseFetchSolutionLinkApi,"responseFetchSolutionLinkApi")
                # if responseFetchSolutionLinkApi.status_code == 200:
                #     print('Fetch solution Link Api Success')
                #     responseProjectUploadJson = responseFetchSolutionLinkApi.json()
                #     ActualsolutionLink = responseProjectUploadJson["result"]
                #     messageArr.append("Response : " + str(responseFetchSolutionLinkApi.text))
                #     ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)

                #     if os.path.exists(MainFilePath + "/" + str(programFile).replace(".xlsx", "") + '-SuccessSheet.xlsx'):
                #         xfile = openpyxl.load_workbook(
                #             MainFilePath + "/" + str(programFile).replace(".xlsx", "") + '-SuccessSheet.xlsx')
                #     else:
                #         xfile = openpyxl.load_workbook(programFile)
                #     print(xfile.sheetnames)

                #     #sheet_name = 'details'.strip()
                #     sheet_name_primary = 'Resource Details'.strip()
                #     sheet_name_fallback = 'details'.strip()

                #     try:
                #         resourceDetailsSheet = xfile[sheet_name_primary]
                #     except KeyError:
                #         resourceDetailsSheet = xfile[sheet_name_fallback]

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
                #     solutionLink = "Solution created successfully."
                #     return solutionLink
                # else:
                #     print("Fetch solution link API Failed")
                #     error_message = ""
                #     if responseFetchSolutionLinkApi.status_code in [400, 401, 403, 404, 422]:
                #         error_message = f"FetchSolutionLinkApi-Client Error {responseFetchSolutionLinkApi.status_code}: {responseFetchSolutionLinkApi.text}"
                #     elif responseFetchSolutionLinkApi.status_code in [500, 502, 503, 504]:
                #         error_message = f"FetchSolutionLinkApi-Server Error {responseFetchSolutionLinkApi.status_code}: {responseFetchSolutionLinkApi.text}"
                #     else:
                #         error_message = f"FetchSolutionLinkApi-Unexpected Error {responseFetchSolutionLinkApi.status_code}: {responseFetchSolutionLinkApi.text}"
                #     errorVar = error_message
                #     messageArr.append("Response : " + str(responseFetchSolutionLinkApi.text))
                #     ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                #     return False
            else:
                print("Fetch solution link API Failed")
                error_message = ""
                if responseFetchSolutionApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"FetchSolutionLinkApi-Client Error {responseFetchSolutionApi.status_code}: {responseFetchSolutionApi.text}"
                elif responseFetchSolutionApi.status_code in [500, 502, 503, 504]:
                    error_message = f"FetchSolutionLinkApi-Server Error {responseFetchSolutionApi.status_code}: {responseFetchSolutionApi.text}"
                else:
                    error_message = f"FetchSolutionLinkApi-Unexpected Error {responseFetchSolutionApi.status_code}: {responseFetchSolutionApi.text}"
                errorVar = error_message
                messageArr.append("Response : " + str(responseFetchSolutionApi.text))
                ElevateObservation.createAPILog(solutionName_for_folder_path, messageArr)
                return False
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")

    def check_sequence(arr):
        for i in range(1, len(arr)):
            if arr[i] != arr[i - 1] + 1:
                return False
        return True
    
    def assignTenantOrgValuesToGlobalVariables(tenantIdFromTheSheets, orgIdsFromTheSheets):
        print(orgIdsFromTheSheets,"orgIdsFromTheSheets--------------")
        print("swaping tenantId")
        global tenantID 
        tenantID = ElevateObservation.clean_single_value(tenantIdFromTheSheets)
        global orgIDFromTemplate 
        orgIDFromTemplate = ElevateObservation.clean_single_value(orgIdsFromTheSheets)
        print(orgIDFromTemplate,"orgIDFromTemplate--------------")

    def validateTenantAndOrgIdsFromProgramSheet(programFileContent):
            
            tenantIdFromProgramFile = None
            orgIdsFromProgramFile = []
                        
            sheetNames = programFileContent.sheet_names()
            # iterate through the sheets 
            for sheetEnv in sheetNames:

                if sheetEnv == "Instructions":
                    # skip Instructions sheet 
                    pass
                elif sheetEnv.strip().lower() == 'program details':
                    print("--->Checking Program details sheet...")
                    detailsEnvSheet = programFileContent.sheet_by_name(sheetEnv)
                    keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                            range(detailsEnvSheet.ncols)]
                    for row_index_env in range(2, detailsEnvSheet.nrows):
                        dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                        for
                                        col_index_env in range(detailsEnvSheet.ncols)}
                        tenantIdFromProgramFile = dictDetailsEnv.get('Tenant ID')
                        # orgIdsFromProgramFile = dictDetailsEnv.get('Org ID')
                        global orgIdForScope
                        if tenantIdFromProgramFile == "shikshalokam":
                            orgIds_str = dictDetailsEnv.get('Org ID', '')
                            orgIds = [oid.strip() for oid in orgIds_str.split(',') if oid.strip()]
                            orgIdForScope = orgIds
                            orgIdsFromProgramFile = orgIds[0] if orgIds else None
                        else:
                            # if tenantIdFromProgramFile == "shikshagrahanew":
                            orgIds_str = dictDetailsEnv.get('Targeted state at program level', '')
                            orgIds = [oid.strip().lower() for oid in orgIds_str.split(',') if oid.strip()]
                            orgIdForScope = orgIds
                            orgIdsFromProgramFile = orgIds[0] if orgIds else None

            global roleOfResourceCreator
            if roleOfResourceCreator not in ['org_admin', 'tenant_admin'] and not tenantIdFromProgramFile:
                raise ValueError("Tenant ID is required in program template for role 'admin', it cannot be empty")

            # if roleOfResourceCreator not in ['org_admin'] and not orgIdsFromProgramFile:
            #     raise ValueError("Org ID is required for role 'admin' and 'tenant_admin' in program template and cannot be empty")
            
            ElevateObservation.assignTenantOrgValuesToGlobalVariables(tenantIdFromProgramFile, orgIdsFromProgramFile)

    def validateTenantAndOrgIdsFromResourceSheet(resourceFileContent):
            print('validating resourceFileconetnt .....')
            tenantIdFromresourceFile = None
            orgIdsFromresourceFile = None
                        
            sheetNames1 = resourceFileContent.sheet_names()
            for sheetEnv in sheetNames1:
                if sheetEnv.strip().lower() == 'details':
                    detailsEnvSheet = resourceFileContent.sheet_by_name(sheetEnv)
                    keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                        range(detailsEnvSheet.ncols)]

                    for row_index_env in range(2, detailsEnvSheet.nrows):
                        dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                    for
                                    col_index_env in range(detailsEnvSheet.ncols)}
                        tenantIdFromresourceFile = dictDetailsEnv.get('Tenant ID')
                        orgIdsFromresourceFile = dictDetailsEnv.get('Targeted District at program level')

            global roleOfResourceCreator
            if roleOfResourceCreator not in ['org_admin', 'tenant_admin'] and not tenantIdFromresourceFile:
                raise ValueError("Tenant ID is required in program template for role 'admin', it cannot be empty")

            if roleOfResourceCreator not in ['org_admin'] and not orgIdsFromresourceFile:
                raise ValueError("Org ID is required for role 'admin' and 'tenant_admin' in program template and cannot be empty")
            
            ElevateObservation.assignTenantOrgValuesToGlobalVariables(tenantIdFromresourceFile, orgIdsFromresourceFile)

    def ObsWRValidate(wbObservation1, accessToken, parentFolder,typeofSolution):
        print("Validating Observation temp....")
        global errorVar, entityType, solutionName, solutionDescription, scopeEntityType, dikshaLoginId, pointBasedValue, criteriaLevels,allow_multiple_submissions,creator, question_sequence_arr,solutionLanguage, keyWords
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
                        detailsCols = ["observation_solution_name", "observation_solution_description", "Username/user id/email id/phone no. of the Content creator","Name_of_the_creator", "language", "allow_multiple_submissions", "keywords","scoring_system", "entity_type"]
                        detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                        keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                                range(detailsEnvSheet.ncols)]
                        for row_index_env in range(2, detailsEnvSheet.nrows):
                            dictDetailsEnv = {
                                keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                col_index_env in range(detailsEnvSheet.ncols)}
                            if set(detailsCols) == set(dictDetailsEnv.keys()):
                                if dictDetailsEnv['observation_solution_name']:
                                    solutionName = dictDetailsEnv['observation_solution_name'].encode('utf-8').decode('utf-8')
                                else:
                                    errorVar = "validation failed :observation_solution_name column must not be Empty in details sheet"
                                # solutionName = dictDetailsEnv['observation_solution_name'].encode('utf-8').decode('utf-8') if dictDetailsEnv['observation_solution_name'] else ElevateObservation.terminatingMessage("\"observation_solution_name\" must not be Empty in \"details\" sheet")
                                if dictDetailsEnv['Username/user id/email id/phone no. of the Content creator']:
                                    dikshaLoginId = dictDetailsEnv['Username/user id/email id/phone no. of the Content creator'].encode('utf-8').decode('utf-8')
                                else:
                                    errorVar = "validation failed :Username/user id/email id/phone no. of the Content creator column must not be Empty in details sheet"
                                # dikshaLoginId = dictDetailsEnv['Elevate_loginId'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Elevate_loginId'] else ElevateObservation.terminatingMessage("\"Elevate_loginId\" must not be Empty in \"details\" sheet")
                                # ccUserDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dikshaLoginId)
                                # if not "CONTENT_CREATOR" in ccUserDetails[3]:
                                #     terminatingMessage("---> "+dikshaLoginId +" is not a CONTENT_CREATOR in Diksha " + environment)
                                # ccRootOrgName = ccUserDetails[4]
                                # ccRootOrgId = ccUserDetails[5]
                                if dictDetailsEnv['observation_solution_description']:
                                    solutionDescription = dictDetailsEnv['observation_solution_description'].encode('utf-8').decode('utf-8')
                                else:
                                    errorVar = "validation failed :observation_solution_description column must not be Empty in details sheet"
                                if dictDetailsEnv['Name_of_the_creator']:
                                    creator = dictDetailsEnv['Name_of_the_creator'].encode('utf-8').decode('utf-8')
                                else:
                                    errorVar = "validation failed :Name_of_the_creator column must not be Empty in details sheet"
                                if dictDetailsEnv['language']:
                                    solutionLanguage = dictDetailsEnv['language'].encode('utf-8').decode('utf-8')
                                else:
                                    errorVar = "validation failed :language column must not be Empty in details sheet"
                                if dictDetailsEnv['keywords']:
                                    keywords = dictDetailsEnv['keywords'].encode('utf-8').decode('utf-8')
                                else:
                                    errorVar = "validation failed :keywords column must not be Empty in details sheet"
                                if dictDetailsEnv['entity_type']:
                                    entityType = dictDetailsEnv['entity_type'].encode('utf-8').decode('utf-8')
                                else:
                                    errorVar = "validation failed :entity_type column must not be Empty in details sheet"
                                if dictDetailsEnv['scoring_system']:
                                    pointBasedValue = dictDetailsEnv['scoring_system'].encode('utf-8').decode('utf-8')
                                else :
                                    errorVar = "\"scoring_system\" must not be Empty in \"details\" sheet"
                                # solutionDescription = dictDetailsEnv['observation_solution_description'].encode('utf-8').decode('utf-8')
                                # pointBasedValue = str(dictDetailsEnv['scoring_system']).encode('utf-8').decode('utf-8') if dictDetailsEnv['scoring_system'] else ElevateObservation.terminatingMessage("\"scoring_system\" must not be Empty in \"details\" sheet")
                                # entityType = dictDetailsEnv['entity_type'].encode('utf-8').decode('utf-8') if dictDetailsEnv['entity_type'] else ElevateObservation.terminatingMessage("\"entity_type\" must not be Empty in \"details\" sheet")
                                # solutionLanguage = dictDetailsEnv['language'].split(",") if dictDetailsEnv['language'] else [""]
                                # keyWords = dictDetailsEnv['keywords'].encode('utf-8').decode('utf-8')
                                # creator = dictDetailsEnv['Name_of_the_creator'].encode('utf-8').decode('utf-8')  if dictDetailsEnv['Name_of_the_creator'] else ElevateObservation.terminatingMessage("\"Name_of_the_creator\" must not be Empty in \"details\" sheet")
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
                                errorVar = "--->Columns Mismatch in Details Sheet."
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
                            print(criteriaLevels,"criteriaLevels")

                            if dictDetailsEnv['Criteria ID'].encode('utf-8').decode('utf-8'):
                                if not [dictDetailsEnv['Domain ID'], dictDetailsEnv['Criteria ID']] in listOfThemeCriteria:
                                    listOfThemeCriteria.append([dictDetailsEnv['Domain ID'], dictDetailsEnv['Criteria ID']])
                                else:
                                    errorVar = "Theme , criteria combo repeating in framework sheet."
                            if not dictDetailsEnv['Domain ID']:
                                errorVar ="Domain ID cannot be empty in framework sheet."
                            if not dictDetailsEnv['Domain Name']:
                                errorVar = "Theme cannot be empty in framework sheet."

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
                                errorVar = "ECM Id/Domian ID cannot be empty in ecm\'s sheet."
                            if not dictDetailsEnv['section_id']:
                                errorVar = "section_id cannot be empty in ecm\'s sheet."
                            if not dictDetailsEnv['section_name']:
                                errorVar = "section_name cannot be empty in ecm\'s sheet."
                            if not dictDetailsEnv['ECM Name/Domain Name']:
                                errorVar = "ECM Name/Domain Name cannot be empty in ecm\'s sheet."
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
                                errorVar = "Mandatory Key: " + "Score for R" + str(n) + " or " + "response(R" + str(
                                    n) + ")_hint is missing"
                        for row_index_env in range(2, detailsEnvSheet.nrows):
                            dictDetailsEnv = {
                                keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                col_index_env in range(detailsEnvSheet.ncols)}
                            quesExtIds.append(dictDetailsEnv['question_id'].encode('utf-8').decode('utf-8').lower())

                            if not dictDetailsEnv['criteria_id']:
                                errorVar = "criteria_id cannot be empty in questions sheet."
                            if not dictDetailsEnv['criteria_id'].lower() in criteriaExternalIds:
                                errorVar = "Criteria ID : " + dictDetailsEnv['criteria_id'] + " in question sheet not present in criteria sheet."
                            # question_sequence = dictDetailsEnv['question_sequence'] if dictDetailsEnv['question_sequence'] else ElevateObservation.terminatingMessage("\"question_sequence\" must not be Empty in \"questions\" sheet")
                            if dictDetailsEnv.get('question_sequence'):
                                question_sequence = dictDetailsEnv['question_sequence']
                            else:
                                errorVar = "\"question_sequence\" must not be Empty in \"questions\" sheet"
                            questionsequenceArr.append(question_sequence)
                            question_sequence_arr = questionsequenceArr
                            if not dictDetailsEnv['question_primary_language']:
                                errorVar = "question_primary_language cannot be empty in questions sheet."
                            if not dictDetailsEnv['question_response_type']:
                                errorVar = "question_response_type cannot be empty in questions sheet."
                            if not dictDetailsEnv['question_id']:
                                errorVar = "question_id cannot be empty in questions sheet."
                            if not dictDetailsEnv['criteria_id']:
                                errorVar = "criteria_id : " + str(
                                    dictDetailsEnv['criteria_id']) + "  cannot be empty in questions sheet."
                            if not dictDetailsEnv['criteria_id'].lower() in criteriaExternalIds:
                                errorVar = "criteria_id : " + str(dictDetailsEnv['criteria_id']) + " in questions sheet is not matching the criteria upload."
                        if not len(question_sequence_arr) == len(set(question_sequence_arr)):
                                errorVar = "\"question_sequence\" must be Unique in \"questions\" sheet"
                        if not len(quesExtIds) == len(set(quesExtIds)):
                                errorVar = "Duplicate question_id detected in questions sheet."
                        if not ElevateObservation.check_sequence(question_sequence_arr): 
                            errorVar = "\"question_sequence\" must be in sequence in \"questions\" sheet"
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
                        if sheetEnv.strip().lower() == 'criteria_rubric-scoring':
                            print("--->Checking Criteria Rubrics sheet")
                            cR_extIds = list()
                            detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                            keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                                    range(detailsEnvSheet.ncols)]
                            listOfCRs = ["criteriaId", "weightage"]
                            for cl in criteriaLevels:
                                listOfCRs.append("L" + str(cl)+" "+"SCORE")
                            listOfCRs.append("Ln SCORE")

                            print(listOfCRs,"listOfCRs")
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
                                    if not dictDetailsEnv["L" + str(cl)+" "+"SCORE"]:
                                        errorVar = "L" + str(cl) + " must not be empty in criteria_rubric."
                                if not dictDetailsEnv['criteriaId']:
                                    errorVar = "criteriaId must not be empty in criteria_rubric sheet."
                                if not dictDetailsEnv['weightage']:
                                    errorVar = "weightage cannot be empty in criteria_rubric sheet."
                            if not len(cR_extIds) == len(set(cR_extIds)):
                                errorVar = "Duplicate externalId detected in criteria_rubric sheet."
                        # sys.exit()
                        if sheetEnv.strip().lower() == 'domain(theme)_rubric_scoring':
                            print("--->Checking Theme Rubrics sheet")
                            detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                            keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                                    range(detailsEnvSheet.ncols)]
                            for row_index_env in range(1, detailsEnvSheet.nrows):
                                dictDetailsEnv = {
                                    keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                    col_index_env in range(detailsEnvSheet.ncols)}
                                if not dictDetailsEnv['domain_Id']:
                                    errorVar = "domain_Id cannot be empty in theme_rubric sheet."
                                if not dictDetailsEnv['domain_name']:
                                    errorVar = "domain_name cannot be empty in theme_rubric sheet."
                                if not dictDetailsEnv['weightage']:
                                    errorVar = "weightage cannot be empty in theme_rubric sheet."
            if errorVar == "":
                return True
            else:
                print(errorVar,"2950")
                return False
        except Exception as e:
            print(f"Error during ECM processing: {str(e)}")
            print(errorVar,"2954")

    def ObsWORValidate(wbObservation1, accessToken, parentFolder):
        print("Validating Observation temp....")
        global errorVar, entityType, solutionName, solutionDescription, creator,solutionLanguage
        try:
            questionsequenceArr =[]
            sheetNames1 = wbObservation1.sheet_names()
            observation_sheet_names = ['Instructions', 'details', 'criteria', 'questions']
            if (len(observation_sheet_names) == len(sheetNames1)) and ((set(observation_sheet_names) == set(sheetNames1))):
                print("--->Observation without rubrics file detected.<---")
            questionsequenceArr =[]
            # Point based value set as null by default for observation without rubrics
            pointBasedValue = "null"
            criteria_id_arr = []
            detailsColNames = ['observation_solution_name', 'observation_solution_description', 'Name_of_the_creator','Username/user id/email id/phone no. of the Content creator','language', 'keywords', 'entity_type',"start_date","end_date"]
            criteriaColNames = ['criteria_id', 'criteria_name']
            questionsColNames = ["criteria_id","question_sequence","question_id","instance_parent_question_id","parent_question_id","show_when_parent_question_value_is","parent_question_value","page","question_number","question_primary_language","question_secondory_language","question_tip","question_hint","instance_identifier","question_response_type","date_auto_capture","response_required","min_number_value","max_number_value","file_upload","show_remarks","response(R1)","response(R1)_hint","response(R2)","response(R2)_hint","response(R3)","response(R3)_hint","response(R4)","response(R4)_hint","response(R5)","response(R5)_hint","response(R6)","response(R6)_hint","response(R7)","response(R7)_hint","response(R8)","response(R8)_hint","response(R9)","response(R9)_hint","response(R10)","response(R10)_hint","response(R11)","response(R11)_hint","response(R12)","response(R12)_hint","response(R13)","response(R13)_hint","response(R14)","response(R14)_hint","response(R15)","response(R15)_hint","response(R16)","response(R16)_hint","response(R17)","response(R17)_hint","response(R18)","response(R18)_hint","response(R19)","response(R19)_hint","response(R20)","response(R20)_hint","question_weightage","section_header"]
            for sheetColCheck in sheetNames1:
                if sheetColCheck.strip().lower() == 'details':
                    detailsColCheck = wbObservation1.sheet_by_name(sheetColCheck)
                    keysColCheckDetai = [detailsColCheck.cell(0, col_index_check).value for col_index_check in
                                        range(detailsColCheck.ncols)]
                    if len(keysColCheckDetai) != len(detailsColNames):
                        print("Details sheet columns mismatch")
                        print("keysColCheckDetai",keysColCheckDetai)
                        print("detailsColNames",detailsColNames)
                        print("len(keysColCheckDetai)",len(keysColCheckDetai))
                        print("len(detailsColNames)",len(detailsColNames))
                        print("keysColCheckDetai != detailsColNames")   
                        errorVar = 'Some Columns are missing in details sheet'
                if sheetColCheck.strip().lower() == 'criteria':
                    criteriaColCheck = wbObservation1.sheet_by_name(sheetColCheck)
                    keysColCheckCrit = [criteriaColCheck.cell(0, col_index_check1).value for col_index_check1 in
                                        range(criteriaColCheck.ncols)]
                    if len(keysColCheckCrit) != len(criteriaColNames):
                        errorVar = 'Columns is missing in criteria sheet'
                if sheetColCheck.strip().lower() == 'questions':
                    questionsColCheck = wbObservation1.sheet_by_name(sheetColCheck)
                    keysColCheckQues = [questionsColCheck.cell(0, col_index_check2).value for col_index_check2 in
                                        range(questionsColCheck.ncols)]
                    if len(keysColCheckQues) != len(questionsColNames):
                        errorVar = 'Columns is missing in questions sheet'
            for sheetEnv in sheetNames1:
                if sheetEnv == "Instructions":
                    pass
                else:
                    if sheetEnv.strip().lower() == 'details':
                        print("--->Checking details sheet...")
                        detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                        keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                                range(detailsEnvSheet.ncols)]
                        for row_index_env in range(2, detailsEnvSheet.nrows):
                            dictDetailsEnv = {
                                keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                col_index_env in range(detailsEnvSheet.ncols)}
                            if dictDetailsEnv['observation_solution_name']:
                                solutionName = dictDetailsEnv['observation_solution_name'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "validation failed :observation_solution_name column must not be Empty in details sheet"
                            if dictDetailsEnv['observation_solution_description']:
                                solutionDescription = dictDetailsEnv['observation_solution_description'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "validation failed :observation_solution_description column must not be Empty in details sheet"
                            # solutionName = dictDetailsEnv['observation_solution_name'].encode('utf-8').decode('utf-8') if dictDetailsEnv['observation_solution_name'] else terminatingMessage("\"observation_solution_name\" must not be Empty in \"details\" sheet")
                            # solutionDescription = dictDetailsEnv['observation_solution_description'].encode('utf-8').decode('utf-8') if dictDetailsEnv['observation_solution_description'] else terminatingMessage("\"observation_solution_description\" must not be Empty in \"details\" sheet")
                            if dictDetailsEnv['Username/user id/email id/phone no. of the Content creator']:
                                dikshaLoginId = dictDetailsEnv['Username/user id/email id/phone no. of the Content creator'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "validation failed :Username/user id/email id/phone no. of the Content creator column must not be Empty in details sheet"
                            # dikshaLoginId = dictDetailsEnv['Elevate_loginId'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Elevate_loginId'] else terminatingMessage("\"Elevate_loginId\" must not be Empty in \"details\" sheet")
                            if dictDetailsEnv['Name_of_the_creator']:
                                creator = dictDetailsEnv['Name_of_the_creator'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "validation failed :Name_of_the_creator column must not be Empty in details sheet"

                            if dictDetailsEnv['Username/user id/email id/phone no. of the Content creator']:
                                dikshaLoginId = dictDetailsEnv['Username/user id/email id/phone no. of the Content creator'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "validation failed :Username/user id/email id/phone no. of the Content creator column must not be Empty in details sheet"
                            # creator = dictDetailsEnv['Name_of_the_creator'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Name_of_the_creator'] else terminatingMessage("\"Name_of_the_creator\" must not be Empty in \"details\" sheet")
                            # ccUserDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dikshaLoginId)
                            
                            # if not "CONTENT_CREATOR" in ccUserDetails[3]:
                            #     terminatingMessage("---> "+dikshaLoginId +" is not a CONTENT_CREATOR in Diksha " + environment)
                            # ccRootOrgName = ccUserDetails[4]
                            # ccRootOrgId = ccUserDetails[5]
                            if dictDetailsEnv['language']:
                                solutionLanguage = dictDetailsEnv['language'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "validation failed :language column must not be Empty in details sheet"    
                            if dictDetailsEnv['entity_type']:
                                entityType = dictDetailsEnv['entity_type'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "validation failed :entity_type column, please select from the given drop down in details sheet"
                            # entityType = dictDetailsEnv['entity_type'].encode('utf-8').decode('utf-8') if dictDetailsEnv['entity_type'] else terminatingMessage("\"entity_type\" must not be Empty in \"details\" sheet")
                            # solutionLanguage = dictDetailsEnv['language'].encode('utf-8').decode('utf-8').split(",") if dictDetailsEnv['language'] else [""]
                            ElevateObservation.getProgramInfo(accessToken, parentFolder, programNameInp)
                    elif sheetEnv.strip().lower() == 'criteria':
                        print("--->Checking criteria sheet...")
                        detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                        keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                                range(detailsEnvSheet.ncols)]
                        for row_index_env in range(2, detailsEnvSheet.nrows):
                            dictDetailsEnv = {
                                keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                col_index_env in range(detailsEnvSheet.ncols)}
                            if dictDetailsEnv['criteria_id']:
                                criteria_id = dictDetailsEnv['criteria_id'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "validation failed :criteria_id column must not be Empty in criteria sheet"
                            if dictDetailsEnv['criteria_name']:
                                criteria_name = dictDetailsEnv['criteria_name'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "validation failed :criteria_name column must not be Empty in criteria sheet"
                            # criteria_id = dictDetailsEnv['criteria_id'].encode('utf-8').decode('utf-8') if dictDetailsEnv['criteria_id'] else terminatingMessage("\"criteria_id\" must not be Empty in \"criteria\" sheet")
                            # criteria_name = dictDetailsEnv['criteria_name'].encode('utf-8').decode('utf-8') if dictDetailsEnv['criteria_name'] else terminatingMessage("\"criteria_name\" must not be Empty in \"criteria\" sheet")
                            criteria_id_arr.append(criteria_id)
                        if not len(criteria_id_arr) == len(set(criteria_id_arr)):
                            errorVar = "criteria_id must be Unique in criteria sheet"
                    elif sheetEnv.strip().lower() == 'questions':
                        print("--->Checking question sheet...")
                        detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                        ques_id_arr = list()
                        keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                                range(detailsEnvSheet.ncols)]
                        for row_index_env in range(2, detailsEnvSheet.nrows):
                            dictDetailsEnv = {
                                keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                col_index_env in range(detailsEnvSheet.ncols)}
                            if dictDetailsEnv['criteria_id']:
                                criteria_id = dictDetailsEnv['criteria_id'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "validation failed :criteria_id column must not be Empty in questions sheet"
                            if dictDetailsEnv['question_sequence']:
                                question_sequence = dictDetailsEnv['question_sequence']
                            else:
                                errorVar = "validation failed :question_sequence column must not be Empty in questions sheet"
                            # criteria_id = dictDetailsEnv['criteria_id'].encode('utf-8').decode('utf-8') if dictDetailsEnv['criteria_id'] else terminatingMessage("\"criteria_id\" must not be Empty in \"questions\" sheet")
                            # question_sequence = dictDetailsEnv['question_sequence'] if dictDetailsEnv['question_sequence'] else terminatingMessage("\"question_sequence\" must not be Empty in \"questions\" sheet")

                            questionsequenceArr.append(question_sequence)
                            question_sequence_arr = questionsequenceArr

                            if not criteria_id in criteria_id_arr:
                                errorVar = "\"criteria_id\" in \"Questions\" sheet must be declared in \"criteria\" sheet"
                            if not criteria_id in criteria_id_arr:
                                errorVar = "criteria_id in Questions sheet must be declared in questions sheet"
                            if dictDetailsEnv['page']:
                                page = dictDetailsEnv['page'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "validation failed :page column must not be Empty in questions sheet"
                            if dictDetailsEnv['question_number']:
                                question_number = dictDetailsEnv['question_number']
                            else:
                                errorVar = "validation failed :question_number column must not be Empty in questions sheet"
                            if dictDetailsEnv['question_primary_language']:
                                question_primary_language = dictDetailsEnv['question_primary_language'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "validation failed :question_primary_language column must not be Empty in questions sheet"
                            # page = dictDetailsEnv['page'].encode('utf-8').decode('utf-8') if dictDetailsEnv['page'] else terminatingMessage("\"page\" must not be Empty in \"questions\" sheet")
                            # question_number = dictDetailsEnv['question_number'] if dictDetailsEnv['question_number'] else terminatingMessage("\"question_number\" must not be Empty in \"questions\" sheet")
                            # question_primary_language = dictDetailsEnv['question_primary_language'].encode('utf-8').decode('utf-8') if dictDetailsEnv['question_primary_language'] else terminatingMessage("\"question_primary_language\" must not be Empty in \"questions\" sheet")
                            # if dictDetailsEnv['response_required'] in (True,False):
                            #     response_required = dictDetailsEnv['response_required']
                            # else:
                            #     errorVar = "validation failed :response_required column must not be Empty in questions sheet"
                            if dictDetailsEnv['question_id']:
                                question_id = dictDetailsEnv['question_id']
                            else:
                                errorVar = "validation failed :question_id column must not be Empty in questions sheet"    
                            # response_required = dictDetailsEnv['response_required'] if str(dictDetailsEnv['response_required']) else terminatingMessage("\"response_required\" must not be Empty in \"questions\" sheet")

                            # question_id = dictDetailsEnv['question_id'] if dictDetailsEnv['question_id'] else terminatingMessage("\"question_id\" must not be Empty in \"questions\" sheet")
                            ques_id_arr.append(question_id)
                            parent_question_id = dictDetailsEnv['question_id']
                            if parent_question_id and not parent_question_id in ques_id_arr:
                                errorVar = "parent_question_id referenced before assigning in questions sheet."
                            if dictDetailsEnv['question_response_type']:
                                question_response_type = dictDetailsEnv['question_response_type'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "validation failed :question_response_type column must not be Empty in questions sheet"
                            # question_response_type = dictDetailsEnv['question_response_type'].encode('utf-8').decode('utf-8') if dictDetailsEnv[
                                # 'question_response_type'] else terminatingMessage(
                                # "\"question_response_type\" must not be Empty in \"questions\" sheet")
                        if not len(question_sequence_arr) == len(set(question_sequence_arr)):
                            errorVar = "question_sequence must be Unique in questions sheet"
                        if not ElevateObservation.check_sequence(question_sequence_arr): 
                            errorVar = "question_sequence must be in sequence in questions sheet"
            if errorVar == "":
                return True
            else:
                print(errorVar,"3292")
                return False
        except Exception as e:
            print(f"Error during ECM processing: {str(e)}")
            print(errorVar,"3270")        

    def surveyValidate(filePathAddObs, accessToken, parentFolder):
        print("Validating survey temp....")
        global errorVar
        try:
            wbObservation1 = xlrd.open_workbook(filePathAddObs, on_demand=True)
            sheetNames1 = wbObservation1.sheet_names()
            survey_sheet_names = ['Instructions', 'details', 'questions']
            if (len(survey_sheet_names) == len(sheetNames1)) and ((set(survey_sheet_names) == set(sheetNames1))):
                print("--->Survey file detected.<---")
            
            for sheetEnvCheck in sheetNames1:
                if sheetEnvCheck.strip().lower() == 'instructions' or sheetEnvCheck.strip().lower() == 'details' or sheetEnvCheck.strip().lower() == 'questions':
                    pass
                else:
                    errorVar = 'Sheet Names in excel file is wrong , Sheet Names are details,questions'

            detailsColNames = ["survey_solution_name", "survey_solution_description", "Name_of_the_creator","Username/user id/email id/phone no. of the Content creator", "survey_start_date", "survey_end_date"]
            questionsColNames = ["question_sequence", "question_id", "section_header", "instance_parent_question_id",
                                "parent_question_id", "show_when_parent_question_value_is", "parent_question_value",
                                "page", "question_number", "question_language1", "question_language2", "question_tip",
                                "question_hint", "instance_identifier", "question_response_type", "date_auto_capture",
                                "response_required","question_response_validation", "min_number_value", "max_number_value", "file_upload", "show_remarks",
                                "response(R1)", "response(R2)", "response(R3)", "response(R4)", "response(R5)",
                                "response(R6)", "response(R7)", "response(R8)", "response(R9)", "response(R10)",
                                "response(R11)", "response(R12)", "response(R13)", "response(R14)", "response(R15)",
                                "response(R16)", "response(R17)", "response(R18)", "response(R19)", "response(R20)",
                                "response(R1)_hint", "response(R2)_hint", "response(R3)_hint", "response(R4)_hint",
                                "response(R5)_hint", "response(R6)_hint", "response(R7)_hint", "response(R8)_hint",
                                "response(R9)_hint", "response(R10)_hint", "response(R11)_hint", "response(R12)_hint",
                                "response(R13)_hint", "response(R14)_hint", "response(R15)_hint", "response(R16)_hint",
                                "response(R17)_hint", "response(R18)_hint", "response(R19)_hint", "response(R20)_hint"]

            for sheetColCheck in sheetNames1:
                # print(sheetColCheck,"sheetColCheck 2717")
                if sheetColCheck.strip().lower() == 'details':
                    detailsColCheck = wbObservation1.sheet_by_name(sheetColCheck)
                    keysColCheckDetai = [detailsColCheck.cell(1, col_index_check).value for col_index_check in
                                        range(detailsColCheck.ncols)]
                    # print(keysColCheckDetai,detailsColNames)
                    # print(len(keysColCheckDetai),len(detailsColNames))
                    if len(keysColCheckDetai) != len(detailsColNames) :
                        errorVar = 'Some Columns are missing in details sheet'

                    detailsEnvSheet = wbObservation1.sheet_by_name(sheetColCheck)
                    keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                                range(detailsEnvSheet.ncols)]

                    for row_index_env in range(2, detailsEnvSheet.nrows):
                        dictDetailsEnv = {
                            keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                            col_index_env in range(detailsEnvSheet.ncols)}
                        if dictDetailsEnv['survey_solution_name']:
                            surveysolutionname = dictDetailsEnv['survey_solution_name']
                        else:
                            errorVar = "validation failed :survey_solution_name column must not be Empty in details sheet"
                        if dictDetailsEnv['survey_solution_description']:
                            surveysolutiondescription = dictDetailsEnv['survey_solution_description']
                        else:
                            errorVar = "validation failed :survey_solution_description column must not be Empty in details sheet"
                        if dictDetailsEnv['Name_of_the_creator']:
                            Nameofthecreator = dictDetailsEnv['Name_of_the_creator']
                        else:
                            errorVar = "validation failed :Name_of_the_creator column must not be Empty in details sheet"
                        if dictDetailsEnv['Username/user id/email id/phone no. of the Content creator']:
                            surveycreatorusername = dictDetailsEnv['Username/user id/email id/phone no. of the Content creator']
                        else:
                            errorVar = "validation failed :Username/user id/email id/phone no. of the Content creator column must not be Empty in details sheet"
                        if dictDetailsEnv['survey_start_date']:
                            surveystartdate = dictDetailsEnv['survey_start_date']
                        else:
                            errorVar = "validation failed :survey_start_date column must not be Empty in details sheet"
                        if dictDetailsEnv['survey_end_date']:
                            surveyenddate = dictDetailsEnv['survey_end_date']
                        else:
                            errorVar = "validation failed :survey_end_date column must not be Empty in details sheet"

                if sheetColCheck.strip().lower() == 'questions':
                    questionsColCheck = wbObservation1.sheet_by_name(sheetColCheck)
                    keysColCheckQues = [questionsColCheck.cell(1, col_index_check2).value for col_index_check2 in
                                        range(questionsColCheck.ncols)]
                    # print(keysColCheckQues)
                    if len(keysColCheckQues) != len(questionsColNames):
                        errorVar = 'Some Columns are missing in questions sheet'
                    for row_index_env in range(2, questionsColCheck.nrows):
                        dictDetailsEnv = {
                            keysColCheckQues[col_index_env]: questionsColCheck.cell(row_index_env, col_index_env).value for
                            col_index_env in range(questionsColCheck.ncols)}
                        if dictDetailsEnv['question_sequence']:
                            question_sequenceSUR = dictDetailsEnv['question_sequence']
                        else:
                            errorVar = "validation failed :question_sequence column must not be Empty in questions sheet"
                        if dictDetailsEnv['question_id']:
                            question_idSUR = dictDetailsEnv['question_id'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "validation failed :question_id column must not be Empty in questions sheet"
                        if dictDetailsEnv['page']:
                            pageSUR = dictDetailsEnv['page']
                        else:
                            errorVar = "validation failed :page column must not be Empty in questions sheet"
                        if dictDetailsEnv['question_number']:
                            question_numberSUR = dictDetailsEnv['question_number']
                        else:
                            errorVar = "validation failed :question_number column must not be Empty in questions sheet"
                        if dictDetailsEnv['question_language1']:
                            question_language1SUR = dictDetailsEnv['question_language1'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "validation failed :question_language1 column must not be Empty in questions sheet"
                        if dictDetailsEnv['question_response_type']:
                            question_response_typeSUR = dictDetailsEnv['question_response_type']
                        else:
                            errorVar = "validation failed :question_response_type column must not be Empty in questions sheet"
                        # print(dictDetailsEnv['response_required'])
                        # if dictDetailsEnv['response_required']:
                        #     response_required = dictDetailsEnv['response_required']
                        # else:
                        #     errorVar = "validation failed :response_required column must not be Empty in questions sheet"
            if errorVar == "":
                return True
            else:
                print(errorVar,"3415")
                return False            
        except Exception as e:
            print(f"Error during ECM processing: {str(e)}")
            print(errorVar,"3419")

    def createSurveySolution(parentFolder, wbSurvey, accessToken):
        global errorVar,creator
        error_message = ""
        print("Create Survey Solution Func Called....")
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
                    surveySolutionCreationReqBody['name'] = dictDetailsEnv['survey_solution_name'].encode('utf-8').decode('utf-8')
                    surveySolutionCreationReqBody["description"] = dictDetailsEnv['survey_solution_description'].encode('utf-8').decode('utf-8')
                    surveySolutionExternalId = str(uuid.uuid1())
                    surveySolutionCreationReqBody["externalId"] = surveySolutionExternalId
                    if dictDetailsEnv['Name_of_the_creator']== "":
                        exceptionHandlingFlag = True
                        print('survey_creator_username column should not be empty in the details sheet column should not be empty in the details sheet')
                        # sys.exit()
                    else:
                        surveySolutionCreationReqBody['creator'] = dictDetailsEnv['Name_of_the_creator']

                    surveySolutionCreationReqBody['isExternalProgram'] = True
                    userDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dictDetailsEnv['Username/user id/email id/phone no. of the Content creator'])
                    surveySolutionCreationReqBody['author'] = userDetails[0]
                    global SurveyTemplateStartDate, SurveyTemplateEndDate
                    SurveyTemplateStartDate = dictDetailsEnv["survey_start_date"]
                    SurveyTemplateEndDate = dictDetailsEnv["survey_end_date"]
                    try: 
                        urlCreateSolutionApi = internal_kong_ip+ surveysolutioncreationapiurl
                        headerCreateSolutionApi = {
                            'Content-Type': content_type,
                            "internal-access-token": internal_access_token,
                            # 'Authorization': config.get(environment, 'Authorization'),
                            'X-auth-token': accessToken,
                            # 'X-Channel-id': config.get(environment, 'X-Channel-id'),
                            # 'appName': config.get(environment, 'appName'),
                            'tenantId': tenantID ,
                            'orgid': orgIDFromTemplate,
                            adminTokenHeaderName: adminAccessToken
                        }
                        print(headerCreateSolutionApi,"headerCreateSolutionApi")
                        responseCreateSolutionApi = requests.post(url=urlCreateSolutionApi,
                                                                headers=headerCreateSolutionApi,
                                                                data=json.dumps(surveySolutionCreationReqBody))
                        responseInText = responseCreateSolutionApi.text
                        messageArr = ["********* Create Survey Solution *********", "URL : " + urlCreateSolutionApi,
                                    "BODY : " + str(surveySolutionCreationReqBody),
                                    "Status code : " + str(responseCreateSolutionApi.status_code),
                                    "Response : " + responseCreateSolutionApi.text]
                        fileheader = [surveySolutionCreationReqBody['name'].encode('utf-8').decode('utf-8'),'Program Sheet Validation'," "]
                        ElevateObservation.createAPILog(parentFolder, messageArr)
                        ElevateObservation.apicheckslog(parentFolder,fileheader)
                        print(responseCreateSolutionApi.text,"responseCreateSolutionApi")
                        if responseCreateSolutionApi.status_code == 200:
                            responseCreateSolutionApi = responseCreateSolutionApi.json()
                            urlSearchSolution = internal_kong_ip + fetchsolutiondetails + "survey&page=1&limit=10&search=" + str(surveySolutionExternalId)
                            print(urlSearchSolution,"urlSearchSolution")
                            responseSearchSolution = requests.post(urlSearchSolution,
                                                                    headers=headerCreateSolutionApi)
                            print(headerCreateSolutionApi,"headerCreateSolutionApi")
                            print(responseSearchSolution.text,"responseSearchSolution")
                            messageArr = ["********* Search Survey Solution *********", "URL : " + urlSearchSolution,
                                        "Status code : " + str(responseSearchSolution.status_code),
                                        "Response : " + responseSearchSolution.text]
                            ElevateObservation.createAPILog(parentFolder, messageArr)
                            ElevateObservation.apicheckslog(parentFolder, messageArr)
                            if responseSearchSolution.status_code == 200:
                                responseSearchSolutionApi = responseSearchSolution.json()
                                surveySolutionExternalId = None
                                surveySolutionExternalId = responseSearchSolutionApi['result']['data'][0]['externalId']
                                # return True
                            else:
                                error_message = ""
                                if responseSearchSolution.status_code in [400, 401, 403, 404, 422]:
                                    error_message = f"SearchSolution-Client Error {responseSearchSolution.status_code}: {responseSearchSolution.text}"
                                elif responseSearchSolution.status_code in [500, 502, 503, 504]:
                                    error_message = f"SearchSolution-Server Error {responseSearchSolution.status_code}: {responseSearchSolution.text}"
                                else:
                                    error_message = f"SearchSolution-Unexpected Error {responseSearchSolution.status_code}: {responseSearchSolution.text}"
                                errorVar = error_message
                                messageArr.append(f"Error Response: {error_message}")
                                ElevateObservation.createAPILog(messageArr,messageArr) 
                                # return False
                            solutionId = None
                            solutionId = responseCreateSolutionApi["result"]["solutionId"]
                            bodySolutionUpdate = {"creator": dictDetailsEnv['Name_of_the_creator'].encode('utf-8').decode('utf-8')}
                            if ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                                return [solutionId, surveySolutionExternalId]
                            else:
                                print("solution update failed...")
                                print(errorVar)
                                return errorVar
                        else:
                            error_message = ""
                            if responseCreateSolutionApi.status_code in [400, 401, 403, 404, 422]:
                                error_message = f"CreateSolutionApi-Client Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}"
                            elif responseCreateSolutionApi.status_code in [500, 502, 503, 504]:
                                error_message = f"CreateSolutionApi-Server Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}"
                            else:
                                error_message = f"CreateSolutionApi-Unexpected Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}"
                            errorVar = error_message
                            print(error_message)
                            messageArr.append(f"Error Response: {error_message}")
                            ElevateObservation.createAPILog(parentFolder,messageArr) 
                            return False 
    
                    except Exception as e:
                        errorVar = error_message
                        print(error_message,"5591")
                        print(errorVar,"5592")
                        ElevateObservation.createAPILog([parentFolder,f"Exception: {str(e)}"])
    
    def convert_to_date(date_str):
        return datetime.strptime(date_str, "%d-%m-%Y")
    
    def uploadSurveyQuestions(MainFilePath, parentFolder, wbSurvey, addObservationSolution, accessToken, surTempExtID, surTempSolID, millisecond, programFile):
        print("Upload Survey Questions Func Called....")
        # print(parentFolder,"4854")
        # wbSurvey = xlrd.open_workbook(wbSurvey, on_demand=True)
        # print(f"Type of wbSurvey: {type(wbSurvey)}")
        sheetNam = wbSurvey.sheet_names()
        # print(sheetNam,"4854")
        global surveySolutionlink, errorVar, entityHierarchy
        error_message = ""
        stDt = None
        enDt = None
        shCnt = 0
        for i in sheetNam:
            if i.strip().lower() == 'questions':
                sheetNam1 = wbSurvey.sheets()[shCnt]
            shCnt = shCnt + 1
        dataSort = [sheetNam1.row_values(i) for i in range(sheetNam1.nrows)]
        labels = dataSort[1]
        dataSort = dataSort[2:]
        dataSort.sort(key=lambda x: int(x[0]))
        openWorkBookSort1 = xl_copy(wbSurvey)
        sheet1 = openWorkBookSort1.add_sheet('questions_sequence_sorted')

        for idx, label in enumerate(labels):
            sheet1.write(0, idx, label)

        for idx_r, row in enumerate(dataSort):
            for idx_c, value in enumerate(row):
                sheet1.write(idx_r + 1, idx_c, value)
        newFileName = str(addObservationSolution)
        openWorkBookSort1.save(newFileName)
        openNewFile = xlrd.open_workbook(newFileName, on_demand=True)
        wbSurvey = openNewFile
        sheetNames = wbSurvey.sheet_names()
        # print("reached till here 4881")
        for sheet2 in sheetNames:
            if sheet2.strip().lower() == 'questions_sequence_sorted':
                questionsList = []
                questionsSheet = wbSurvey.sheet_by_name(sheet2.lower())
                keys2 = [questionsSheet.cell(0, col_index2).value for col_index2 in
                        range(questionsSheet.ncols)]
                for row_index2 in range(1, questionsSheet.nrows):
                    d2 = {keys2[col_index2]: questionsSheet.cell(row_index2, col_index2).value
                        for col_index2 in range(questionsSheet.ncols)}
                    questionsList.append(d2)
                questionSeqByEcmArr = []
                quesSeqCnt = 1.0
                questionUploadFieldnames = []
                questionUploadFieldnames = ['solutionId', 'instanceParentQuestionId','hasAParentQuestion', 'parentQuestionOperator','parentQuestionValue', 'parentQuestionId','externalId', 'question0', 'question1', 'tip','hint', 'instanceIdentifier', 'responseType','dateFormat', 'autoCapture', 'validation','validationIsNumber', 'validationRegex','validationMax', 'validationMin', 'file','fileIsRequired', 'fileUploadType','allowAudioRecording', 'minFileCount','maxFileCount', 'caption', 'questionGroup','modeOfCollection', 'accessibility', 'showRemarks','rubricLevel', 'isAGeneralQuestion', 'R1','R1-hint', 'R2', 'R2-hint', 'R3', 'R3-hint', 'R4','R4-hint', 'R5', 'R5-hint', 'R6', 'R6-hint', 'R7','R7-hint', 'R8', 'R8-hint', 'R9', 'R9-hint', 'R10','R10-hint', 'R11', 'R11-hint', 'R12', 'R12-hint','R13', 'R13-hint', 'R14', 'R14-hint', 'R15','R15-hint', 'R16', 'R16-hint', 'R17', 'R17-hint','R18', 'R18-hint', 'R19', 'R19-hint', 'R20','R20-hint', 'sectionHeader', 'page','questionNumber', '_arrayFields']

                for ques in questionsList:

                    questionFilePath = parentFolder + '/questionUpload/'
                    file_exists_ques = os.path.isfile(
                        parentFolder + '/questionUpload/uploadSheet.csv')
                    # print(questionFilePath,"4904")
                    if not os.path.exists(questionFilePath):
                        os.mkdir(questionFilePath)
                    with open(parentFolder + '/questionUpload/uploadSheet.csv', 'a',
                            encoding='utf-8') as questionUploadFile:
                        writerQuestionUpload = csv.DictWriter(questionUploadFile, fieldnames=questionUploadFieldnames, lineterminator='\n')
                        if not file_exists_ques:
                            writerQuestionUpload.writeheader()
                        questionFileObj = {}
                        surveyExternalId = None
                        questionFileObj['solutionId'] = surTempExtID
                        if ques['instance_parent_question_id'].encode('utf-8').decode('utf-8'):
                            questionFileObj['instanceParentQuestionId'] = ques[
                                                                            'instance_parent_question_id'].strip() + '_' + str(
                                millisecond)
                        else:
                            questionFileObj['instanceParentQuestionId'] = 'NA'
                        if ques['parent_question_id'].encode('utf-8').decode('utf-8').strip():
                            questionFileObj['hasAParentQuestion'] = 'YES'
                            if ques['show_when_parent_question_value_is'] == 'or':
                                questionFileObj['parentQuestionOperator'] = '||'
                            else:
                                questionFileObj['parentQuestionOperator'] = ques['show_when_parent_question_value_is']
                            if type(ques['parent_question_value']) != str:
                                if (ques['parent_question_value'] and ques[
                                    'parent_question_value'].is_integer() == True):
                                    questionFileObj['parentQuestionValue'] = int(ques['parent_question_value'])
                                elif (ques['parent_question_value'] and ques[
                                    'parent_question_value'].is_integer() == False):
                                    questionFileObj['parentQuestionValue'] = ques['parent_question_value']
                            else:
                                questionFileObj['parentQuestionValue'] = ques['parent_question_value']
                                questionFileObj['parentQuestionId'] = ques['parent_question_id'].encode('utf-8').decode('utf-8').strip() + '_' + str(
                                    millisecond)
                        else:
                            questionFileObj['hasAParentQuestion'] = 'NO'
                            questionFileObj['parentQuestionOperator'] = None
                            questionFileObj['parentQuestionValue'] = None
                            questionFileObj['parentQuestionId'] = None
                        questionFileObj['externalId'] = ques['question_id'].strip() + '_' + str(millisecond)
                        if quesSeqCnt == ques['question_sequence']:
                            questionSeqByEcmArr.append(ques['question_id'].strip() + '_' + str(millisecond))
                            quesSeqCnt = quesSeqCnt + 1.0
                        if ques['question_language1']:
                            questionFileObj['question0'] = ques['question_language1'].encode('utf-8').decode('utf-8')
                        else:
                            questionFileObj['question0'] = None
                        if ques['question_language2']:
                            questionFileObj['question1'] = ques['question_language2'].encode('utf-8').decode('utf-8')
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
                        if ques['question_response_type'].strip().lower() == 'date':
                            questionFileObj['dateFormat'] = "DD-MM-YYYY"
                        else:
                            questionFileObj['dateFormat'] = None
                        if ques['question_response_type'].strip().lower() == 'date':
                            if ques['date_auto_capture'] and ques['date_auto_capture'] == 1:
                                questionFileObj['autoCapture'] = 'TRUE'
                            elif ques['date_auto_capture'] and ques['date_auto_capture'] == 0:
                                questionFileObj['autoCapture'] = 'false'
                            else:
                                questionFileObj['autoCapture'] = 'false'
                        else:
                            questionFileObj['autoCapture'] = None
                        if ques['response_required']:
                            if ques['response_required'] == 1:
                                questionFileObj['validation'] = 'TRUE'
                            elif ques['response_required'] == 0:
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
                        if ques['file_upload'] == 1:
                            questionFileObj['file'] = 'Snapshot'
                            questionFileObj['fileIsRequired'] = 'TRUE'
                            questionFileObj['fileUploadType'] = 'image/jpeg,docx,pdf,ppt'
                            questionFileObj['minFileCount'] = 0
                            questionFileObj['maxFileCount'] = 10
                        elif ques['file_upload'] == 0:
                            questionFileObj['file'] = 'NA'
                            questionFileObj['fileIsRequired'] = None
                            questionFileObj['fileUploadType'] = None
                            questionFileObj['minFileCount'] = None
                            questionFileObj['maxFileCount'] = None

                        questionFileObj['caption'] = 'FALSE'
                        questionFileObj['questionGroup'] = 'A1'
                        questionFileObj['modeOfCollection'] = 'onfield'
                        questionFileObj['accessibility'] = 'No'
                        if ques['show_remarks'] == 1:
                            questionFileObj['showRemarks'] = 'TRUE'
                        elif ques['show_remarks'] == 0:
                            questionFileObj['showRemarks'] = 'FALSE'
                        questionFileObj['rubricLevel'] = None
                        questionFileObj['isAGeneralQuestion'] = None
                        if ques['question_response_type'].strip().lower() == 'radio' or ques[
                            'question_response_type'].strip() == 'multiselect':
                            for quesIndex in range(1, 21):
                                if type(ques['response(R' + str(quesIndex) + ')']) != str:
                                    if (ques['response(R' + str(quesIndex) + ')'] and ques[
                                        'response(R' + str(quesIndex) + ')'].is_integer() == True):
                                        questionFileObj['R' + str(quesIndex) + ''] = int(
                                            ques['response(R' + str(quesIndex) + ')'])
                                    elif (ques['response(R' + str(quesIndex) + ')'] and ques[
                                        'response(R' + str(quesIndex) + ')'].is_integer() == False):
                                        questionFileObj['R' + str(quesIndex) + ''] = ques[
                                            'response(R' + str(quesIndex) + ')']
                                else:
                                    questionFileObj['R' + str(quesIndex) + ''] = ques[
                                        'response(R' + str(quesIndex) + ')']

                                if type(ques['response(R' + str(quesIndex) + ')_hint']) != str:
                                    if (ques['response(R' + str(quesIndex) + ')_hint'] and ques[
                                        'response(R' + str(quesIndex) + ')_hint'].is_integer() == True):
                                        questionFileObj['R' + str(quesIndex) + '-hint'] = int(
                                            ques['response(R' + str(quesIndex) + ')_hint'])
                                    elif (ques['response(R' + str(quesIndex) + ')_hint'] and ques[
                                        'response(R' + str(quesIndex) + ')_hint'].is_integer() == False):
                                        questionFileObj['R' + str(quesIndex) + '-hint'] = ques[
                                            'response(R' + str(quesIndex) + ')_hint']
                                else:
                                    questionFileObj['R' + str(quesIndex) + '-hint'] = ques[
                                        'response(R' + str(quesIndex) + ')_hint']
                                questionFileObj['_arrayFields'] = 'parentQuestionValue'
                        else:
                            for quesIndex in range(1, 21):
                                questionFileObj['R' + str(quesIndex)] = None
                                questionFileObj['R' + str(quesIndex) + '-hint'] = None
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
                        writerQuestionUpload.writerow(questionFileObj)
                try:        
                    urlQuestionsUploadApi = internal_kong_ip + questionuploadapiurl
                    headerQuestionUploadApi = {
                        "internal-access-token": internal_access_token,
                        'Authorization': authorization,
                        'X-auth-token': accessToken,
                        'X-Channel-id': x_channel_id,
                        'tenantId': tenantID ,
                        'orgid': orgIDFromTemplate,
                        adminTokenHeaderName: adminAccessToken
                    }
                    filesQuestion = {
                        'questions': open(parentFolder + '/questionUpload/uploadSheet.csv', 'rb')
                    }
                    responseQuestionUploadApi = requests.post(url=urlQuestionsUploadApi,
                                                            headers=headerQuestionUploadApi, files=filesQuestion)
                    print(responseQuestionUploadApi.text,"responseQuestionUploadApi")
                    if responseQuestionUploadApi.status_code == 200:
                        print('Question upload Success')

                        messageArr = ["********* Question Upload api *********", "URL : " + urlQuestionsUploadApi,
                                    "Path : " + str(parentFolder) + str('/questionUpload/uploadSheet.csv'),
                                    "Status code : " + str(responseQuestionUploadApi.status_code),
                                    "Response : " + responseQuestionUploadApi.text]
                        ElevateObservation.createAPILog(parentFolder, messageArr)
                        messageArr1 = ["Questions","Question upload Success","Passed",str(responseQuestionUploadApi.status_code)]
                        ElevateObservation.apicheckslog(parentFolder,messageArr1)

                        with open(parentFolder + '/questionUpload/uploadInternalIdsSheet.csv', 'w+',encoding='utf-8') as questionRes:
                            questionRes.write(responseQuestionUploadApi.text)
                        urlImportSoluTemplate = internal_kong_ip + importsurveysolutiontemplateurl + str(surTempSolID) + "?appName=manage-learn&programId=" + programID
                        print(urlImportSoluTemplate,"urlImportSoluTemplate")
                        headerImportSoluTemplateApi = {
                            'X-auth-token': accessToken,
                            'X-Channel-id': x_channel_id,
                            'internal-access-token': internal_access_token,
                            'tenantId': tenantID ,
                            'orgid': orgIDFromTemplate,
                            adminTokenHeaderName: adminAccessToken
                        }
                        print(headerImportSoluTemplateApi,"headerImportSoluTemplateApi")
                        responseImportSoluTemplateApi = requests.post(url=urlImportSoluTemplate,
                                                                    headers=headerImportSoluTemplateApi)
                        if responseImportSoluTemplateApi.status_code == 200:
                            print('Creating Child Success')

                            messageArr = ["********* Creating Child api *********", "URL : " + urlImportSoluTemplate,
                                        "Status code : " + str(responseImportSoluTemplateApi.status_code),
                                        "Response : " + responseImportSoluTemplateApi.text]
                            ElevateObservation.createAPILog(parentFolder, messageArr)
                            print(responseImportSoluTemplateApi.text,"responseImportSoluTemplateApi")
                            responseImportSoluTemplateApi = responseImportSoluTemplateApi.json()
                            solutionIdSuc = responseImportSoluTemplateApi["result"]["solutionId"]
                            urlSurveyProgramMapping = internal_kong_ip + importsurveysolutiontoprogramurl + str(solutionIdSuc) + "?programId=" + programExternalId.lstrip().rstrip()
                            print(urlSurveyProgramMapping,"urlSurveyProgramMapping")
                            headeSurveyProgramMappingApi = {
                                'X-auth-token': accessToken,
                                'X-Channel-id': x_channel_id,
                                'internal-access-token': internal_access_token,
                                'tenantId': tenantID ,
                                'orgid': orgIDFromTemplate,
                                adminTokenHeaderName: adminAccessToken
                            }
                            responseSurveyProgramMappingApi = requests.post(url=urlSurveyProgramMapping,headers=headeSurveyProgramMappingApi)
                            print(responseSurveyProgramMappingApi.text,"responseSurveyProgramMappingApi")
                            if responseSurveyProgramMappingApi.status_code == 200:
                                print('Program Mapping Success')
                                
                                messageArr = ["********* Program mapping api *********", "URL : " + urlSurveyProgramMapping,
                                            "Status code : " + str(responseSurveyProgramMappingApi.status_code),
                                            "Response : " + responseSurveyProgramMappingApi.text]
                                ElevateObservation.createAPILog(parentFolder, messageArr)
                                surveyLink = None
                                solutionIdSuc = None
                                surveyExternalIdSuc = None
                                surveyLink = responseImportSoluTemplateApi["result"]["link"]
                                solutionIdSuc = responseImportSoluTemplateApi["result"]["solutionId"]
                                solutionExtIdSuc = responseImportSoluTemplateApi["result"]["solutionExternalId"]
                                print("Survey Child Id : " + str(solutionExtIdSuc))
                                solutionDetails = ElevateObservation.fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, solutionIdSuc,
                                                                                    accessToken)  
                                print(solutionDetails,"solutionDetailssurvey") 
                                scopeRoles = solutionDetails[0]
                                scopeSubRoles = solutionDetails[1]
                                verifiedRoles = ElevateObservation.validate_roles_against_api(scopeRoles, scopeSubRoles)
                                mainRoleproff = verifiedRoles[0]
                                rolesPGMID = verifiedRoles[1]
                                print("mainRole", mainRoleproff)
                                print("rolesPGMID4444", rolesPGMID)
                                scopeEntities = entitiesPGMID
                                print("scopeEntities4444", scopeEntities)
                                scope = {}
                                print(scope)
                                scope.update(entityHierarchy)
                                scope["professional_subroles"] = rolesPGMID
                                scope["professional_role"] = mainRoleproff
                                bodySolutionUpdate = {
                                "scope": scope
                                }
                                print("scope", bodySolutionUpdate)
                                ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionIdSuc, bodySolutionUpdate)
                                print(solutionDetails[2],solutionDetails[3],SurveyTemplateStartDate,SurveyTemplateEndDate)
                                solutionStartDate1 = ElevateObservation.convert_to_date(solutionDetails[2])
                                solutionEndDate1 = ElevateObservation.convert_to_date(solutionDetails[3])
                                SurveyTemplateStartDate1 = ElevateObservation.convert_to_date(SurveyTemplateStartDate)
                                SurveyTemplateEndDate1 = ElevateObservation.convert_to_date(SurveyTemplateEndDate)
                                print(solutionStartDate1,solutionEndDate1,SurveyTemplateStartDate1,SurveyTemplateEndDate1)

                                if SurveyTemplateStartDate1 == solutionStartDate1 and SurveyTemplateEndDate1 == solutionEndDate1:
                                    if solutionDetails[2]:
                                        startDateArr = str(solutionDetails[2]).split("-")
                                        bodySolutionUpdate = {
                                            "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                                        ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionIdSuc, bodySolutionUpdate)
                                    if solutionDetails[3]:
                                        endDateArr = str(solutionDetails[3]).split("-")
                                        bodySolutionUpdate = {
                                            "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                                        ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionIdSuc, bodySolutionUpdate)
                                    print('Survey Successfully Added')
                                    # print(surveySolutionlink)
                                    # surveySolutionlink = ElevateObservation.prepareProgramSuccessSheet(MainFilePath, parentFolder, programFile, solutionExtIdSuc,
                                    #                     solutionIdSuc, accessToken)
                                    surveySolutionlink = "https: Deeplink for survey created successfully."
                                    return surveySolutionlink
                                else:
                                    errorVar = "The survey Template start date and end date do not match the start date and end date at the Program Template."
                                    return errorVar
                            else:
                                print('Program Mapping Failed')
                                error_message = ""
                                if responseSurveyProgramMappingApi.status_code in [400, 401, 403, 404, 422]:
                                    error_message = f"SurveyProgramMappingApi-Client Error {responseSurveyProgramMappingApi.status_code}: {responseSurveyProgramMappingApi.text}"
                                elif responseSurveyProgramMappingApi.status_code in [500, 502, 503, 504]:
                                    error_message = f"SurveyProgramMappingApi-Server Error {responseSurveyProgramMappingApi.status_code}: {responseSurveyProgramMappingApi.text}"
                                else:
                                    error_message = f"SurveyProgramMappingApi-Unexpected Error {responseSurveyProgramMappingApi.status_code}: {responseSurveyProgramMappingApi.text}"

                                messageArr = ["********* Program mapping api *********", "URL : " + urlSurveyProgramMapping,
                                            "Status code : " + str(responseSurveyProgramMappingApi.status_code),
                                            "Response : " + responseSurveyProgramMappingApi.text]
                                ElevateObservation.createAPILog(parentFolder, messageArr)
                                errorVar = error_message
                                print(error_message)
                                messageArr.append(f"Error Response: {error_message}")
                        else:
                            print('Creating Child API Failed')
                            error_message = ""
                            if responseImportSoluTemplateApi.status_code in [400, 401, 403, 404, 422]:
                                error_message = f"ImportSoluTemplateApi-Client Error {responseImportSoluTemplateApi.status_code}: {responseImportSoluTemplateApi.text}"
                            elif responseImportSoluTemplateApi.status_code in [500, 502, 503, 504]:
                                error_message = f"ImportSoluTemplateApi-Server Error {responseImportSoluTemplateApi.status_code}: {responseImportSoluTemplateApi.text}"
                            else:
                                error_message = f"ImportSoluTemplateApi-Unexpected Error {responseImportSoluTemplateApi.status_code}: {responseImportSoluTemplateApi.text}"

                            messageArr = ["********* Program mapping api *********", "URL : " + urlImportSoluTemplate,
                                        "Status code : " + str(responseImportSoluTemplateApi.status_code),
                                        "Response : " + responseImportSoluTemplateApi.text]
                            ElevateObservation.createAPILog(parentFolder, messageArr)
                            errorVar = error_message
                            print(error_message)
                            messageArr.append(f"Error Response: {error_message}")
                    else:
                        if responseQuestionUploadApi.status_code in [400, 401, 403, 404, 422]:
                            error_message = f"QuestionUploadApi-Client Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}"
                        elif responseQuestionUploadApi.status_code in [500, 502, 503, 504]:
                            error_message = f"QuestionUploadApi-Server Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}"
                        else:
                            error_message = f"QuestionUploadApi-Unexpected Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}"
                        print('QuestionUploadApi Failed')
                        messageArr = ["********* Question Upload api *********", "URL : " + urlQuestionsUploadApi,
                                    "Path : " + str(parentFolder) + str('/questionUpload/uploadSheet.csv'),
                                    "Status code : " + str(responseQuestionUploadApi.status_code),
                                    "Response : " + responseQuestionUploadApi.text]
                        ElevateObservation.createAPILog(parentFolder, messageArr)
                        errorVar = error_message
                        messageArr.append(f"Error Response: {error_message}")
            
                except Exception as e:
                    errorVar = error_message
                    ElevateObservation.createAPILog(parentFolder, [f"Exception: {str(e)}"])
        if errorVar == "":
            return surveySolutionlink
        else:
            return errorVar

    def NOPuploadSurveyQuestions(MainFilePath, parentFolder, wbSurvey, addObservationSolution, accessToken, surTempExtID, surTempSolID, millisecond, programFile):
        print("Upload Survey Questions Func Called....")
        # print(parentFolder,"4854")
        # wbSurvey = xlrd.open_workbook(wbSurvey, on_demand=True)
        # print(f"Type of wbSurvey: {type(wbSurvey)}")
        sheetNam = wbSurvey.sheet_names()
        # print(sheetNam,"4854")
        global surveySolutionlink, errorVar, programExternalId
        error_message = ""
        stDt = None
        enDt = None
        shCnt = 0
        for i in sheetNam:
            if i.strip().lower() == 'questions':
                sheetNam1 = wbSurvey.sheets()[shCnt]
            shCnt = shCnt + 1
        dataSort = [sheetNam1.row_values(i) for i in range(sheetNam1.nrows)]
        labels = dataSort[1]
        dataSort = dataSort[2:]
        dataSort.sort(key=lambda x: int(x[0]))
        openWorkBookSort1 = xl_copy(wbSurvey)
        sheet1 = openWorkBookSort1.add_sheet('questions_sequence_sorted')

        for idx, label in enumerate(labels):
            sheet1.write(0, idx, label)

        for idx_r, row in enumerate(dataSort):
            for idx_c, value in enumerate(row):
                sheet1.write(idx_r + 1, idx_c, value)
        newFileName = str(addObservationSolution)
        openWorkBookSort1.save(newFileName)
        openNewFile = xlrd.open_workbook(newFileName, on_demand=True)
        wbSurvey = openNewFile
        sheetNames = wbSurvey.sheet_names()
        # print("reached till here 4881")
        for sheet2 in sheetNames:
            if sheet2.strip().lower() == 'questions_sequence_sorted':
                questionsList = []
                questionsSheet = wbSurvey.sheet_by_name(sheet2.lower())
                keys2 = [questionsSheet.cell(0, col_index2).value for col_index2 in
                        range(questionsSheet.ncols)]
                for row_index2 in range(1, questionsSheet.nrows):
                    d2 = {keys2[col_index2]: questionsSheet.cell(row_index2, col_index2).value
                        for col_index2 in range(questionsSheet.ncols)}
                    questionsList.append(d2)
                questionSeqByEcmArr = []
                quesSeqCnt = 1.0
                questionUploadFieldnames = []
                questionUploadFieldnames = ['solutionId', 'instanceParentQuestionId','hasAParentQuestion', 'parentQuestionOperator','parentQuestionValue', 'parentQuestionId','externalId', 'question0', 'question1', 'tip','hint', 'instanceIdentifier', 'responseType','dateFormat', 'autoCapture', 'validation','validationIsNumber', 'validationRegex','validationMax', 'validationMin', 'file','fileIsRequired', 'fileUploadType','allowAudioRecording', 'minFileCount','maxFileCount', 'caption', 'questionGroup','modeOfCollection', 'accessibility', 'showRemarks','rubricLevel', 'isAGeneralQuestion', 'R1','R1-hint', 'R2', 'R2-hint', 'R3', 'R3-hint', 'R4','R4-hint', 'R5', 'R5-hint', 'R6', 'R6-hint', 'R7','R7-hint', 'R8', 'R8-hint', 'R9', 'R9-hint', 'R10','R10-hint', 'R11', 'R11-hint', 'R12', 'R12-hint','R13', 'R13-hint', 'R14', 'R14-hint', 'R15','R15-hint', 'R16', 'R16-hint', 'R17', 'R17-hint','R18', 'R18-hint', 'R19', 'R19-hint', 'R20','R20-hint', 'sectionHeader', 'page','questionNumber', '_arrayFields']

                for ques in questionsList:

                    questionFilePath = parentFolder + '/questionUpload/'
                    file_exists_ques = os.path.isfile(
                        parentFolder + '/questionUpload/uploadSheet.csv')
                    # print(questionFilePath,"4904")
                    if not os.path.exists(questionFilePath):
                        os.mkdir(questionFilePath)
                    with open(parentFolder + '/questionUpload/uploadSheet.csv', 'a',
                            encoding='utf-8') as questionUploadFile:
                        writerQuestionUpload = csv.DictWriter(questionUploadFile, fieldnames=questionUploadFieldnames, lineterminator='\n')
                        if not file_exists_ques:
                            writerQuestionUpload.writeheader()
                        questionFileObj = {}
                        surveyExternalId = None
                        questionFileObj['solutionId'] = surTempExtID
                        if ques['instance_parent_question_id'].encode('utf-8').decode('utf-8'):
                            questionFileObj['instanceParentQuestionId'] = ques[
                                                                            'instance_parent_question_id'].strip() + '_' + str(
                                millisecond)
                        else:
                            questionFileObj['instanceParentQuestionId'] = 'NA'
                        if ques['parent_question_id'].encode('utf-8').decode('utf-8').strip():
                            questionFileObj['hasAParentQuestion'] = 'YES'
                            if ques['show_when_parent_question_value_is'] == 'or':
                                questionFileObj['parentQuestionOperator'] = '||'
                            else:
                                questionFileObj['parentQuestionOperator'] = ques['show_when_parent_question_value_is']
                            if type(ques['parent_question_value']) != str:
                                if (ques['parent_question_value'] and ques[
                                    'parent_question_value'].is_integer() == True):
                                    questionFileObj['parentQuestionValue'] = int(ques['parent_question_value'])
                                elif (ques['parent_question_value'] and ques[
                                    'parent_question_value'].is_integer() == False):
                                    questionFileObj['parentQuestionValue'] = ques['parent_question_value']
                            else:
                                questionFileObj['parentQuestionValue'] = ques['parent_question_value']
                                questionFileObj['parentQuestionId'] = ques['parent_question_id'].encode('utf-8').decode('utf-8').strip() + '_' + str(
                                    millisecond)
                        else:
                            questionFileObj['hasAParentQuestion'] = 'NO'
                            questionFileObj['parentQuestionOperator'] = None
                            questionFileObj['parentQuestionValue'] = None
                            questionFileObj['parentQuestionId'] = None
                        questionFileObj['externalId'] = ques['question_id'].strip() + '_' + str(millisecond)
                        if quesSeqCnt == ques['question_sequence']:
                            questionSeqByEcmArr.append(ques['question_id'].strip() + '_' + str(millisecond))
                            quesSeqCnt = quesSeqCnt + 1.0
                        if ques['question_language1']:
                            questionFileObj['question0'] = ques['question_language1'].encode('utf-8').decode('utf-8')
                        else:
                            questionFileObj['question0'] = None
                        if ques['question_language2']:
                            questionFileObj['question1'] = ques['question_language2'].encode('utf-8').decode('utf-8')
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
                        if ques['question_response_type'].strip().lower() == 'date':
                            questionFileObj['dateFormat'] = "DD-MM-YYYY"
                        else:
                            questionFileObj['dateFormat'] = None
                        if ques['question_response_type'].strip().lower() == 'date':
                            if ques['date_auto_capture'] and ques['date_auto_capture'] == 1:
                                questionFileObj['autoCapture'] = 'TRUE'
                            elif ques['date_auto_capture'] and ques['date_auto_capture'] == 0:
                                questionFileObj['autoCapture'] = 'false'
                            else:
                                questionFileObj['autoCapture'] = 'false'
                        else:
                            questionFileObj['autoCapture'] = None
                        if ques['response_required']:
                            if ques['response_required'] == 1:
                                questionFileObj['validation'] = 'TRUE'
                            elif ques['response_required'] == 0:
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
                        if ques['file_upload'] == 1:
                            questionFileObj['file'] = 'Snapshot'
                            questionFileObj['fileIsRequired'] = 'TRUE'
                            questionFileObj['fileUploadType'] = 'image/jpeg,docx,pdf,ppt'
                            questionFileObj['minFileCount'] = 0
                            questionFileObj['maxFileCount'] = 10
                        elif ques['file_upload'] == 0:
                            questionFileObj['file'] = 'NA'
                            questionFileObj['fileIsRequired'] = None
                            questionFileObj['fileUploadType'] = None
                            questionFileObj['minFileCount'] = None
                            questionFileObj['maxFileCount'] = None

                        questionFileObj['caption'] = 'FALSE'
                        questionFileObj['questionGroup'] = 'A1'
                        questionFileObj['modeOfCollection'] = 'onfield'
                        questionFileObj['accessibility'] = 'No'
                        if ques['show_remarks'] == 1:
                            questionFileObj['showRemarks'] = 'TRUE'
                        elif ques['show_remarks'] == 0:
                            questionFileObj['showRemarks'] = 'FALSE'
                        questionFileObj['rubricLevel'] = None
                        questionFileObj['isAGeneralQuestion'] = None
                        if ques['question_response_type'].strip().lower() == 'radio' or ques[
                            'question_response_type'].strip() == 'multiselect':
                            for quesIndex in range(1, 21):
                                if type(ques['response(R' + str(quesIndex) + ')']) != str:
                                    if (ques['response(R' + str(quesIndex) + ')'] and ques[
                                        'response(R' + str(quesIndex) + ')'].is_integer() == True):
                                        questionFileObj['R' + str(quesIndex) + ''] = int(
                                            ques['response(R' + str(quesIndex) + ')'])
                                    elif (ques['response(R' + str(quesIndex) + ')'] and ques[
                                        'response(R' + str(quesIndex) + ')'].is_integer() == False):
                                        questionFileObj['R' + str(quesIndex) + ''] = ques[
                                            'response(R' + str(quesIndex) + ')']
                                else:
                                    questionFileObj['R' + str(quesIndex) + ''] = ques[
                                        'response(R' + str(quesIndex) + ')']

                                if type(ques['response(R' + str(quesIndex) + ')_hint']) != str:
                                    if (ques['response(R' + str(quesIndex) + ')_hint'] and ques[
                                        'response(R' + str(quesIndex) + ')_hint'].is_integer() == True):
                                        questionFileObj['R' + str(quesIndex) + '-hint'] = int(
                                            ques['response(R' + str(quesIndex) + ')_hint'])
                                    elif (ques['response(R' + str(quesIndex) + ')_hint'] and ques[
                                        'response(R' + str(quesIndex) + ')_hint'].is_integer() == False):
                                        questionFileObj['R' + str(quesIndex) + '-hint'] = ques[
                                            'response(R' + str(quesIndex) + ')_hint']
                                else:
                                    questionFileObj['R' + str(quesIndex) + '-hint'] = ques[
                                        'response(R' + str(quesIndex) + ')_hint']
                                questionFileObj['_arrayFields'] = 'parentQuestionValue'
                        else:
                            for quesIndex in range(1, 21):
                                questionFileObj['R' + str(quesIndex)] = None
                                questionFileObj['R' + str(quesIndex) + '-hint'] = None
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
                        writerQuestionUpload.writerow(questionFileObj)
                try:        
                    urlQuestionsUploadApi = internal_kong_ip + questionuploadapiurl
                    headerQuestionUploadApi = {
                        "internal-access-token": internal_access_token,
                        'Authorization': authorization,
                        'X-authenticated-user-token': accessToken,
                        'X-Channel-id': x_channel_id
                    }
                    filesQuestion = {
                        'questions': open(parentFolder + '/questionUpload/uploadSheet.csv', 'rb')
                    }
                    responseQuestionUploadApi = requests.post(url=urlQuestionsUploadApi,
                                                            headers=headerQuestionUploadApi, files=filesQuestion)
                    if responseQuestionUploadApi.status_code == 200:
                        print('Question upload Success')

                        messageArr = ["********* Question Upload api *********", "URL : " + urlQuestionsUploadApi,
                                    "Path : " + str(parentFolder) + str('/questionUpload/uploadSheet.csv'),
                                    "Status code : " + str(responseQuestionUploadApi.status_code),
                                    "Response : " + responseQuestionUploadApi.text]
                        ElevateObservation.createAPILog(parentFolder, messageArr)
                        messageArr1 = ["Questions","Question upload Success","Passed",str(responseQuestionUploadApi.status_code)]
                        ElevateObservation.apicheckslog(parentFolder,messageArr1)

                        with open(parentFolder + '/questionUpload/uploadInternalIdsSheet.csv', 'w+',encoding='utf-8') as questionRes:
                            questionRes.write(responseQuestionUploadApi.text)
                        urlImportSoluTemplate = internal_kong_ip + importsurveysolutiontemplateurl + str(surTempSolID) + "?appName=manage-learn"
                        headerImportSoluTemplateApi = {
                            'X-auth-token': accessToken,
                            'X-Channel-id': x_channel_id,
                            'internal-access-token': internal_access_token,
                            'tenantId': tenantID ,
                            'orgid': orgIDFromTemplate,
                            adminTokenHeaderName: adminAccessToken
                        }
                        responseImportSoluTemplateApi = requests.get(url=urlImportSoluTemplate,
                                                                    headers=headerImportSoluTemplateApi)
                        if responseImportSoluTemplateApi.status_code == 200:
                            print('Creating Child Success')

                            messageArr = ["********* Creating Child api *********", "URL : " + urlImportSoluTemplate,
                                        "Status code : " + str(responseImportSoluTemplateApi.status_code),
                                        "Response : " + responseImportSoluTemplateApi.text]
                            ElevateObservation.createAPILog(parentFolder, messageArr)
                            responseImportSoluTemplateApi = responseImportSoluTemplateApi.json()
                            solutionIdSuc = responseImportSoluTemplateApi["result"]["solutionId"]
                            urlSurveyProgramMapping = internal_kong_ip + importsurveysolutiontoprogramurl + str(solutionIdSuc) + "?programId=" + programExternalId.lstrip().rstrip()
                            headeSurveyProgramMappingApi = {
                                'X-auth-token': accessToken,
                                'X-Channel-id': x_channel_id,
                                'internal-access-token': internal_access_token,
                                'tenantId': tenantID ,
                                'orgid': orgIDFromTemplate,
                                adminTokenHeaderName: adminAccessToken
                            }
                            responseSurveyProgramMappingApi = requests.get(url=urlSurveyProgramMapping,headers=headeSurveyProgramMappingApi)
                            if responseSurveyProgramMappingApi.status_code == 200:
                                print('Program Mapping Success')
                                
                                messageArr = ["********* Program mapping api *********", "URL : " + urlSurveyProgramMapping,
                                            "Status code : " + str(responseSurveyProgramMappingApi.status_code),
                                            "Response : " + responseSurveyProgramMappingApi.text]
                                ElevateObservation.createAPILog(parentFolder, messageArr)
                                surveyLink = None
                                solutionIdSuc = None
                                surveyExternalIdSuc = None
                                surveyLink = responseImportSoluTemplateApi["result"]["link"]
                                solutionIdSuc = responseImportSoluTemplateApi["result"]["solutionId"]
                                solutionExtIdSuc = responseImportSoluTemplateApi["result"]["solutionExternalId"]
                                print("Survey Child Id : " + str(solutionExtIdSuc))
                                solutionDetails = ElevateObservation.fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, solutionIdSuc,
                                                                                    accessToken)
                                solutionStartDate1 = ElevateObservation.convert_to_date(solutionDetails[2])
                                solutionEndDate1 = ElevateObservation.convert_to_date(solutionDetails[3])
                                SurveyTemplateStartDate1 = ElevateObservation.convert_to_date(SurveyTemplateStartDate)
                                SurveyTemplateEndDate1 = ElevateObservation.convert_to_date(SurveyTemplateEndDate)
                                if SurveyTemplateStartDate1 == solutionStartDate1 and SurveyTemplateEndDate1 == solutionEndDate1:
                                    scopeRoles = solutionDetails[0]
                                    scopeSubRoles = solutionDetails[1]
                                    verifiedRoles = ElevateObservation.validate_roles_against_api(scopeRoles, scopeSubRoles)
                                    mainRoleproff = verifiedRoles[0]
                                    rolesPGMID = verifiedRoles[1]
                                    print("mainRole", mainRoleproff)
                                    print("rolesPGMID-------13", rolesPGMID)
                                    scopeEntities = entitiesPGMID
                                    scope = {}
                                    for i in range(len(entitiesType)):
                                        entity_type = entitiesType[i]
                                        entity_value = scopeEntities[i]
                                        if entity_type in scope:
                                            scope[entity_type].append(entity_value)
                                        else:
                                            scope[entity_type] = [entity_value]
                                    scope["professional_subroles"] = rolesPGMID
                                    scope["professional_role"] = mainRoleproff
                                    bodySolutionUpdate = {
                                    "scope": scope
                                    }
                                    print(bodySolutionUpdate,"bodySolutionUpdate")
                                    ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionIdSuc, bodySolutionUpdate)
                                    if solutionDetails[2]:
                                        startDateArr = str(solutionDetails[2]).split("-")
                                        bodySolutionUpdate = {
                                            "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                                        ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionIdSuc, bodySolutionUpdate)
                                    if solutionDetails[3]:
                                        endDateArr = str(solutionDetails[3]).split("-")
                                        bodySolutionUpdate = {
                                            "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                                        ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionIdSuc, bodySolutionUpdate)
                                    surveySolutionlink = "Survey solution Created Successfully."
                                
                                    print('Survey Successfully Added')
                                    print(surveySolutionlink)
                                else:
                                    errorVar = "The survey Template start date and end date do not match the start date and end date at the Program Template."
                                    return errorVar
                            else:
                                print('Program Mapping Failed')
                                error_message = ""
                                if responseSurveyProgramMappingApi.status_code in [400, 401, 403, 404, 422]:
                                    error_message = f"SurveyProgramMappingApi-Client Error {responseSurveyProgramMappingApi.status_code}: {responseSurveyProgramMappingApi.text}"
                                elif responseSurveyProgramMappingApi.status_code in [500, 502, 503, 504]:
                                    error_message = f"SurveyProgramMappingApi-Server Error {responseSurveyProgramMappingApi.status_code}: {responseSurveyProgramMappingApi.text}"
                                else:
                                    error_message = f"SurveyProgramMappingApi-Unexpected Error {responseSurveyProgramMappingApi.status_code}: {responseSurveyProgramMappingApi.text}"

                                messageArr = ["********* Program mapping api *********", "URL : " + urlSurveyProgramMapping,
                                            "Status code : " + str(responseSurveyProgramMappingApi.status_code),
                                            "Response : " + responseSurveyProgramMappingApi.text]
                                ElevateObservation.createAPILog(parentFolder, messageArr)
                                errorVar = error_message
                                print(error_message)
                                messageArr.append(f"Error Response: {error_message}")
                        else:
                            print('Creating Child API Failed')
                            error_message = ""
                            if responseImportSoluTemplateApi.status_code in [400, 401, 403, 404, 422]:
                                error_message = f"ImportSoluTemplateApi-Client Error {responseImportSoluTemplateApi.status_code}: {responseImportSoluTemplateApi.text}"
                            elif responseImportSoluTemplateApi.status_code in [500, 502, 503, 504]:
                                error_message = f"ImportSoluTemplateApi-Server Error {responseImportSoluTemplateApi.status_code}: {responseImportSoluTemplateApi.text}"
                            else:
                                error_message = f"ImportSoluTemplateApi-Unexpected Error {responseImportSoluTemplateApi.status_code}: {responseImportSoluTemplateApi.text}"

                            messageArr = ["********* Program mapping api *********", "URL : " + urlImportSoluTemplate,
                                        "Status code : " + str(responseImportSoluTemplateApi.status_code),
                                        "Response : " + responseImportSoluTemplateApi.text]
                            ElevateObservation.createAPILog(parentFolder, messageArr)
                            errorVar = error_message
                            print(error_message)
                            messageArr.append(f"Error Response: {error_message}")
                    else:
                        if responseQuestionUploadApi.status_code in [400, 401, 403, 404, 422]:
                            error_message = f"QuestionUploadApi-Client Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}"
                        elif responseQuestionUploadApi.status_code in [500, 502, 503, 504]:
                            error_message = f"QuestionUploadApi-Server Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}"
                        else:
                            error_message = f"QuestionUploadApi-Unexpected Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}"
                        print('QuestionUploadApi Failed')
                        messageArr = ["********* Question Upload api *********", "URL : " + urlQuestionsUploadApi,
                                    "Path : " + str(parentFolder) + str('/questionUpload/uploadSheet.csv'),
                                    "Status code : " + str(responseQuestionUploadApi.status_code),
                                    "Response : " + responseQuestionUploadApi.text]
                        ElevateObservation.createAPILog(parentFolder, messageArr)
                        errorVar = error_message
                        messageArr.append(f"Error Response: {error_message}")
            
                except Exception as e:
                    errorVar = error_message
                    ElevateObservation.createAPILog(parentFolder, [f"Exception: {str(e)}"])
        if errorVar == "":
            return surveySolutionlink
        else:
            return errorVar
        
    def validate_roles_against_api(mainRoles, subRoles):
        print(mainRoles, "mainRoles")
        print(subRoles, "subRoles")

        urlFetchRoleList = userLoginHost + fetchprofessionalRole
        headers = {
            'Content-Type': content_type,
            'tenantId': tenantID,
            'X-Channel-id': x_channel_id,
        }
        payload = {}
        print(urlFetchRoleList,"urlFetchRoleList")
        print(headers,"headers")
        response = requests.request("GET", urlFetchRoleList, headers=headers, data=payload)
        # response = requests.get(urlFetchRoleList, headers=headers, data=json.dumps({}))
        print(response.text,"response4670")
        messageArr = []
        messageArr.append("Fetched professional roles from: " + urlFetchRoleList)
        messageArr.append("Status Code: " + str(response.status_code))

        if response.status_code != 200:
            messageArr.append("Error fetching roles.")
            print("Error fetching roles.")
            return [], []

        role_data = response.json()
        validated_main_role_ids = []
        validated_subrole_ids_list = []

        remaining_subroles = [s.strip() for s in subRoles]

        for role in mainRoles:
            matched = next((r for r in role_data['result']
                            if r.get('externalId', '').strip() == role.strip() or
                            r.get('name', '').strip() == role.strip()), None)

            if matched:
                main_role_id = matched['_id']
                validated_main_role_ids.append(main_role_id)
                # subrole_url = requests.request("GET", userLoginHost+ "entity-management/v1/entities/subEntityList/" +main_role_id+"?type=professional_subroles", headers=headers, data=payload)
                subrole_url = f"{userLoginHost}entity-management/v1/entities/subEntityList/{main_role_id}?type=professional_subroles"
                subrole_resp = requests.request("GET",subrole_url, headers=headers,data=payload)
                if subrole_resp.status_code == 200:
                    subroles_data = subrole_resp.json()
                    for item in subroles_data['result']['data']:
                        sub_external_id = item.get('externalId', '').strip()
                        sub_name = item.get('name', '').strip()
                        if sub_external_id in remaining_subroles or sub_name in remaining_subroles:
                            validated_subrole_ids_list.append(item['_id'])
                            if sub_external_id in remaining_subroles:
                                remaining_subroles.remove(sub_external_id)
                            elif sub_name in remaining_subroles:
                                remaining_subroles.remove(sub_name)
                            print(f"Subrole '{sub_external_id}' validated under mainRole '{role}'")
                else:
                    messageArr.append(f"Failed to fetch subroles for mainRole '{role}'")
            else:
                messageArr.append(f"MainRole '{role}' not found in API")

        for s in remaining_subroles:
            messageArr.append(f"Subrole '{s}' not found in any of the provided mainRoles.")

        for msg in messageArr:
            print(msg)

        return validated_main_role_ids, validated_subrole_ids_list   

    def mainFunc(MainFilePath, programFile, addObservationSolution, millisecond, isProgramnamePresent, isCourse,
             scopeEntityType=scopeEntityType):
        print("entering mainFUnc")
        global errorVar,pointBasedValue,solutionDict,allow_multiple_submissions,creator,userEntity,criteriaLevelsReport,isExternalProgram,orgIDFromTemplate,tenantID,orgIdForScope
        errorVar = ""
        scopeEntityType = scopeEntityType
        if not isCourse:
            parentFolder = ElevateObservation.createFileStruct(MainFilePath, addObservationSolution)
            accessToken = ElevateObservation.generateAccessToken(parentFolder)
            print(accessToken,"4538")
            wbPgm =xlrd.open_workbook(programFile, on_demand=True)
            ElevateObservation.validateTenantAndOrgIdsFromProgramSheet(wbPgm)
            typeofSolution = ElevateObservation.typeofresource(addObservationSolution, accessToken, parentFolder)
            if typeofSolution == 0:
                result = {
                    "solutionDict": solutionDict,
                    "programName": programName 
                }
                return json.dumps(result)
            print(typeofSolution,"this is type of solution")
            
            # typeofSolution = validateSheets(addObservationSolution, accessToken, parentFolder)
            # sys.exit()
            wbObservation = xlrd.open_workbook(addObservationSolution, on_demand=True)
            projectSheetNames = wbObservation.sheet_names()
            wbPgm=xlrd.open_workbook(programFile, on_demand=True)
            programSheetNames = wbPgm.sheet_names()
            for programSheets in programSheetNames:
                if programSheets.strip().lower() == 'program details':
                    print("Checking program details sheet...")
                    programDetailsSheet = wbPgm.sheet_by_name(programSheets)
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
            if not ElevateObservation.programsFileCheck(programFile, accessToken, parentFolder, MainFilePath):
                print("---> no program found / unable to create program....")
                result = {programName : errorVar}
                return result
            
            if typeofSolution == 1 or typeofSolution == 5:
                if typeofSolution == 5:
                    impLedObsFlag = True
                else:
                    impLedObsFlag = False
            # print(impLedObsFlag,"impLedObsFlag")
            for sheets in projectSheetNames:
                if sheets.strip().lower() == 'details'.lower() and typeofSolution in [1, 5]:
                    ResourceSheet = wbObservation.sheet_by_name(sheets)
                    keysEnv = [ResourceSheet.cell(1, col_index_env).value for col_index_env in range(ResourceSheet.ncols)]
                    dictDetailsEnv = {keysEnv[col_index_env]: ResourceSheet.cell(row_index_env, col_index_env).value
                                    for col_index_env in range(ResourceSheet.ncols)}
                    ObsWRResourceName = dictDetailsEnv['observation_solution_name'].encode('utf-8').decode('utf-8')
                    Entity_To_Upload = dictDetailsEnv['entity_type'].encode('utf-8').decode('utf-8')
                    print("Entity_To_Upload", Entity_To_Upload)
                    try:
                        if not ElevateObservation.ObsWRValidate(wbObservation, accessToken, parentFolder,typeofSolution,):
                            print("Error during validation of Observation with Rubric file ....")
                            finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                            return finalObsRubricSolutionLink   
                        print("validation successful---------")
                        def addObsWRFunc(parentFolder, wbObservation, millisecond, accessToken):
                            if not ElevateObservation.criteriaUpload(parentFolder, wbObservation, millisecond, accessToken, "framework", impLedObsFlag):
                                finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                return finalObsRubricSolutionLink
                            print("Criteria Upload success....")
                            # userDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dikshaLoginId)
                            # if not userDetails:
                            #     finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                            #     return finalObsRubricSolutionLink
                            # matchedShikshalokamLoginId = userDetails[0]
                            
                            frameworkExternalId = ElevateObservation.frameWorkUpload(parentFolder, wbObservation, millisecond, accessToken)
                            if not frameworkExternalId:
                                finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                return finalObsRubricSolutionLink
                            observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
                            if not ElevateObservation.themesUpload(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, False):
                                finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                return finalObsRubricSolutionLink
                            solutionId = ElevateObservation.createSolutionFromFramework(parentFolder, accessToken, frameworkExternalId)
                            if not solutionId:
                                finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                return finalObsRubricSolutionLink
                            
                        
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
                                if 'Is ECM Mandatory?' in dictECMs and dictECMs['Is ECM Mandatory?'] is not None:
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
                                print(ecm_update[EMC_ID])
                                ecmSeqCount += 1
                            ecm_dict['evidenceMethods'] = ecm_update
                            bodySolutionUpdate = ecm_dict
                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                                finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                return finalObsRubricSolutionLink
                            bodySolutionUpdate = {"sections": section}
                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                                finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                return finalObsRubricSolutionLink
                            print(Entity_To_Upload,"Entity_To_Upload222")
                            if Entity_To_Upload.strip().lower() in ['state', 'district', 'block', 'cluster', 'school']:
                                parentEntityKey = "state"

                            else:
                                parentEntityKey = None
                               
                            bodySolutionUpdate = {"status": "active", "isDeleted": False, "criteriaLevelReport": criteriaLevelsReport,"parentEntityKey": parentEntityKey}
                            print("bodySolutionUpdate", bodySolutionUpdate)
                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                                finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                return finalObsRubricSolutionLink
                            excelBook = open_workbook(addObservationSolution)
                            if not ElevateObservation.questionUpload(addObservationSolution, parentFolder, frameworkExternalId, millisecond, accessToken,solutionId,typeofSolution):
                                finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                return finalObsRubricSolutionLink
                            if not pointBasedValue.lower() == "null":
                                bodySolutionUpdate = {"isRubricDriven": True}
                                if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                    return finalObsRubricSolutionLink
                                if not ElevateObservation.fetchSolutionCriteria(parentFolder, observationExternalId, accessToken):
                                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                    return finalObsRubricSolutionLink
                                if not ElevateObservation.uploadCriteriaRubrics(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, True):
                                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                    return finalObsRubricSolutionLink
                                if not ElevateObservation.uploadThemeRubrics(parentFolder, wbObservation, accessToken, frameworkExternalId, True):
                                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                    return finalObsRubricSolutionLink
                            else:
                                print("Observation with scoring system : null.")
                            bodySolutionUpdate = {'allowMultipleAssessemts': allow_multiple_submissions, "creator": creator,"parentEntityKey": parentEntityKey}
                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                                finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                return finalObsRubricSolutionLink
                            # solutionDetails = ElevateObservation.fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, solutionId, accessToken)
                            # if not solutionDetails:
                            #     finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                            #     return finalObsRubricSolutionLink
                            # if solutionDetails[1]:
                            #     startDateArr = str(solutionDetails[1]).split("-")
                            #     bodySolutionUpdate = {
                            #         "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                            #     if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                            #         finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                            #         return finalObsRubricSolutionLink
                            # if solutionDetails[2]:
                            #     print(solutionDetails[2],"this is 5294")
                            #     endDateArr = str(solutionDetails[2]).split("-")
                            #     bodySolutionUpdate = {
                            #         "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                            #     if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                            #         finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                            #         return finalObsRubricSolutionLink
                            if isProgramnamePresent:
                                childId = ElevateObservation.createChild(parentFolder, observationExternalId, accessToken)
                                if not childId:
                                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                    return finalObsRubricSolutionLink
                                if childId[0]:
                                    print("fetching details")
                                    solutionDetails = ElevateObservation.fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, childId[0],
                                                                                        accessToken)
                                    print(solutionDetails,"solutionDetails")
                                    if not solutionDetails:
                                        finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                        return finalObsRubricSolutionLink
                                    scopeRoles = solutionDetails[0]
                                    scopeSubRoles = solutionDetails[1]
                                    verifiedRoles = ElevateObservation.validate_roles_against_api(scopeRoles, scopeSubRoles)
                                    mainRoleproff = verifiedRoles[0]
                                    rolesPGMID = verifiedRoles[1]
                                    print("mainRole", mainRoleproff)
                                    print("rolesPGMID--------22", rolesPGMID)
                                    scopeEntities = entitiesPGMID
                                    print("scopeEntities", scopeEntities)
                                    print("entitiesType", entitiesType)
                                    scope = {}
                                    scope.update(entityHierarchy)
                                    scope["professional_subroles"] = rolesPGMID
                                    scope["professional_role"] = mainRoleproff
                                    bodySolutionUpdate = {
                                    "scope": scope
                                    }

                                    if not ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate):
                                        finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                        return finalObsRubricSolutionLink
                                    ReffstartDateOfProgram1 = ElevateObservation.convert_to_date(ReffstartDateOfProgram)
                                    ReffendDateOfProgram1 = ElevateObservation.convert_to_date(ReffendDateOfProgram)
                                    solutionDetails2 = ElevateObservation.convert_to_date(solutionDetails[2])
                                    solutionDetails3 = ElevateObservation.convert_to_date(solutionDetails[3])
                                    if ReffstartDateOfProgram1 <= solutionDetails2 <= ReffendDateOfProgram1 and ReffstartDateOfProgram1 <= solutionDetails3 <= ReffendDateOfProgram1:
                                        if solutionDetails[2]:
                                            startDateArr = str(solutionDetails[2]).split("-")
                                            bodySolutionUpdate = {
                                                "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[
                                                    0] + " 00:00:00"}
                                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate):
                                                finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                                return finalObsRubricSolutionLink
                                        if solutionDetails[3]:
                                            endDateArr = str(solutionDetails[3]).split("-")
                                            bodySolutionUpdate = {
                                                "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate):
                                                finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                                return finalObsRubricSolutionLink
                                    else:
                                        errorVar = "Date Mismatched! Creation Stopped."
                                        finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                        return finalObsRubricSolutionLink
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
                                finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                return finalObsRubricSolutionLink
                        millisecond = int(time.time() * 1000)
                        ObsWRSolutionLink = addObsWRFunc(parentFolder, wbObservation, millisecond, accessToken)
                        return ObsWRSolutionLink
                    except Exception as e:
                        print(f"Error occurred during project creation: {str(e)}")
                        # raise RuntimeError("The project creation failed due to an unexpected error")
                        solutionError = str(e)
                        print(errorVar,"3266-0---")
                        if errorVar == "":
                            finalObsRubricSolutionLink = {ObsWRResourceName: solutionError}
                        else:
                            finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                        # print(projectSolutionLink, "3247")
                        return finalObsRubricSolutionLink
                
                elif typeofSolution == 2 and sheets.strip().lower() == 'details'.lower():
                    ResourceSheet = wbObservation.sheet_by_name(sheets)
                    keysEnv = [ResourceSheet.cell(1, col_index_env).value for col_index_env in range(ResourceSheet.ncols)]
                    
                    # Collect observation solution details
                    dictDetailsEnv = {
                        keysEnv[col_index_env]: ResourceSheet.cell(row_index_env, col_index_env).value
                        for col_index_env in range(ResourceSheet.ncols)
                    }
                    ObsWORResourceName = dictDetailsEnv['observation_solution_name'].encode('utf-8').decode('utf-8')
                    pointBasedValue = "null"
                    Entity_To_Upload = dictDetailsEnv['entity_type'].encode('utf-8').decode('utf-8')
                    print("Entity_To_Upload", Entity_To_Upload)

                    if Entity_To_Upload.strip().lower() in ['state', 'district', 'block', 'cluster', 'school']:
                        parentEntityKey = "state"

                    else:
                        parentEntityKey = None
                    try:
                        def addObsWORFunc(parentFolder, wbObservation, millisecond, accessToken):
                            if not ElevateObservation.ObsWORValidate(wbObservation, accessToken, parentFolder):
                                print(errorVar,"---->validation error 3931")
                                ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                return ObsWORSolutionLink
                            print("Create Observation Function called ....")
                            if not ElevateObservation.criteriaUpload(parentFolder, wbObservation, millisecond, accessToken, "criteria", False):
                                ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                return ObsWORSolutionLink
                            
                            # userDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dikshaLoginId)
                            # if not userDetails:
                            #     ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                            #     return ObsWORSolutionLink
                            # matchedShikshalokamLoginId = userDetails[0]
                            
                            frameworkExternalId = ElevateObservation.frameWorkUpload(parentFolder, wbObservation, millisecond, accessToken)
                            if not frameworkExternalId:
                                ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                return ObsWORSolutionLink
                            observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
                            if not ElevateObservation.themesUpload(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, True):
                                ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                return ObsWORSolutionLink
                            solutionId = ElevateObservation.createSolutionFromFramework(parentFolder, accessToken, frameworkExternalId)
                            if not solutionId:
                                ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                return ObsWORSolutionLink
                            sectionsObj = {"sections": {'S1': 'Observation Question'}}
                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, sectionsObj):
                                ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                return ObsWORSolutionLink
                            ecmObj = {}
                            ecmExternalId = None
                            ecmObj = {
                                "evidenceMethods": {'OB': {'externalId': 'OB', 'tip': None, 'name': 'Observation', 'description': None,
                                                        'modeOfCollection': 'onfield', 'canBeNotApplicable': False,
                                                        'notApplicable': False, 'canBeNotAllowed': False, 'remarks': None}}}
                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, ecmObj):
                                ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                return ObsWORSolutionLink
                            if not ElevateObservation.questionUpload(addObservationSolution, parentFolder, frameworkExternalId, millisecond, accessToken,
                                        solutionId, typeofSolution):
                                ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                return ObsWORSolutionLink

                            if not pointBasedValue.lower() == "null":
                                if not ElevateObservation.uploadCriteriaRubrics(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, False):
                                    ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                    return ObsWORSolutionLink
                                if not ElevateObservation.uploadThemeRubrics(parentFolder, wbObservation, accessToken, frameworkExternalId, False):
                                    ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                    return ObsWORSolutionLink
                            bodySolutionUpdate = {"status": "active", "isDeleted": False, "allowMultipleAssessemts": True,
                                                "creator": creator,"parentEntityKey": parentEntityKey}
                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                                ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                return ObsWORSolutionLink

                            # solutionDetails = ElevateObservation.fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, solutionId, accessToken)
                            # # Below script will convert date DD-MM-YYYY TO YYYY-MM-DD 00:00:00 to match the code syntax
                            # if not solutionDetails:
                            #     ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                            #     return ObsWORSolutionLink
                            # if solutionDetails[1]:
                            #     startDateArr = str(solutionDetails[1]).split("-")
                            #     bodySolutionUpdate = {
                            #         "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                            #     if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                            #         ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                            #     return ObsWORSolutionLink                            #     endDateArr = str(solutionDetails[2]).split("-")

                            # print(solutionDetails[2],"solutionDetails[2]")
                            # if solutionDetails[2]:
                            #     bodySolutionUpdate = {
                            #         "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                            #     if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                            #         ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                            #     return ObsWORSolutionLink
                            if isProgramnamePresent:
                                childId = ElevateObservation.createChild(parentFolder, observationExternalId, accessToken)
                                if not childId:
                                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                    return finalObsRubricSolutionLink
                                if childId[0]:
                                    solutionDetails = ElevateObservation.fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, childId[0],
                                                                                        accessToken)
                                    
                                    print(solutionDetails,"solutionDetails line 5064")
                                    if not solutionDetails:
                                        ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                        return ObsWORSolutionLink
                                    scopeRoles = solutionDetails[0]

                                    print("scopeRoles line 5070", scopeRoles)
                                    scopeSubRoles = solutionDetails[1]
                                    verifiedRoles = ElevateObservation.validate_roles_against_api(scopeRoles, scopeSubRoles)
                                    mainRoleproff = verifiedRoles[0]
                                    rolesPGMID = verifiedRoles[1]
                                    print("mainRole", mainRoleproff)
                                    print("rolesPGMID1111", rolesPGMID)
                                    scopeEntities = entitiesPGMID
                                    scope = {}
                                    scope.update(entityHierarchy)
                                    scope["professional_subroles"] = rolesPGMID
                                    scope["professional_role"] = mainRoleproff
                                    bodySolutionUpdate = {
                                    "scope": scope
                                    }
                                    if not ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate):
                                        ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                        return ObsWORSolutionLink
                                    ReffstartDateOfProgram1 = ElevateObservation.convert_to_date(ReffstartDateOfProgram)
                                    ReffendDateOfProgram1 = ElevateObservation.convert_to_date(ReffendDateOfProgram)
                                    solutionDetails2 = ElevateObservation.convert_to_date(solutionDetails[2])
                                    solutionDetails3 = ElevateObservation.convert_to_date(solutionDetails[3])
                                    if ReffstartDateOfProgram1 <= solutionDetails2 <= ReffendDateOfProgram1 and ReffstartDateOfProgram1 <= solutionDetails3 <= ReffendDateOfProgram1:
                                        if solutionDetails[2]:
                                            startDateArr = str(solutionDetails[2]).split("-")
                                            bodySolutionUpdate = {
                                                "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[
                                                    0] + " 00:00:00"}
                                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate):
                                                ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                                return ObsWORSolutionLink
                                        if solutionDetails[3]:
                                            endDateArr = str(solutionDetails[3]).split("-")
                                            bodySolutionUpdate = {
                                                "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate):
                                                ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                                return ObsWORSolutionLink
                                    else:
                                        errorVar = "Date Mismatched! Creation Stopped."
                                        ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                        return ObsWORSolutionLink
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
                                ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                return ObsWORSolutionLink

                        millisecond = int(time.time() * 1000)
                        ObsWORSolutionLink = addObsWORFunc(parentFolder, wbObservation, millisecond, accessToken)
                        return ObsWORSolutionLink
                    except Exception as e:
                        print(f"Error occurred: {str(e)}")
                        solutionError = str(e)
                        print(errorVar,"3266---1111")
                        if errorVar == "":
                            ObsWORSolutionLink = {ObsWORResourceName: solutionError}
                        else:
                            ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                            return ObsWORSolutionLink

                elif typeofSolution == 3 and sheets.strip().lower() == 'details'.lower():
                    # ElevateObservation.programsFileCheck(programFile, accessToken, parentFolder, MainFilePath)
                    wbPgm= xlrd.open_workbook(programFile, on_demand=True)
                    programSheetNames = wbPgm.sheet_names()
                    wbSurvey = xlrd.open_workbook(addObservationSolution, on_demand=True)
                    ResourceSheet = wbObservation.sheet_by_name(sheets)
                    keysEnv = [ResourceSheet.cell(1, col_index_env).value for col_index_env in range(ResourceSheet.ncols)]
                    dictDetailsEnv = {keysEnv[col_index_env]: ResourceSheet.cell(row_index_env, col_index_env).value for col_index_env in range(ResourceSheet.ncols)}
                    SurveyResourceName = dictDetailsEnv['survey_solution_name'].encode('utf-8').decode('utf-8')
                    try:
                        def addsurveyFunc(parentFolder, wbObservation, millisecond, accessToken):
                            if not ElevateObservation.surveyValidate(addObservationSolution, accessToken, parentFolder): 
                                finalsurveySolutionlink = {SurveyResourceName: errorVar}
                                return finalsurveySolutionlink
                            # Create survey solution
                            surveyResp = ElevateObservation.createSurveySolution(parentFolder, wbSurvey, accessToken)
                            print("4865")
                            if not surveyResp:
                                finalsurveySolutionlink = {SurveyResourceName: errorVar}
                                return finalsurveySolutionlink
                            surTempExtID = surveyResp[1]
                            surTempSolID = surveyResp[0]
                            # Update solution status
                            bodySolutionUpdate = {"status": "active", "isDeleted": False}
                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, surveyResp[0], bodySolutionUpdate):
                                finalsurveySolutionlink = {SurveyResourceName: errorVar}
                                return finalsurveySolutionlink
                            # if errorVar == "":
                                # Upload survey questions
                            surveySolutionlink =  ElevateObservation.uploadSurveyQuestions(MainFilePath, parentFolder, wbSurvey, addObservationSolution, accessToken, surTempExtID, surTempSolID, millisecond, programFile)
    
                            finalsurveySolutionlink = {SurveyResourceName: surveySolutionlink}
                            return finalsurveySolutionlink
                        millisecond = int(time.time() * 1000)
                        surveySollink = addsurveyFunc(parentFolder, wbObservation, millisecond, accessToken)
                        return surveySollink
                        
                    except Exception as e:
                        print(f"Error occurred during survey creation: {str(e)}")
                        solutionError = str(e)
                        print(errorVar,"3772")
                        if errorVar == "":
                            surveySollink = {SurveyResourceName: solutionError}
                        else:
                            surveySollink = {SurveyResourceName: errorVar}
                        return surveySollink
                
        else :
            parentFolder = ElevateObservation.createFileStruct(MainFilePath, addObservationSolution)
            accessToken = ElevateObservation.generateAccessToken(parentFolder)
            typeofSolution = ElevateObservation.typeofresource(addObservationSolution, accessToken, parentFolder)
            if typeofSolution == 0:
                result = {}
                return result
            ElevateObservation.validateTenantAndOrgIdsFromResourceSheet(wbObservation)
            wbObservation = xlrd.open_workbook(addObservationSolution, on_demand=True)
            print(typeofSolution,"this is type of solution")
            if typeofSolution == 1 or typeofSolution == 5:
                if typeofSolution == 5:
                    impLedObsFlag = True
                else:
                    impLedObsFlag = False
                if not ElevateObservation.criteriaUpload(parentFolder, wbObservation, millisecond, accessToken, "framework", impLedObsFlag):
                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                    return finalObsRubricSolutionLink
                
                # userDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dikshaLoginId)
                # if not userDetails:
                #     finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                #     return finalObsRubricSolutionLink
                # matchedShikshalokamLoginId = userDetails[0]
                
                frameworkExternalId = ElevateObservation.frameWorkUpload(parentFolder, wbObservation, millisecond, accessToken)
                if not frameworkExternalId:
                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                    return finalObsRubricSolutionLink
                observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
                if not ElevateObservation.themesUpload(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, False):
                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                    return finalObsRubricSolutionLink
                solutionId = ElevateObservation.createSolutionFromFramework(parentFolder, accessToken, frameworkExternalId)
                if not solutionId:
                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                    return finalObsRubricSolutionLink

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
                if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                    return finalObsRubricSolutionLink
                bodySolutionUpdate = {"sections": section}
                if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                    return finalObsRubricSolutionLink
                bodySolutionUpdate = {"status": "active", "isDeleted": False, "criteriaLevelReport": criteriaLevelsReport}
                if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                    return finalObsRubricSolutionLink
                excelBook = open_workbook(addObservationSolution)
                if not ElevateObservation.questionUpload(addObservationSolution, parentFolder, frameworkExternalId, millisecond, accessToken,solutionId,typeofSolution):
                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                    return finalObsRubricSolutionLink
                if not pointBasedValue.lower() == "null":
                    bodySolutionUpdate = {"isRubricDriven": True}
                    if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                        finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                        return finalObsRubricSolutionLink
                    if not ElevateObservation.fetchSolutionCriteria(parentFolder, observationExternalId, accessToken):
                        finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                        return finalObsRubricSolutionLink
                    if not ElevateObservation.uploadCriteriaRubrics(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, True):
                        finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                        return finalObsRubricSolutionLink
                    if not ElevateObservation.uploadThemeRubrics(parentFolder, wbObservation, accessToken, frameworkExternalId, True):
                        finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                        return finalObsRubricSolutionLink
                else:
                    print("Observation with scoring system : null.")
                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                    return finalObsRubricSolutionLink
                bodySolutionUpdate = {'allowMultipleAssessemts': allow_multiple_submissions, "creator": creator}
                if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                    return finalObsRubricSolutionLink
                global solutionRolesArray, solutionStartDate, solutionEndDate
                xfile = openpyxl.load_workbook(programFile)
                sheet_name = 'details'.strip()
                resourceDetailsSheet = xfile[sheet_name]
                solutionDetails = ElevateObservation.fetchSolutionDetailsFromResourceSheet(parentFolder, programFile, solutionId, accessToken,typeofSolution)
                if not solutionDetails:
                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                    return finalObsRubricSolutionLink
                if solutionDetails[1]:
                    startDateArr = str(solutionDetails[1]).split("-")
                    bodySolutionUpdate = {
                        "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                    if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                        finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                        return finalObsRubricSolutionLink
                if solutionDetails[2]:
                    endDateArr = str(solutionDetails[2]).split("-")
                    bodySolutionUpdate = {
                        "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                    if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                        finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                        return finalObsRubricSolutionLink
                if isProgramnamePresent:
                    childId = ElevateObservation.createChild(parentFolder, observationExternalId, accessToken)
                    if not childId:
                        finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                        return finalObsRubricSolutionLink
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
                        
                        if not ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate):
                            finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                            return finalObsRubricSolutionLink
                        if solutionDetails[1]:
                            startDateArr = str(solutionDetails[1]).split("-")
                            bodySolutionUpdate = {
                                "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[
                                    0] + " 00:00:00"}
                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate):
                                finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                return finalObsRubricSolutionLink
                        if solutionDetails[2]:
                            endDateArr = str(solutionDetails[2]).split("-")
                            bodySolutionUpdate = {
                                "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate):
                                finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                                return finalObsRubricSolutionLink
                        else:
                            result = "The Solution has been successfully created."
                            finalObsRubricSolutionLink = {ObsWRResourceName: result}
                            return finalObsRubricSolutionLink
                else:
                    print("No program name detected.")
                    finalObsRubricSolutionLink = {ObsWRResourceName: errorVar}
                    return finalObsRubricSolutionLink
            elif typeofSolution == 2:
                if not ElevateObservation.criteriaUpload(parentFolder, wbObservation, millisecond, accessToken, "criteria", False):
                    ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                    return ObsWORSolutionLink
                
                # userDetails = ElevateObservation.fetchUserDetails(environment, accessToken, dikshaLoginId)
                # if not userDetails:
                #     ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                #     return ObsWORSolutionLink
                # matchedShikshalokamLoginId = userDetails[0]
                
                frameworkExternalId = ElevateObservation.frameWorkUpload(parentFolder, wbObservation, millisecond, accessToken)
                if not frameworkExternalId:
                    ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                    return ObsWORSolutionLink
                observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
                if not ElevateObservation.themesUpload(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, True):
                    ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                    return ObsWORSolutionLink
                solutionId = ElevateObservation.createSolutionFromFramework(parentFolder, accessToken, frameworkExternalId)
                if not solutionId:
                    ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                    return ObsWORSolutionLink
                sectionsObj = {"sections": {'S1': 'Observation Question'}}
                if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, sectionsObj):
                    ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                    return ObsWORSolutionLink
                ecmObj = {}
                ecmExternalId = None
                ecmObj = {
                    "evidenceMethods": {'OB': {'externalId': 'OB', 'tip': None, 'name': 'Observation', 'description': None,
                                            'modeOfCollection': 'onfield', 'canBeNotApplicable': False,
                                            'notApplicable': False, 'canBeNotAllowed': False, 'remarks': None}}}
                if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, ecmObj):
                    ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                    return ObsWORSolutionLink
                if not ElevateObservation.questionUpload(addObservationSolution, parentFolder, frameworkExternalId, millisecond, accessToken,
                            solutionId, typeofSolution):
                    ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                    return ObsWORSolutionLink
                if not ElevateObservation.fetchSolutionCriteria(parentFolder, observationExternalId, accessToken):
                    ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                    return ObsWORSolutionLink
                if not pointBasedValue.lower() == "null":
                    if not ElevateObservation.uploadCriteriaRubrics(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, False):
                        ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                        return ObsWORSolutionLink
                    if not ElevateObservation.uploadThemeRubrics(parentFolder, wbObservation, accessToken, frameworkExternalId, False):
                        ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                        return ObsWORSolutionLink
                bodySolutionUpdate = {"status": "active", "isDeleted": False, "allowMultipleAssessemts": True,
                                    "creator": creator}
                if not ElevateObservation.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate):
                    ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                    return ObsWORSolutionLink

                solutionDetails = ElevateObservation.fetchSolutionDetailsFromResourceSheet(parentFolder, programFile, solutionId, accessToken,typeofSolution)
                # Below script will convert date DD-MM-YYYY TO YYYY-MM-DD 00:00:00 to match the code syntax
                if not solutionDetails:
                    ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                    return ObsWORSolutionLink
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
                    if not childId:
                        ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                        return ObsWORSolutionLink
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
                        if not ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate):
                            ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                            return ObsWORSolutionLink
                        if solutionDetails[1]:
                            startDateArr = str(solutionDetails[1]).split("-")
                            bodySolutionUpdate = {
                                "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[
                                    0] + " 00:00:00"}
                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate):
                                ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                return ObsWORSolutionLink
                        if solutionDetails[2]:
                            endDateArr = str(solutionDetails[2]).split("-")
                            bodySolutionUpdate = {
                                "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                            if not ElevateObservation.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate):
                                ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                                return ObsWORSolutionLink
                        result = "The solution has been created successfully."
                        ObsWORSolutionLink = {ObsWORResourceName: result}
                        return ObsWORSolutionLink
                else:
                    print("No program name detected.")
                    ObsWORSolutionLink = {ObsWORResourceName: errorVar}
                    return ObsWORSolutionLink
            elif typeofSolution == 3:
                wbPgm= xlrd.open_workbook(programFile, on_demand=True)
                programSheetNames = wbPgm.sheet_names()
                wbSurvey = xlrd.open_workbook(addObservationSolution, on_demand=True)
                ResourceSheet = wbObservation.sheet_by_name(sheets)
                keysEnv = [ResourceSheet.cell(1, col_index_env).value for col_index_env in range(ResourceSheet.ncols)]
                dictDetailsEnv = {keysEnv[col_index_env]: ResourceSheet.cell(row_index_env, col_index_env).value for col_index_env in range(ResourceSheet.ncols)}
                SurveyResourceName = dictDetailsEnv['survey_solution_name'].encode('utf-8').decode('utf-8')
                try:
                    def addsurveyFunc(parentFolder, wbObservation, millisecond, accessToken):
                        if not ElevateObservation.surveyValidate(addObservationSolution, accessToken, parentFolder): 
                            finalsurveySolutionlink = {SurveyResourceName: errorVar}
                            return finalsurveySolutionlink
                        # Create survey solution
                        surveyResp = ElevateObservation.createSurveySolution(parentFolder, wbSurvey, accessToken)
                        if not surveyResp:
                            finalsurveySolutionlink = {SurveyResourceName: errorVar}
                            return finalsurveySolutionlink
                        surTempExtID = surveyResp[1]
                        surTempSolID = surveyResp[0]
                        # Update solution status
                        bodySolutionUpdate = {"status": "active", "isDeleted": False}
                        if not ElevateObservation.solutionUpdate(parentFolder, accessToken, surveyResp[0], bodySolutionUpdate):
                            finalsurveySolutionlink = {SurveyResourceName: errorVar}
                            return finalsurveySolutionlink
                        # if errorVar == "":
                            # Upload survey questions
                        if not ElevateObservation.NOPuploadSurveyQuestions(MainFilePath, parentFolder, wbSurvey, addObservationSolution, accessToken, surTempExtID, surTempSolID, millisecond, programFile):
                            finalsurveySolutionlink = {SurveyResourceName: errorVar}
                            return finalsurveySolutionlink
                        if errorVar == "":
                            finalsurveySolutionlink = {SurveyResourceName: surveySolutionlink}
                            return finalsurveySolutionlink
                    millisecond = int(time.time() * 1000)
                    surveySollink = addsurveyFunc(parentFolder, wbObservation, millisecond, accessToken)
                    return surveySollink
                    
                except Exception as e:
                    print(f"Error occurred during survey creation: {str(e)}")
                    solutionError = str(e)
                    print(errorVar,"3772")
                    if errorVar == "":
                        surveySollink = {SurveyResourceName: solutionError}
                    else:
                        surveySollink = {SurveyResourceName: errorVar}
                    return surveySollink
                
    def loadSurveyFile(programFile,resourceName):
        print("entering the loadfile")
        global downloaded_file,userEntity
        downloaded_file = ""
        global solutionDict
        # start_time = time.time()
        # parser = argparse.ArgumentParser()
        # parser.add_argument('--programFile', '--resourceFile', type=ElevateObservation.valid_file)
        # parser.add_argument('--env', '--env')
        # argument = parser.parse_args()
        # programFile = argument.programFile
        # environment = argument.env
        # millisecond = int(time.time() * 1000)
        MainFilePath = ElevateObservation.createFileStructForProgram(programFile)
        wbPgm = xlrd.open_workbook(programFile, on_demand=True)
        sheetNames = wbPgm.sheet_names()
        pgmSheets = ["Instructions", "Program Details", "Resource Details","Program Manager Details","Role-Subrole Mapping"]
        if len(sheetNames) == len(pgmSheets) and sheetNames == pgmSheets:
            print("--->Program Template detected.<---")
            
            for sheetEnv in sheetNames:
                if sheetEnv.strip().lower() == 'program details':
                    print("Checking program details sheet...")
                    programDetailsSheet = wbPgm.sheet_by_name(sheetEnv)
                    keysEnv = [programDetailsSheet.cell(1, col_index_env).value for col_index_env in
                            range(programDetailsSheet.ncols)]

                    print("11111111111111111111111111111111111111111111111111")
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
                        # scopeEntityType = scopeEntityType
                        userEntity = dictProgramDetails['Targeted state at program level'].encode('utf-8').decode('utf-8').lstrip().rstrip().split(",")
                        print(userEntity,"userentity")
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
                        print(resourceNamePGM,"resourceNamePGM")
                        resourceTypePGM = dictDetailsEnv['Type of resources'].encode('utf-8').decode('utf-8')
                        resourceLinkOrExtPGM = dictDetailsEnv['Resource Link']
                        if str(dictDetailsEnv['Type of resources']).lower().strip() == "course":
                            isCourse = False
                        else:
                            isCourse = False
                            print(resourceName)
                            print(resourceNamePGM)
                            if resourceNamePGM == resourceName:
                                resourceStatus = dictDetailsEnv['Resource Status']
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
                                    downloaded_file = download_file
                                    break
                            else:
                                continue
                                # Result = solutionDict[resourceName] = "Not Found"
                                # return Result

            print("--->Solution input file successfully downloaded: " + str(downloaded_file))
            # for addObservationSolution in downloaded_file:
            print(f"Processing file: {downloaded_file}")
            solutionSL = ElevateObservation.mainFunc(MainFilePath, programFile, downloaded_file, millisecond, isProgramnamePresent, isCourse,scopeEntityType=scopeEntityType)
            print(solutionSL)
            print(solutionSL.items(),"3400")
            for resourceName, solutionLink in solutionSL.items():
                solutionDict[resourceName] = solutionLink
            downloaded_file = ""
        else :
            MainFilePath = ElevateObservation.createFileStructForProgram(programFile)
            print(programFile,"58222")
            downloaded_file = programFile
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
        solutionDict = {}
        print(solutionDict,"5473")
        return json.dumps(result)

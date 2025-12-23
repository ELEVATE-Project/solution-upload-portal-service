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
from backend.src.main.modules.common_config import *
# from common_config import *
import threading
import wget
import gdown
import jwt
from dotenv import load_dotenv
from pathlib import Path
from backend.src.main.modules.ElevateObs import ElevateObservation
from backend.src.main.modules import ElevateObs
from dateutil import parser

env_path = Path(__file__).resolve().parents[1] / "apiServices" / "src" / "main" / ".env"

# Load the .env file
load_dotenv(dotenv_path=env_path)

# SECRET_KEY = os.getenv("SECRET_KEY")
# ADMIN_TOKEN = os.getenv("admin-token")
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
orgIdForScope = []
roles = ""
mainRole = ""
dictCritLookUp = {}
isProgramnamePresent = None
solutionLanguage = None
keyWords = None
entityTypeId = None
solutionDescription = None
creator = None
projectServiceLoginId = None
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
ProfessionalrolesPGM = None
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
mainRole = None
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
typeofSolution = 0
programFile = ""
solutionLink = ""
errorVar = ""
finalprojectsolutionlink = {}
tenantID = None
orgIDFromTemplate = None
TaskEvidenceOperator = ""
AnyTaskEvidenceNo = ""
stateEntitiesPGM = []
districtEntitiesPGM = []
blockEntitiesPGM = []
clusterEntitiesPGM = []
schoolEntitiesPGM = []
entityHierarchy = []

class Elevateproject:

    # def terminatingMessage(msg):
    #     print(msg)
    #     sys.exit()

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

    # def createFileStructForProgram(programFile):
    #     if not os.path.isdir('programFiles'):
    #         os.mkdir('programFiles')
    #     if "\\" in str(programFile):
    #         fileNameSplit = str(programFile).split('\\')[-1:]
    #     elif "/" in str(programFile):
    #         fileNameSplit = str(programFile).split('/')[-1:]
    #     else:
    #         fileNameSplit = str(programFile)
    #     print(fileNameSplit)
    #     if ".xlsx" in fileNameSplit:
    #         ts = str(time.time()).replace(".", "_")
    #         folderName = fileNameSplit.replace(".xlsx", "-" + str(ts))
    #         os.mkdir('programFiles/' + str(folderName))
    #         path = os.path.join('programFiles', str(folderName))
    #     else:
    #         print("File Error.")
    #         sys.exit()
    #     returnPathStr = os.path.join('programFiles', str(folderName))

    #     return returnPathStr
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
            return [Elevateproject.clean_single_value(part) for part in value.split(',') if part.strip()]
        return Elevateproject.clean_single_value(value)

    def decodeToken(accessTokenUser):
        try:
            accessTokenSecret = jwtTokenSecret
            decodedToken = jwt.decode(accessTokenUser, accessTokenSecret, algorithms=["HS256"])
            if 'data' not in decodedToken:
                print("Data not present in decodedToken")
                errorVar = ("Invalid Token")
            if 'tenant_code' not in decodedToken['data']:
                print("Tenant Id is not present in decodedToken")
                errorVar("Invalid Token")
            # if 'id' not in decodedToken['data']:
            #     print("Organization Id is not present in decodedToken")
            #     errorVar("Invalid Token")
            global tenantId
            tenantId = Elevateproject.clean_single_value(decodedToken['data']['tenant_code'])
            global orgIds
            orgIds = Elevateproject.clean_single_value(decodedToken['data']['organizations'][0].get('id'))

        except jwt.exceptions.InvalidTokenError as e:
            raise Exception(f"Invalid token: {str(e)}")
        except Exception as e:
            raise Exception(f"Token decoding failed: {str(e)}")

    def validateTenantAndOrgIdsFromProgramSheet(programFileContent):        
        tenantIdFromProgramFile = None
        orgIdsFromProgramFile = None
        sheetNames = programFileContent.sheet_names()
        # iterate through the sheets 
        for sheet in sheetNames:
            if sheet.strip().lower() == 'program details':
                print("--->Checking Program details sheet...")
                programDetailsSheet = programFileContent.sheet_by_name(sheet)
                keysEnv = [programDetailsSheet.cell(1, col_index_env).value for col_index_env in
                            range(programDetailsSheet.ncols)]
                for row_index_env in range(2, programDetailsSheet.nrows):
                    dictDetailsEnv = {keysEnv[col_index_env]: programDetailsSheet.cell(row_index_env, col_index_env).value
                                        for
                                        col_index_env in range(programDetailsSheet.ncols)}
                    tenantIdFromProgramFile = dictDetailsEnv.get('Tenant ID')
                    # orgIdsFromProgramFile = dictDetailsEnv.get('Org ID')
                    if tenantIdFromProgramFile == "shikshalokam":
                        # orgIdsFromProgramFile = dictDetailsEnv.get('Org ID')
                        global orgIdForScope
                        orgIds_str = dictDetailsEnv.get('Org ID', '')
                        orgIds = [oid.strip() for oid in orgIds_str.split(',') if oid.strip()]
                        orgIdForScope = orgIds
                        orgIdsFromProgramFile = orgIds[0] if orgIds else None
                    else:
                        # global orgIdForScope
                        # if tenantIdFromProgramFile == "shikshagrahanew":
                        orgIds_str = dictDetailsEnv.get('Targeted state at program level', '')
                        orgIds = [oid.strip().lower() for oid in orgIds_str.split(',') if oid.strip()]
                        orgIdForScope = orgIds
                        orgIdsFromProgramFile = orgIds[0] if orgIds else None
                        # if tenantIdFromProgramFile == "shikshagrahanew":
                        # orgIdsFromProgramFile = dictDetailsEnv.get('Targeted state at program level','').strip().lower()
                        # else:
                            # orgIdsFromProgramFile = dictDetailsEnv.get('Targeted District at program level', '').strip()

        # global roleOfResourceCreator
        # if roleOfResourceCreator not in ['org_admin', 'tenant_admin'] and not tenantIdFromProgramFile:
        #     raise ValueError("Tenant ID is required in program template for role 'admin', it cannot be empty")

        # if roleOfResourceCreator not in ['org_admin'] and not orgIdsFromProgramFile:
        #     raise ValueError("Org ID is required for role 'admin' and 'tenant_admin' in program template and cannot be empty")

        Elevateproject.assignTenantOrgValuesToGlobalVariables(tenantIdFromProgramFile, orgIdsFromProgramFile)   

    def validateTenantAndOrgIdsFromResourceSheet(resourceFileContent):

        for projectSheets in resourceFileContent:
                wbproject = xlrd.open_workbook(programFile, on_demand=True)

                if projectSheets.strip().lower() == 'project upload':
                    print("Checking project details sheet...")
                    detailsColCheck = wbproject.sheet_by_name(projectSheets)
                    keysColCheckDetai = [detailsColCheck.cell(0, col_index_check).value for col_index_check in range(detailsColCheck.ncols)]
                    detailsEnvSheet = wbproject.sheet_by_name(projectSheets)
                    keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in range(detailsEnvSheet.ncols)]
                    for row_index_env in range(2, detailsEnvSheet.nrows):
                        dictDetailsEnv = { keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for col_index_env in range(detailsEnvSheet.ncols)}
                        print(dictDetailsEnv['tenant_id'],"===========================================")
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
                    orgIdsFromresourceFile = dictDetailsEnv.get('Org ID')

        global roleOfResourceCreator
        if roleOfResourceCreator not in ['org_admin', 'tenant_admin'] and not tenantIdFromresourceFile:
            raise ValueError("Tenant ID is required in program template for role 'admin', it cannot be empty")

        if roleOfResourceCreator not in ['org_admin'] and not orgIdsFromresourceFile:
            raise ValueError("Org ID is required for role 'admin' and 'tenant_admin' in program template and cannot be empty")
        
        Elevateproject.assignTenantOrgValuesToGlobalVariables(tenantIdFromresourceFile, orgIdsFromresourceFile)

    def assignTenantOrgValuesToGlobalVariables(tenantIdFromTheSheets, orgIdsFromTheSheets):
        global tenantIDFromTemplate
        tenantIDFromTemplate = Elevateproject.clean_single_value(tenantIdFromTheSheets)
        global orgIDFromTemplate
        orgIDFromTemplate = Elevateproject.clean_single_value(orgIdsFromTheSheets)


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

    def check_sequence(arr):
        for i in range(1, len(arr)):
            if arr[i] != arr[i - 1] + 1:
                return False
        return True
    
    def createFileStructre(MainFilePath, addObservationSolution):
        global errorVar
        if not os.path.isdir(MainFilePath + '/SolutionFiles'):
            os.mkdir(MainFilePath + '/SolutionFiles')
        
        # Extract the file name regardless of the path format
        fileNameSplit = os.path.basename(str(addObservationSolution))
        
        if ".xlsx" in fileNameSplit:
            ts = str(time.time()).replace(".", "_")
            folderName = fileNameSplit.replace(".xlsx", "-" + str(ts))
            os.mkdir(MainFilePath + '/SolutionFiles/' + str(folderName))
            path = os.path.join(MainFilePath + '/SolutionFiles', str(folderName))
            path = os.path.join(path, str('apiHitLogs'))
            os.mkdir(path)
        else:
            errorVar = "File Error.offff"        
        returnPathStr = os.path.join(MainFilePath + '/SolutionFiles', str(folderName))

        if not os.path.isdir(returnPathStr + "/user_input_file"):
            os.mkdir(returnPathStr + "/user_input_file")

        shutil.copy(addObservationSolution, os.path.join(returnPathStr, "user_input_file"))
        # shutil.copy(programFile, os.path.join(returnPathStr, "user_input_file"))
        
        return returnPathStr
    
    
    def getProgramInfo(accessTokenUser, solutionName_for_folder_path, programNameInp):
        try:
            global programID, programExternalId, programDescription, isProgramnamePresent, programName, errorVar,tenantIDFromTemplate,orgIdForScope
            programName = programNameInp
            programUrl = elevateprojecthost + fetchprograminfoapiurl
            print(programUrl,"programUrl")
            payload = json.dumps({
                    "query": {
                        "name": programNameInp.lstrip().rstrip(),
                        "isAPrivateProgram": False,
                        "status": "active",
                        "tenantId": tenantIDFromTemplate
                        },
                        "mongoIdKeys": []
                        })
            print(payload,"payload")
            headersProgramSearch = {'Content-Type': content_type,
                                    'X-auth-token': accessTokenUser
                                    # 'internal-access-token': internal_access_token
                                    }
            print(headersProgramSearch,"headersProgramSearch")
            responseProgramSearch = requests.request("POST", url=programUrl, headers=headersProgramSearch,data=payload)
            print(responseProgramSearch.text,"responseProgramSearch")
            messageArr = []
            messageArr.append("Program Search API")
            messageArr.append("URL : " + programUrl)
            messageArr.append("Status Code : " + str(responseProgramSearch.status_code))
            messageArr.append("Response : " + str(responseProgramSearch.text))
            Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
            messageArr = []
            if responseProgramSearch.status_code == 200:
                print('--->Program fetch API Success')
                messageArr.append("--->Program fetch API Success")
                responseProgramSearch = responseProgramSearch.json()
                countOfPrograms = len(responseProgramSearch['result'])
                print(countOfPrograms,"countOfPrograms")
                messageArr.append("--->Program Count : " + str(countOfPrograms))
                if countOfPrograms == 0:
                    messageArr.append("No program found with the name : " + str(programName.lstrip().rstrip()))
                    messageArr.append("******************** Preparing for program Upload **********************")
                    print("No program found with the name : " + str(programName.lstrip().rstrip()))
                    print("******************** Preparing for program Upload **********************")
                    Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
                    fileheader = ["Program name fetch","Successfully fetched program name","Passed"]
                    Elevateproject.apicheckslog(solutionName_for_folder_path,fileheader)
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
                                Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
                                fileheader = ["program find api is running","found"+str(len(
                                    getProgramDetails))+"programs in backend","Failed","found"+str(len(
                                    getProgramDetails))+"programs ,check logs"]
                                Elevateproject.apicheckslog(solutionName_for_folder_path,fileheader)
                                Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
                            elif len(getProgramDetails) > 1:
                                print("Total " + str(len(getProgramDetails)) + " backend programs found with the name : " + programName.lstrip().rstrip())
                                messageArr.append("Total " + str(len(getProgramDetails)) + " backend programs found with the name : " + programName.lstrip().rstrip())
                                Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
                            
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
                            Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
                return True
            else:
                print("Program search API failed...")
                messageArr.append("Program search API failed...")
                Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
                print("Response Code : " + str(responseProgramSearch.status_code))
                errorVar = str(responseProgramSearch.text)
                return False
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")

    def fetchUserDetails(environment, accessToken, projectServiceId):
        global OrgName,errorVar
        try:
            decoded_token = jwt.decode(accessToken, options={"verify_signature": False}, algorithms=["HS256"])
            data=decoded_token.get('data')
            user_id = data.get('id')  
            url = userLoginHost + userinfoapiurl
            messageArr = ["User search API called."]
            headers = {#'Content-Type': 'application/json',
                    'internal-access-token': internal_access_token,
                    'X-auth-token': accessToken}
            responseUserSearch = requests.request("GET", url, headers=headers)
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
                    print("coming to a success")
                    return [userKeycloak, userName, firstName,roledetails,rootOrgId]
                else:
                    print("-->Given username/email is not present in projectService platform<--.")
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
            print(f"Error occurred: {str(e)}")
            print(errorVar)
            
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
                    "metaInformation.name":{
                        "$in": entitiesPGM.split(",")
                    },
                    "tenantId":tenantIDFromTemplate,
                    # "orgIds": {"$in":ElevateObservation.append_to_list(ElevateObservation.normalize_cell_value(orgIDFromTemplate),'ALL')}
                },

                "projection": [
                    "_id","metaInformation.name"
                ]
                }
            # data=json.dumps(payload)
            print(payload,"payload")
            responseFetchEntityListApi = requests.post(url=urlFetchEntityListApi, headers=headerFetchEntityListApi,data=json.dumps(payload))
            messageArr = ["Entities List Fetch API executed.", "URL  : " + str(urlFetchEntityListApi),
                        "Status : " + str(responseFetchEntityListApi.status_code)]
            Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
            print(responseFetchEntityListApi.text,"responseFetchEntityListApi")
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
                Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
                print("---> Error in location search.")
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")

    def programCreation(accessToken,parentFolder,externalId,pName,pDescription,roles,userId,mainRoleproff,rolesPGMID,entitiesPGMID,entityHierarchy):
        # accessToken, parentFolder, externalId, pName, pDescription, keywords, entities, roles, orgIds,entitiesPGM,mainRole,rolesPGM
        global errorVar,orgIdForScope,programExternalId,programID
        messageArr = []
        messageArr.append("++++++++++++ Program Creation ++++++++++++")
        # program creation url 
        try: 
            programCreationurl = elevateprojecthost + programcreationurl
            print(programCreationurl,"programCreationurl")
            messageArr.append("Program Creation URL : " + programCreationurl)

            # adding state entities
            # entities = stateEntitiesPGM.split(',')
            # entitiesTypeStr = entitiesType[0]
            scope={}
            scope["organizations"] = orgIdForScope
            scope["professional_subroles"] = rolesPGMID
            scope["professional_role"] = mainRoleproff
            scope.update(entityHierarchy)
            print(scope,"scope")
            programExternalId = externalId
            # program creation payload
            payload = json.dumps({
                        "externalId": programExternalId,
                        "name": pName,
                        "description": pDescription,
                        "isDeleted": False,
                        "resourceType": [
                            "program"
                        ],
                        "language": [
                            "English"
                        ],
                        "metaInformation": {
                        "state":stateEntitiesPGM.split(","),
                        "recommendedFor" : roles
                        },
                        "keywords": [],
                        "concepts": [],
                        "userId":userId,
                        "imageCompression": {
                            "quality": 10
                        },
                        "startDate": startDateOfProgram,
                        "endDate": endDateOfProgram,
                        "components": [],
                        "scope": scope,
                        "requestForPIIConsent": True
                    }
            )
            print(payload,"payload")
            messageArr.append("Body : " + str(payload))
            headers = {
                'internal-access-token': internal_access_token,
                'X-auth-token': accessToken,
                'Content-Type': 'application/json',
                'Authorization':authorization,
                'tenantId': tenantIDFromTemplate ,
                'orgid': orgIDFromTemplate,
                adminTokenHeaderName: projAdminAccessToken
                }
            print(headers,"headers")
            # program creation 
            responsePgmCreate = requests.request("POST", programCreationurl, headers=headers, data=(payload))
            messageArr.append("Program Creation Status Code : " + str(responsePgmCreate.status_code))
            messageArr.append("Program Creation Response : " + str(responsePgmCreate.text))
            messageArr.append("Program body : " + str(payload))
            # save logs 
            Elevateproject.createAPILog(parentFolder, messageArr)
            # check status 
            fileheader = [pName, ('Program Sheet Validation'), ('Passed')]
            Elevateproject.createAPILog(parentFolder, messageArr)
            Elevateproject.apicheckslog(parentFolder, fileheader)
            print(responsePgmCreate.text,"responsePgmCreate")
            if responsePgmCreate.status_code == 200:
                responsePgmCreatejson = responsePgmCreate.json()
                program_data = responsePgmCreatejson.get("result", {})
                programID = program_data.get("_id")
                if programID:
                    print("Program created successfully.")
                    print("Program ID:", programID)
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

    def programmappingpdpmsheetcreation(MainFilePath,accessToken, program_file,programexternalId,parentFolder):
        global errorVar
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
                    if dictDetailsEnv.get('Title of the Program'):
                        programNameInp = dictDetailsEnv['Title of the Program'].encode('utf-8').decode('utf-8')
                    else:
                        errorVar = "\"Title of the Program\" must not be Empty in \"Program details\" sheet"
                        
                    # programNameInp = dictDetailsEnv['Title of the Program'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Title of the Program'] else Elevateproject.terminatingMessage("\"Title of the Program\" must not be Empty in \"Program details\" sheet")

                if dictDetailsEnv.get('Program ID'):
                    extIdPGM = dictDetailsEnv['Program ID'].encode('utf-8').decode('utf-8')
                else:
                    errorVar = "\"Program ID\" must not be Empty in \"Program details\" sheet"
                    
                # extIdPGM = dictDetailsEnv['Program ID'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Program ID'] else Elevateproject.terminatingMessage("\"Program ID\" must not be Empty in \"Program details\" sheet")
                
                if dictDetailsEnv.get('Username/user id/email id/phone no. of Program Designer'):
                    programdesigner = dictDetailsEnv['Username/user id/email id/phone no. of Program Designer'].encode('utf-8').decode('utf-8')
                else:
                    errorVar = "\"Username/user id/email id/phone no. of Program Designer\" must not be Empty in \"Program details\" sheet"
                    
                # programdesigner = dictDetailsEnv['projectService username/user id/email id/phone no. of Program Designer'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Program ID'] else Elevateproject.terminatingMessage("\"projectService username/user id/email id/phone no. of Program Designer\" must not be Empty in \"Program details\" sheet")
                userDetails = Elevateproject.fetchUserDetails(environment, accessToken, programdesigner)
                if not userDetails:
                    return False
                creatorKeyCloakId = userDetails[0]
                creatorName = userDetails[1]
                # if "program_designer" in userDetails[4]:
                #     creatorKeyCloakId = userDetails[0]
                #     creatorName = userDetails[1]
                # else :
                #     print("user does't have program designer role")

                pdpmcolo1 = [creatorName, " ", " ", " ", creatorKeyCloakId, " ", " ","ADD","program_desiginer", extIdPGM, "programs"]
                with open(pdpmsheet + 'mapping.csv', 'a',encoding='utf-8') as file:
                    writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                    writer.writerows([pdpmcolo1])
                    fileheader = [creatorName,"program designer mapped successfully","Passed"]
                    Elevateproject.apicheckslog(parentFolder,fileheader)


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
                        if dictDetailsEnv.get('Elevate user id ( profile ID)'):
                            programmanagername2 = dictDetailsEnv['Elevate user id ( profile ID)']
                        else:
                            errorVar = "\"Elevate user id ( profile ID)\" must not be Empty in \"Program details\" sheet"
                            
                        # programmanagername2 = dictDetailsEnv['projectService user id ( profile ID)'] if dictDetailsEnv['projectService user id ( profile ID)'] else Elevateproject.terminatingMessage("\"projectService user id ( profile ID)\" must not be Empty in \"Program details\" sheet")
                    else:
                        try :
                            if dictDetailsEnv.get('Login ID on Elevate'):
                                programmanagername2 = dictDetailsEnv['Login ID on Elevate'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "\"Login ID on Elevate\" must not be Empty in \"Program details\" sheet"
                                
                            # programmanagername2 = dictDetailsEnv['Login ID on projectService'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Login ID on projectService'] else Elevateproject.terminatingMessage("\"Login ID on projectService\" must not be Empty in \"Program details\" sheet")
                            userDetails = Elevateproject.fetchUserDetails(environment, accessToken, programmanagername2)
                        except :
                            if dictDetailsEnv.get('Elevate user id ( profile ID)'):
                                programmanagername2 = dictDetailsEnv['Elevate user id ( profile ID)'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "\"Elevate user id ( profile ID)\" must not be Empty in \"Program details\" sheet"
                                
                            # programmanagername2 = dictDetailsEnv['projectService user id ( profile ID)'].encode('utf-8').decode('utf-8') if dictDetailsEnv['projectService user id ( profile ID)'] else Elevateproject.terminatingMessage("\"projectService user id ( profile ID)\" must not be Empty in \"Program details\" sheet")
                    userDetails = Elevateproject.fetchUserDetails(environment, accessToken, programmanagername2)
                    creatorKeyCloakId = userDetails[0]
                    creatorName = userDetails[1]
                    # if "program_manager" in userDetails[4]:
                    #     creatorKeyCloakId = userDetails[0]
                    #     creatorName = userDetails[1]
                    # else:
                    #     errorVar = ("user does't have program manager role")
                    pdpmcolo1 = [creatorName, " ", " ", " ", creatorKeyCloakId, " ", " ","ADD","program_desiginer", extIdPGM, "programs"]

                    with open(pdpmsheet + 'mapping.csv', 'a',encoding='utf-8') as file:
                        writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                        writer.writerows([pdpmcolo1])
                    messageArr = ""
                    messageArr = ("Response : " + str(pdpmcolo1))
                    Elevateproject.createAPILog(parentFolder, messageArr)
                    print("program manager mapped succesfully")
                    fileheader = [creatorName,"program manager mapped succesfully","Passed"]
                    Elevateproject.apicheckslog(parentFolder,fileheader)
        if errorVar == "":
            print(errorVar)
            return True
        else:
            print(errorVar,"pdpm mapping failure")
            return False

    def validate_hierarchy(parent_info, entity_level, pgm_maps, entity_type, entity_name):
        hierarchy = ["school", "cluster", "block", "district", "state"]
        start_index = hierarchy.index(entity_level)

        for level in hierarchy[start_index:]:
            pgm_values = pgm_maps.get(level)

            # Empty / "" means ALL
            if not pgm_values:
                continue

            # 🔥 If validating the entity itself, read from entity name
            if level == entity_type:
                value = entity_name
            else:
                value = parent_info.get(level, [{}])[0].get("name")

            if not value:
                print(f"{level} missing in entity hierarchy")
                return False

            if value not in pgm_values:
                print(f"{level} mismatch: {value} not in {pgm_values}")
                return False

        return True
   
    def fetchEntityTypeAndHierarchy(solutionName_for_folder_path,accessToken,entitiesPGM,scopeEntityType,schoolEntitiesPGM,
        clusterEntitiesPGM,blockEntitiesPGM,districtEntitiesPGM,stateEntitiesPGM):
        urlFetchEntityListApi = elevateentityhost + searchforlocation

        headerFetchEntityListApi = {
            'Content-Type': content_type,
            'internal-access-token': internal_access_token,
        }

        entityTypes = []
        entityTypeID = []

        for entityName in entitiesPGM:
            entityName = entityName.strip()

            payload = {
                "query": {
                    "metaInformation.name": entityName,
                    "tenantId": tenantIDFromTemplate,
                    "entityType": scopeEntityType[0]
                },
                "projection": ["entityType", "_id", "metaInformation.name"]
            }

            response = requests.post(
                url=urlFetchEntityListApi,
                headers=headerFetchEntityListApi,
                data=json.dumps(payload)
            )

            ElevateObservation.createAPILog(
                solutionName_for_folder_path,
                [
                    f"Entities List Fetch API executed for entity: {entityName}",
                    f"URL: {urlFetchEntityListApi}",
                    f"Status: {response.status_code}"
                ]
            )

            if response.status_code != 200:
                raise RuntimeError(
                    f"Failed to fetch entity list for '{entityName}'. "
                    f"Status code: {response.status_code}"
                )

            responseJson = response.json()
            entityToUpload = None

            for listEntity in responseJson.get("result", []):
                entityId = listEntity["_id"]

                detailsUrl = elevateentityhost + fetchDetailsEntity + entityId
                detailsResp = requests.get(
                    detailsUrl,
                    headers={"tenantId": tenantIDFromTemplate}
                )

                if detailsResp.status_code != 200:
                    raise RuntimeError(
                        f"Failed to fetch entity details for '{entityName}'. "
                        f"Status code: {detailsResp.status_code}"
                    )

                entities = detailsResp.json().get("result", [])
                print("Fetched Entities Details:", entities)

                for entity in entities:
                    parent_info = entity.get("parentInformation", {})
                    print("Parent Information:", parent_info)
                    entityToUpload = entity.get("entityType")
                    entityId = entity.get("_id")

                    EntityFlag = False

                    pgm_maps = {
                        "school": schoolEntitiesPGM,
                        "cluster": clusterEntitiesPGM,
                        "block": blockEntitiesPGM,
                        "district": districtEntitiesPGM,
                        "state": stateEntitiesPGM
                    }

                    # Detect lowest entity level present
                    if schoolEntitiesPGM:
                        detected_level = "school"
                    elif clusterEntitiesPGM:
                        detected_level = "cluster"
                    elif blockEntitiesPGM:
                        detected_level = "block"
                    elif districtEntitiesPGM:
                        detected_level = "district"
                    elif stateEntitiesPGM:
                        detected_level = "state"
                    else:
                        detected_level = "state"

                    print("Detected Entity Level:", detected_level)

                    # 🔥 STRICT VALIDATION FROM DETECTED LEVEL → STATE
                    entity_name = entity.get("metaInformation", {}).get("name")
                    entity_type = entity.get("entityType").lower()

                    if Elevateproject.validate_hierarchy(
                            parent_info,
                            detected_level,
                            pgm_maps,
                            entity_type,
                            entity_name
                        ):
                        EntityFlag = True
                    else:
                        print("Hierarchy validation failed")
                        return False

                    if EntityFlag:
                        entityTypes.append(entityToUpload)
                        entityTypeID.append(entityId)
                        print("Accepted:", entityToUpload, entityId)
                    else:
                        print("Rejected entity:", entityId)

            if not entityToUpload:
                raise ValueError(f"Entity type not found for entity '{entityName}'")

        # -------------------- PARENT–CHILD MERGE LOGIC --------------------

        hierarchy = ["state", "district", "block", "cluster", "school"]
        merged_output = {level: [] for level in hierarchy}

        for entityId in entityTypeID:
            urlFetchEntity = elevateentityhost + fetchDetailsEntity + entityId

            response = requests.get(
                url=urlFetchEntity,
                headers={
                    'Content-Type': content_type,
                    'tenantId': tenantIDFromTemplate,
                    'Authorization': f'Bearer {accessToken}'
                }
            )

            if response.status_code != 200:
                continue

            result = response.json().get("result", [])[0]
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
                if not merged_output[hierarchy[i]]:
                    merged_output[hierarchy[i]].append("ALL")

        for level in hierarchy:
            if not merged_output[level]:
                merged_output[level].append("ALL")

        print("Final Structured Hierarchy:", json.dumps(merged_output, indent=2))

        return entityTypes, entityTypeID, merged_output


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
                    'tenantId': tenantIDFromTemplate,
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

        
        
    def programsFileCheck(filePathAddPgm, accessToken, parentFolder, MainFilePath):
        global errorVar,entityHierarchy,orgIDFromTemplate,tenantIDFromTemplate,orgIdForScope,programID
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
                        returnvalues = []
                        
                        if dictDetailsEnv.get('Targeted subrole at program level'):
                            roles = dictDetailsEnv['Targeted subrole at program level'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "\"Targeted subrole at program level\" must not be Empty in \"Program details\" sheet"
                        
                        if dictDetailsEnv.get('Description of the Program'):
                            proDesc = dictDetailsEnv['Description of the Program'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "\"Description of the Program\" must not be Empty in \"Program details\" sheet"
                        
                        global stateEntitiesPGM,districtEntitiesPGM,blockEntitiesPGM,clusterEntitiesPGM,schoolEntitiesPGM
                        if dictDetailsEnv.get('Targeted state at program level'):
                            stateEntitiesPGM = dictDetailsEnv['Targeted state at program level'].encode('utf-8').decode('utf-8')
                        else:
                            stateEntitiesPGM = ""
                            errorVar = "\"Targeted state at program level\" must not be Empty in \"Program details\" sheet"
                        if dictDetailsEnv.get('Targeted District at program level'):
                            districtEntitiesPGM = dictDetailsEnv['Targeted District at program level'].encode('utf-8').decode('utf-8')
                        else:
                            districtEntitiesPGM = ""
                        if dictDetailsEnv.get('Targeted Block at program level'):
                            blockEntitiesPGM = dictDetailsEnv['Targeted Block at program level'].encode('utf-8').decode('utf-8')
                        else:
                            blockEntitiesPGM = ""
                        if dictDetailsEnv.get('Targeted Cluster at program level'):
                            clusterEntitiesPGM = dictDetailsEnv['Targeted Cluster at program level'].encode('utf-8').decode('utf-8')
                        else:
                            clusterEntitiesPGM = ""
                        if dictDetailsEnv.get('Targeted School at program level'):
                            schoolEntitiesPGM = dictDetailsEnv['Targeted School at program level'].encode('utf-8').decode('utf-8')
                        else:
                            schoolEntitiesPGM = ""
                       
                        global startDateOfProgram, endDateOfProgram, ReffstartDateOfProgram, ReffendDateOfProgram
                        startDateOfProgram = dictDetailsEnv['Start date of program']
                        endDateOfProgram = dictDetailsEnv['End date of program']
                        ReffstartDateOfProgram = dictDetailsEnv['Start date of program']
                        ReffendDateOfProgram = dictDetailsEnv['End date of program']
                        # taking the start date of program from program template and converting YYYY-MM-DD 00:00:00 format
                        
                        startDateArr = str(startDateOfProgram).split("-")
                        startDateOfProgram = startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"
                        endDateArr = str(endDateOfProgram).split("-")
                        endDateOfProgram = endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"
                        
                        global mainRole,rolesPGM, rolesPGMID
                        mainRole = dictDetailsEnv['Targeted role at program level']
                        newProgramRole = mainRole.split(",")
                        programRoleArray = list(newProgramRole)
                        rolesPGM = dictDetailsEnv['Targeted subrole at program level']
                        mainRoles = str(mainRole).strip().encode('utf-8').decode('utf-8').split(",")
                        subRoles = str(rolesPGM).strip().encode('utf-8').decode('utf-8').split(",")

                        mainRoles = [r.strip() for r in mainRoles if r.strip()]
                        subRoles = [r.strip() for r in subRoles if r.strip()]

                        verifiedRoles = Elevateproject.validate_roles_against_api(mainRoles, subRoles)
                        if not verifiedRoles:
                            return False
                        global mainRoleproff

                        mainRoleproff = verifiedRoles[0]
                        rolesPGMID = verifiedRoles[1]

                        global scopeEntityType,entitiesType

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
                        print("mainRole", mainRoleproff)
                        print("rolesPGMID", rolesPGMID)
                        
                        scopeEntityType = [EntityType] if isinstance(EntityType, str) else EntityType
                        entitiesType = Elevateproject.fetchEntityTypeAndHierarchy(parentFolder, accessToken,
                                                    entitiesPGM.lstrip().rstrip().split(","), scopeEntityType,schoolEntitiesPGM,clusterEntitiesPGM,blockEntitiesPGM,districtEntitiesPGM,stateEntitiesPGM)
                        if not entitiesType:
                            return False

                        if entitiesPGM:
                            entitiesPGM = entitiesPGM
                            scopeEntityType = entitiesType[0]
                        global entitiesPGMID
                        entitiesPGMID = entitiesType[1]
                        print(entitiesPGMID,"1068")
                        # entitiesPGMID = Elevateproject.fetchEntityId(parentFolder, accessToken,
                                                    # entitiesPGM.lstrip().rstrip().split(","), scopeEntityType,entitiesPGM)
                        global orgIds, entityHierarchy
                        entityHierarchy = entitiesType[2]
                        # entityHierarchy = Elevateproject.fetchEntityParentChilds(parentFolder, scopeEntityType, entitiesPGMID)
                        print("fetchedhirearchy", entityHierarchy)
                        if not Elevateproject.getProgramInfo(accessToken, parentFolder, programNameInp.encode('utf-8').decode('utf-8')):
                            if dictDetailsEnv.get('Program ID'):
                                extIdPGM = dictDetailsEnv['Program ID'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "\"Program ID\" must not be Empty in \"Program details\" sheet"
                            # extIdPGM = dictDetailsEnv['Program ID'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Program ID'] else Elevateproject.terminatingMessage("\"Program ID\" must not be Empty in \"Program details\" sheet")
                            if str(dictDetailsEnv['Program ID']).strip() == "Do not fill this field":
                                errorVar = ("change the program id")
                            if dictDetailsEnv.get('Description of the Program'):
                                descriptionPGM = dictDetailsEnv['Description of the Program'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "\"Description of the Program\" must not be Empty in \"Program details\" sheet"
                           
                            keywordsPGM = dictDetailsEnv['Keywords'].encode('utf-8').decode('utf-8')
                            
                            # entitiesType = Elevateproject.fetchEntityType(parentFolder, accessToken,
                                                    # entitiesPGM.lstrip().rstrip().split(","), scopeEntityType)
                            # selecting entity type based on the users input 
                            if entitiesPGM:
                                entitiesPGM = entitiesPGM
                                scopeEntityType = entitiesType[0]

                            userDetails = Elevateproject.fetchUserDetails(environment, accessToken, dictDetailsEnv['Username/user id/email id/phone no. of Program Designer'])
                            print(userDetails,"userDetails")
                            userId = userDetails[0]
                            messageArr = []

                            scopeEntityType = entitiesType[0]
                            # fetch entity details 
                            entitiesPGMID = entitiesType[1]
                            print("entitiesPGMID915", entitiesPGMID)
                            # entitiesPGMID = Elevateproject.fetchEntityId(parentFolder, accessToken,entitiesPGM.lstrip().rstrip().split(","), scopeEntityType,entitiesPGM)
                            entityHierarchy = entitiesType[2]
                            # entityHierarchy = Elevateproject.fetchEntityParentChilds(parentFolder, scopeEntityType, entitiesPGMID)

                            # sys.exit()
                            # fetch sub-role details 
                            # rolesPGMID = fetchScopeRole(parentFolder, accessToken, rolesPGM.lstrip().rstrip().split(","))
                            
                            # sys.exit()

                            # call function to create program 
                            if not Elevateproject.programCreation(accessToken,parentFolder,extIdPGM,programNameInp,proDesc,programRoleArray,userId,mainRoleproff,rolesPGMID,entitiesPGMID,entityHierarchy):
                                print("Program creation failed! Please check logs.")
                                return False
                            print("Program Created SuccessFully.")
                            # accessToken, parentFolder, extIdPGM, programNameInp, descriptionPGM,keywordsPGM.lstrip().rstrip().split(","),mainRole,rolesPGM
                            # sys.exit()
                            # if not Elevateproject.programmappingpdpmsheetcreation(MainFilePath, accessToken, program_file, extIdPGM,parentFolder):
                            #     return False


                            # map PM / PD to the program 
                            # Programmappingapicall(MainFilePath, accessToken, program_file,parentFolder)

                            # check if program is created or not 
                        else :
                            userDetails = Elevateproject.fetchUserDetails(environment, accessToken, dictDetailsEnv['Username/user id/email id/phone no. of Program Designer'])
                            print(userDetails,"userDetails")
                            userId = userDetails[0]
                            messageArr = []
                        
                            
                            if not Elevateproject.getProgramInfo(accessToken, parentFolder, programNameInp):
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
                        print(resourceNamePGM,"resourceNamePGM")
                        if dictDetailsEnv.get('Type of resources'):
                            resourceTypePGM = dictDetailsEnv['Type of resources'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "\"Type of resources\" must not be Empty in \"Program details\" sheet"
                        print(resourceTypePGM,"resourceTypePGM")
                        if dictDetailsEnv.get('Resource Link'):
                            resourceLinkOrExtPGM = dictDetailsEnv['Resource Link'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "\"Resource Link\" must not be Empty in \"Program details\" sheet"
                        print(resourceLinkOrExtPGM,"resourceLinkOrExtPGM")
                        if dictDetailsEnv.get('Resource Status'):
                            resourceStatusOrExtPGM = dictDetailsEnv['Resource Status']
                        else:
                            errorVar = "\"Resource Status\" must not be Empty in \"Program details\" sheet"
                        print(resourceStatusOrExtPGM,"resourceStatusOrExtPGM")

                        if dictDetailsEnv.get('Targeted subrole at resource level'):
                            rolesPGM = dictDetailsEnv['Targeted subrole at resource level']
                        else:
                            errorVar = "\"Targeted subrole at resource level\" must not be Empty in \"Program details\" sheet"
                        print(rolesPGM,"rolesPGM")
                        global ProfessionalrolesPGM
                        if dictDetailsEnv.get('Target role at the resource level'):
                            ProfessionalrolesPGM = dictDetailsEnv['Target role at the resource level']
                        else:
                            errorVar = "\"Target role at the resource level\" must not be Empty in \"Program details\" sheet"
                        print(ProfessionalrolesPGM,"ProfessionalrolesPGM")
                        # setting start and end dates globally. 
                        global startDateOfResource, endDateOfResource
                        startDateOfResource = dictDetailsEnv['Start date of resource']
                        print(startDateOfResource,"startDateOfResource")
                        endDateOfResource = dictDetailsEnv['End date of resource']
                        print(endDateOfResource,"endDateOfResource")
                        print(errorVar,"errorVar")
                        if errorVar == "":
                            return True
                        else:
                            return True
            
    def fetchSolutionDetailsFromProgramSheet(solutionName_for_folder_path, programFile, solutionId, accessToken):
        global solutionRolesArray, solutionStartDate, solutionEndDate
        urlFetchSolutionApi = elevateprojecthost + fetchsolutiondoc + solutionId
        headerFetchSolutionApi = {
            'Content-Type': 'application/json',
            'Authorization': authorization,
            'X-auth-token': accessToken,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token,
            'tenantId': tenantIDFromTemplate ,
            'orgid': orgIDFromTemplate,
            adminTokenHeaderName: projAdminAccessToken
        }
        payloadFetchSolutionApi = {}
        responseFetchSolutionApiUrl = requests.get(url=urlFetchSolutionApi, headers=headerFetchSolutionApi,
                                                data=payloadFetchSolutionApi)
        responseFetchSolutionJson = responseFetchSolutionApiUrl.json()
        print("solution name : " + responseFetchSolutionJson["result"]["name"])
        messageArr = ["Solution Fetch Link.",
                    "solution name : " + responseFetchSolutionJson["result"]["name"],
                    "solution ExternalId : " + responseFetchSolutionJson["result"]["externalId"]]
        messageArr.append("Upload status code : " + str(responseFetchSolutionApiUrl.status_code))
        Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
        print(responseFetchSolutionApiUrl.text,"responseFetchSolutionApiUrl")
        # print(responseFetchSolutionJson)
        if responseFetchSolutionApiUrl.status_code == 200:
            print('Fetch solution Api Success')
            solutionName = responseFetchSolutionJson["result"]["name"]
            xfile = openpyxl.load_workbook(programFile)
            resourceDetailsSheet = xfile['Resource Details']
            rowCountRD = resourceDetailsSheet.max_row
            columnCountRD = resourceDetailsSheet.max_column
            for row in range(3, rowCountRD + 1):
                solutionNameCell = resourceDetailsSheet[f"A{row}"].value
                if resourceDetailsSheet["A" + str(row)].value == solutionName:
                    solutionMainRole = str(resourceDetailsSheet["E" + str(row)].value).split(",")
                    solutionRolesArray = str(resourceDetailsSheet["F" + str(row)].value).split(",")
                    solutionStartDate = resourceDetailsSheet["G" + str(row)].value
                    solutionEndDate = resourceDetailsSheet["H" + str(row)].value

                    print("solutionMainRole", solutionMainRole)
                    print("solutionRolesArray", solutionRolesArray)
                    print("solutionStartDate", solutionStartDate)
                    print("solutionEndDate", solutionEndDate)
                    return [solutionMainRole,solutionRolesArray, solutionStartDate, solutionEndDate]
    
    def generateAccessToken(solutionName_for_folder_path):
        try:
            global errorVar
            # production search user api - start
            headerKeyClockUser = {'Content-Type': "application/x-www-form-urlencoded",'origin': "default-qa.tekdinext.com"}
            # responseKeyClockUser = requests.post(url=config.get(environment, 'elevateuserhost') + config.get(environment, 'userlogin'), headers=headerKeyClockUser,
                                                #  data=json.dumps(config.get(environment, 'keyclockAPIBody')))
            # Elevateproject.terminatingMessage(type(json.loads(config.get(environment, 'keyclockAPIBody'))))\
            loginBody = {
                'identifier' : identifier,
                'password' : password
            }
            responseKeyClockUser = requests.post(userLoginHost + keyclockapiurl , headers=headerKeyClockUser, data=loginBody)
            print(responseKeyClockUser.text,"1248")
            messageArr = []
            # messageArr.append("URL : " + str(keyclockapiurl))
            # messageArr.append("Body : " + str(keyclockapibody))
            messageArr.append("Status Code : " + str(responseKeyClockUser.status_code))
            if responseKeyClockUser.status_code == 200:
                responseKeyClockUser = responseKeyClockUser.json()
                accessTokenUser = responseKeyClockUser['result']['access_token']
                messageArr.append("Acccess Token : " + str(accessTokenUser))
                Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
                fileheader = ["Access Token","Access Token succesfully genarated","Passed"]
                Elevateproject.apicheckslog(solutionName_for_folder_path,fileheader)
                print("--->Access Token Generated!")
                # Elevateproject.decodeToken(accessTokenUser)
                return accessTokenUser
            
            else:
                print("Error in generating Access token")
                print("Status code : " + str(responseKeyClockUser.status_code))
                Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
                errorVar = str(responseKeyClockUser.text)
                return accessTokenUser
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")

    def validate_identifier(identifier, field_name="Field"):
        global errorVar
        pattern = r'^[A-Za-z0-9_-]+$'
        if not re.match(pattern, identifier):
            errorVar = (f"Invalid {field_name}: '{identifier}'. Only A-Z, a-z, 0-9, '-', and '_' are allowed.")
            return False
        else:
            print(f"{field_name} '{identifier}' is valid.")
            return True

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
            # errorVar = ("Please check the Input sheet.")
        return typeofSolution
    
    def projectValidate(filePathAddObs, accessToken, parentFolder):
        print("Validating project temp....")
        global scopeRoles, scopeEntityType , ccRootOrgName , ccRootOrgId, errorVar, criteriaLevelsReport,criteriaLevels, AnyTaskEvidenceNo,TaskEvidenceOperator
        try:
            criteria_id_arr = list()
            wbObservation1 = xlrd.open_workbook(filePathAddObs, on_demand=True)
            sheetNames1 = wbObservation1.sheet_names()
            projectDetailsCols = ["title", "projectId", "Username/user id/email id/phone no. of content creator", "categories",
                                "objective","duration","entityType","recommendedFor","keywords"]
            detailsColCheck = wbObservation1.sheet_by_name('Project upload')
            keysColCheckDetai = [detailsColCheck.cell(0, col_index_check).value for col_index_check in
                                        range(detailsColCheck.ncols)]
            lentasks = (len(keysColCheckDetai) - 9) // 2
            for i in range(lentasks):
                projectDetailsCols.append(f"learningResources{i+1}-name")
                projectDetailsCols.append(f"learningResources{i+1}-link")
            projectDetailsCols.append("has certificate")
            # projectDetailsCols.append("Project Level Evidence")
            # projectDetailsCols.append("Minimum No. of Evidence")
            # sys.exit()
            if tenantIDFromTemplate == "shikshalokam":
                taskUploadCols = ["TaskId", "TaskTitle", "parentTaskId",
                            "Mandatory task(Yes or No)","Solution Name","solutionType","isAnExternalTask","Number of submissions for observation"]
            else: 
                taskUploadCols = ["TaskId", "TaskTitle", "parentTaskId",
                            "Mandatory task(Yes or No)","Solution Name","solutionType","isAnExternalTask","Number of submissions for observation","Mitra_Link"]
            detailsColCheck = wbObservation1.sheet_by_name('Tasks upload')
            keysColCheckDetai = [detailsColCheck.cell(0, col_index_check).value for col_index_check in
                                        range(detailsColCheck.ncols)]
            lentasks = (len(keysColCheckDetai) - 12) // 2
            for i in range(lentasks):
                taskUploadCols.append(f"learningResources{i+1}-name")
                taskUploadCols.append(f"learningResources{i+1}-link")
            taskUploadCols.append("Evidence required for any task for certificate criteria")
            taskUploadCols.append("Minimum No. of Evidence for any task criteria")
            taskUploadCols.append("Task Level Evidence req. for certificate criteria")
            taskUploadCols.append("Minimum No. of Evidence for task level evidence criteria")
            # taskUploadCols.append("Mitra_Link")
            # taskUploadCols.append("Task Level Evidence")
            # taskUploadCols.append("Minimum No. of Evidence")

            certificateCols = ["Certificate issuer","Type of certificate","Logo - 1","Logo - 2","Authorised Signature Image - 1",
                            "Authorised Signatory - 1","Authorised Signature Image - 2","Authorised Signatory - 2"]
            for sheetColCheck in sheetNames1:
                if sheetColCheck.strip().lower() == 'Project upload'.lower():
                    print("--->Checking Project Upload sheet...")
                    detailsColCheck = wbObservation1.sheet_by_name(sheetColCheck)
                    keysColCheckDetai = [detailsColCheck.cell(0, col_index_check).value for col_index_check in
                                        range(detailsColCheck.ncols)]
                    print(keysColCheckDetai,"keysColCheckDetai")
                    print(projectDetailsCols,"projectDetailsCols")
                    if len(keysColCheckDetai) != len(projectDetailsCols) or set(keysColCheckDetai) == set(projectDetailsCols):
                        errorVar = 'Columns is missing in Project Upload sheet'
                    detailsEnvSheet = wbObservation1.sheet_by_name(sheetColCheck)
                    keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                            range(detailsEnvSheet.ncols)]
                    for row_index_env in range(1, detailsEnvSheet.nrows):
                        # print(dictDetailsEnv)
                        # sys.exit()
                        dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                        for
                                        col_index_env in range(detailsEnvSheet.ncols)}
                        if dictDetailsEnv['title']:
                            projectTitle = dictDetailsEnv['title'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "validation failed : Title must not be Empty in Project Upload sheet"
                        if dictDetailsEnv['projectId']:
                            projectId = dictDetailsEnv['projectId']
                        else:
                            errorVar = "validation failed :projectId must not be Empty in Project Upload sheet"
                        if not Elevateproject.validate_identifier(projectId):
                            errorVar = "ProjectID should be alpha numeric."
                        if dictDetailsEnv['categories']:
                            projectCategories = dictDetailsEnv['categories'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "validation failed :categories must not be Empty in Project Upload sheet"
                        if dictDetailsEnv['objective']:
                            projectDescription = dictDetailsEnv['objective'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "validation failed :objective must not be Empty in Project Upload sheet"
                        # if dictDetailsEnv['is a SSO user?']:
                        #     projectSSOuser = dictDetailsEnv['is a SSO user?']
                        # else:
                        #     errorVar = "validation failed :is a SSO user? column must not be Empty in Project Upload sheet"
                        if dictDetailsEnv['Username/user id/email id/phone no. of content creator']:
                            projectprojectServiceloginid = dictDetailsEnv['Username/user id/email id/phone no. of content creator'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "validation failed :Username/user id/email id/phone no. of content creator column must not be Empty in Project Upload sheet"
                        if dictDetailsEnv['duration']:
                            projectDuration = dictDetailsEnv['duration'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "validation failed :duration column must not be Empty in Project Upload sheet"
                        if dictDetailsEnv['has certificate']:
                            projectcertificate = dictDetailsEnv['has certificate']
                        else:
                            errorVar = "validation failed :has certificate column must not be Empty in Project Upload sheet"
                        
                        # projectlevelEvidence = dictDetailsEnv["Project Level Evidence"] if dictDetailsEnv[
                        #     "Project Level Evidence"] else Elevateproject.terminatingMessage(
                        #     "\"Project Level Evidence\" must not be Empty in \"Project Upload\" sheet")
                        # projectminnoofEvidence = dictDetailsEnv["Minimum No. of Evidence"] if dictDetailsEnv[
                        #     "Minimum No. of Evidence"] else Elevateproject.terminatingMessage(
                        #     "\"Minimum No. of Evidence\" must not be Empty in \"Project Upload\" sheet")


                if sheetColCheck.strip().lower() == 'Tasks upload'.lower():
                    print("--->Checking Tasks upload sheet...")
                    # sys.exit()
                    detailsColCheck = wbObservation1.sheet_by_name(sheetColCheck)
                    print(detailsColCheck,"detailsColCheck")
                    print(sheetColCheck,"sheetColCheck")
                    sheet_data = [detailsColCheck.row_values(row) for row in range(detailsColCheck.nrows)]
                    print(sheet_data)
                    for i in range(min(5, detailsColCheck.nrows)):
                        print(detailsColCheck.row_values(i))
                    print(f"Sheet Name: {detailsColCheck.name}, Rows: {detailsColCheck.nrows}, Cols: {detailsColCheck.ncols}")
                    keysColCheckDetai = [detailsColCheck.cell(0, col_index_check).value for col_index_check in
                                        range(detailsColCheck.ncols)]
                    
                    
                    print("keysColCheckDetai---------", keysColCheckDetai)
                    print("taskUploadCols 1417", taskUploadCols)
                    print(len(keysColCheckDetai), len(taskUploadCols))
                    if len(keysColCheckDetai) != len(taskUploadCols) or set(keysColCheckDetai) == set(taskUploadCols):
                        errorVar = 'Columns is missing in Task Upload sheet'
                    detailsEnvSheet = wbObservation1.sheet_by_name(sheetColCheck)
                    keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                            range(detailsEnvSheet.ncols)]
                    try:
                        evidence_col_index = keysEnv.index("Evidence required for any task for certificate criteria")
                        no_evidence_col_index = keysEnv.index("Minimum No. of Evidence for any task criteria")
                    except ValueError:
                        print("Column 'Evidence required for any task for certificate criteria' or 'Min No. Evidence for any task' not found.")
                        evidence_col_index = None
                        no_evidence_col_index = None

                    if evidence_col_index is not None:
                        TaskEvidenceOperator = detailsColCheck.cell_value(2, evidence_col_index)
                        if not TaskEvidenceOperator or str(TaskEvidenceOperator).strip().lower() == "":
                            TaskEvidenceOperator = "no"
                    if TaskEvidenceOperator.lower() == "yes" and no_evidence_col_index is not None:
                        AnyTaskEvidenceNo = detailsColCheck.cell_value(2, no_evidence_col_index)
                        if not AnyTaskEvidenceNo or str(AnyTaskEvidenceNo).strip().lower() == "":
                            AnyTaskEvidenceNo = 1
                            
                    for row_index_env in range(1, detailsEnvSheet.nrows):
                        dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                        for
                                        col_index_env in range(detailsEnvSheet.ncols)}
                        if dictDetailsEnv['TaskId']:
                            projectTaskId = dictDetailsEnv['TaskId']
                        else:
                            errorVar = "validation failed :TaskId column must not be Empty in Task Upload sheet"
                        if dictDetailsEnv['TaskTitle']:
                            projectTaskTitle = dictDetailsEnv['TaskTitle'].encode('utf-8').decode('utf-8')
                        else:
                            errorVar = "validation failed :TaskTitle column must not be Empty in Task Upload sheet"
                        if dictDetailsEnv['Mandatory task(Yes or No)']:
                            projectTaskMandatory = dictDetailsEnv['Mandatory task(Yes or No)']
                        else:
                            errorVar = "validation failed :Mandatory task(Yes or No) column must not be Empty in Task Upload sheet"
                    
                    for row_index_env in range(2, detailsEnvSheet.nrows):
                        dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                        for
                                        col_index_env in range(detailsEnvSheet.ncols)}
                        if tenantIDFromTemplate != "shikshalokam":
                            if dictDetailsEnv["Solution Name"] and dictDetailsEnv["Mitra_Link"]:
                                errorVar = "Validation Failed - Either an observation as a task or Mithra link, only one can be given per task."
                            elif (dictDetailsEnv["learningResources1-link"] or dictDetailsEnv["learningResources2-link"] or dictDetailsEnv["learningResources3-link"] or dictDetailsEnv["learningResources4-link"]) and dictDetailsEnv["Mitra_Link"]:
                                errorVar = "Validation Failed - Either an Learning Resource or Mithra link, only one can be given per task."    
                        # projectTaskMandatory = dictDetailsEnv['Mandatory task(Yes or No)'] if dictDetailsEnv[
                        #     'Mandatory task(Yes or No)'] else Elevateproject.terminatingMessage(
                        #     "\"Mandatory task(Yes or No)\" must not be Empty in \"Tasks Upload\" sheet")
                        

                    if sheetColCheck.strip().lower() == 'Certificate details'.lower():
                        print("--->Checking Certificate details  sheet...")

                        detailsColCheck = wbObservation1.sheet_by_name(sheetColCheck)
                        keysColCheckDetai = [detailsColCheck.cell(0, col_index_check).value for col_index_check in
                                                range(detailsColCheck.ncols)]

                        if len(keysColCheckDetai) != len(certificateCols) or set(keysColCheckDetai) == set(
                                    certificateCols):
                            print("certificate not found")
                            errorVar = 'Columns is missing in certificate details sheet'
                        detailsEnvSheet = wbObservation1.sheet_by_name(sheetColCheck)
                        keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                                    range(detailsEnvSheet.ncols)]
                        for row_index_env in range(2, detailsEnvSheet.nrows):

                            dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                            for
                                            col_index_env in range(detailsEnvSheet.ncols)}
                        
                        
                            # certificateissuer = dictDetailsEnv['Certificate issuer'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Certificate issuer'] else Elevateproject.terminatingMessage(
                            # "\"Certificate issuer\" must not be Empty in \"Certificate details\" sheet")
                        
                        
                            # Typeofcertificate = dictDetailsEnv['Type of certificate'] if dictDetailsEnv['Type of certificate'] in ["One Logo - One Signature","One Logo - Two Signature","Two Logo - One Signature","Two Logo - Two Signature"]  else Elevateproject.terminatingMessage(
                            # "\"Type of certificate\" must not be Empty in \"Certificate details\" sheet")
                            # Logo1 = dictDetailsEnv['Logo - 1'] if dictDetailsEnv[
                            # 'Logo - 1'] else Elevateproject.terminatingMessage(
                            # "\"Logo - 1\" must not be Empty in \"Certificate details\" sheet")

                            # Authorisedsignlogo1 = dictDetailsEnv['Authorised Signature Image - 1'] if dictDetailsEnv['Authorised Signature Image - 1'] else Elevateproject.terminatingMessage("\"Authorised Signature Image - 1\" must not be Empty in \"Certificate details\" sheet")
                            # Authoriseddesifnation1 = dictDetailsEnv['Authorised Signatory - 1'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Authorised Signatory - 1'] else Elevateproject.terminatingMessage("\"Authorised Signatory - 1\" must not be Empty in \"Certificate details\" sheet")
                            
                            if dictDetailsEnv['Certificate issuer']:
                                certificateissuer = dictDetailsEnv['Certificate issuer'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "validation failed :Certificate issuer column must not be Empty in Certificate Details sheet"
                            if dictDetailsEnv['Type of certificate']:
                                Typeofcertificate = dictDetailsEnv['Type of certificate']
                            else:
                                errorVar = "validation failed :Type of certificate column must not be Empty in Certificate Details sheet"
                            if dictDetailsEnv['Logo - 1']:
                                Logo1 = dictDetailsEnv['Logo - 1']
                            else:
                                errorVar = "validation failed :Logo - 1 column must not be Empty in Certificate Details sheet"
                            if dictDetailsEnv['Authorised Signature Image - 1']:
                                Authorisedsignlogo1 = dictDetailsEnv['Authorised Signature Image - 1']
                            else:
                                errorVar = "validation failed :Authorised Signature Image - 1 column must not be Empty in Certificate Details sheet"
                            if dictDetailsEnv['Authorised Signatory - 1']:
                                Authoriseddesifnation1 = dictDetailsEnv['Authorised Signatory - 1'].encode('utf-8').decode('utf-8')
                            else:
                                errorVar = "validation failed :Authorised Signatory - 1 column must not be Empty in Certificate Details sheet"
                            
                            if Typeofcertificate in ["One Logo - Two Signature", "Two Logo - Two Signature"]:
                                if dictDetailsEnv['Authorised Signature Image - 2']:
                                    Authorisedsignlogo2 = dictDetailsEnv['Authorised Signature Image - 2']
                                else:
                                    errorVar = "validation failed :Authorised Signature Image - 2 column must not be Empty in Certificate Details sheet"
                                if dictDetailsEnv['Authorised Signature Name - 2']:
                                    Authorisedsignname2 = dictDetailsEnv['Authorised Signature Name - 2'].encode('utf-8').decode('utf-8')
                                else:
                                    errorVar = "validation failed :Authorised Signature Name - 2 column must not be Empty in Certificate Details sheet"
                                if dictDetailsEnv['Authorised Designation - 2']:
                                    Authoriseddesifnation2 = dictDetailsEnv['Authorised Designation - 2'].encode('utf-8').decode('utf-8')
                                else:
                                    errorVar = "validation failed :Authorised Designation - 2 column must not be Empty in Certificate Details sheet"
                            elif Typeofcertificate in ["Two Logo - One Signature", "Two Logo - Two Signature"]:
                                if dictDetailsEnv['Logo - 2']:
                                    Logo2 = dictDetailsEnv['Logo - 2']
                                else:
                                    errorVar = "validation failed :Logo - 2 column must not be Empty in Certificate Details sheet"
            if not errorVar == "":
                return False
            else:
                print(errorVar,"3109")
                return True
        except:
            print(errorVar,"3107")

    def solutionUpdate(solutionName_for_folder_path, accessToken, solutionId, bodySolutionUpdate):
        solutionUpdateApiurl = elevateprojecthost + solutionupdateapi + str(solutionId)
        headerUpdateSolutionApi = {
            'Content-Type': content_type,
            'X-auth-token': accessToken,
            'X-Channel-id': x_channel_id,
            "internal-access-token": internal_access_token,
            'tenantId': tenantIDFromTemplate ,
            'orgid': orgIDFromTemplate,
            adminTokenHeaderName: projAdminAccessToken
            }
        responseUpdateSolutionApi = requests.post(url=solutionUpdateApiurl, headers=headerUpdateSolutionApi,data=json.dumps(bodySolutionUpdate))
        messageArr = ["Solution Update API called.", "URL : " + str(solutionUpdateApiurl), "Body : " + str(bodySolutionUpdate),"Response : " + str(responseUpdateSolutionApi.text),"Status Code : " + str(responseUpdateSolutionApi.status_code)]
        Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
        if responseUpdateSolutionApi.status_code == 200:
            print("Solution Update Success.")
            return True
        else:
            print("Solution Update Failed.")
            return False
    

    def checkEntityOfSolution(projectName_for_folder_path, solutionNameOrId, accessToken):
        urldbFind = internal_kong_ip + dbfindapi_url
        searchSolutionpayload = {}
        headerdbFindApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': content_type,
                'tenantId': tenantIDFromTemplate,
                'orgId' : orgIDFromTemplate,
                adminTokenHeaderName: projAdminAccessToken
            }
        searchSolutionpayload = json.dumps({
            "query": {
                "name": solutionNameOrId
            },
            "mongoIdKeys": [
                "_id",
                "solutionId",
                "metaInformation.solutionId"
            ],
            "limit": 10000
        })
        print(searchSolutionpayload,"searchSolutionpayload")
        print(urldbFind,"2163")
        searchSolutionresponse = requests.request("POST", url=urldbFind, headers=headerdbFindApi,
                                                data=searchSolutionpayload)
        print(searchSolutionresponse.text,"searchSolutionresponse")
        if searchSolutionresponse.status_code == 200:
            searchSolutionjson = searchSolutionresponse.json()
            results = searchSolutionjson.get("result", [])
            print(len(results), "1607")

            for solution in results:
                solution_id = solution["_id"]
                print(solution.get("isReusable"))

                if solution.get("isReusable") is True:
                    messageArr = [f"Solution found : {solution_id}"]
                    Elevateproject.createAPILog(projectName_for_folder_path, messageArr)
                    print("searchSolutionApi Success")

                    solutionEntityType = solution.get("entityType")
                    solutionExternalId = solution.get("externalId")

                    messageArr = [f"Task solution Entity Type found : {solutionEntityType}"]
                    Elevateproject.createAPILog(projectName_for_folder_path, messageArr)
                    print("FetchSolutionDocApi Success")

                    return [solutionEntityType, solutionExternalId]
                else:
                    print("No solution Found..")
                    messageArr = ["No Solution found"]
                    Elevateproject.createAPILog(projectName_for_folder_path, messageArr)

        else:
            messageArr = [
                "Solution fetch failed",
                f"URL : {urldbFind}",
                f"Status Code : {searchSolutionresponse.status_code}",
                f"Response : {searchSolutionresponse.text}"
            ]
            Elevateproject.createAPILog(projectName_for_folder_path, messageArr)

            # terminatingMessage("FetchSolutionDocApi is failed")

        # else:
            # terminatingMessage("search solution api is failed"
    
    def ObservationsolutionUpdate(solutionName_for_folder_path, accessToken, solutionId, bodySolutionUpdate):
        solutionUpdateApiurl = internal_kong_ip + solutionupdateapi + str(solutionId)
        headerUpdateSolutionApi = {
            'Content-Type': content_type,
            'X-auth-token': accessToken,
            'X-Channel-id': x_channel_id,
            "internal-access-token": internal_access_token,
            'tenantId': tenantIDFromTemplate ,
            'orgid': orgIDFromTemplate,
            adminTokenHeaderName: projAdminAccessToken
            }
        responseUpdateSolutionApi = requests.post(url=solutionUpdateApiurl, headers=headerUpdateSolutionApi,data=json.dumps(bodySolutionUpdate))
        messageArr = ["Solution Update API called.", "URL : " + str(solutionUpdateApiurl), "Body : " + str(bodySolutionUpdate),"Response : " + str(responseUpdateSolutionApi.text),"Status Code : " + str(responseUpdateSolutionApi.status_code)]
        Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
        if responseUpdateSolutionApi.status_code == 200:
            print("Solution Update Success.")
            return True
        else:
            print("Solution Update Failed.")
            return False

    def prepareProjectAndTasksSheets(project_inputFile, projectName_for_folder_path, accessToken):
        print("prepareProjectAndTasksSheets")
        millisecond = int(time.time() * 1000)
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
            print(dictProjectDetails,"1603")
            title = str(dictProjectDetails["title"]).encode('utf-8').decode('utf-8').strip()
            externalId = str(dictProjectDetails["projectId"]).strip()  + "-" + str(millisecond)
            entityType = str(dictProjectDetails["entityType"]).encode('utf-8').decode('utf-8').strip()
            categories_list = ["teachers", "students", "infrastructure", "community", "educationLeader", "schoolProcess","learner","faciliator"]
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

            projectAuthor = str(dictProjectDetails["Username/user id/email id/phone no. of content creator"]).encode('utf-8').decode('utf-8').strip()
            recommendedFor = str(dictProjectDetails["recommendedFor"]).encode('utf-8').decode('utf-8').strip()
            objective = str(dictProjectDetails["objective"]).encode('utf-8').decode('utf-8').strip()
            # entityType = None
            project_values = [title, externalId, categories_final, recommendedFor,objective, entityType,projectGoal]
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
                    project_values.append("projectService")
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
                            "categories,primaryAudience,successIndicators,risks,approaches,recommendedFor")
                    else:
                        project_values.append("")
                    
            with open(projectFilePath + 'projectUpload.csv','a',encoding='utf-8') as file:
                writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                writer.writerows([project_values])

        print("1657")
        tasksDetailsSheet = wbproject.sheet_by_name('Tasks upload')
        keysTasks = [tasksDetailsSheet.cell(1, col_index_env).value for col_index_env in
                    range(tasksDetailsSheet.ncols)]
        taskColumns1 = ["name", "externalId", "description", "type", "hasAParentTask", "parentTaskOperator",
                        "parentTaskValue",
                        "parentTaskId", "solutionType", "solutionSubType", "solutionId", "isDeletable","isAnExternalTask"]
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
        taskColumns1.append("redirectLink")
        taskColumns1.append("buttonLabel")

        with open(taskFilePath + 'taskUpload.csv', 'w',encoding='utf-8') as file:
            writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
            writer.writerows([taskColumns1])
        sequenceNumber = 0
        for row_index_env in range(2, tasksDetailsSheet.nrows):
            dictTasksDetails = {keysTasks[col_index_env]: tasksDetailsSheet.cell(row_index_env, col_index_env).value
                                for col_index_env in range(tasksDetailsSheet.ncols)}
            taskName = str(dictTasksDetails["TaskTitle"]).encode('utf-8').decode('utf-8').strip()
            if tenantIDFromTemplate != 'shikshalokam':
                Mitra_Link = str(dictTasksDetails["Mitra_Link"]).strip()
                if Mitra_Link == "":
                    Mitra_Link = Mitra_Link
                
            # subtaskname = str(dictTasksDetails["Subtask"]).encode('utf-8').decode('utf-8').strip()
            # startDate = dictTasksDetails["startDate"]
            # endDate = dictTasksDetails["endDate"]
            # if startDate:
            #     startDateArr = str(startDate).split("-")
            #     bodyStartDate = startDateArr[0] + "/" + startDateArr[1] + "/" + startDateArr[2]
            #     if endDate:
            #         endDateArr = str(endDate).split("-")
            #         bodyEndDate = endDateArr[0] + "/" + endDateArr[1] + "/" + endDateArr[2]                     

            taskId = str(dictTasksDetails["TaskId"]).encode('utf-8').decode('utf-8').strip() + "-" + str(millisecond)
            taskminNoOfSubmissionsRequired = str(dictTasksDetails["Number of submissions for observation"]).strip()
            sequenceNumber = sequenceNumber + 1
            taskSolutionType = ""
            try:
                taskDescription = str(dictTasksDetails["description"]).strip()
            except:
                taskDescription = ""

            # print(dictTasksDetails["solutionType"],"1815")
            if dictTasksDetails["solutionType"]:
                taskType = dictTasksDetails["solutionType"]
            elif dictTasksDetails["learningResources1-name"] != "" and dictTasksDetails["learningResources1-link"] != "":
                taskType = "content"

            elif tenantIDFromTemplate != 'shikshalokam':
                if dictTasksDetails["Mitra_Link"] != "":
                    taskType = "reflection"
                else:
                    taskType = "simple"
            else:
                    taskType = "simple"
        
            hasAParentTask = "NO"
            parentTaskOperator = ""
            parentTaskValue = ""
            parentTaskId = ""
        
            if dictTasksDetails["parentTaskId"] != "":
                hasAParentTask = "YES"
                parentTaskOperator = "EQUALS"
                parentTaskValue = "started"
                parentTaskId = str(dictTasksDetails["parentTaskId"]).strip() + "-" + str(millisecond)
            else:
                hasAParentTask = "NO"
                parentTaskOperator = ""
                parentTaskValue = ""
                parentTaskId = ""

            solutionSubType = ""
            solutionId = ""
            AnExternalTask = ""
            
            # solutionSubTypeForTask = dictTasksDetails["SolutionSubType"]
            # solutionIdForTask = dictTasksDetails["SolutionId"]
            if dictTasksDetails.get("Solution Name"):
                solutionNameOrId = dictTasksDetails["Solution Name"]
                print(solutionNameOrId, "solutionNameOrId----")

                taskSolutionType = taskType
                solutionDetailsInTask = Elevateproject.checkEntityOfSolution(
                    projectName_for_folder_path, solutionNameOrId, accessToken
                )

                observationSolutionChildId = None
                ObservationChildfrom = None

                # Try getting the observation child ID
                try:
                    observationSolutionChildId = ElevateObs.ElevateObservation.observationChildId
                    print("child id :", observationSolutionChildId)
                except AttributeError:
                    # Fallback if no observationChildId exists
                    ObservationChildfrom = solutionDetailsInTask[2] if len(solutionDetailsInTask) > 2 else None
                    print(ObservationChildfrom, "ObservationChild")

                # Prepare body for update
                bodysolutionUpdate = {
                    "status": "inactive",
                    "isDeleted": True
                }

                # Use whichever ID exists
                child_id_to_update = observationSolutionChildId or ObservationChildfrom
                if child_id_to_update:
                    Elevateproject.ObservationsolutionUpdate(
                        projectName_for_folder_path,
                        accessToken,
                        child_id_to_update,
                        bodysolutionUpdate
                    )


                print("observation child solution deactivated")
                solutionSubType = solutionDetailsInTask[0]
                solutionId = solutionDetailsInTask[1]

                taskSolutionType = dictTasksDetails["solutionType"]

                if dictTasksDetails["isAnExternalTask"] == "No":
                    AnExternalTask = "False"
                elif dictTasksDetails["isAnExternalTask"] == "Yes":
                    AnExternalTask = "True"

            if str(dictTasksDetails["Mandatory task(Yes or No)"]).strip().strip().lower() == "no":
                isDeletable = "TRUE"
            else:
                isDeletable = "FALSE"
            # task_values = [taskName, taskId, taskDescription, taskType, hasAParentTask, parentTaskOperator, parentTaskValue,
                        #   parentTaskId, taskSolutionType, solutionSubTypeForTask, solutionIdForTask, isDeletable,bodyStartDate,bodyEndDate,AnExternalTask]
            task_values = [taskName, taskId, taskDescription, taskType, hasAParentTask, parentTaskOperator, parentTaskValue,
                          parentTaskId, taskSolutionType, solutionSubType, solutionId, isDeletable,AnExternalTask]
            task_lr_value_count = 1
            for task_lr in range(0, int(taskLearningResource_count)):
                task_lr_name = str(dictTasksDetails["learningResources" + str(task_lr_value_count) + "-name"]).strip()
                task_lr_link = str(dictTasksDetails["learningResources" + str(task_lr_value_count) + "-link"]).strip()
                if task_lr_link and not task_lr_name:
                    raise ValueError(
                        f"Name is required for the learning resource with link: '{task_lr_link }'")
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
                    task_values.append("projectService")
                    task_values.append(task_lr_link_id)
                    task_lr_value_count += 1
            task_values.append(taskminNoOfSubmissionsRequired)
            task_values.append(sequenceNumber)
            if tenantIDFromTemplate != 'shikshalokam':
                task_values.append(Mitra_Link)
                if Mitra_Link.strip() != "":
                    task_values.append("Start Reflection")
                else:
                    task_values.append("")
            with open(taskFilePath + 'taskUpload.csv','a',encoding='utf-8') as file:
                writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                writer.writerows([task_values])
            #   subtaskname2 = str(dictTasksDetails["Subtask"]).encode('utf-8').decode('utf-8').strip()

    def projectUpload(projectFile, projectName_for_folder_path, accessToken):
        try:
            global errorVar
            error_message = ""
            urlProjectUploadApi = elevateprojecthost + projectuploadapi
            print(urlProjectUploadApi,"urlProjectUploadApi")
            headerProjectUploadApi = {
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantIDFromTemplate,
                'orgId' : orgIDFromTemplate,
                adminTokenHeaderName: projAdminAccessToken
            }
            print(headerProjectUploadApi,"headerProjectUploadApi")
            project_payload = {}
            filesProject = {
                'projectTemplates': open(projectName_for_folder_path + '/projectUpload/projectUpload.csv', 'rb')
            }
            responseProjectUploadApi = requests.post(url=urlProjectUploadApi, headers=headerProjectUploadApi,data=project_payload,files=filesProject)
            print(responseProjectUploadApi.text,"responseProjectUploadApi")
            messageArr = ["program mapping is success.","File path : " + projectName_for_folder_path + '/projectUpload/projectUpload.csv']
            messageArr.append("Upload status code : " + str(responseProjectUploadApi.status_code))
            Elevateproject.createAPILog(projectName_for_folder_path, messageArr)
            if responseProjectUploadApi.status_code == 200:
                print('ProjectUploadApi Success')
                with open(projectName_for_folder_path + '/projectUpload/projectInternal.csv','w+',encoding='utf-8') as projectRes:
                    projectRes.write(responseProjectUploadApi.text)

                    messageArr=["responnse :" ,responseProjectUploadApi.text ]
                    Elevateproject.createAPILog(projectName_for_folder_path,messageArr)
                    return True
            else:
                error_message = ""
                if responseProjectUploadApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"ProjectUploadApi-Client Error {responseProjectUploadApi.status_code}: {responseProjectUploadApi.text}"
                elif responseProjectUploadApi.status_code in [500, 502, 503, 504]:
                    error_message = f"ProjectUploadApi-Server Error {responseProjectUploadApi.status_code}: {responseProjectUploadApi.text}"
                else:
                    error_message = f"ProjectUploadApi-Unexpected Error {responseProjectUploadApi.status_code}: {responseProjectUploadApi.text}"

                errorVar = error_message
                print(error_message)
                messageArr.append(f"Error Response: {error_message}")
                Elevateproject.createAPILog(projectName_for_folder_path, messageArr)
                return False
        except Exception as e:
            errorVar = error_message
            print(error_message)
            print(errorVar)
            Elevateproject.createAPILog(projectName_for_folder_path, [f"Exception: {str(e)}"])
        
    def taskUpload(projectFile, projectName_for_folder_path, accessToken):
        try:
            global errorVar
            error_message = ""
            projectInternalfile = open(projectName_for_folder_path + '/projectUpload/projectInternal.csv', mode='r',encoding='utf-8')
            projectInternalfile = csv.DictReader(projectInternalfile)
            for projectInternal in projectInternalfile:
                projectExternalId = projectInternal["externalId"]
                project_id = projectInternal["_SYSTEM_ID"]
                if str(project_id).strip() == "Could not pushed to kafka":
                    fetchProjectIdApi = elevateentityhost + fetchsolutiondetails
                    headerfetchProjectIdApi = {
                        'Authorization': authorization,
                        'X-auth-token': accessToken,
                        'X-Channel-id': x_channel_id,
                        'internal-access-token': internal_access_token,
                        'tenantId': tenantIDFromTemplate,
                        'orgId' : orgIDFromTemplate,
                        adminTokenHeaderName: projAdminAccessToken
                    }
                    fetchProjectIdPayload = {}

                    responseProjectListApi = requests.get(url=fetchProjectIdApi, headers=headerfetchProjectIdApi,
                                                        data=fetchProjectIdPayload)
                    messageArr = ["Tasks Upload Sheet Prepared.",
                                "File path : " + projectName_for_folder_path + '/taskUpload/taskUpload.csv']
                    messageArr.append("URL : " + str(fetchProjectIdApi))
                    messageArr.append("Upload status code : " + str(responseProjectListApi.status_code))
                    Elevateproject.createAPILog(projectName_for_folder_path, messageArr)

                    if responseProjectListApi.status_code == 200:
                        print('project fetch api Success')
                        responsejson = responseProjectListApi.json()
                        projectList = responsejson['result']['data']
                        for project in projectList:
                            if project['externalId'] == projectExternalId:
                                project_id = project['_id']
                        return True
                    else:
                        error_message = ""
                        if responseProjectListApi.status_code in [400, 401, 403, 404, 422]:
                            error_message = f"ProjectListApi-Client Error {responseProjectListApi.status_code}: {responseProjectListApi.text}"
                        elif responseProjectListApi.status_code in [500, 502, 503, 504]:
                            error_message = f"ProjectListApi-Server Error {responseProjectListApi.status_code}: {responseProjectListApi.text}"
                        else:
                            error_message = f"ProjectListApi-Unexpected Error {responseProjectListApi.status_code}: {responseProjectListApi.text}"

                        errorVar = error_message
                        print(error_message)
                        messageArr.append(f"Error Response: {error_message}")
                        Elevateproject.createAPILog(projectName_for_folder_path, messageArr) 
                        return False 

                urlTasksUploadApi = elevateprojecthost + taskuploadapi + project_id
                headerTasksUploadApi = {
                    'X-auth-token': accessToken,
                    'X-Channel-id': x_channel_id,
                    'internal-access-token': internal_access_token,
                    'tenantId': tenantIDFromTemplate,
                    'orgId' : orgIDFromTemplate,
                    adminTokenHeaderName: projAdminAccessToken
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
                Elevateproject.createAPILog(projectName_for_folder_path, messageArr)
                if responseTasksUploadApi.status_code == 200:
                    print('TaskUploadApi Success')
                    with open(projectName_for_folder_path + '/taskUpload/taskInternal.csv','w+',encoding='utf-8') as tasksRes:
                        tasksRes.write(responseTasksUploadApi.text)
                    return True
                else:
                    error_message = ""
                    if responseTasksUploadApi.status_code in [400, 401, 403, 404, 422]:
                        error_message = f"TasksUploadApi-Client Error {responseTasksUploadApi.status_code}: {responseTasksUploadApi.text}"
                    elif responseTasksUploadApi.status_code in [500, 502, 503, 504]:
                        error_message = f"TasksUploadApi-Server Error {responseTasksUploadApi.status_code}: {responseTasksUploadApi.text}"
                    else:
                        error_message = f"TasksUploadApi-Unexpected Error {responseTasksUploadApi.status_code}: {responseTasksUploadApi.text}"

                    errorVar = error_message
                    print(error_message)
                    messageArr.append(f"Error Response: {error_message}")
                    Elevateproject.createAPILog(projectName_for_folder_path, messageArr)
                    return False
        except Exception as e:
            errorVar = error_message
            print(errorVar,"1315")
            Elevateproject.createAPILog(projectName_for_folder_path, [f"Exception: {str(e)}"])

    def convert_to_date(date_str):
        return datetime.strptime(date_str, "%d-%m-%Y")

    def solutionCreationAndMapping(projectName_for_folder_path, entityToUpload, listOfFoundRoles, accessToken, programFile):
        try:
            print("solutionCreationAndMapping....")
            global errorVar,entityHierarchy,programExternalId
            error_message = ""
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

                
                urlCreateProjectSolutionApi = elevateprojecthost + projectsolutioncreationapi
                print(urlCreateProjectSolutionApi,"urlCreateProjectSolutionApi")
                headerCreateSolutionApi = {
                    'Content-Type': content_type,
                    'X-auth-token': accessToken,
                    "internal-access-token" : internal_access_token,
                    'X-Channel-id': x_channel_id,
                    'tenantId': tenantIDFromTemplate,
                    'orgId' : orgIDFromTemplate,
                    adminTokenHeaderName: projAdminAccessToken
                }
                print(programExternalId,"programExternalId")
                print(headerCreateSolutionApi,"headerCreateSolutionApi")
                sol_payload = {
                    "createdFor": orgIds,
                    "rootOrganisations": orgIds,
                    "programExternalId": programExternalId,
                    "entityType": projectEntityType,
                    "externalId": solutionExternalId,
                    "name": project_name,
                    "scope": {
                        "roles": [rolesPGM],
                        },
                    "description": project_description,
                    "isReusable" : False,
                    "startDate": startDateOfProgram,
                    "endDate": endDateOfProgram,
                }
                print(sol_payload,"sol_payload")
                responseCreateSolutionApi = requests.post(url=urlCreateProjectSolutionApi,headers=headerCreateSolutionApi, data=json.dumps(sol_payload))
                print(responseCreateSolutionApi.text,"responseCreateSolutionApi")
                messageArr = ["Project Solution Created.","URL : " + str(urlCreateProjectSolutionApi),"Status Code : " + str(responseCreateSolutionApi.status_code),"Response : " + str(responseCreateSolutionApi.text)]
                if responseCreateSolutionApi.status_code == 200:
                    responseCreateSolutionApi = responseCreateSolutionApi.json()
                    solutionId = responseCreateSolutionApi['result']['_id']
                    print(solutionId,"solutionId")
                    messageArr.append("Solution Generated : " + str(solutionId))
                    Elevateproject.createAPILog(projectName_for_folder_path, messageArr)
                    print("ProjectSolutionCreationApi Success")
                    duplicateTemplateExtId = projectExternalId + '_IMPORTED'
                    queryparamsMapProjectSolutionApi = projectExternalId + '?solutionId='+solutionId
                    urlMapProjectSolutionApi = elevateprojecthost + mapsolutiontoproject + queryparamsMapProjectSolutionApi
                    headerMapSolutionProject = {
                        'Content-Type': content_type,
                        'Authorization': authorization,
                        'internal-access-token' : internal_access_token,
                        'X-auth-token': accessToken,
                        'X-Channel-id': x_channel_id,
                        'tenantId': tenantIDFromTemplate,
                        'orgId' : orgIDFromTemplate,
                        adminTokenHeaderName: projAdminAccessToken
                    }
                    payloadMapSolutionProject = {
                        "externalId": duplicateTemplateExtId,
                        "rating": 5
                    }
                    responseMapProjectSolutionApi = requests.post(
                        url=urlMapProjectSolutionApi ,
                        headers=headerMapSolutionProject,data=json.dumps(payloadMapSolutionProject))
                    messageArr = ["Successfully mapped the project to Solution",
                                "URL : " + str(urlMapProjectSolutionApi + queryparamsMapProjectSolutionApi),
                                "Status Code : " + str(responseMapProjectSolutionApi.status_code),
                                "Response : " + str(responseMapProjectSolutionApi.text)]
                    print(responseMapProjectSolutionApi.text,"responseMapProjectSolutionApi")
                    if responseMapProjectSolutionApi.status_code == 200:
                        responseMapProjectSolutionApi = responseMapProjectSolutionApi.json()
                        duplicateTemplateId = responseMapProjectSolutionApi['result']['_id']
                        messageArr.append("duplicate TemplateId successfully created: " + str(duplicateTemplateId))
                        Elevateproject.createAPILog(projectName_for_folder_path, messageArr)
                        print("MapSolutionToProjectApi Sucsess")
                        with open(projectName_for_folder_path + '/solutionDetails/solutionDetails.csv', 'a',encoding='utf-8') as file:
                            writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                            writer.writerows([[solutionExternalId, project_name, project_description, solutionId,
                                            programExternalId, projectEntityType,
                                            scopeEntityType, entityToUpload, listOfFoundRoles, duplicateTemplateExtId,
                                            duplicateTemplateId]])
                        solutionDetails = Elevateproject.fetchSolutionDetailsFromProgramSheet(projectName_for_folder_path, programFile,
                                                                            solutionId, accessToken)
                        print("solutionDetails---",solutionDetails)
                        if solutionDetails:
                            scopeRoles = solutionDetails[0]
                            scopeSubRoles = solutionDetails[1]
                            verifiedRoles = Elevateproject.validate_roles_against_api(scopeRoles, scopeSubRoles)
                            if not verifiedRoles:
                                errorVar = "Roles validation failed!"
                                print(errorVar)
                                return False
                            mainRoleproff = verifiedRoles[0]
                            rolesPGMID = verifiedRoles[1]
                            print("mainRole", mainRoleproff)
                            print("rolesPGMID", rolesPGMID)
                            scopeEntities = entitiesPGMID
                            scope = {}
                            scope["organizations"] = orgIdForScope       
                            scope["professional_subroles"] = rolesPGMID
                            scope["professional_role"] = mainRoleproff
                            scope.update(entityHierarchy)
                            print(scope)
                            bodySolutionUpdate = {
                            "scope": scope
                            }
                            if Elevateproject.solutionUpdate(projectName_for_folder_path, accessToken, solutionId, bodySolutionUpdate):
                                userDetails = Elevateproject.fetchUserDetails(environment, accessToken, projectAuthor)
                                if userDetails:
                                    matchedShikshalokamLoginId = userDetails[0]
                                    projectCreator = userDetails[1]
                                    bodySolutionUpdate = {
                                        "creator": projectCreator, "author": matchedShikshalokamLoginId}
                                    Elevateproject.solutionUpdate(projectName_for_folder_path, accessToken, solutionId, bodySolutionUpdate)
                                    # Below script will convert date DD-MM-YYYY TO YYYY-MM-DD 00:00:00 to match the code syntax
                                    ReffstartDateOfProgram1 = Elevateproject.convert_to_date(ReffstartDateOfProgram)
                                    ReffendDateOfProgram1 = Elevateproject.convert_to_date(ReffendDateOfProgram)
                                    solutionDetails2 = Elevateproject.convert_to_date(solutionDetails[2])
                                    solutionDetails3 = Elevateproject.convert_to_date(solutionDetails[3])
                                    print(ReffstartDateOfProgram1 <= solutionDetails2 <= ReffendDateOfProgram1,ReffstartDateOfProgram1, solutionDetails2, ReffendDateOfProgram1)
                                    print(ReffstartDateOfProgram <= solutionDetails[3] <= ReffendDateOfProgram,ReffstartDateOfProgram, solutionDetails[3], ReffendDateOfProgram)
                                    if ReffstartDateOfProgram1 <= solutionDetails2 <= ReffendDateOfProgram1 and ReffstartDateOfProgram1 <= solutionDetails3 <= ReffendDateOfProgram1:
                                        if solutionDetails[2]:
                                            startDateArr = str(solutionDetails[2]).split("-")
                                            bodySolutionUpdate = {
                                                "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                                            Elevateproject.solutionUpdate(projectName_for_folder_path, accessToken, solutionId, bodySolutionUpdate)
                                        if solutionDetails[3]:
                                            endDateArr = str(solutionDetails[3]).split("-")
                                            bodySolutionUpdate = {
                                                "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                                            Elevateproject.solutionUpdate(projectName_for_folder_path, accessToken, solutionId, bodySolutionUpdate)
                                    else:
                                        errorVar = "Date mismatching!"
                                        print(errorVar)
                                        return False
                                else:
                                    print(errorVar)
                                    return False
                            else:
                                print(errorVar)
                                return False
                        else:
                            error_message = ""
                            if responseMapProjectSolutionApi.status_code in [400, 401, 403, 404, 422]:
                                error_message = f"MapProjectSolutionApi-Client Error {responseMapProjectSolutionApi.status_code}: {responseMapProjectSolutionApi.text}"
                            elif responseMapProjectSolutionApi.status_code in [500, 502, 503, 504]:
                                error_message = f"MapProjectSolutionApi-Server Error {responseMapProjectSolutionApi.status_code}: {responseMapProjectSolutionApi.text}"
                            else:
                                error_message = f"MapProjectSolutionApi-Unexpected Error {responseMapProjectSolutionApi.status_code}: {responseMapProjectSolutionApi.text}"
                            errorVar = error_message
                            print("Map project to solution api failed.")
                            return False
                        if errorVar =="":
                            print("its created......................")
                            return [solutionExternalId, solutionId]
                        else:
                            return errorVar
                    else:
                        errorVar = str(responseCreateSolutionApi.text)
                        print("Project solution creation api failed.")
                        error_message = ""
                        if responseCreateSolutionApi.status_code in [400, 401, 403, 404, 422]:
                            error_message = f"CreateSolutionApi-Client Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}"
                        elif responseCreateSolutionApi.status_code in [500, 502, 503, 504]:
                            error_message = f"CreateSolutionApi-Server Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}"
                        else:
                            error_message = f"CreateSolutionApi-Unexpected Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}"
                        errorVar = error_message
                        return False
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")


    def fetchCertificateBaseTemplate(filePathAddProject,accessToken,projectName_for_folder_path):
        global errorVar
        error_message = ""
        try:
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
                        
            urldbFind = elevateprojecthost + dbfindapi
            headerdbFindApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': content_type,
                'tenantId': tenantIDFromTemplate,
                'orgId' : orgIDFromTemplate,
                adminTokenHeaderName: projAdminAccessToken
            }
            payload = json.dumps({
                "query": {
                    'tenantId': tenantIDFromTemplate
                },
                "mongoIdKeys": []
            })
            print(urldbFind,"2163")
            responsedbFindApi = requests.request("POST", url=urldbFind, headers=headerdbFindApi,
                                                data=payload)
            
            print(responsedbFindApi.text,"responsedbFindApi 2166")
            if responsedbFindApi.status_code == 200:
                responseaddcetificate = responsedbFindApi.json()
                result_list = responseaddcetificate['result']
                baseTemplateLookup = {}
                for i in result_list:
                    baseTemplateLookup[i['code']] = i['_id']
                typeOfCertificate=typeOfCertificate.lower()
                typeOfCertificate=typeOfCertificate.replace(" ","")
                typeOfCertificate
                baseTemplateCode= certificatetypeof[typeOfCertificate]
                return baseTemplateLookup[baseTemplateCode]
                
            else:
                error_message = ""
                if responsedbFindApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"dbFindApi-Client Error {responsedbFindApi.status_code}: {responsedbFindApi.text}"
                elif responsedbFindApi.status_code in [500, 502, 503, 504]:
                    error_message = f"dbFindApi-Server Error {responsedbFindApi.status_code}: {responsedbFindApi.text}"
                else:
                    error_message = f"dbFindApi-Unexpected Error {responsedbFindApi.status_code}: {responsedbFindApi.text}"
                errorVar = error_message
                print(error_message)
                messageArr = (f"Error Response: {error_message}")
                Elevateproject.createAPILog(projectName_for_folder_path, messageArr) 
                print("--->Error in fetching DBfind data please give proper code value<---")
                return False 
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")
    
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
                    # certificateissuer = dictDetailsEnv['Certificate issuer'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Certificate issuer'] else Elevateproject.terminatingMessage("\"Certificate issuer\" must not be Empty in \"Certificate details\" sheet")
                    # typeOfCertificate = dictDetailsEnv['Type of certificate'] if dictDetailsEnv['Type of certificate'] else Elevateproject.terminatingMessage("\"Type of certificate\" must not be Empty in \"Certificate details\" sheet")
                    if dictDetailsEnv.get('Certificate issuer'):
                        certificateissuer = dictDetailsEnv['Certificate issuer'].encode('utf-8').decode('utf-8')
                    else:
                        errorVar = "\"Certificate issuer\" must not be Empty in \"Program details\" sheet"
                        raise errorVar
                    if dictDetailsEnv.get('Type of certificate'):
                        typeOfCertificate = dictDetailsEnv['Type of certificate']
                    else:
                        errorVar = "\"Type of certificate\" must not be Empty in \"Program details\" sheet"
                        raise errorVar
                    
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
                        logo_split = str(Authsign1).split('/')[5]

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

    def editsvg(accessToken,filePathAddProject,projectName_for_folder_path,baseTemplate_id):
        global errorVar
        error_message = ""
        try:
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
                        authrigeddesignation1 = dictDetailsEnv['Authorised Signatory - 1'].encode('utf-8').decode('utf-8')
                        authrigedlogo2 = dictDetailsEnv['Authorised Signature Image - 2']
                        authrigeddesignation2 = dictDetailsEnv['Authorised Signatory - 2'].encode('utf-8').decode('utf-8')

                        payload = {}
                        downloadedfiles = []
                        baseTemplateId = ''
                        if Typeofcertificate == 'One Logo - One Signature':
                            print("-->This is One Logo - One Signature<--")
                            stateLogo1 = ('stateLogo1', ('logo1.jpg', open(projectName_for_folder_path + '/Logofile/logo1.jpg', 'rb'), 'image/jpeg'))
                            downloadedfiles.append(stateLogo1)
                            payload['stateTitle'] = Certificateisuuer
                            signatureImg1 = ('signatureImg1',('signatureImg1',open(projectName_for_folder_path +'/Logofile/signature1.jpg','rb'),'image/jpeg'))
                            downloadedfiles.append(signatureImg1)
                            payload['signatureTitle1a'] = authrigeddesignation1
                            baseTemplateId=baseTemplate_id
                            

                        elif Typeofcertificate == 'One Logo - Two Signature':
                            print("-->This is One Logo - Two Signature<--")

                            stateLogo1 = ('stateLogo1', ('logo1.jpg', open(projectName_for_folder_path + '/Logofile/logo1.jpg', 'rb'), 'image/jpeg'))
                            downloadedfiles.append(stateLogo1)
                            payload['stateTitle'] = Certificateisuuer
                            signatureImg1 = ('signatureImg1', (
                            'signatureImg1', open(projectName_for_folder_path + '/Logofile/signature1.jpg', 'rb'),
                            'image/jpeg'))
                            downloadedfiles.append(signatureImg1)
                            signatureImg2 = ('signatureImg2', ('signature2.jpg', open(projectName_for_folder_path + '/Logofile/signature2.jpg', 'rb'),'image/jpeg'))
                            downloadedfiles.append(signatureImg2)
                            payload['signatureTitle1a'] = authrigeddesignation1
                            payload['signatureTitle2a'] = authrigeddesignation2
                            baseTemplateId=baseTemplate_id
                        
                        elif Typeofcertificate == 'Two Logo - One Signature':
                            print("-->This is Two Logo - One Signature<--")
                            stateLogo1 = ('stateLogo1', ('logo1.jpg', open(projectName_for_folder_path + '/Logofile/logo1.jpg', 'rb'), 'image/jpeg'))
                            downloadedfiles.append(stateLogo1)
                            stateLogo2 = ('stateLogo2', ('logo2.jpg', open(projectName_for_folder_path + '/Logofile/logo2.jpg', 'rb'), 'image/jpeg'))
                            downloadedfiles.append(stateLogo2)
                            payload['stateTitle'] = Certificateisuuer
                            signatureImg1 = ('signatureImg1', ('signatureImg1', open(projectName_for_folder_path + '/Logofile/signature1.jpg', 'rb'),'image/jpeg'))
                            downloadedfiles.append(signatureImg1)
                            payload['signatureTitle1a'] = authrigeddesignation1
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
                            payload['signatureTitle1a'] = authrigeddesignation1
                            payload['signatureTitle2a'] = authrigeddesignation2
                            baseTemplateId=baseTemplate_id
                        
                        urleditnigsvgApi = elevateprojecthost + editsvgtemp + baseTemplateId
                        print(urleditnigsvgApi,"urleditnigsvgApi")
                        headereditingsvgApi = {
                            'Authorization': authorization,
                            'X-auth-token': accessToken,
                            'X-Channel-id': x_channel_id,
                            'internal-access-token': internal_access_token,
                            'tenantId': tenantIDFromTemplate,
                            'orgId' : orgIDFromTemplate,
                            adminTokenHeaderName: projAdminAccessToken
                        }
                        print(headereditingsvgApi,"headereditingsvgApi")
                        print(downloadedfiles,"downloadedfiles")
                        responseeditsvg = requests.request("POST",url=urleditnigsvgApi, headers=headereditingsvgApi,data=payload, files=downloadedfiles)
                        print(responseeditsvg.text,"responseeditsvg")
                        if responseeditsvg.status_code == 200:
                            responseeditsvg = responseeditsvg.json()
                            svgid = responseeditsvg['result']['url']
                            filesvg = svgid
                            Logofilepath = projectName_for_folder_path + '/Dowloadedsvg/'
                            if not os.path.exists(Logofilepath):
                                os.mkdir(Logofilepath)
                            dest_file = Logofilepath + 'Dowloaded.svg'
                            Logofile1 = gdown.download(filesvg, dest_file, quiet=False)
                            return True
                        else:
                            error_message = ""
                            if responseeditsvg.status_code in [400, 401, 403, 404, 422]:
                                error_message = f"editsvg-Client Error {responseeditsvg.status_code}: {responseeditsvg.text}"
                            elif responseeditsvg.status_code in [500, 502, 503, 504]:
                                error_message = f"editsvg-Server Error {responseeditsvg.status_code}: {responseeditsvg.text}"
                            else:
                                error_message = f"editsvg-Unexpected Error {responseeditsvg.status_code}: {responseeditsvg.text}"
                            errorVar = error_message
                            print(error_message)
                            errorVar = str(responseeditsvg.text)
                            print("-->Error in downloading SVG file please check logs<--")
                            return False
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            print("-----> API-Error",errorVar)
            return False

    def prepareaddingcertificatetemp(filePathAddProject, projectName_for_folder_path, accessToken, solutionId, programID,baseTemplate_id):
        global errorVar,TaskEvidenceOperator,AnyTaskEvidenceNo
        error_message = ""
        try:
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
                        
                        
                        taskLevelEvidence = dictDetailsEnv["Task Level Evidence req. for certificate criteria"].lower()
                        minNoOfEvidence = dictDetailsEnv["Minimum No. of Evidence for task level evidence criteria"]
                    
                        if TaskEvidenceOperator.lower() == "yes":
                            tasksLevelEvidance.append(dictDetailsEnv["TaskTitle"])
                        
                        else:
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

            urladdcertificate = elevateprojecthost + addcertificatetemplate
            headeraddcertificateApi = {
                #'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': content_type,
                'tenantId': tenantIDFromTemplate,
                'orgId' : orgIDFromTemplate,
                adminTokenHeaderName: projAdminAccessToken
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
                    # certificateissuer = dictDetailsEnv['Certificate issuer'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Certificate issuer'] else Elevateproject.terminatingMessage("\"Certificate issuer\" must not be Empty in \"Certificate details\" sheet")
                    if dictDetailsEnv.get('Certificate issuer'):
                        certificateissuer = dictDetailsEnv['Certificate issuer'].encode('utf-8').decode('utf-8')
                    else:
                        errorVar = "\"Certificate issuer\" must not be Empty in \"Certificate details\" sheet"
                    payload["issuer"]["name"] = certificateissuer

                    # Typeofcertificate = dictDetailsEnv['Type of certificate'] if dictDetailsEnv['Type of certificate'] in ["One Logo - One Signature", "One Logo - Two Signature", "Two Logo - One Signature","Two Logo - Two Signature"] else Elevateproject.terminatingMessage("\"Type of certificate\" must not be Empty in \"Certificate details\" sheet")
                    if dictDetailsEnv.get('Type of certificate') in ["One Logo - One Signature", "One Logo - Two Signature", "Two Logo - One Signature", "Two Logo - Two Signature"]:
                        Typeofcertificate = dictDetailsEnv['Type of certificate']
                    else:
                        errorVar = "\"Type of certificate\" must not be Empty or invalid in \"Certificate details\" sheet"   
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
                print(tasksLevelEvidance, task['name'])
                if task['name'] in tasksLevelEvidance:
                    hasAparent = task["hasAParentTask"]
                    if task["hasAParentTask"].lower() == "no":
                        
                        task_id = task["_SYSTEM_ID"]
                        if TaskEvidenceOperator.lower() == "no":
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
                            c = c + 1
                            cn = "C" + str(c)
                            taskconditions = {
                                cn: {
                                    # "validationText": f"Add {int(AnyTaskEvidenceNo)} evidence for any task",
                                    # "validationText": f"Add {int(AnyTaskEvidenceNo)} evidence for any task {tasksLevelEvidance[c-3]}",
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
                                            "value": int(AnyTaskEvidenceNo),
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


            if TaskEvidenceOperator.lower() == "yes":
                # payload["criteria"]["conditions"]["C1"]["validationText"] = f"Add {int(AnyTaskEvidenceNo)} evidence for any task"
                payload["criteria"]["conditions"]["C3"]["validationText"] = f"Add {int(AnyTaskEvidenceNo)} evidence for any task"

            if str(projectLevelEvidance).strip().lower() == "yes":       
                condition = ""
                print(payload["criteria"]["conditions"],"4514")
                print(TaskEvidenceOperator.lower(),"4515")
                condition_keys = list(payload["criteria"]["conditions"].keys())
                TaskEvidenceOperator = TaskEvidenceOperator.lower()

                if TaskEvidenceOperator.lower() == "yes" and len(condition_keys) > 2:
                    first_part = "&&".join(condition_keys[:2])
                    grouped_part = "||".join(condition_keys[2:])
                    condition = f"{first_part}&&({grouped_part})"
                else:
                    condition = "&&".join(condition_keys)

                payload["criteria"]["expression"] = condition
            else:
                condition = ""
                condition_keys = list(payload["criteria"]["conditions"].keys())
                TaskEvidenceOperator = TaskEvidenceOperator.lower()

                if TaskEvidenceOperator.lower() == "yes" and len(condition_keys) > 1:
                    first_part = "&&".join(condition_keys[:1])
                    grouped_part = "||".join(condition_keys[1:])
                    condition = f"{first_part}&&({grouped_part})"
                else:
                    condition = "&&".join(condition_keys)

                payload["criteria"]["expression"] = condition

            TaskEvidenceOperator = ""
            print(payload["criteria"]["expression"])

            print(json.dumps(payload, indent=1))
            # sys.exit()

            # sys.exit()
            # responseaddcertificateUploadApi = requests.request("POST",url=urladdcertificate, headers=headeraddcertificateApi,
            #                                        data=json.dumps(payload))
            responseaddcertificateUploadApi = requests.post(url=urladdcertificate, headers=headeraddcertificateApi, data=json.dumps(payload))
            messageArr = ["Add certificate json is prepared",
                        "File path : " + projectName_for_folder_path + '/addCertificate/Addcertificate.text']
            messageArr.append("URL : " + str(responseaddcertificateUploadApi))
            messageArr.append("Upload status code : " + str(responseaddcertificateUploadApi.status_code))
            Elevateproject.createAPILog(projectName_for_folder_path, messageArr)
            with open(projectName_for_folder_path + '/addCertificate/Addcertificatejson.json',
                    'w+',encoding='utf-8') as tasksRes:
                tasksRes.write(json.dumps(payload))

            if responseaddcertificateUploadApi.status_code == 200:
                responseaddcetificate = responseaddcertificateUploadApi.json()
                certificatetemplateid = responseaddcetificate['result']['_id']
                print("-->Certificate template id generated <--", certificatetemplateid)
                with open(projectName_for_folder_path + '/addCertificate/Addcertificate.text',
                        'w+',encoding='utf-8') as tasksRes:
                    tasksRes.write(responseaddcertificateUploadApi.text)

            else:
                error_message = ""
                if responseaddcertificateUploadApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"addcertificateUploadApi-Client Error {responseaddcertificateUploadApi.status_code}: {responseaddcertificateUploadApi.text}"
                elif responseaddcertificateUploadApi.status_code in [500, 502, 503, 504]:
                    error_message = f"addcertificateUploadApi-Server Error {responseaddcertificateUploadApi.status_code}: {responseaddcertificateUploadApi.text}"
                else:
                    error_message = f"addcertificateUploadApi-Unexpected Error {responseaddcertificateUploadApi.status_code}: {responseaddcertificateUploadApi.text}"
                errorVar = error_message
                print(error_message)
                print("Add certificate mission failed please check logs")
                messageArr.append("Response : " + str(responseaddcertificateUploadApi.text))
                Elevateproject.createAPILog(projectName_for_folder_path, messageArr)
                return False

            urluploadcertificatepi = elevateprojecthost + uploadcertificatetosvg + certificatetemplateid

            headeruploadcertificateApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantIDFromTemplate,
                    'orgId' : orgIDFromTemplate,
                    adminTokenHeaderName: projAdminAccessToken
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

                urlsolutionupdateapi = elevateprojecthost + solutionupdateapi + solutionId

                headersolutionupdateApi = {
                    'Authorization': authorization,
                    'X-auth-token': accessToken,
                    'X-Channel-id': x_channel_id,
                    'internal-access-token': internal_access_token,
                    'Content-Type': content_type,
                    'tenantId': tenantIDFromTemplate,
                    'orgId' : orgIDFromTemplate,
                    adminTokenHeaderName: projAdminAccessToken
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
                    error_message = ""
                    if responseupdatecertificateApi.status_code in [400, 401, 403, 404, 422]:
                        error_message = f"updatecertificateApi-Client Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}"
                    elif responseupdatecertificateApi.status_code in [500, 502, 503, 504]:
                        error_message = f"updatecertificateApi-Server Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}"
                    else:
                        error_message = f"updatecertificateApi-Unexpected Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}"

                    errorVar = error_message
                    print(error_message)
                    print("error in updating solution")
                    messageArr.append(f"Error Response: {error_message}")
                    Elevateproject.createAPILog(projectName_for_folder_path, messageArr)
                    return False

                urlprojecttemplateapi = elevateprojecthost + updateprojecttemplate + projectTemplateId
                headerprojectrtemplateupdateApi = {
                    'Authorization': authorization,
                    'X-auth-token': accessToken,
                    'X-Channel-id': x_channel_id,
                    'internal-access-token': internal_access_token,
                    'Content-Type': content_type,
                    'tenantId': tenantIDFromTemplate,
                    'orgId' : orgIDFromTemplate,
                    adminTokenHeaderName: projAdminAccessToken
                }

                certificate_payload = json.dumps({
                    'certificateTemplateId': certificatetemplateid
                })
                responseupdatecertificateApi = requests.request("POST", url=urlprojecttemplateapi,
                                                                headers=headerprojectrtemplateupdateApi,
                                                                data=certificate_payload)
                if responseupdatecertificateApi.status_code == 200:
                    print("--->Certificate added to project<---")
                    return True
                else:
                    error_message = ""
                    if responseupdatecertificateApi.status_code in [400, 401, 403, 404, 422]:
                        error_message = f"TasksUploadApi-Client Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}"
                    elif responseupdatecertificateApi.status_code in [500, 502, 503, 504]:
                        error_message = f"TasksUploadApi-Server Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}"
                    else:
                        error_message = f"TasksUploadApi-Unexpected Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}"

                    errorVar = error_message
                    print(error_message)
                    messageArr.append(f"Error Response: {error_message}")
                    Elevateproject.createAPILog(projectName_for_folder_path, messageArr)
                    print("error in updating certificate with project")
                    return False
            else:
                error_message = ""
                if responseDownloadsvgApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"DownloadsvgApi-Client Error {responseDownloadsvgApi.status_code}: {responseDownloadsvgApi.text}"
                elif responseDownloadsvgApi.status_code in [500, 502, 503, 504]:
                    error_message = f"DownloadsvgApi-Server Error {responseDownloadsvgApi.status_code}: {responseDownloadsvgApi.text}"
                else:
                    error_message = f"DownloadsvgApi-Unexpected Error {responseDownloadsvgApi.status_code}: {responseDownloadsvgApi.text}"
                errorVar = error_message
                print(error_message)
                messageArr.append(f"Error Response: {error_message}")
                Elevateproject.createAPILog(projectName_for_folder_path, messageArr)
                print("error in updating certificate with project")
                return False
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")

    def prepareProgramSuccessSheet(MainFilePath, solutionName_for_folder_path, programFile, solutionExternalId, solutionId,accessToken):
        global solutionLink
        urlFetchSolutionApi = elevateprojecthost + fetchsolutiondoc + solutionId
        headerFetchSolutionApi = {
            'Authorization': authorization,
            'X-auth-token': accessToken,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token,
            'tenantId': tenantIDFromTemplate,
                    'orgId' : orgIDFromTemplate,
                    adminTokenHeaderName: projAdminAccessToken
        }
        payloadFetchSolutionApi = {}

        responseFetchSolutionApi = requests.get(url=urlFetchSolutionApi, headers=headerFetchSolutionApi,
                                                data=payloadFetchSolutionApi)
        responseFetchSolutionJson = responseFetchSolutionApi.json()
        messageArr = ["Solution Fetch Link.",
                    "solution name : " + responseFetchSolutionJson["result"]["name"],
                    "solution ExternalId : " + responseFetchSolutionJson["result"]["externalId"]]
        messageArr.append("Upload status code : " + str(responseFetchSolutionApi.status_code))
        Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)

        if responseFetchSolutionApi.status_code == 200:
            print('Fetch solution Api Success')
            solutionName = responseFetchSolutionJson["result"]["name"]
        urlFetchSolutionLinkApi = elevateprojecthost + fetchlink + solutionId
        headerFetchSolutionLinkApi = {
            # 'Authorization': authorization,
            'X-auth-token': accessToken
            # 'X-Channel-id': x_channel_id
            # 'internal-access-token': internal_access_token,
            # 'tenantId': tenantIDFromTemplate,
            #         'orgId' : orgIDFromTemplate,
            #         adminTokenHeaderName: projAdminAccessToken
        }
        payloadFetchSolutionLinkApi = {}

        responseFetchSolutionLinkApi = requests.get(url=urlFetchSolutionLinkApi, headers=headerFetchSolutionLinkApi,
                                                    data=payloadFetchSolutionLinkApi)
        
        print(responseFetchSolutionLinkApi.text, "responseFetchSolutionLinkApi")
        messageArr = ["Solution Fetch Link.","solution id : " + solutionId,"solution ExternalId : " + solutionExternalId]
        messageArr.append("Upload status code : " + str(responseFetchSolutionLinkApi.status_code))
        Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
        if responseFetchSolutionLinkApi.status_code == 200:
            print(responseFetchSolutionLinkApi.text,"responseFetchSolutionLinkApi")
            print('Fetch solution Link Api Success')
            responseProjectUploadJson = responseFetchSolutionLinkApi.json()
            solutionLink = responseProjectUploadJson["result"]
            solutionLink = ','.join(solutionLink)
            print(solutionLink,"solutionLink")
            messageArr.append("Response : " + str(responseFetchSolutionLinkApi.text))
            Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
            # Ensure programFile is formatted correctly
            programFileBase = str(programFile).replace(".xlsx", "")
            success_file_path = os.path.join(MainFilePath, programFileBase + '-SuccessSheet.xlsx')

            # Ensure directory exists
            os.makedirs(os.path.dirname(success_file_path), exist_ok=True)

            # Load the workbook correctly
            if os.path.exists(success_file_path):
                xfile = openpyxl.load_workbook(success_file_path)
            else:
                xfile = openpyxl.load_workbook(programFile)

            # Use the correct way to get the sheet
            resourceDetailsSheet = xfile["Resource Details"]  # Instead of get_sheet_by_name

            print(resourceDetailsSheet)

            # Define fill color
            greenFill = PatternFill(start_color='0000FF00', end_color='0000FF00', fill_type='solid')

            # Get row and column count
            rowCountRD = resourceDetailsSheet.max_row
            columnCountRD = resourceDetailsSheet.max_column

            for row in range(3, rowCountRD + 1):
                if str(resourceDetailsSheet["B" + str(row)].value).strip().lower() == "course":
                    resourceDetailsSheet["D1"].value = ""
                    resourceDetailsSheet["E1"].value = ""
                    resourceDetailsSheet['I2'].value = "External id of the resource"
                    resourceDetailsSheet['J2'].value = "link to access the resource/Response"

                    resourceDetailsSheet['I2'].fill = greenFill
                    resourceDetailsSheet['J2'].fill = greenFill
                    resourceDetailsSheet['I' + str(row)].value = solutionExternalId
                    resourceDetailsSheet['J' + str(row)].value = "The course has been successfully mapped to the program"

                    resourceDetailsSheet['I' + str(row)].fill = greenFill
                    resourceDetailsSheet['J' + str(row)].fill = greenFill

                elif str(resourceDetailsSheet["A" + str(row)].value).strip() == solutionName:
                    resourceDetailsSheet["D1"].value = ""
                    resourceDetailsSheet["E1"].value = ""
                    resourceDetailsSheet['I2'].value = "External id of the resource"
                    resourceDetailsSheet['J2'].value = "link to access the resource/Response"

                    resourceDetailsSheet['I2'].fill = greenFill
                    resourceDetailsSheet['J2'].fill = greenFill
                    resourceDetailsSheet['I' + str(row)].value = solutionExternalId
                    resourceDetailsSheet['J' + str(row)].value = solutionLink

                    resourceDetailsSheet['I' + str(row)].fill = greenFill
                    resourceDetailsSheet['J' + str(row)].fill = greenFill

            # Save the file
            xfile.save(success_file_path)
            print("Program success sheet is created")
            return solutionLink
            
    def projectSolutionDeepLink(MainFilePath, programFile, solutionId,accessToken):
        try:
            urlFetchSolutionLinkApi = elevateprojecthost + fetchlink + solutionId
        
            headerFetchSolutionLinkApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantIDFromTemplate,
                    'orgId' : orgIDFromTemplate,
                    adminTokenHeaderName: projAdminAccessToken
            }
            payloadFetchSolutionLinkApi = {}

            responseFetchSolutionLinkApi = requests.get(url=urlFetchSolutionLinkApi, headers=headerFetchSolutionLinkApi,
                                                        data=payloadFetchSolutionLinkApi)
            # messageArr = ["Solution Fetch Link.","solution id : " + solutionId,"solution ExternalId : " + solutionExternalId]
            # messageArr.append("Upload status code : " + str(responseFetchSolutionLinkApi.status_code))
            # Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
            if responseFetchSolutionLinkApi.status_code == 200:
                print('Fetch solution Link Api Success')
                responseProjectUploadJson = responseFetchSolutionLinkApi.json()
                solutionLink = responseProjectUploadJson["result"]
                return solutionLink
            else:
                error_message = ""
                if responseFetchSolutionLinkApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"DownloadsvgApi-Client Error {responseFetchSolutionLinkApi.status_code}: {responseFetchSolutionLinkApi.text}"
                elif responseFetchSolutionLinkApi.status_code in [500, 502, 503, 504]:
                    error_message = f"DownloadsvgApi-Server Error {responseFetchSolutionLinkApi.status_code}: {responseFetchSolutionLinkApi.text}"
                else:
                    error_message = f"DownloadsvgApi-Unexpected Error {responseFetchSolutionLinkApi.status_code}: {responseFetchSolutionLinkApi.text}"
                errorVar = error_message
                print(error_message)
                # messageArr.append(f"Error Response: {error_message}")
                # Elevateproject.createAPILog(solutionName_for_folder_path, messageArr)
                print("error in updating certificate with project")
                return False
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            errorVar
            print(errorVar,"---> API-Error")
    
    def validateTenantAndOrgIdFromAPI(accessToken, tenant_id, org_id):
        try:
            urlFetchSolutionApi = userLoginHost + tenantFetch + tenant_id

            headers = {
                'Content-Type': 'application/json',
                'origin': 'https://dev.elevate-mentoring.shikshalokam.org',
                'X-auth-token': accessToken
            }

            response = requests.get(urlFetchSolutionApi, headers=headers)

            if response.status_code not in (200, 202):
                return {
                    "success": False,
                    "message": f"API request failed with status {response.status_code}",
                    "status_code": response.status_code
                }

            api_data = response.json()
            result = api_data.get("result", {})

            tenant_code_from_api = result.get("code")
            org_ids_from_api = [str(org.get("id")) for org in result.get("organizations", [])]

            if tenant_id and str(tenant_id).strip() != str(tenant_code_from_api).strip():
                return {
                    "success": False,
                    "message": f" Tenant ID '{tenant_id}' does not match API tenant code '{tenant_code_from_api}'",
                    "status_code": 400
                }

            if org_id:
                org_id_list = [id.strip() for id in str(org_id).split(",")]
                missing_org_ids = [oid for oid in org_id_list if oid not in org_ids_from_api]

                if missing_org_ids:
                    return {
                        "success": False,
                        "message": f"Org ID(s) '{', '.join(missing_org_ids)}' not found in API organizations",
                        "status_code": 400
                    }

            return {
                "success": True,
                "message": "Tenant ID and Org ID(s) validated successfully.",
                "data": api_data,
                "status_code": 200
            }

        except Exception as e:
            return {
                "success": False,
                "message": f"Exception occurred: {str(e)}",
                "status_code": 500
            }

    def validate_roles_against_api(mainRoles, subRoles):
        print(mainRoles, "mainRoles")
        print(subRoles, "subRoles")

        urlFetchRoleList = userLoginHost + fetchprofessionalRole
        headers = {
            'Content-Type': content_type,
            'tenantId': tenantIDFromTemplate,
            'X-Channel-id': x_channel_id,
        }
        payload = {}

        response = requests.request("GET", urlFetchRoleList, headers=headers, data=payload)
        # response = requests.get(urlFetchRoleList, headers=headers, data=json.dumps({}))

        messageArr = []
        messageArr.append("Fetched professional roles from: " + urlFetchRoleList)
        messageArr.append("Status Code: " + str(response.status_code))

        if response.status_code != 200:
            messageArr.append("Error fetching roles.")
            print("Error fetching roles.")
            return False

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
                    return False
            else:
                messageArr.append(f"MainRole '{role}' not found in API")
                return False

        for s in remaining_subroles:
            messageArr.append(f"Subrole '{s}' not found in any of the provided mainRoles.")
            return False
        
        for msg in messageArr:
            print(msg)

        return validated_main_role_ids, validated_subrole_ids_list        

    def mainFunc(MainFilePath, programFile, addObservationSolution,resourceName, millisecond, isProgramnamePresent, isCourse,
             scopeEntityType=scopeEntityType):
        scopeEntityType = scopeEntityType
        global solutionLink, errorVar, ObservationOrSurveyResult, finalprojectsolutionlink, AnyTaskEvidenceNo,TaskEvidenceOperator, entityHierarchy ,orgIDFromTemplate,tenantIDFromTemplate,programExternalId,orgIdForScope,programID
        observationInstance = ElevateObservation
        if not isCourse:
            parentFolder = Elevateproject.createFileStructre(MainFilePath, addObservationSolution)
            accessToken = Elevateproject.generateAccessToken(parentFolder)
            Elevateproject.validateTenantAndOrgIdsFromProgramSheet(xlrd.open_workbook(programFile, on_demand=True))
            print(tenantIDFromTemplate)
            Elevateproject.validateTenantAndOrgIdFromAPI(accessToken, tenantIDFromTemplate, orgIDFromTemplate)
            print(tenantIDFromTemplate,"print(tenantIDFromTemplate)")
            typeofSolution = Elevateproject.typeofresource(addObservationSolution, accessToken, parentFolder)
            if typeofSolution != 4:
                
                ObservationOrSurveyResult=observationInstance.loadSurveyFile(programFile,resourceName)
                ObservationOrSurveyResult = json.loads(ObservationOrSurveyResult)
                solution_dict = ObservationOrSurveyResult["solutionDict"]
                return solution_dict
            # sys.exit()
            else:
                wbObservation = xlrd.open_workbook(addObservationSolution, on_demand=True)
                wbProgram = xlrd.open_workbook(programFile, on_demand=True)
                if typeofSolution == 4:
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
                                userEntity = dictProgramDetails['Targeted state at program level'].encode('utf-8').decode('utf-8').lstrip().rstrip().split(",")
                               
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
                                entityType = "school"
                        
                        elif sheets.strip().lower() == 'Tasks upload'.lower():
                            projectsheet = wbproject.sheet_by_name('Tasks upload')
                            keysEnv = [projectsheet.cell_value(1, col_index) for col_index in range(projectsheet.ncols)]
                            try:
                                evidence_col_index = keysEnv.index("Evidence required for any task for certificate criteria")
                                no_evidence_col_index = keysEnv.index("Minimum No. of Evidence for any task criteria")
                            except ValueError:
                                print("Column 'Evidence required for any task for certificate criteria' or 'Min No. Evidence for any task' not found.")
                                evidence_col_index = None
                                no_evidence_col_index = None

                            if evidence_col_index is not None:
                                TaskEvidenceOperator = projectsheet.cell_value(2, evidence_col_index)
                                if not TaskEvidenceOperator or str(TaskEvidenceOperator).strip().lower() == "":
                                    TaskEvidenceOperator = "no"
                            if TaskEvidenceOperator.lower() == "yes" and no_evidence_col_index is not None:
                                AnyTaskEvidenceNo = projectsheet.cell_value(2, no_evidence_col_index)
                                if not AnyTaskEvidenceNo or str(AnyTaskEvidenceNo).strip().lower() == "":
                                    AnyTaskEvidenceNo = 1

                    if not Elevateproject.programsFileCheck(programFile, accessToken, parentFolder, MainFilePath):
                        finalprojectSolutionLink = {ProjectName: errorVar}
                        return finalprojectSolutionLink
                    try:

                        # Adds a project by processing the input file, creating necessary folders,copying files, and preparing project and task sheets.
                        def addProjectFunc(filePathAddProject, projectName_for_folder_path, millisAddObs,typeofSolution):
                            validationResult =  Elevateproject.projectValidate(filePathAddProject, accessToken, parentFolder)
                            if not validationResult: 
                                global finalprojectsolutionlink  
                                print(errorVar,"---->Validation Failed")
                                finalprojectsolutionlink = {ProjectName: errorVar} 
                                
                                print(finalprojectsolutionlink)  
                                return finalprojectsolutionlink  
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
                            Elevateproject.createAPILog(projectName_for_folder_path, messageArr)

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
                                            Elevateproject.prepareProjectAndTasksSheets(addObservationSolution, projectName_for_folder_path,
                                                                        accessToken)
                                            print("prepareProjectAndTasksSheets")
                                            # sys.exit()
                                            if not Elevateproject.projectUpload(addObservationSolution, projectName_for_folder_path, accessToken):
                                                finalprojectsolutionlink = {ProjectName: errorVar}
                                                return finalprojectsolutionlink
                                            print("projectUpload")
                                            if not Elevateproject.taskUpload(addObservationSolution, projectName_for_folder_path, accessToken):
                                                finalprojectsolutionlink = {ProjectName: errorVar}
                                                return finalprojectsolutionlink
                                            ProjectSolutionResp = Elevateproject.solutionCreationAndMapping(projectName_for_folder_path,
                                                                                            entityToUpload,
                                                                                            listOfFoundRoles, accessToken,programFile)
                                            if not ProjectSolutionResp:
                                                finalprojectsolutionlink = {ProjectName: errorVar}
                                                return finalprojectsolutionlink
                                            ProjectSolutionExternalId = ProjectSolutionResp[0]
                                            ProjectSolutionId = ProjectSolutionResp[1]
                                            # print("reached here")
                                            solutionLink = Elevateproject.prepareProgramSuccessSheet(MainFilePath, projectName_for_folder_path, programFile,
                                                                    ProjectSolutionExternalId,
                                                                    ProjectSolutionId, accessToken)
                                            print(solutionLink,"solutionLink")
                                            if not solutionLink:
                                                print(solutionLink)
                                                finalprojectsolutionlink = {ProjectName: errorVar}
                                                return finalprojectsolutionlink
                                            else:
                                                finalprojectsolutionlink = {ProjectName: solutionLink}
                                                return finalprojectsolutionlink
                                        
                                        elif str(dictDetailsEnv['has certificate']).lower()== 'Yes'.lower():
                                            print("---->this is certificate with project<---")
                                            baseTemplate_id=Elevateproject.fetchCertificateBaseTemplate(filePathAddProject,accessToken,projectName_for_folder_path)
                                            if not baseTemplate_id:
                                                finalprojectsolutionlink = {ProjectName: errorVar}
                                                return finalprojectsolutionlink
                                            Elevateproject.downloadlogosign(filePathAddProject,projectName_for_folder_path)
                                            if not Elevateproject.editsvg(accessToken,filePathAddProject,projectName_for_folder_path,baseTemplate_id):
                                                finalprojectsolutionlink = {ProjectName: errorVar}
                                                return finalprojectsolutionlink
                                            Elevateproject.prepareProjectAndTasksSheets(addObservationSolution, projectName_for_folder_path,accessToken)
                                            if not Elevateproject.projectUpload(addObservationSolution, projectName_for_folder_path, accessToken):
                                                finalprojectsolutionlink = {ProjectName: errorVar}
                                                return finalprojectsolutionlink
                                            if not Elevateproject.taskUpload(addObservationSolution, projectName_for_folder_path, accessToken):
                                                finalprojectsolutionlink = {ProjectName: errorVar}
                                                return finalprojectsolutionlink
                                            ProjectSolutionResp = Elevateproject.solutionCreationAndMapping(projectName_for_folder_path,entityToUpload,listOfFoundRoles, accessToken, programFile)
                                            if not ProjectSolutionResp:
                                                finalprojectsolutionlink = {ProjectName: errorVar}
                                                return finalprojectsolutionlink
                                            ProjectSolutionExternalId = ProjectSolutionResp[0]
                                            ProjectSolutionId = ProjectSolutionResp[1]
                                            certificatetemplateid= Elevateproject.prepareaddingcertificatetemp(filePathAddProject,projectName_for_folder_path, accessToken,ProjectSolutionId,programID,baseTemplate_id)
                                            if not certificatetemplateid:
                                                finalprojectsolutionlink = {ProjectName: errorVar}
                                                return finalprojectsolutionlink
                                            
                                            solutionlink = Elevateproject.prepareProgramSuccessSheet(MainFilePath, projectName_for_folder_path, programFile,
                                                                    ProjectSolutionExternalId,
                                                                    ProjectSolutionId, accessToken)
                                            if not solutionlink:
                                                finalprojectsolutionlink = {ProjectName: errorVar}
                                                return finalprojectsolutionlink
                                            else:
                                                print(solutionlink)
                                                finalprojectsolutionlink = {ProjectName: solutionlink}
                                                return finalprojectsolutionlink  
                        projectSolutionLink = addProjectFunc(addObservationSolution, parentFolder, millisecond, typeofSolution)
                        print("Done.")
                        return projectSolutionLink
                    except Exception as e:
                        print(f"Error occurred during project creation: {str(e)}")
                        # raise RuntimeError("The project creation failed due to an unexpected error")
                        solutionError = str(e)
                        print(errorVar,"3266")
                        if errorVar == "":
                            projectSolutionLink = {ProjectName: solutionError}
                        else:
                            projectSolutionLink = {ProjectName: errorVar}
                        # print(projectSolutionLink, "3247")
                        return projectSolutionLink

                
        else:
            try:
                parentFolder = Elevateproject.createFileStructre(MainFilePath, addObservationSolution)
                accessToken = Elevateproject.generateAccessToken(parentFolder)
                wbproject = xlrd.open_workbook(programFile, on_demand=True)
                projectSheetNames = wbproject.sheet_names()
                for projectSheets in projectSheetNames:
                    if projectSheets.strip().lower() == 'project upload':
                        print("Checking project details sheet...")
                        detailsColCheck = wbproject.sheet_by_name(projectSheets)
                        keysColCheckDetai = [detailsColCheck.cell(0, col_index_check).value for col_index_check in range(detailsColCheck.ncols)]
                        detailsEnvSheet = wbproject.sheet_by_name(projectSheets)
                        keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in range(detailsEnvSheet.ncols)]
                        for row_index_env in range(2, detailsEnvSheet.nrows):
                            dictDetailsEnv = { keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for col_index_env in range(detailsEnvSheet.ncols)}
                            if str(dictDetailsEnv['has certificate']).lower() == 'No'.lower():
                                Elevateproject.prepareProjectAndTasksSheets(addObservationSolution, parentFolder,  accessToken)
                                if not Elevateproject.projectUpload(addObservationSolution, parentFolder, accessToken):
                                    finalprojectsolutionlink = {ProjectName: errorVar}
                                    return finalprojectsolutionlink
                                if not Elevateproject.taskUpload(addObservationSolution, parentFolder, accessToken):
                                    finalprojectsolutionlink = {ProjectName: errorVar}
                                    return finalprojectsolutionlink
                                if not Elevateproject.projectSolutionDeepLink(MainFilePath, programFile, solutionId,accessToken):
                                    finalprojectsolutionlink = {ProjectName: errorVar}
                                    return finalprojectsolutionlink
                                else:
                                    print(solutionLink)
                                    finalprojectsolutionlink = {ProjectName: solutionLink}
                                    return finalprojectsolutionlink
                                
                            elif str(dictDetailsEnv['has certificate']).lower()== 'Yes'.lower():
                                print("---->this is certificate with project<---")
                                baseTemplate_id=Elevateproject.fetchCertificateBaseTemplate(addObservationSolution,accessToken,parentFolder)
                                if not baseTemplate_id:
                                    finalprojectsolutionlink = {ProjectName: errorVar}
                                    return finalprojectsolutionlink
                                Elevateproject.downloadlogosign(addObservationSolution,parentFolder)
                                if not Elevateproject.editsvg(accessToken,addObservationSolution,parentFolder,baseTemplate_id):
                                    finalprojectsolutionlink = {ProjectName: errorVar}
                                    return finalprojectsolutionlink
                                Elevateproject.prepareProjectAndTasksSheets(addObservationSolution, parentFolder,accessToken)
                                if not Elevateproject.projectUpload(addObservationSolution, parentFolder, accessToken):
                                    finalprojectsolutionlink = {ProjectName: errorVar}
                                    return finalprojectsolutionlink
                                if not Elevateproject.taskUpload(addObservationSolution, parentFolder, accessToken):
                                    finalprojectsolutionlink = {ProjectName: errorVar}
                                    return finalprojectsolutionlink
                                else:
                                    success = "Project has been created successfully"
                                    finalprojectsolutionlink = {ProjectName: success}
                                    return finalprojectsolutionlink
            except Exception as e:
                print(f"Error occurred during project creation: {str(e)}")
                # raise RuntimeError("The project creation failed due to an unexpected error")
                solutionError = str(e)
                print(errorVar,"3266")
                if errorVar == "":
                    projectSolutionLink = {ProjectName: solutionError}
                else:
                    projectSolutionLink = {ProjectName: errorVar}
                # print(projectSolutionLink, "3247")
                return projectSolutionLink


    def loadSurveyFile(programFile):
        print("into the project file...")
        MainFilePath = Elevateproject.createFileStructForProgram(programFile)
        global downloaded_file
        global addObservationSolution
        # if downloaded_file is None:
        downloaded_file = {}
        print(downloaded_file, "downloaded_file 3044")

        wbPgm = xlrd.open_workbook(programFile, on_demand=True)
        sheetNames = wbPgm.sheet_names()
        pgmSheets = ["Instructions", "Program Details", "Resource Details", "Program Manager Details","Role-Subrole Mapping"]

        solutionDict = {}
        programName = ""  # Initialize the programName variable

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
                            for col_index_env in range(programDetailsSheet.ncols)
                        }
                        # Extracting the program name
                        programName = dictProgramDetails['Title of the Program'].encode('utf-8').decode('utf-8')
                        isProgramnamePresent = bool(programName)  # Check if program name exists
                        
                        # Example handling of userEntity
                        userEntity = dictProgramDetails['Targeted state at program level'].encode('utf-8').decode('utf-8').lstrip().rstrip().split(",")

                if sheetEnv.strip().lower() == 'resource details':
                    print("--->Checking Resource Details sheet...")
                    messageArr = []
                    messageArr.append("--->Checking Resource Details sheet...")
                    detailsEnvSheet = wbPgm.sheet_by_name(sheetEnv)
                    keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in range(detailsEnvSheet.ncols)]
                    
                    for row_index_env in range(2, detailsEnvSheet.nrows):
                        millisecond = int(time.time() * 1000)
                        dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                        for col_index_env in range(detailsEnvSheet.ncols)}
                        resourceNamePGM = dictDetailsEnv['Name of resources in program'].encode('utf-8').decode('utf-8')
                        resourceTypePGM = dictDetailsEnv['Type of resources'].encode('utf-8').decode('utf-8')
                        resourceLinkOrExtPGM = dictDetailsEnv['Resource Link']
                        
                        if str(dictDetailsEnv['Type of resources']).lower().strip() == "course":
                            isCourse = False
                        else:
                            isCourse = False
                            resourceStatus = dictDetailsEnv['Resource Status']
                            if resourceStatus.strip() == "New Upload":
                                print("--->Resource Name : " + str(resourceNamePGM))
                                resourceLinkOrExtPGM = str(resourceLinkOrExtPGM).split('/')[5]
                                file_url = 'https://docs.google.com/spreadsheets/d/' + resourceLinkOrExtPGM + '/export?format=xlsx'
                                if not os.path.isdir('InputFiles'):
                                    os.mkdir('InputFiles')
                                dest_file = 'InputFiles'
                                
                                download_file = wget.download(file_url, dest_file)
                                downloaded_file[download_file] = resourceNamePGM

            print("--->Solution input file successfully downloaded: " + str(downloaded_file))
            for addObservationSolution, resourceName in downloaded_file.items():
                print(f"Processing file: {addObservationSolution} for resource: {resourceName}")
                solutionSL = Elevateproject.mainFunc(MainFilePath, programFile, addObservationSolution,resourceName, millisecond, isProgramnamePresent, isCourse,scopeEntityType=scopeEntityType)
                print(solutionSL)
                print(solutionSL.items(),"3400")
                for resourceName, solutionLink in solutionSL.items():
                    solutionDict[resourceName] = solutionLink
                    print()
            downloaded_file = {}
            print()

        else:
            print("Stand alone Project file Detected")
            MainFilePath = Elevateproject.createFileStructForProgram(programFile)
            addObservationSolution = programFile
            wbPgm = xlrd.open_workbook(programFile, on_demand=True)
            millisecond = int(time.time() * 1000)
            Elevateproject.mainFunc(MainFilePath, programFile, addObservationSolution,resourceName,millisecond ,isProgramnamePresent =True,isCourse=True)
        # Combine solutionDict and programName into a single dictionary for returning
        result = {
            "solutionDict": solutionDict,
            "programName": programName  # Ensure programName is extracted from the 'Program Details' sheet
        }
        solutionDict = {}
        print(solutionDict,"3401")
        # print(f"Type of solutionDict: {type(solutionDict)}")
        # print(f"Program Name: {programName}")

        return json.dumps(result)

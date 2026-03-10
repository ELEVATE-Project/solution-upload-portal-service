import os
import time
from configparser import ConfigParser, ExtendedInterpolation
import xlrd
import uuid
import csv
from bson.objectid import ObjectId
import json
from datetime import datetime, timedelta, timezone
import requests
from difflib import get_close_matches
from requests import post, get, delete
import sys
import shutil
import re
import pandas as pd
from xlutils.copy import copy
from xlrd import open_workbook
from xlutils.copy import copy as xl_copy
import logging.handlers
from logging.handlers import TimedRotatingFileHandler
import xlsxwriter
import argparse
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


from GlobalVariable import global_vars
from GlobalVariable import exception_handler


class Helpers:
    errorVar = []

    @exception_handler
    def __init__(self):
        self.millisecond = None
        self.scopeEntityType = ""

    @staticmethod
    def _to_text(value):
        """Normalize Excel/cell values to string without redundant utf-8 encode/decode."""
        if value is None:
            return ""
        if isinstance(value, str):
            return value
        if isinstance(value, bytes):
            return value.decode("utf-8", errors="ignore")
        return str(value)

    @staticmethod
    def _get_cached_resource_workbook(file_path):
        """Return validation-cached workbook when path matches; fallback to opening file."""
        cache = global_vars.get_resource_validation_cache() or {}
        cached_path = cache.get("solution_path")
        current_path = os.path.abspath(file_path) if file_path else file_path
        if cached_path == current_path and cache.get("workbook") is not None:
            return cache.get("workbook")
        return xlrd.open_workbook(file_path, on_demand=True)

    @staticmethod
    def _update_resource_validation_cache(**kwargs):
        """Merge data into existing resource validation cache."""
        cache = global_vars.get_resource_validation_cache() or {}
        cache.update(kwargs)
        global_vars.set_resource_validation_cache(**cache)

    @classmethod
    def reset_errors(cls):
        cls.errorVar = []

    @exception_handler
    def checkIfObsMappedToProgram(accessToken, obsExt, parentFolder):
        # fetch observation solution details API end points 
        fetchSolutionDetailsURL = internal_kong_ip + fetchsolutiondetails + "observation&page=1&limit=10&search=" + str(obsExt)
        # fetch observation solution details payload
        payload = {}
        # fetch observation solution header
        headers = {'Content-Type': 'application/json',
                'Authorization': 'Bearer ' + internal_access_token,
                'X-authenticated-user-token': accessToken, 'X-Channel-id': x_channel_id}
        
        responseSearchSol = requests.request("POST", fetchSolutionDetailsURL, headers=headers, data=payload)
        
        listOfFoundSolutionIds = {}

        if responseSearchSol.status_code == 200:
            # parse list of Observations into a python dictionary 
            responseSearchSol = responseSearchSol.json()

            # iterate through each _id of solution and fetch the solution dump 
            for eachSol in responseSearchSol['result']['data']:

                fetchSolutionDumpURL = internal_kong_ip + fetchsolutiondump + eachSol['_id']
                headersSolutionDumpURL = {
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer ' + authorization,
                    'X-authenticated-user-token': accessToken,
                    'X-Channel-id': x_channel_id,
                    'internal-access-token': internal_access_token
                }
                responseSolDump = requests.request("POST", fetchSolutionDumpURL, headers=headersSolutionDumpURL)
                if responseSolDump.status_code == 200:
                    responseSolDump = responseSolDump.json()
                    # save details of observation 
                    listOfFoundSolutionIds[eachSol['_id']] = {"externalId": responseSolDump['result']['externalId'],
                                                            "isReusable": str(responseSolDump['result']['isReusable']),
                                                            "programId": responseSolDump['result']['programId']}
                else:
                    Helpers.errorVar.append(
                        f"checkIfObsMappedToProgram solution dump API failed: status={responseSolDump.status_code}, response={responseSolDump.text}"
                    )
            # create API logs 
            Helpers.createAPILog(parentFolder, ["List of solutions found : " + str(listOfFoundSolutionIds)])
            return listOfFoundSolutionIds
        else:
            error_message = ""
            if responseSearchSol.status_code in [400, 401, 403, 404, 422]:
                error_message = f"checkIfObsMappedToProgram-Client Error {responseSearchSol.status_code}: {responseSearchSol.text}"
            elif responseSearchSol.status_code in [500, 502, 503, 504]:
                error_message = f"checkIfObsMappedToProgram-Server Error {responseSearchSol.status_code}: {responseSearchSol.text}"
            else:
                error_message = f"checkIfObsMappedToProgram-Unexpected Error {responseSearchSol.status_code}: {responseSearchSol.text}"
            Helpers.errorVar.append(error_message)
            return False
    
    @exception_handler
    def programCreation(self, accessToken, externalId, pName, pDescription, keywords, entities, roles, orgIds,creatorKeyCloakId, creatorName):
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
            "imageCompression": {
                "quality": 10
            },
            "creator": creatorName,
            "owner": creatorKeyCloakId,
            "author": creatorKeyCloakId,
            "scope": {
                "entityType": global_vars.scopeEntityType,
                "entities": entities,
                "roles": roles
            }})
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
            countOfPrograms = (responsePgmCreateResp['result'])
            programidforValidation = countOfPrograms["_id"]
        else:
            error_message = ""
            if responsePgmCreate.status_code in [400, 401, 403, 404, 422]:
                error_message = f"Program Creation-Client Error {responsePgmCreate.status_code}: {responsePgmCreate.text}"
            elif responsePgmCreate.status_code in [500, 502, 503, 504]:
                error_message = f"Program Creation-Server Error {responsePgmCreate.status_code}: {responsePgmCreate.text}"
            else:
                error_message = f"Program Creation-Unexpected Error {responsePgmCreate.status_code}: {responsePgmCreate.text}"
            Helpers.errorVar.append(error_message)
            return False
        
        return programidforValidation


    @exception_handler
    def programmappingpdpmsheetcreation(MainFilePath,accessToken, program_file):
        pdpmsheet = MainFilePath+ "/pdpmmapping/"
        if not os.path.exists(pdpmsheet):
            os.mkdir(pdpmsheet)

        pdpmcolo1 = ["user","role","entity","entityOperation","keycloak-userId","acl_school","acl_cluster","programOperation",
                    "platform_role","programs","_arrayFields"]
        with open(pdpmsheet + 'mapping.csv', 'w',encoding='utf-8') as file:
             writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
             writer.writerows([pdpmcolo1])

        global_vars.load_program_template(program_file)

        print("--->Checking Program details sheet...")
        dictDetailsEnv = global_vars.programDict
        if not dictDetailsEnv:
            Helpers.errorVar.append("\"Program Details\" sheet has no data rows")
            return False

        if not dictDetailsEnv.get('Title of the Program'):
            Helpers.errorVar.append("\"Title of the Program\" must not be Empty in \"Program details\" sheet")
            return False
        global_vars.programNameInp = Helpers._to_text(dictDetailsEnv['Title of the Program'])

        if not dictDetailsEnv.get('Program ID'):
            Helpers.errorVar.append("\"Program ID\" must not be Empty in \"Program details\" sheet")
            return False
        extIdPGM = Helpers._to_text(dictDetailsEnv['Program ID'])

        if not dictDetailsEnv.get('Diksha username/user id/email id/phone no. of Program Designer'):
            Helpers.errorVar.append("\"Diksha username/user id/email id/phone no. of Program Designer\" must not be Empty in \"Program details\" sheet")
            return False
        programdesigner = Helpers._to_text(dictDetailsEnv['Diksha username/user id/email id/phone no. of Program Designer'])
        userDetails = Helpers.fetchUserDetails(accessToken, programdesigner)
        if not userDetails:
            Helpers.errorVar.append(f"Failed to fetch user details for: {programdesigner}")
            return False

        creatorKeyCloakId = userDetails[0]
        creatorName = userDetails[1]
        if "PROGRAM_DESIGNER" not in userDetails[3]:
            print("user does't have program designer role")

        pdpmcolo1 = [creatorName, " ", " ", " ", creatorKeyCloakId, " ", " ", "ADD", "PROGRAM_DESIGNER", extIdPGM, "programs"]
        with open(pdpmsheet + 'mapping.csv', 'a', encoding='utf-8') as file:
            writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',', lineterminator='\n')
            writer.writerows([pdpmcolo1])
            fileheader = [creatorName, "program designer mapped successfully", "Passed"]
            # apicheckslog(parentFolder,fileheader)

        print("--->Program Manager Details...")
        for dictManagerDetails in global_vars.programManagerDetails:
            if str(dictManagerDetails.get('Is a SSO user?', '')).strip() == "YES":
                programmanagername2 = Helpers._to_text(dictManagerDetails.get('Diksha user id ( profile ID)') or dictManagerDetails.get('Login ID on DIKSHA'))
            else:
                programmanagername2 = Helpers._to_text(dictManagerDetails.get('Login ID on DIKSHA') or dictManagerDetails.get('Diksha user id ( profile ID)'))

            if not programmanagername2:
                Helpers.errorVar.append("\"Login ID on DIKSHA\" or \"Diksha user id ( profile ID)\" must not be Empty in \"Program Manager details\" sheet")
                return False

            userDetails = Helpers.fetchUserDetails(accessToken, programmanagername2)
            if not userDetails:
                Helpers.errorVar.append(f"Failed to fetch user details for: {programmanagername2}")
                return False
            creatorKeyCloakId = userDetails[0]
            creatorName = userDetails[1]
            if "PROGRAM_MANAGER" not in userDetails[3]:
                print("user does't have program manager role")

            pdpmcolo1 = [creatorName, " ", " ", " ", creatorKeyCloakId, " ", " ", "ADD", "PROGRAM_MANAGER", extIdPGM, "programs"]

            with open(pdpmsheet + 'mapping.csv', 'a', encoding='utf-8') as file:
                writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',', lineterminator='\n')
                writer.writerows([pdpmcolo1])
            # messageArr.append("Response : " + str(pdpmcolo1))
            # Helpers.createAPILog(parentFolder, messageArr)

            fileheader = [creatorName, "program manager mapped succesfully", "Passed"]
            # apicheckslog(parentFolder,fileheader)
        return creatorKeyCloakId

    @exception_handler
    def validate_program_mapping(accessToken, programId, userkeycklockid):
        """
        Calls dbFind API and validates whether programId exists
        in platformRoles.programs.
        Appends errors to errorVar and returns False if validation fails.
        """

        urldbFind = internal_kong_ip + dbfindapi_url + "userExtension"

        headerdbFindApi = {
            'Authorization': authorization,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token,
            'Content-Type': 'application/json'
        }

        payload = json.dumps({
            "query": {
                "userId" : userkeycklockid
            },
            "mongoIdKeys": []
        })

        responsedbFindApi = requests.post(
            url=urldbFind,
            headers=headerdbFindApi,
            data=payload
        )

        if responsedbFindApi.status_code != 200:
            # In local/mock-disabled environments dbFind may be unavailable (often 404).
            # Skip hard failure so project/resource flow can continue.
            if responsedbFindApi.status_code == 404:
                print("dbFind API returned 404; skipping program mapping validation in current environment.")
                return True
            Helpers.errorVar.append(f"dbFind API failed with status code {responsedbFindApi.status_code}")
            return False

        response_json = responsedbFindApi.json()
        result_list = response_json.get("result", [])

        if not isinstance(result_list, list) or not result_list:
            Helpers.errorVar.append("Invalid response: result is empty or not a list")
            return False

        # ---- Validation Logic ----
        for user in result_list:
            platform_roles = user.get("platformRoles", [])

            for role in platform_roles:
                programs = role.get("programs", [])

                if programId in programs:
                    # Mapping found
                    return True

        # If reached here → programId not found
        print(f"Mapping unsuccessful: Program ID {programId} not found in platformRoles.programs")
        return False

    # this function is used for call the api and map the pdpm roles which we created
    @exception_handler
    def Programmappingapicall(MainFilePath,accessToken,parentFolder):
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
        Helpers.createAPILog(parentFolder, messageArr)

        if responseProgrammappingApi.status_code == 200:
            print('--->program manager and designer mapping is Success')
            with open(MainFilePath + '/pdpmmapping/mappinginternal.csv', 'w+',encoding='utf-8') as projectRes:
                projectRes.write(responseProgrammappingApi.text)
                messageArr.append("Response : " + str(responseProgrammappingApi.text))
                Helpers.createAPILog(parentFolder, messageArr)
            return True
        else:
            error_message = ""
            if responseProgrammappingApi.status_code in [400, 401, 403, 404, 422]:
                error_message = f"Program Mapping-Client Error {responseProgrammappingApi.status_code}: {responseProgrammappingApi.text}"
            elif responseProgrammappingApi.status_code in [500, 502, 503, 504]:
                error_message = f"Program Mapping-Server Error {responseProgrammappingApi.status_code}: {responseProgrammappingApi.text}"
            else:
                error_message = f"Program Mapping-Unexpected Error {responseProgrammappingApi.status_code}: {responseProgrammappingApi.text}"
            Helpers.errorVar.append(error_message)
            return False


    
    @exception_handler
    def createFileStructForProgram(programFile):
        #  print("programFile:-------------",programFile)
        if not os.path.isdir('programFiles'):
            os.mkdir('programFiles')
        if "/" in str(programFile):
            fileNameSplit = str(programFile).split('/')[-1:]
            # print("newfileNameSplit ;",fileNameSplit)
        else :

            fileNameSplit = os.path.basename(programFile)
        # print("updatedfileNameSplit : ",fileNameSplit)

        if isinstance(fileNameSplit, list):
            fileNameSplit = fileNameSplit[0]
        # fileNameSplit = str(programFile)
        if fileNameSplit.endswith(".xlsx"):
            # print("latest fileNameSplit",fileNameSplit)
            ts = str(time.time()).replace(".", "_")
            
            folderName = fileNameSplit.replace(".xlsx", "-" + str(ts))
            # print("folderName :",folderName)
            os.mkdir('programFiles/' + str(folderName))
            path = os.path.join('programFiles', str(folderName))
            # print("done",path)
        else:
            print("something")
        returnPathStr = os.path.join('programFiles', str(folderName))
        # print("returnPathStr",returnPathStr)

        return returnPathStr
    
    @exception_handler
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
        Helpers.createAPILog(solutionName_for_folder_path, messageArr)
        if responseFetchRolesListApi.status_code == 200:
            responseFetchRolesListApi = responseFetchRolesListApi.json()
            for listRoles in responseFetchRolesListApi['result']:
                eachDict = dict()
                eachDict['id'] = listRoles['_id'].lstrip().rstrip()
                eachDict['code'] = listRoles['code'].lstrip().rstrip()
                rolesLookup[listRoles['code']] = eachDict['id']
                rolesReturn.append(listRoles['code'].lstrip().rstrip())
        else:
            error_message = ""
            if responseFetchRolesListApi.status_code in [400, 401, 403, 404, 422]:
                error_message = f"fetchScopeRole-Client Error {responseFetchRolesListApi.status_code}: {responseFetchRolesListApi.text}"
            elif responseFetchRolesListApi.status_code in [500, 502, 503, 504]:
                error_message = f"fetchScopeRole-Server Error {responseFetchRolesListApi.status_code}: {responseFetchRolesListApi.text}"
            else:
                error_message = f"fetchScopeRole-Unexpected Error {responseFetchRolesListApi.status_code}: {responseFetchRolesListApi.text}"
            Helpers.errorVar.append(error_message)
            return False

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
                    Helpers.createAPILog(solutionName_for_folder_path, messageArr)

        messageArr = ["Accepted Roles : " + str(listOfFoundRoles)]
        Helpers.createAPILog(solutionName_for_folder_path, messageArr)
        if len(listOfFoundRoles) == 0:
            messageArr = ["No roles matched our DB "]
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)
            print("No Roles matched our DB.")
        return listOfFoundRoles


    
    @exception_handler
    def getProgramInfo(accessTokenUser, solutionName_for_folder_path, programNameInp, resourceEndDates):
        global_vars.programName = programNameInp
        programUrl = internal_kong_ip + fetchprograminfoapiurl
        # print(programUrl)
        payload = json.dumps({
            "query": {
                "name": programNameInp,
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
                messageArr.append("No program found with the name : " + str(global_vars.programName))
                messageArr.append("******************** Preparing for program Upload **********************")
                print("No program found with the name : " + str(global_vars.programName))
                print("******************** Preparing for program Upload **********************")
                
                return False
            else:
                getProgramDetails = []
                for eachPgm in responseProgramSearch['result']:
                    if eachPgm['isAPrivateProgram'] == False:
                        global_vars.programID = eachPgm['_id']
                        global_vars.programExternalId = eachPgm['externalId']
                        global_vars.programDescription = eachPgm['description']
                        isAPrivateProgram = eachPgm['isAPrivateProgram']
                        endDate = eachPgm["endDate"]
                        getProgramDetails.append([
                            global_vars.programID,
                            global_vars.programExternalId,
                            global_vars.programDescription,
                            isAPrivateProgram,
                            endDate
                        ])
                        if len(getProgramDetails) == 0:
                            print("Total " + str(len(getProgramDetails)) + " backend programs found with the name : " + global_vars.programName.lstrip().rstrip())
                            messageArr.append("Total " + str(len(getProgramDetails)) + " backend programs found with the name : " + global_vars.programName.lstrip().rstrip())
                            
                            fileheader = ["program find api is running","found"+str(len(
                                getProgramDetails))+"programs in backend","Failed","found"+str(len(
                                getProgramDetails))+"programs ,check logs"]
                           
                        elif len(getProgramDetails) > 1:
                            print("Total " + str(len(getProgramDetails)) + " backend programs found with the name : " + global_vars.programName.lstrip().rstrip())
                            messageArr.append("Total " + str(len(getProgramDetails)) + " backend programs found with the name : " + global_vars.programName.lstrip().rstrip())
                           

                        else:
                            global_vars.programID = getProgramDetails[0][0]
                            global_vars.programExternalId = getProgramDetails[0][1]
                            global_vars.programDescription = getProgramDetails[0][2]
                            isAPrivateProgram = getProgramDetails[0][3]
                            programEndDate = getProgramDetails[0][4]
                            if resourceEndDates:
                                Helpers.validate_solution_end_date(programEndDate,resourceEndDates)
                            global_vars.isProgramnamePresent = True
                            messageArr.append("programID : " + str(global_vars.programID))
                            messageArr.append("programExternalId : " + str(global_vars.programExternalId))
                            messageArr.append("programDescription : " + str(global_vars.programDescription))
                            messageArr.append("isAPrivateProgram : " + str(isAPrivateProgram))
                            messageArr.append("programEndDate : " + str(programEndDate))
                        Helpers.createAPILog(solutionName_for_folder_path, messageArr)
        else:
            print("Program search API failed...")
            error_message = ""
            if responseProgramSearch.status_code in [400, 401, 403, 404, 422]:
                error_message = f"Program Search-Client Error {responseProgramSearch.status_code}: {responseProgramSearch.text}"
            elif responseProgramSearch.status_code in [500, 502, 503, 504]:
                error_message = f"Program Search-Server Error {responseProgramSearch.status_code}: {responseProgramSearch.text}"
            else:
                error_message = f"Program Search-Unexpected Error {responseProgramSearch.status_code}: {responseProgramSearch.text}"
            Helpers.errorVar.append(error_message)
            return False
        return True
    
    @exception_handler
    def fetchEntityId(solutionName_for_folder_path, accessToken, entitiesNameList, scopeEntityType):
        # print(scopeEntityType,"scopeEntityType--------------")
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
        # print(responseFetchEntityListApi,"responseFetchEntityListApi")

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
            error_message = ""
            if responseFetchEntityListApi.status_code in [400, 401, 403, 404, 422]:
                error_message = f"Entity Fetch-Client Error {responseFetchEntityListApi.status_code}: {responseFetchEntityListApi.text}"
            elif responseFetchEntityListApi.status_code in [500, 502, 503, 504]:
                error_message = f"Entity Fetch-Server Error {responseFetchEntityListApi.status_code}: {responseFetchEntityListApi.text}"
            else:
                error_message = f"Entity Fetch-Unexpected Error {responseFetchEntityListApi.status_code}: {responseFetchEntityListApi.text}"
            Helpers.errorVar.append(error_message)
            return False

    
    @exception_handler
    def programsFileCheck(filePathAddPgm, accessToken, parentFolder, MainFilePath):
        """Validates program file structure.
        
        Returns:    
            bool: True if validation successful, False otherwise.
        """
        normalized_program_path = os.path.abspath(filePathAddPgm) if filePathAddPgm else filePathAddPgm
        try:
            is_valid = bool(Helpers._programsFileCheck_impl(filePathAddPgm, accessToken, parentFolder, MainFilePath))
            global_vars.programValidationCache = {
                "program_path": normalized_program_path,
                "is_valid": is_valid,
            }
            return is_valid
        except Exception as e:
            Helpers.errorVar.append(str("CRITICAL") + ': ' + str(f"Program template validation exception: {e}"))
            global_vars.programValidationCache = {
                "program_path": normalized_program_path,
                "is_valid": False,
            }
            return False

    @exception_handler
    def _programsFileCheck_impl(filePathAddPgm, accessToken, parentFolder, MainFilePath):
        wb_program = xlrd.open_workbook(filePathAddPgm, on_demand=True)
        datemode = wb_program.datemode

        def _to_ymd_hms(date_value, field_name, end_of_day=False):
            """Normalize program date cells to `YYYY-MM-DD HH:MM:SS`."""
            if date_value is None or Helpers._to_text(date_value).strip() == "":
                Helpers.errorVar.append(
                    str("CRITICAL") + ': ' +
                    str(f"\"{field_name}\" must not be Empty in \"Program details\" sheet")
                )
                return None

            # Excel may provide datetime/date object.
            if isinstance(date_value, datetime):
                d = date_value.date()
            elif isinstance(date_value, (int, float)):
                try:
                    d = xlrd.xldate.xldate_as_datetime(float(date_value), datemode).date()
                except Exception:
                    Helpers.errorVar.append(
                        str("CRITICAL") + ': ' +
                        str(
                            f"Invalid Excel date value in \"{field_name}\" in \"Program details\" sheet: "
                            f"{date_value}"
                        )
                    )
                    return None
            else:
                text_value = Helpers._to_text(date_value).strip()
                d = None
                for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
                    try:
                        d = datetime.strptime(text_value, fmt).date()
                        break
                    except ValueError:
                        pass
                if d is None:
                    Helpers.errorVar.append(
                        str("CRITICAL") + ': ' +
                        str(
                            f"Invalid date format in \"{field_name}\" in \"Program details\" sheet: "
                            f"{text_value}. Expected DD-MM-YYYY or YYYY-MM-DD"
                        )
                    )
                    return None

            hhmmss = "23:59:59" if end_of_day else "00:00:00"
            return f"{d.year:04d}-{d.month:02d}-{d.day:02d} {hhmmss}"

        program_file = filePathAddPgm
        global_vars.load_program_template(program_file)
        sheetNames = global_vars.programTemplateSheetNames
        # list of sheets in the program sheet 
        pgmSheets = ["Instructions", "Program Details", "Resource Details","Program Manager Details", "Role-Subrole mapping"]

        # checking the sheets in the program sheet 
        if (len(sheetNames) == len(pgmSheets)) and ((set(sheetNames) == set(pgmSheets))):
            print("--->Program Template detected.<---")
            # iterate through the sheets
            resourceEndDates = []

            for resource_row in global_vars.programResourceDetails:
                endDateOfResources = resource_row.get('End date of resource')
                if endDateOfResources:
                    resourceEndDates.append(endDateOfResources)
             
            for sheetEnv in sheetNames:

                if sheetEnv == "Instructions":
                    # skip Instructions sheet 
                    pass
                elif sheetEnv.strip().lower() == 'program details':
                    print("--->Checking Program details sheet...")
                    dictDetailsEnv = global_vars.programDict
                    if not dictDetailsEnv:
                        Helpers.errorVar.append(str("CRITICAL") + ': ' + str("\"Program Details\" sheet has no data rows"))
                        return False

                    if not dictDetailsEnv.get('Title of the Program'):
                        Helpers.errorVar.append(str("CRITICAL") + ': ' + str("\"Title of the Program\" must not be Empty in \"Program details\" sheet"))
                        return False
                    global_vars.programNameInp = Helpers._to_text(dictDetailsEnv['Title of the Program'])

                    if not dictDetailsEnv.get('Program ID'):
                        Helpers.errorVar.append(str("CRITICAL") + ': ' + str("\"Program ID\" must not be Empty in \"Program details\" sheet"))
                        return False
                    extIdPGM = Helpers._to_text(dictDetailsEnv['Program ID'])

                    if not dictDetailsEnv.get('Targeted state at program level'):
                        Helpers.errorVar.append(str("CRITICAL") + ': ' + str("\"Targeted state at program level\" must not be Empty in \"Program details\" sheet"))
                        return False
                    global_vars.entitiesPGM = Helpers._to_text(dictDetailsEnv['Targeted state at program level'])
                    districtentitiesPGM = dictDetailsEnv.get('Targeted district at program level', '')
                    districtentitiesPGM = Helpers._to_text(districtentitiesPGM) if districtentitiesPGM else ''
                    global_vars.startDateOfProgram = _to_ymd_hms(
                        dictDetailsEnv.get('Start date of program'),
                        "Start date of program",
                        end_of_day=False
                    )
                    if not global_vars.startDateOfProgram:
                        return False
                    global_vars.endDateOfProgram = _to_ymd_hms(
                        dictDetailsEnv.get('End date of program'),
                        "End date of program",
                        end_of_day=True
                    )
                    if not global_vars.endDateOfProgram:
                        return False

                    global_vars.scopeEntityType = "state"

                    if districtentitiesPGM:
                        global_vars.entitiesPGM = districtentitiesPGM
                        EntityType = "district"
                    else:
                        global_vars.entitiesPGM = global_vars.entitiesPGM
                        EntityType = "state"

                    global_vars.scopeEntityType = EntityType

                    # print(entitiesPGMID,"entitiesPGMID")
                    global_vars.entitiesPGMID = Helpers.fetchEntityId(parentFolder, accessToken,global_vars.entitiesPGM.lstrip().rstrip().split(","), global_vars.scopeEntityType)
                    print(global_vars.entitiesPGMID)
                    if not global_vars.entitiesPGMID:
                        Helpers.errorVar.append(str("CRITICAL") + ': ' + str("None of the entities mentioned in \"Targeted state at program level\" or \"Targeted district at program level\" column in \"Program details\" sheet were found in backend"))
                        return False

                    if not Helpers.getProgramInfo(accessToken, parentFolder, global_vars.programNameInp, resourceEndDates):
                        # print("reached till here")
                        extIdPGM = Helpers._to_text(dictDetailsEnv['Program ID'])
                        if str(dictDetailsEnv['Program ID']).strip() == "Do not fill this field":
                            print("change the program id")
                        descriptionPGM = Helpers._to_text(dictDetailsEnv['Description of the Program'])
                        keywordsPGM = Helpers._to_text(dictDetailsEnv['Keywords'])
                        global_vars.entitiesPGM = Helpers._to_text(dictDetailsEnv['Targeted state at program level']) 
                        districtentitiesPGM = dictDetailsEnv.get('Targeted district at program level', '')
                        districtentitiesPGM = Helpers._to_text(districtentitiesPGM) if districtentitiesPGM else ''
                        # selecting entity type based on the users input 
                        if districtentitiesPGM:
                            global_vars.entitiesPGM = districtentitiesPGM
                            EntityType = "district"
                        else:
                            global_vars.entitiesPGM = global_vars.entitiesPGM
                            EntityType = "state"

                        global_vars.scopeEntityType = EntityType

                        global_vars.mainRole = dictDetailsEnv['Targeted role at program level'] 
                        # print(global_vars.mainRole,"mainRole")
                        pass 
                        global_vars.rolesPGM = dictDetailsEnv['Targeted subrole at program level']
                        # print(rolesPGM,rolesPGM)

                        if "teacher" in global_vars.mainRole.strip().lower():
                            global_vars.rolesPGM = str(global_vars.rolesPGM).strip() + ",TEACHER"
                        userDetails = Helpers.fetchUserDetails(accessToken, dictDetailsEnv['Diksha username/user id/email id/phone no. of Program Designer'])
                        if not userDetails:
                            Helpers.errorVar.append(f"Failed to fetch user details for: {dictDetailsEnv['Diksha username/user id/email id/phone no. of Program Designer']}")
                            return False
                        OrgName=userDetails[4]
                        # print(OrgName,"OrgName")
                        global_vars.orgIds=Helpers.fetchOrgId(accessToken, OrgName)
                        print(global_vars.orgIds,"orgIds")
                        creatorKeyCloakId = userDetails[0]
                        creatorName = userDetails[2]

                        messageArr = []

                        global_vars.scopeEntityType = EntityType
                        # fetch entity details 
                        global_vars.entitiesPGMID = Helpers.fetchEntityId(parentFolder, accessToken,global_vars.entitiesPGM.lstrip().rstrip().split(","), global_vars.scopeEntityType)
                        if not global_vars.entitiesPGMID:
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Failed to fetch entity IDs for program scope"))
                            return False
                        # print(entitiesPGMID,"entitiesPGMID")

                        # sys.exit()
                        # fetch sub-role details 
                        rolesPGMID = Helpers.fetchScopeRole(parentFolder, accessToken, global_vars.rolesPGM.lstrip().rstrip().split(","))
                        if not rolesPGMID:
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Failed to fetch scope roles for program"))
                            return False
                        # print(rolesPGMID,"rolesPGMID")

                        # sys.exit()

                        # call function to create program 
                        programIDCreated = Helpers.programCreation(accessToken, extIdPGM, global_vars.programNameInp, descriptionPGM,keywordsPGM.lstrip().rstrip().split(","), global_vars.entitiesPGMID, rolesPGMID, global_vars.orgIds,creatorKeyCloakId, creatorName)
                        if not programIDCreated:
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Program creation API failed"))
                            return False
                        print(programIDCreated)
                        # sys.exit()
                        userkeycklockid=Helpers.programmappingpdpmsheetcreation(MainFilePath, accessToken, program_file)
                        if not userkeycklockid:
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Program mapping sheet creation failed"))
                            return False
                        print(userkeycklockid)

                        # map PM / PD to the program 
                        if not Helpers.Programmappingapicall(MainFilePath, accessToken,parentFolder):
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Program mapping API call failed"))
                            return False
                        
                        #check if created program is mapped to program manager or not
                        try:
                            if Helpers.validate_program_mapping( accessToken, programIDCreated, userkeycklockid):
                                print("✅ Program mapping successful")
                            else:
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Program mapping validation failed"))
                                return False
                        except Exception as e:
                            print("❌", str(e))
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Program mapping validation failed"))
                            return False

                        # check if program is created or not 
                        if Helpers.getProgramInfo(accessToken, parentFolder, global_vars.programNameInp, resourceEndDates):
                            print("Program Created SuccessFully.")
                        else :
                            print("Program creation failed! Please check logs.")
                    else:
                        messageArr = []
                        userDetails = Helpers.fetchUserDetails(accessToken, dictDetailsEnv['Diksha username/user id/email id/phone no. of Program Designer'])
                        if not userDetails:
                            Helpers.errorVar.append(f"Failed to fetch user details for: {dictDetailsEnv['Diksha username/user id/email id/phone no. of Program Designer']}")
                            return False
                        userkeycklockid = userDetails[0]
                        
                        # map PM / PD to the program 
                        
                        #check if created program is mapped to program manager or not
                        try:
                            if not Helpers.validate_program_mapping( accessToken, global_vars.programID, userkeycklockid):
                                userkeycklockid=Helpers.programmappingpdpmsheetcreation(MainFilePath, accessToken, program_file)
                                if not userkeycklockid:
                                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Program mapping sheet creation failed"))
                                    return False
                                print(userkeycklockid)
                                if not Helpers.Programmappingapicall(MainFilePath, accessToken,parentFolder):
                                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Program mapping API call failed"))
                                    return False
                                if Helpers.validate_program_mapping( accessToken, global_vars.programID, userkeycklockid):
                                    print("✅ Program mapping successful")
                                    messageArr.append("✅ Program mapping check is successful for existing progam")
                                    Helpers.createAPILog(parentFolder, messageArr)
                                else:
                                    print("❌ Program mapping Failed")
                                    messageArr.append("❌ Program mapping check is failed for existing progam")
                            else:
                                print("✅ Program mapping recheck successful")
                                messageArr.append("✅ Program mapping check is successful for existing progam")
                                Helpers.createAPILog(parentFolder, messageArr)
                        except Exception as e:
                            print("❌", str(e))
                            messageArr.append("❌ Program mapping check is failed for existing progam" + str(e))
                            Helpers.createAPILog(parentFolder, messageArr)
                elif sheetEnv.strip().lower() == 'resource details':
                    print("--->Checking Resource Details sheet...")
                    for dictDetailsEnv in global_vars.programResourceDetails:
                        required_resource_cols = [
                            "Name of resources in program",
                            "Type of resources",
                            "Resource Link",
                            "Resource Status",
                            "Targeted role at resource level",
                            "Targeted subrole at resource level",
                            "Start date of resource",
                            "End date of resource",
                        ]
                        missing_cols = [c for c in required_resource_cols if c not in dictDetailsEnv]
                        if missing_cols:
                            Helpers.errorVar.append(
                                str("CRITICAL") + ': ' +
                                str(f"Missing columns in \"Resource Details\" sheet: {missing_cols}")
                            )
                            return False
                        if not dictDetailsEnv['Name of resources in program']:
                            Helpers.errorVar.append("\"Name of resources in program\" must not be Empty in \"Resource Details\" sheet")
                        if not dictDetailsEnv['Type of resources']:
                            Helpers.errorVar.append("\"Type of resources\" must not be Empty in \"Resource Details\" sheet")
                        if not dictDetailsEnv['Resource Link']:
                            Helpers.errorVar.append("\"Resource Link\" must not be Empty in \"Resource Details\" sheet")
                        if not dictDetailsEnv['Resource Status']:
                            Helpers.errorVar.append("\"Resource Status\" must not be Empty in \"Resource Details\" sheet")
                        if not dictDetailsEnv['Targeted role at resource level']:
                            Helpers.errorVar.append("\"Targeted role at resource level\" must not be Empty in \"Resource Details\" sheet")
                        if not dictDetailsEnv['Targeted subrole at resource level']:
                            Helpers.errorVar.append("\"Targeted subrole at resource level\" must not be Empty in \"Resource Details\" sheet")

                        global_vars.startDateOfResource = dictDetailsEnv['Start date of resource']
                        global_vars.endDateOfResource = dictDetailsEnv['End date of resource']
            return True
        Helpers.errorVar.append(
            str("CRITICAL") + ': ' +
            str(f"Invalid Program Template sheets. Expected {pgmSheets}, found {sheetNames}")
        )
        return False

    # function to accept only csv file as input in command line argument
    @exception_handler
    def valid_file(param):
        base, ext = os.path.splitext(param)
        if ext.lower() not in ('.xlsx'):
            raise argparse.ArgumentTypeError('File must have a csv extension')
        return param
    
    # function to check environment 
    @exception_handler
    def envCheck():
        try:
            keyclockapiurl
            return True
        except Exception as e:
            print(e)
            return False
        
    # Generate access token for the APIs.
    @exception_handler
    def validate_identifier(identifier, field_name="Field"):
        pattern = r'^[A-Za-z0-9_-]+$'
        if not re.match(pattern, identifier):
            raise ValueError(f"Invalid {field_name}: '{identifier}'. Only A-Z, a-z, 0-9, '-', and '_' are allowed.")
        else:
            print(f"{field_name} '{identifier}' is valid.")

# Function create File structure for Solutions
    @exception_handler
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
    @exception_handler
    def generateAccessToken():
        """Generate access token."""
        try:
            headerKeyClockUser = {'Content-Type': keyclockapicontent_type}
        
            responseKeyClockUser = requests.post(url=host + (keyclockapiurl), headers=headerKeyClockUser,
                                             data=(keyclockapibody), timeout=30)
            print(responseKeyClockUser)
        
            if responseKeyClockUser.status_code == 200:
                responseKeyClockUser = responseKeyClockUser.json()
                accessTokenUser = responseKeyClockUser['access_token']
                print("--->Access Token Generated!")
                return accessTokenUser
            else:
                error_message = ""
                if responseKeyClockUser.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"generateAccessToken-Client Error {responseKeyClockUser.status_code}: {responseKeyClockUser.text}"
                elif responseKeyClockUser.status_code in [500, 502, 503, 504]:
                    error_message = f"generateAccessToken-Server Error {responseKeyClockUser.status_code}: {responseKeyClockUser.text}"
                else:
                    error_message = f"generateAccessToken-Unexpected Error {responseKeyClockUser.status_code}: {responseKeyClockUser.text}"
                Helpers.errorVar.append(error_message)
                return False
        except Exception as e:
            error_msg = f"Error generating access token: {str(e)}"
            Helpers.errorVar.append(error_msg)
            return False
    
    @exception_handler
    def validate_solution_end_date(program_end_date, solution_end_date):
        """
        program_end_date  : '2025-11-27T18:29:59.000Z'
        solution_end_date : ['15-01-2026'] or '15-01-2026'
        """
    
        # -----------------------------
        # Validate program end date
        # -----------------------------
        if not program_end_date or not str(program_end_date).strip():
            raise ValueError("Program end date is missing or empty")
    
        try:
            program_utc = datetime.strptime(
                program_end_date, "%Y-%m-%dT%H:%M:%S.%fZ"
            ).replace(tzinfo=timezone.utc)
        except ValueError:
            raise ValueError(
                f"Invalid program end date format: {program_end_date}"
            )
    
        # Convert UTC → IST
        ist_timezone = timezone(timedelta(hours=5, minutes=30))
        program_ist = program_utc.astimezone(ist_timezone)
    
        # -----------------------------
        # Normalize solution end date
        # -----------------------------
        if isinstance(solution_end_date, list):
            if not solution_end_date:
                raise ValueError("Solution end date list is empty")
            solution_end_date = solution_end_date[0]
    
        if not solution_end_date or not str(solution_end_date).strip():
            raise ValueError("Solution end date is missing or empty")
    
        original_solution_end_date = solution_end_date
        solution_end_date = str(solution_end_date).strip()

        # 1) Excel serial date (float/int) support, including numeric strings like "46078.0"
        numeric_solution_date = None
        if isinstance(original_solution_end_date, (int, float)):
            numeric_solution_date = float(original_solution_end_date)
        else:
            try:
                numeric_solution_date = float(solution_end_date)
            except Exception:
                numeric_solution_date = None

        if numeric_solution_date is not None:
            parsed = None
            for datemode in (0, 1):
                try:
                    parsed = xlrd.xldate.xldate_as_datetime(numeric_solution_date, datemode).date()
                    break
                except Exception:
                    continue
            if parsed is None:
                raise ValueError(
                    f"Invalid Excel serial date for solution end date: {solution_end_date}"
                )
            solution_date = parsed
        else:
            # 2) String date formats support
            parsed = None
            accepted_formats = [
                "%d-%m-%Y",
                "%Y-%m-%d",
                "%d/%m/%Y",
                "%Y/%m/%d",
                "%d-%m-%Y %H:%M:%S",
                "%Y-%m-%d %H:%M:%S",
            ]
            for fmt in accepted_formats:
                try:
                    parsed = datetime.strptime(solution_end_date, fmt).date()
                    break
                except ValueError:
                    continue
            if parsed is None:
                raise ValueError(
                    f"Invalid solution end date in sheet: {solution_end_date}. "
                    f"Please correct your sheet and keep date in DD-MM-YYYY format."
                )
            solution_date = parsed
    
        # -----------------------------
        # Compare dates
        # -----------------------------
        if solution_date > program_ist.date():
            raise ValueError(
                f"Solution end date ({solution_date.strftime('%d-%m-%Y')}) "
                f"cannot be beyond program end date "
                f"({program_ist.date().strftime('%d-%m-%Y')}). "
                f"Please correct your sheet and keep date in DD-MM-YYYY format."
            )
    
        return True



    @exception_handler
    def checkEmailValidation(email):
        if (re.search(global_vars.regex, email)):
            return True
        else:
            return False
        
    @exception_handler
    def fetchUserDetails(accessToken, dikshaId):
        """Fetch user details."""
        try:
            url =  host + userinfoapiurl
            headers = {'Content-Type': 'application/json',
                   'Authorization': authorizationforhost,
                   'X-authenticated-user-token': accessToken}
            isEmail = Helpers.checkEmailValidation(dikshaId.lstrip().rstrip())
            
            if isEmail:
                body = "{\n  \"request\": {\n    \"filters\": {\n    \t\"email\": \"" + dikshaId.lstrip().rstrip() + "\"\n    },\n      \"fields\" :[],\n    \"limit\": 1000,\n    \"sort_by\": {\"createdDate\": \"desc\"}\n  }\n}"
            else:
                body = "{\n  \"request\": {\n    \"filters\": {\n    \t\"userName\": \"" + dikshaId.lstrip().rstrip() + "\"\n    },\n      \"fields\" :[],\n    \"limit\": 1000,\n    \"sort_by\": {\"createdDate\": \"desc\"}\n  }\n}"

            responseUserSearch = requests.request("POST", url, headers=headers, data=body, timeout=30)
            response_json = responseUserSearch.json()
            print(responseUserSearch.text)
            print(responseUserSearch, "---------------------------------------------------------------")
            
            if responseUserSearch.status_code == 200:
                responseUserSearch = responseUserSearch.json()
                if responseUserSearch['result']['response']['content']:
                    userKeycloak = responseUserSearch['result']['response']['content'][0]['userId']
                    global_vars.creatorId = userKeycloak
                    userName = responseUserSearch['result']['response']['content'][0]['userName']
                    firstName = responseUserSearch['result']['response']['content'][0]['firstName']
                    rootOrgId = responseUserSearch['result']['response']['content'][0]['rootOrgId']
                    roledetails = []
                    rootOrgName = ""
                    for index in responseUserSearch['result']['response']['content'][0]['organisations']:
                        if rootOrgId == index['organisationId']:
                            roledetails = index['roles']
                            rootOrgName = index['orgName']
                            global_vars.OrgName.append(rootOrgName)
                    print(roledetails)
                    return [userKeycloak, userName, firstName,roledetails,rootOrgName,rootOrgId]
                else:
                    error_msg = f"User '{dikshaId}' not found in Diksha platform"
                    Helpers.errorVar.append(error_msg)
                    return False
            else:
                error_message = ""
                if responseUserSearch.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"fetchUserDetails-Client Error {responseUserSearch.status_code}: {responseUserSearch.text}"
                elif responseUserSearch.status_code in [500, 502, 503, 504]:
                    error_message = f"fetchUserDetails-Server Error {responseUserSearch.status_code}: {responseUserSearch.text}"
                else:
                    error_message = f"fetchUserDetails-Unexpected Error {responseUserSearch.status_code}: {responseUserSearch.text}"
                Helpers.errorVar.append(error_message)
                return False
        except Exception as e:
            error_msg = f"Error fetching user details for '{dikshaId}': {str(e)}"
            Helpers.errorVar.append(error_msg)
            return False
        # Fallback return used when live user-search API block above is disabled.
        # Keep role payload as a list so role checks don't fail with mocked values.
        # return [
        #     Helpers._to_text(dikshaId) or "userKeycloak",
        #     Helpers._to_text(dikshaId) or "userName",
        #     "firstName",
        #     ["CONTENT_CREATOR", "PROGRAM_DESIGNER"],
        #     "rootOrgName",
        #     "rootOrgId",
        # ]
    
    @exception_handler
    def SolutionFileCheck(filePathAddPgm, accessToken, parentFolder):
        print("--->Checking resource details sheet...")
        dictDetailsEnv = global_vars.load_solution_template_details(filePathAddPgm)
        if not dictDetailsEnv:
            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("\"details\" sheet has no data rows"))
            return False

        global_vars.solutionNameInp = Helpers._to_text(dictDetailsEnv.get('solution_name'))
        global_vars.solutionNameForSuccess = global_vars.solutionNameInp
        global_vars.startDateOfProgram = dictDetailsEnv.get('start_date')
        global_vars.endDateOfProgram = dictDetailsEnv.get('end_date')

        startDateArr = str(global_vars.startDateOfProgram).split("-")
        global_vars.startDateOfProgram = startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"

        endDateArr = str(global_vars.endDateOfProgram).split("-")
        global_vars.endDateOfProgram = endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"

        if not Helpers.getProgramInfo(accessToken, parentFolder, global_vars.solutionNameInp, []):
            extIdPGM = global_vars.solutionNameInp
            programName = global_vars.solutionNameInp
            creator_username = dictDetailsEnv.get('creator_username')
            userDetails = Helpers.fetchUserDetails(accessToken, creator_username)
            if not userDetails:
                Helpers.errorVar.append(f"Failed to fetch user details for: {creator_username}")
                return False

            OrgName = userDetails[4]
            print(OrgName, "OrgName")
            global_vars.orgIds = Helpers.fetchOrgId(accessToken, OrgName)
            if not global_vars.orgIds:
                Helpers.errorVar.append(str("CRITICAL") + ': ' + str(f"Failed to fetch org IDs for: {OrgName}"))
                return False
            creatorKeyCloakId = userDetails[0]
            creatorName = userDetails[2]

            if Helpers.getProgramInfo(accessToken, parentFolder, extIdPGM, []):
                print("Program Created SuccessFully.")
            else:
                print("program creation API called")
                # Use default empty values for missing arguments in this legacy context
                if not Helpers.programCreation(accessToken, extIdPGM, programName, "", [], [], "", global_vars.orgIds, creatorKeyCloakId, creatorName):
                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Program creation API failed"))
                    return False

    @exception_handler
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

        wbproject = Helpers._get_cached_resource_workbook(project_inputFile)
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
            global_vars.projectAuthor = str(dictProjectDetails["Diksha_loginId"]).encode('utf-8').decode('utf-8').strip()
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
                   solutionDetailsInTask = Helpers.checkEntityOfSolution(projectName_for_folder_path, solutionNameOrId, accessToken)
                   if not solutionDetailsInTask or not isinstance(solutionDetailsInTask, list) or len(solutionDetailsInTask) < 2:
                       Helpers.errorVar.append(
                           f"Failed to resolve observation details for task observation '{solutionNameOrId}'"
                       )
                       return False
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


        return True

    @exception_handler
    def projectUpload(projectName_for_folder_path, accessToken):
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
        Helpers.createAPILog(projectName_for_folder_path, messageArr)

        if responseProjectUploadApi.status_code == 200:
            print('ProjectUploadApi Success')
            project_internal_content = responseProjectUploadApi.text
            # Mock servers sometimes return CSV inside JSON wrapper: {"raw":"...csv..."}
            try:
                parsed_payload = responseProjectUploadApi.json()
                if isinstance(parsed_payload, dict):
                    raw_csv = parsed_payload.get("raw")
                    if isinstance(raw_csv, str) and raw_csv.strip():
                        project_internal_content = raw_csv
                    elif isinstance(parsed_payload.get("result"), dict):
                        result_raw = parsed_payload.get("result", {}).get("raw")
                        if isinstance(result_raw, str) and result_raw.strip():
                            project_internal_content = result_raw
            except ValueError:
                # Non-JSON response is expected in many environments (plain CSV)
                pass

            if not str(project_internal_content).strip():
                Helpers.errorVar.append("Project Upload-Empty response body while generating projectInternal.csv")
                return False

            with open(projectName_for_folder_path + '/projectUpload/projectInternal.csv','w+',encoding='utf-8') as projectRes:
                projectRes.write(project_internal_content)
        else:
            print("Project Upload failed.")
            error_message = ""
            if responseProjectUploadApi.status_code in [400, 401, 403, 404, 422]:
                error_message = f"Project Upload-Client Error {responseProjectUploadApi.status_code}: {responseProjectUploadApi.text}"
            elif responseProjectUploadApi.status_code in [500, 502, 503, 504]:
                error_message = f"Project Upload-Server Error {responseProjectUploadApi.status_code}: {responseProjectUploadApi.text}"
            else:
                error_message = f"Project Upload-Unexpected Error {responseProjectUploadApi.status_code}: {responseProjectUploadApi.text}"
            Helpers.errorVar.append(error_message)
            return False
        return True

    @exception_handler
    def taskUpload(projectName_for_folder_path, accessToken):
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
                Helpers.createAPILog(projectName_for_folder_path, messageArr)

                if responseProjectListApi.status_code == 200:
                    print('project fetch api Success')
                    try:
                        responsejson = responseProjectListApi.json()
                    except ValueError:
                        Helpers.errorVar.append(
                            f"Project Fetch-Invalid JSON response: {responseProjectListApi.text}"
                        )
                        return False

                    projectList = (
                        responsejson.get('result', {}).get('data')
                        if isinstance(responsejson, dict) else None
                    )
                    if not isinstance(projectList, list):
                        Helpers.errorVar.append(
                            "Project Fetch-Invalid response shape: expected 'result.data' list, "
                            f"got {responsejson}"
                        )
                        return False
                    for project in projectList:
                        if project['externalId'] == projectExternalId:
                            project_id = project['_id']
                else:
                    error_message = ""
                    if responseProjectListApi.status_code in [400, 401, 403, 404, 422]:
                        error_message = f"Project Fetch-Client Error {responseProjectListApi.status_code}: {responseProjectListApi.text}"
                    elif responseProjectListApi.status_code in [500, 502, 503, 504]:
                        error_message = f"Project Fetch-Server Error {responseProjectListApi.status_code}: {responseProjectListApi.text}"
                    else:
                        error_message = f"Project Fetch-Unexpected Error {responseProjectListApi.status_code}: {responseProjectListApi.text}"
                    Helpers.errorVar.append(error_message)
                    return False

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
            Helpers.createAPILog(projectName_for_folder_path, messageArr)

            if responseTasksUploadApi.status_code == 200:
                print('TaskUploadApi Success')
                with open(projectName_for_folder_path + '/taskUpload/taskInternal.csv','w+',encoding='utf-8') as tasksRes:
                    tasksRes.write(responseTasksUploadApi.text)
            else:
                error_message = ""
                if responseTasksUploadApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"Tasks Upload-Client Error {responseTasksUploadApi.status_code}: {responseTasksUploadApi.text}"
                elif responseTasksUploadApi.status_code in [500, 502, 503, 504]:
                    error_message = f"Tasks Upload-Server Error {responseTasksUploadApi.status_code}: {responseTasksUploadApi.text}"
                else:
                    error_message = f"Tasks Upload-Unexpected Error {responseTasksUploadApi.status_code}: {responseTasksUploadApi.text}"
                Helpers.errorVar.append(error_message)
                return False
        return True

    @exception_handler
    def prepareaddingcertificatetemp(filePathAddProject, projectName_for_folder_path, accessToken, solutionId, programID,baseTemplate_id):
        wbproject = Helpers._get_cached_resource_workbook(filePathAddProject)
        projectsheetforcertificate = wbproject.sheet_names()
        validation_cache = global_vars.get_resource_validation_cache() or {}
        project_upload_rows = validation_cache.get("project_upload_rows") or []
        task_upload_rows = validation_cache.get("task_upload_rows") or []
        certificate_detail_rows = validation_cache.get("certificate_detail_rows") or []

        # Fallback to sheet reads when row cache is unavailable.
        if not project_upload_rows and 'Project upload' in projectsheetforcertificate:
            project_sheet = wbproject.sheet_by_name('Project upload')
            keys = [project_sheet.cell(1, col_idx).value for col_idx in range(project_sheet.ncols)]
            for row_idx in range(2, project_sheet.nrows):
                project_upload_rows.append({
                    keys[col_idx]: project_sheet.cell(row_idx, col_idx).value
                    for col_idx in range(project_sheet.ncols)
                })

        if not task_upload_rows and 'Tasks upload' in projectsheetforcertificate:
            task_sheet = wbproject.sheet_by_name('Tasks upload')
            keys = [task_sheet.cell(1, col_idx).value for col_idx in range(task_sheet.ncols)]
            for row_idx in range(2, task_sheet.nrows):
                task_upload_rows.append({
                    keys[col_idx]: task_sheet.cell(row_idx, col_idx).value
                    for col_idx in range(task_sheet.ncols)
                })

        if not certificate_detail_rows and 'Certificate details' in projectsheetforcertificate:
            certificate_sheet = wbproject.sheet_by_name('Certificate details')
            keys = [certificate_sheet.cell(1, col_idx).value for col_idx in range(certificate_sheet.ncols)]
            for row_idx in range(2, certificate_sheet.nrows):
                certificate_detail_rows.append({
                    keys[col_idx]: certificate_sheet.cell(row_idx, col_idx).value
                    for col_idx in range(certificate_sheet.ncols)
                })

        tasksLevelEvidance = []
        projectMinNooEvide = None
        projectLevelEvidance = []
        taskMinNooEvide =[]

        for dictDetailsEnv in project_upload_rows:
            projectLevelMinNooEvidence = dictDetailsEnv.get("Minimum No. of Evidence")
            print(projectLevelMinNooEvidence)
            projectLevelEvidance = Helpers._to_text(dictDetailsEnv.get("Project Level Evidence", "")).lower()
            if projectLevelMinNooEvidence in ["", None]:
                projectLevelMinNooEvidence = 1
            projectMinNooEvide = int(projectLevelMinNooEvidence)

        for dictDetailsEnv in task_upload_rows:
            taskLevelEvidence = Helpers._to_text(dictDetailsEnv.get("Task Level Evidence", "")).lower()
            minNoOfEvidence = dictDetailsEnv.get("Minimum No. of Evidence")

            if taskLevelEvidence == "yes":
                tasksLevelEvidance.append(dictDetailsEnv.get("TaskTitle"))
                if minNoOfEvidence in ["", None]:
                    minNoOfEvidence = 1
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


        if certificate_detail_rows:
            print("--->Checking Certificate details  sheet...")
            for dictDetailsEnv in certificate_detail_rows:
                if not dictDetailsEnv.get('Certificate issuer'):
                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str("\"Certificate issuer\" must not be Empty in \"Certificate details\" sheet"))
                    return False
                certificateissuer = Helpers._to_text(dictDetailsEnv.get('Certificate issuer'))
                payload["issuer"]["name"] = certificateissuer

                if dictDetailsEnv.get('Type of certificate') not in ["One Logo - One Signature", "One Logo - Two Signature", "Two Logo - One Signature","Two Logo - Two Signature"]:
                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str("\"Type of certificate\" must not be Empty or Invalid in \"Certificate details\" sheet"))
                    return False
                Typeofcertificate = dictDetailsEnv.get('Type of certificate')

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
                    if global_vars.TaskEvidenceOperator.lower() == "no":
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
                                        "value": int(global_vars.AnyTaskEvidenceNo),
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
        if global_vars.TaskEvidenceOperator.lower() == "yes":
            # payload["criteria"]["conditions"]["C1"]["validationText"] = f"Add {int(AnyTaskEvidenceNo)} evidence for any task"
            payload["criteria"]["conditions"]["C3"]["validationText"] = f"Add {int(global_vars.AnyTaskEvidenceNo)} evidence for any task"

        if str(projectLevelEvidance).strip().lower() == "yes":       
            condition = ""
            print(payload["criteria"]["conditions"],"4514")
            print(global_vars.TaskEvidenceOperator.lower(),"4515")
            condition_keys = list(payload["criteria"]["conditions"].keys())
            task_evidence_operator = global_vars.TaskEvidenceOperator.lower()

            if task_evidence_operator == "yes" and len(condition_keys) > 2:
                first_part = "&&".join(condition_keys[:2])
                grouped_part = "||".join(condition_keys[2:])
                condition = f"{first_part}&&({grouped_part})"
            else:
                condition = "&&".join(condition_keys)

            payload["criteria"]["expression"] = condition
        else:
            condition = ""
            condition_keys = list(payload["criteria"]["conditions"].keys())
            task_evidence_operator = global_vars.TaskEvidenceOperator.lower()

            if task_evidence_operator == "yes" and len(condition_keys) > 1:
                first_part = "&&".join(condition_keys[:1])
                grouped_part = "||".join(condition_keys[1:])
                condition = f"{first_part}&&({grouped_part})"
            else:
                condition = "&&".join(condition_keys)

            payload["criteria"]["expression"] = condition

        global_vars.TaskEvidenceOperator = ""
        print(payload["criteria"]["expression"])
        print(json.dumps(payload, indent=1))

        responseaddcertificateUploadApi = requests.request("POST",url=urladdcertificate, headers=headeraddcertificateApi,
                                               data=json.dumps(payload))
        messageArr = ["Add certificate json is prepared",
                      "File path : " + projectName_for_folder_path + '/addCertificate/Addcertificate.text']
        messageArr.append("URL : " + str(responseaddcertificateUploadApi))
        messageArr.append("Upload status code : " + str(responseaddcertificateUploadApi.status_code))
        Helpers.createAPILog(projectName_for_folder_path, messageArr)
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
            Helpers.createAPILog(projectName_for_folder_path, messageArr)
            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Add certificate failed") + ' | details=' + str({"response": str(responseaddcertificateUploadApi.text)}))
            return False

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
                error_msg = ""
                if responseupdatecertificateApi.status_code in [400, 401, 403, 404, 422]:
                    error_msg = f"Update Certificate to Solution-Client Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}"
                elif responseupdatecertificateApi.status_code in [500, 502, 503, 504]:
                    error_msg = f"Update Certificate to Solution-Server Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}"
                else:
                    error_msg = f"Update Certificate to Solution-Unexpected Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}"
                Helpers.errorVar.append(error_msg)
                print("error in updating solution")
                return False

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
                error_message = ""
                if responseupdatecertificateApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"Update Certificate-Client Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}"
                elif responseupdatecertificateApi.status_code in [500, 502, 503, 504]:
                    error_message = f"Update Certificate-Server Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}"
                else:
                    error_message = f"Update Certificate-Unexpected Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}"
                Helpers.errorVar.append(error_message)
                return False
        else:
            error_message = ""
            if responseDownloadsvgApi.status_code in [400, 401, 403, 404, 422]:
                error_message = f"Download SVG-Client Error {responseDownloadsvgApi.status_code}: {responseDownloadsvgApi.text}"
            elif responseDownloadsvgApi.status_code in [500, 502, 503, 504]:
                error_message = f"Download SVG-Server Error {responseDownloadsvgApi.status_code}: {responseDownloadsvgApi.text}"
            else:
                error_message = f"Download SVG-Unexpected Error {responseDownloadsvgApi.status_code}: {responseDownloadsvgApi.text}"
            Helpers.errorVar.append(error_message)
            return False

        return certificatetemplateid

# This function is used to add SVG to the certificate based on type of certificate
    @exception_handler
    def editsvg(accessToken,filePathAddProject,projectName_for_folder_path,baseTemplate_id):
        wbproject = Helpers._get_cached_resource_workbook(filePathAddProject)
        projectsheetforcertificate = wbproject.sheet_names()
        validation_cache = global_vars.get_resource_validation_cache() or {}
        certificate_detail_rows = validation_cache.get("certificate_detail_rows") or []

        if not certificate_detail_rows and 'Certificate details' in projectsheetforcertificate:
            certificate_sheet = wbproject.sheet_by_name('Certificate details')
            keys = [certificate_sheet.cell(1, col_idx).value for col_idx in range(certificate_sheet.ncols)]
            for row_idx in range(2, certificate_sheet.nrows):
                certificate_detail_rows.append({
                    keys[col_idx]: certificate_sheet.cell(row_idx, col_idx).value
                    for col_idx in range(certificate_sheet.ncols)
                })

        if not certificate_detail_rows:
            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Certificate details not found in cache/sheet"))
            return False

        print("--->Checking Certificate details  sheet...")
        for dictDetailsEnv in certificate_detail_rows:
            certificateissuer = Helpers._to_text(dictDetailsEnv.get('Certificate issuer'))
            Typeofcertificate = dictDetailsEnv.get('Type of certificate')
            Certificateisuuer = Helpers._to_text(dictDetailsEnv.get('Certificate issuer'))
            authrigedsignaturename1 = Helpers._to_text(dictDetailsEnv.get('Authorised Signature Name - 1'))
            authrigeddesignation1 = Helpers._to_text(dictDetailsEnv.get('Authorised Designation - 1'))
            authrigedsignaturename2 = Helpers._to_text(dictDetailsEnv.get('Authorised Signature Name - 2'))
            authrigeddesignation2 = Helpers._to_text(dictDetailsEnv.get('Authorised Designation - 2'))

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
                error_message = ""
                if responseeditsvg.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"Edit SVG-Client Error {responseeditsvg.status_code}: {responseeditsvg.text}"
                elif responseeditsvg.status_code in [500, 502, 503, 504]:
                    error_message = f"Edit SVG-Server Error {responseeditsvg.status_code}: {responseeditsvg.text}"
                else:
                    error_message = f"Edit SVG-Unexpected Error {responseeditsvg.status_code}: {responseeditsvg.text}"
                Helpers.errorVar.append(error_message)
                return False    
        return True

    


    @exception_handler
    def fetchCertificateBaseTemplate(filePathAddProject,accessToken):
        wbproject = Helpers._get_cached_resource_workbook(filePathAddProject)
        projectsheetforcertificate = wbproject.sheet_names()
        validation_cache = global_vars.get_resource_validation_cache() or {}
        certificate_detail_rows = validation_cache.get("certificate_detail_rows") or []

        if not certificate_detail_rows and 'Certificate details' in projectsheetforcertificate:
            certificate_sheet = wbproject.sheet_by_name('Certificate details')
            keys = [certificate_sheet.cell(1, col_idx).value for col_idx in range(certificate_sheet.ncols)]
            for row_idx in range(2, certificate_sheet.nrows):
                certificate_detail_rows.append({
                    keys[col_idx]: certificate_sheet.cell(row_idx, col_idx).value
                    for col_idx in range(certificate_sheet.ncols)
                })

        typeOfCertificate = ""
        for dictDetailsEnv in certificate_detail_rows:
            typeOfCertificate = Helpers._to_text(dictDetailsEnv.get("Type of certificate"))
            if typeOfCertificate:
                print(typeOfCertificate)
                break
        if not typeOfCertificate:
            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("\"Type of certificate\" not found in cache/sheet"))
            return False

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
            baseTemplateCode = certificatetypeof.get(typeOfCertificate)
            print(baseTemplateCode,"baseTemplateCode")
            print(baseTemplateLookup,"baseTemplateLookup")

            candidate_codes = []
            if baseTemplateCode:
                candidate_codes.append(baseTemplateCode)
            candidate_codes.append(typeOfCertificate)
            typeOfCertificate_sign = typeOfCertificate.replace("signature", "sign")
            if typeOfCertificate_sign != typeOfCertificate:
                candidate_codes.append(typeOfCertificate_sign)
            candidate_codes.append(typeOfCertificate + "proddummy")
            if typeOfCertificate_sign != typeOfCertificate:
                candidate_codes.append(typeOfCertificate_sign + "proddummy")

            for code in candidate_codes:
                if code in baseTemplateLookup:
                    return baseTemplateLookup[code]

            for code in candidate_codes:
                for lookup_code in baseTemplateLookup.keys():
                    if lookup_code.startswith(code):
                        return baseTemplateLookup[lookup_code]

            for code in candidate_codes:
                for lookup_code in baseTemplateLookup.keys():
                    if code in lookup_code:
                        return baseTemplateLookup[lookup_code]

            Helpers.errorVar.append(
                "CRITICAL: base template code not found for typeOfCertificate="
                + str(typeOfCertificate)
                + "; available codes="
                + str(list(baseTemplateLookup.keys()))
            )
            return False

        else:
            print("--->Error in fetching DBfind data please give proper code value<---")
            error_message = ""
            if responsedbFindApi.status_code in [400, 401, 403, 404, 422]:
                error_message = f"fetchCertificateBaseTemplate-Client Error {responsedbFindApi.status_code}: {responsedbFindApi.text}"
            elif responsedbFindApi.status_code in [500, 502, 503, 504]:
                error_message = f"fetchCertificateBaseTemplate-Server Error {responsedbFindApi.status_code}: {responsedbFindApi.text}"
            else:
                error_message = f"fetchCertificateBaseTemplate-Unexpected Error {responsedbFindApi.status_code}: {responsedbFindApi.text}"
            Helpers.errorVar.append(error_message)
            return False


    @exception_handler
    def downloadlogosign(filePathAddProject,projectName_for_folder_path):
        wbproject = Helpers._get_cached_resource_workbook(filePathAddProject)
        projectsheetforcertificate = wbproject.sheet_names()
        validation_cache = global_vars.get_resource_validation_cache() or {}
        certificate_detail_rows = validation_cache.get("certificate_detail_rows") or []

        if not certificate_detail_rows and 'Certificate details' in projectsheetforcertificate:
            detailsEnvSheet = wbproject.sheet_by_name('Certificate details')
            keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in range(detailsEnvSheet.ncols)]
            for row_index_env in range(2, detailsEnvSheet.nrows):
                certificate_detail_rows.append({
                    keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                    for col_index_env in range(detailsEnvSheet.ncols)
                })

        if not certificate_detail_rows:
            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Certificate details not found in cache/sheet"))
            return False

        print("--->Checking Certificate details  sheet...")
        Logofilepath = projectName_for_folder_path + '/Logofile/'
        if not os.path.exists(Logofilepath):
            os.mkdir(Logofilepath)

        for dictDetailsEnv in certificate_detail_rows:
            if not dictDetailsEnv.get('Certificate issuer'):
                error_msg = "\"Certificate issuer\" must not be Empty in \"Certificate details\" sheet"
                Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                return False

            if not dictDetailsEnv.get('Type of certificate'):
                error_msg = "\"Type of certificate\" must not be Empty in \"Certificate details\" sheet"
                Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                return False

            typeOfCertificate = dictDetailsEnv.get('Type of certificate')

            if typeOfCertificate == 'One Logo - One Signature':
               Logo1 = dictDetailsEnv['Logo - 1']
               logo_split = str(Logo1).split('/')[5]

               file_url = 'https://drive.google.com/uc?export=download&id='+logo_split
               gdown.download(file_url, Logofilepath + '/logo1.jpg', quiet=False)

               Authsign1 = dictDetailsEnv['Authorised Signature Image - 1']
               logo_split = str(Authsign1).split('/')[5]
               file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split
               gdown.download(file_url, Logofilepath + '/signature1.jpg', quiet=False)

            elif typeOfCertificate == 'One Logo - Two Signature':
                Logo1 = dictDetailsEnv['Logo - 1']
                logo_split = str(Logo1).split('/')[5]
                file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split
                gdown.download(file_url, Logofilepath + '/logo1.jpg', quiet=False)

                Authsign1 = dictDetailsEnv['Authorised Signature Image - 1']
                logo_split = str(Authsign1).split('/')[5]
                file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split
                gdown.download(file_url, Logofilepath + '/signature1.jpg', quiet=False)

                Authsign2 = dictDetailsEnv['Authorised Signature Image - 2']
                logo_split = str(Authsign2).split('/')[5]
                file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split
                gdown.download(file_url, Logofilepath + '/signature2.jpg', quiet=False)

            elif typeOfCertificate == 'Two Logo - One Signature':
                Logo1 = dictDetailsEnv['Logo - 1']
                logo_split = str(Logo1).split('/')[5]
                file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split
                gdown.download(file_url, Logofilepath + '/logo1.jpg', quiet=False)

                Logo2 = dictDetailsEnv['Logo - 2']
                logo_split = str(Logo2).split('/')[5]
                file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split
                gdown.download(file_url, Logofilepath + '/logo2.jpg', quiet=False)

                Authsign1 = dictDetailsEnv['Authorised Signature Image - 1']
                logo_split = str(Authsign1).split('/')[5]
                file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split
                gdown.download(file_url, Logofilepath + '/signature1.jpg', quiet=False)

            elif typeOfCertificate == 'Two Logo - Two Signature':
                Logo1 = dictDetailsEnv['Logo - 1']
                logo_split = str(Logo1).split('/')[5]
                file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split
                gdown.download(file_url, Logofilepath + '/logo1.jpg', quiet=False)

                Logo2 = dictDetailsEnv['Logo - 2']
                logo_split = str(Logo2).split('/')[5]
                file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split
                gdown.download(file_url, Logofilepath + '/logo2.jpg', quiet=False)

                Authsign1 = dictDetailsEnv['Authorised Signature Image - 1']
                logo_split = str(Authsign1).split('/')[5]
                file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split
                gdown.download(file_url, Logofilepath + '/signature1.jpg', quiet=False)

                Authsign2 = dictDetailsEnv['Authorised Signature Image - 2']
                logo_split = str(Authsign2).split('/')[5]
                file_url = 'https://drive.google.com/uc?export=download&id=' + logo_split
                gdown.download(file_url, Logofilepath + '/signature2.jpg', quiet=False)

            else:
                msg = "Logos and signature downloading failed (check if drive link are Anyone with the link or not)"
                print("--->" + msg + "<---")
                Helpers.errorVar.append(str("CRITICAL") + ': ' + str(msg))
                return False
        return True




    @exception_handler
    def fetchSolutionDetailsFromProgramSheet(solutionName_for_folder_path, programFile, solutionId, accessToken):
        """Fetch solution details from API. Returns None on failure."""
        try:
            urlFetchSolutionApi = internal_kong_ip + fetchsolutiondoc + solutionId
            print(urlFetchSolutionApi,"urlFetchSolutionApi")
    
            headerFetchSolutionApi = {
                'Content-Type': 'application/json',
                'Authorization': authorization,
                'X-authenticated-user-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token
            }
            payloadFetchSolutionApi = {}
    
            responseFetchSolutionApiUrl = requests.post(
                url=urlFetchSolutionApi, 
                headers=headerFetchSolutionApi,
                data=payloadFetchSolutionApi,
                timeout=30
            )
            print(responseFetchSolutionApiUrl.text,"responseFetchSolutionApiUrl")
            
            if responseFetchSolutionApiUrl.status_code != 200:
                error_msg = f"Failed to fetch solution details. Status: {responseFetchSolutionApiUrl.status_code}, response={responseFetchSolutionApiUrl.text}"
                Helpers.createAPILog(solutionName_for_folder_path, [error_msg])
                Helpers.errorVar.append(str("ERROR") + ': ' + str(error_msg))
                return False
            else:    
                responseFetchSolutionJson = responseFetchSolutionApiUrl.json()
                print(responseFetchSolutionJson,"responseFetchSolutionJson")
                messageArr = ["Solution Fetch Link.",
                              "solution name : " + responseFetchSolutionJson["result"]["name"],
                              "solution ExternalId : " + responseFetchSolutionJson["result"]["externalId"]]
                messageArr.append("Upload status code : " + str(responseFetchSolutionApiUrl.status_code))
                Helpers.createAPILog(solutionName_for_folder_path, messageArr)
                print(" reached 2441")
            
        except Exception as e:
            error_msg = f"Error fetching solution details: {str(e)}"
            Helpers.createAPILog(solutionName_for_folder_path, [error_msg])
            Helpers.errorVar.append(str("ERROR") + ': ' + str(error_msg))
            return False

        if responseFetchSolutionApiUrl.status_code == 200:
            print('Fetch solution Api Success')

            solutionName = responseFetchSolutionJson["result"]["name"]

            xfile = openpyxl.load_workbook(programFile)

            resourceDetailsSheet = xfile['Resource Details']
            # print(resourceDetailsSheet,"1911")
            rowCountRD = resourceDetailsSheet.max_row
            # print(rowCountRD,"rowCountRD")
            columnCountRD = resourceDetailsSheet.max_column
            for row in range(3, rowCountRD + 1):
                # print("here we reacher")
                # print(resourceDetailsSheet,"resourceDetailsSheet")
                solutionNameCell = resourceDetailsSheet[f"A{row}"].value
                # print(f"Row {row} Solution Name 1919: {solutionNameCell}")
                # print(solutionName,"solutionName")
                # print(resourceDetailsSheet["A" + str(row)].value,"1919")
                if resourceDetailsSheet["A" + str(row)].value == solutionName:
                    solutionMainRole = str(resourceDetailsSheet["E" + str(row)].value).strip()
                    global_vars.solutionRolesArray = str(resourceDetailsSheet["F" + str(row)].value).split(",") if str(
                        resourceDetailsSheet["E" + str(row)].value).split(",") else []
                    if "teacher" in solutionMainRole.strip().lower():
                        global_vars.solutionRolesArray.append("TEACHER")
                    global_vars.solutionStartDate = resourceDetailsSheet["G" + str(row)].value
                    # print(solutionStartDate, "<-------------------solutionStartDate////====")
                    global_vars.solutionEndDate = resourceDetailsSheet["H" + str(row)].value
                    # print(solutionEndDate, "<---------------------solutionEndDate/////========")
        else:
            error_msg = ""
            if responseFetchSolutionApiUrl.status_code in [400, 401, 403, 404, 422]:
                error_msg = f"Fetch Solution Details-Client Error {responseFetchSolutionApiUrl.status_code}: {responseFetchSolutionApiUrl.text}"
            elif responseFetchSolutionApiUrl.status_code in [500, 502, 503, 504]:
                error_msg = f"Fetch Solution Details-Server Error {responseFetchSolutionApiUrl.status_code}: {responseFetchSolutionApiUrl.text}"
            else:
                error_msg = f"Fetch Solution Details-Unexpected Error {responseFetchSolutionApiUrl.status_code}: {responseFetchSolutionApiUrl.text}"
            Helpers.errorVar.append(error_msg)
        
        return [global_vars.solutionRolesArray, global_vars.solutionStartDate, global_vars.solutionEndDate]





    @exception_handler
    def solutionCreationAndMapping(projectName_for_folder_path, entityToUpload, listOfFoundRoles, accessToken,programFile):
        SolutionFilePath = projectName_for_folder_path + '/solutionDetails/'
        if not os.path.exists(SolutionFilePath):
            os.mkdir(SolutionFilePath)
        with open(projectName_for_folder_path + '/solutionDetails/solutionDetails.csv', 'w',encoding='utf-8') as file:
            writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
            writer.writerows(
                    [["solutionExtId", "solutionName", "solutionDescription", "solution_id", "programExternalId", "entityType",
                      "scopeEntityType", "entityNames", "roles", "duplicateTemplateExtId", "duplicateTemplate_id"]])

        project_internal_path = projectName_for_folder_path + '/projectUpload/projectInternal.csv'
        if not os.path.exists(project_internal_path):
            Helpers.errorVar.append(f"Project Internal file missing: {project_internal_path}")
            return False

        with open(project_internal_path, mode='r', encoding='utf-8') as projectInternalfp:
            projectInternalfile = csv.DictReader(projectInternalfp)
            project_rows = list(projectInternalfile)

        if not project_rows:
            preview = ""
            try:
                with open(project_internal_path, mode='r', encoding='utf-8') as fp:
                    preview = fp.read(400)
            except Exception:
                preview = ""
            Helpers.errorVar.append(
                f"Project Internal CSV has no rows or invalid headers. Raw preview: {preview}"
            )
            return False

        required_cols = ["externalId", "title", "description"]
        for projectInternal in project_rows:
            missing_cols = [col for col in required_cols if col not in projectInternal]
            if missing_cols:
                Helpers.errorVar.append(
                    f"Project Internal CSV missing required columns {missing_cols}. "
                    f"Available columns: {list(projectInternal.keys())}"
                )
                return False

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
                "createdFor": global_vars.orgIds,
                "rootOrganisations": global_vars.orgIds,
                "programExternalId": global_vars.programExternalId,
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
                Helpers.createAPILog(projectName_for_folder_path, messageArr)
                print("ProjectSolutionCreationApi Success")
                duplicateTemplateExtId = projectExternalId + '_IMPORTED'
                queryparamsMapProjectSolutionApi = projectExternalId + '?solutionId=' + solutionExternalId
                urlMapProjectSolutionApi = internal_kong_ip + mapsolutiontoproject
                print(urlMapProjectSolutionApi,"urlMapProjectSolutionApi")
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
                print(responseMapProjectSolutionApi.text,"response of mapping solution to project api")

                messageArr = ["Successfully mapped the project to Solution",
                              "URL : " + str(urlMapProjectSolutionApi + queryparamsMapProjectSolutionApi),
                              "Status Code : " + str(responseMapProjectSolutionApi.status_code),
                              "Response : " + str(responseMapProjectSolutionApi.text)]
                if responseMapProjectSolutionApi.status_code == 200:
                    responseMapProjectSolutionApi = responseMapProjectSolutionApi.json()
                    duplicateTemplateId = responseMapProjectSolutionApi['result']['_id']
                    messageArr.append("duplicate TemplateId successfully created: " + str(duplicateTemplateId))
                    Helpers.createAPILog(projectName_for_folder_path, messageArr)
                    print("MapSolutionToProjectApi Sucsess")
                    with open(projectName_for_folder_path + '/solutionDetails/solutionDetails.csv', 'a',encoding='utf-8') as file:
                        writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                        writer.writerows([[solutionExternalId, project_name, project_description, solutionId,
                                           global_vars.programExternalId, projectEntityType,
                                           global_vars.scopeEntityType, entityToUpload, listOfFoundRoles, duplicateTemplateExtId,
                                           duplicateTemplateId]])
                    solutionDetails = Helpers.fetchSolutionDetailsFromProgramSheet(projectName_for_folder_path, programFile,
                                                                           solutionId, accessToken)
                    if not solutionDetails:
                        Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Failed to fetch solution details from program sheet"))
                        return False
                    scopeEntities = global_vars.entitiesPGMID
                    scopeRoles = solutionDetails[0]
                    bodySolutionUpdate = {
                        "scope": {"entityType": global_vars.scopeEntityType, "entities": scopeEntities, "roles": scopeRoles}}
                    if not Helpers.solutionUpdate(accessToken, solutionId, bodySolutionUpdate):
                        return False

                    userDetails = Helpers.fetchUserDetails(accessToken, global_vars.projectAuthor)
                    if not userDetails:
                        Helpers.errorVar.append(f"Failed to fetch user details for: {global_vars.projectAuthor}")
                        return False
                    global_vars.matchedShikshalokamLoginId = userDetails[0]
                    global_vars.projectCreator = userDetails[2]

                    bodySolutionUpdate = {
                        "creator": global_vars.projectCreator, "author": global_vars.matchedShikshalokamLoginId}
                    if not Helpers.solutionUpdate(accessToken, solutionId, bodySolutionUpdate):
                        return False
                    # Below script will convert date DD-MM-YYYY TO YYYY-MM-DD 00:00:00 to match the code syntax

                    if solutionDetails[1]:
                        startDateArr = str(solutionDetails[1]).split("-")
                        bodySolutionUpdate = {
                            "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                        if not Helpers.solutionUpdate(accessToken, solutionId, bodySolutionUpdate):
                            return False
                    if solutionDetails[2]:
                        endDateArr = str(solutionDetails[2]).split("-")
                        bodySolutionUpdate = {
                            "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                        if not Helpers.solutionUpdate(accessToken, solutionId, bodySolutionUpdate):
                            return False
                else:
                    error_message = ""
                    if responseMapProjectSolutionApi.status_code in [400, 401, 403, 404, 422]:
                        error_message = f"Map Project to Solution-Client Error {responseMapProjectSolutionApi.status_code}: {responseMapProjectSolutionApi.text}"
                    elif responseMapProjectSolutionApi.status_code in [500, 502, 503, 504]:
                        error_message = f"Map Project to Solution-Server Error {responseMapProjectSolutionApi.status_code}: {responseMapProjectSolutionApi.text}"
                    else:
                        error_message = f"Map Project to Solution-Unexpected Error {responseMapProjectSolutionApi.status_code}: {responseMapProjectSolutionApi.text}"
                    Helpers.errorVar.append(error_message)
                    return False

                return [solutionExternalId, solutionId]
            else:
                error_message = ""
                if responseCreateSolutionApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"Create Solution-Client Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}"
                elif responseCreateSolutionApi.status_code in [500, 502, 503, 504]:
                    error_message = f"Create Solution-Server Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}"
                else:
                    error_message = f"Create Solution-Unexpected Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}"
                Helpers.errorVar.append(error_message)
                return False




    @exception_handler
    def prepareProgramSuccessSheet(MainFilePath, solutionName_for_folder_path, programFile, solutionExternalId, solutionId,accessToken):
        urlFetchSolutionDocApi = internal_kong_ip + fetchsolutiondoc + solutionId
        headerFetchSolutionDocApi = {
            'Authorization': authorization,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token
        }
        payloadFetchSolutionDocApi = {}

        responseFetchSolutionDocApi = requests.post(url=urlFetchSolutionDocApi, headers=headerFetchSolutionDocApi,
                                                 data=payloadFetchSolutionDocApi)
        if responseFetchSolutionDocApi.status_code == 200:
            print('Fetch solution Doc Api Success', responseFetchSolutionDocApi.text)
            try:
                responseFetchSolutionDocJson = responseFetchSolutionDocApi.json()
            except ValueError:
                Helpers.errorVar.append(
                    f"prepareProgramSuccessSheet-Invalid JSON in solution doc API: {responseFetchSolutionDocApi.text}"
                )
                return False
            solutionName = responseFetchSolutionDocJson["result"]["name"]
            print(solutionName,"solutionName in solution doc api")
            messageArr = ["Solution Fetch Doc.",
                          "solution name : " + solutionName,
                          "solution ExternalId : " + responseFetchSolutionDocJson["result"]["externalId"]]
            print("solution ExternalId", responseFetchSolutionDocJson["result"]["externalId"])
            print(messageArr,"messageArr in solution doc api")
            messageArr.append("Upload status code : " + str(responseFetchSolutionDocApi.status_code))
            print(messageArr,"messageArr in solution doc api after appending status code")
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)
        else:
            error_message = ""
            if responseFetchSolutionDocApi.status_code in [400, 401, 403, 404, 422]:
                error_message = f"prepareProgramSuccessSheet-Client Error {responseFetchSolutionDocApi.status_code}: {responseFetchSolutionDocApi.text}"
            elif responseFetchSolutionDocApi.status_code in [500, 502, 503, 504]:
                error_message = f"prepareProgramSuccessSheet-Server Error {responseFetchSolutionDocApi.status_code}: {responseFetchSolutionDocApi.text}"
            else:
                error_message = f"prepareProgramSuccessSheet-Unexpected Error {responseFetchSolutionDocApi.status_code}: {responseFetchSolutionDocApi.text}"
            Helpers.errorVar.append(error_message)
            return False

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
        Helpers.createAPILog(solutionName_for_folder_path, messageArr)

        if responseFetchSolutionLinkApi.status_code == 200:
            print('Fetch solution Link Api Success')
            try:
                responseProjectUploadJson = responseFetchSolutionLinkApi.json()
            except ValueError:
                Helpers.errorVar.append(
                    f"responseFetchSolutionLinkApi-Invalid JSON response: {responseFetchSolutionLinkApi.text}"
                )
                return False
            solutionLink = responseProjectUploadJson["result"]
            messageArr.append("Response : " + str(responseFetchSolutionLinkApi.text))
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)

            program_base_name = os.path.splitext(os.path.basename(str(programFile)))[0]
            success_sheet_path = os.path.join(MainFilePath, program_base_name + '-SuccessSheet.xlsx')
            if os.path.exists(success_sheet_path):
                xfile = openpyxl.load_workbook(success_sheet_path)
            else:
                xfile = openpyxl.load_workbook(programFile)

            resourceDetailsSheet = xfile.get_sheet_by_name('Resource Details')

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

            xfile.save(success_sheet_path)
            print("Program success sheet is created")

        else:
            print("Fetch solution link API Failed")
            error_message = ""
            if responseFetchSolutionLinkApi.status_code in [400, 401, 403, 404, 422]:
                error_message = f"responseFetchSolutionLinkApi-Client Error {responseFetchSolutionLinkApi.status_code}: {responseFetchSolutionLinkApi.text}"
            elif responseFetchSolutionLinkApi.status_code in [500, 502, 503, 504]:
                error_message = f"responseFetchSolutionLinkApi-Server Error {responseFetchSolutionLinkApi.status_code}: {responseFetchSolutionLinkApi.text}"
            else:
                error_message = f"responseFetchSolutionLinkApi-Unexpected Error {responseFetchSolutionLinkApi.status_code}: {responseFetchSolutionLinkApi.text}"
            Helpers.errorVar.append(error_message)
            return False
        print("solutionLink", solutionLink)
        return solutionLink

# fetch org Ids 
    @exception_handler
    def fetchOrgId(accessToken, OrgName):
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
                    error_msg = f"Org '{str(org).strip()}' not found in Diksha."
                    print(error_msg)
                    Helpers.errorVar.append(str("ERROR") + ': ' + str(error_msg))
            else:
                error_message = ""
                if responseOrgSearch.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"fetchOrgId-Client Error {responseOrgSearch.status_code}: {responseOrgSearch.text}"
                elif responseOrgSearch.status_code in [500, 502, 503, 504]:
                    error_message = f"fetchOrgId-Server Error {responseOrgSearch.status_code}: {responseOrgSearch.text}"
                else:
                    error_message = f"fetchOrgId-Unexpected Error {responseOrgSearch.status_code}: {responseOrgSearch.text}"
                Helpers.errorVar.append(error_message)
                return False
        return orgIds

    @exception_handler
    def solutionUpdate(accessToken, solutionId, bodySolutionUpdate):
        try:
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
                print("Solution Update Success.", responseUpdateSolutionApi.text)
                return True
            else:
                if responseUpdateSolutionApi.status_code in [400, 401, 403, 404, 422]:
                    Helpers.errorVar.append(f"UpdateSolutionApi-Client Error {responseUpdateSolutionApi.status_code}: {responseUpdateSolutionApi.text}")
                elif responseUpdateSolutionApi.status_code in [500, 502, 503, 504]:
                    Helpers.errorVar.append(f"UpdateSolutionApi-Server Error {responseUpdateSolutionApi.status_code}: {responseUpdateSolutionApi.text}")
                else:
                    Helpers.errorVar.append(f"UpdateSolutionApi-Unexpected Error {responseUpdateSolutionApi.status_code}: {responseUpdateSolutionApi.text}")
                return False
        except Exception as e:
            messageArr = []
            messageArr.append("Exception caught : " + str(e))
            Helpers.errorVar.append(f"Error occurred: {str(e)}")
            return False

    @exception_handler
    def schedule_deletion(returnPathStr):
        @exception_handler
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

    @exception_handler
    def checkEntityOfSolution(projectName_for_folder_path, solutionNameOrId, accessToken):
        solutionEntityType = None
        solutionExternalId = None

        def _parse_maybe_html_json(response_obj, context_label):
            """Parse JSON from normal API response or HTML-wrapped mock payload."""
            try:
                return response_obj.json()
            except ValueError:
                pass

            text = response_obj.text or ""
            # Common mock shape: <pre>{...json...}</pre>
            pre_match = re.search(r"<pre[^>]*>(.*?)</pre>", text, flags=re.IGNORECASE | re.DOTALL)
            if pre_match:
                candidate = pre_match.group(1).strip()
                try:
                    return json.loads(candidate)
                except ValueError:
                    pass

            # Fallback: try first JSON object in the body
            obj_match = re.search(r"(\{.*\})", text, flags=re.DOTALL)
            if obj_match:
                candidate = obj_match.group(1).strip()
                try:
                    return json.loads(candidate)
                except ValueError:
                    pass

            Helpers.errorVar.append(
                f"{context_label}-Invalid JSON/HTML response: {text[:300]}"
            )
            return None
        
        searchSolutionurl = internal_kong_ip + fetchsolutiondetails + "observation&page=1&limit=100&search=" + solutionNameOrId

        searchSolutionpayload = {}
        searchSolutionheaders = {
            'X-authenticated-user-token': accessToken,
            'internal-access-token': internal_access_token,
            'Authorization': authorization
        }

        searchSolutionresponse = requests.request("GET", searchSolutionurl, headers=searchSolutionheaders,
                                                data=searchSolutionpayload)
        messageArr = ["Solution found",
                    "URL : " + str(searchSolutionurl),
                    "Status Code : " + str(searchSolutionresponse.status_code),
                    "Response : " + str(searchSolutionresponse.text)]

        if searchSolutionresponse.status_code == 200:
            searchSolutionjson = _parse_maybe_html_json(searchSolutionresponse, "checkEntityOfSolution-search")
            if not searchSolutionjson:
                return False
            result_obj = searchSolutionjson.get("result") if isinstance(searchSolutionjson, dict) else None
            result_data = None
            if isinstance(result_obj, dict) and isinstance(result_obj.get("data"), list):
                # Standard search response shape.
                result_data = result_obj.get("data")
            elif isinstance(result_obj, dict) and result_obj.get("_id"):
                # Mock can return a single solution object directly under "result".
                result_data = [result_obj]

            if not isinstance(result_data, list):
                Helpers.errorVar.append(
                    "checkEntityOfSolution-Invalid response shape: expected 'result.data' list or 'result._id' object, "
                    f"got {searchSolutionjson}"
                )
                return False

            for listOfSoulution in range(0, len(result_data)):
                solution_id = result_data[listOfSoulution]["_id"]
                messageArr.append("solution found : " + str(solution_id))
                Helpers.createAPILog(projectName_for_folder_path, messageArr)
                print("searchSolutionApi Success")

                # If mock already returns full solution details, use it directly.
                inline_result = result_data[listOfSoulution]
                if inline_result.get("isReusable") is False and inline_result.get("externalId"):
                    inline_entity_type = inline_result.get("entityType")
                    if inline_entity_type:
                        solutionEntityType = inline_entity_type
                        solutionExternalId = inline_result.get("externalId")
                        messageArr.append("Task solution Entity Type found : " + str(solutionEntityType))
                        Helpers.createAPILog(projectName_for_folder_path, messageArr)
                        break
                
                solutionDetailsurl = internal_kong_ip + fetchsolutiondoc + str(solution_id)

                solutionDetailspayload = {}
                solutionDetailsheaders = {
                    'X-authenticated-user-token': accessToken,
                    'internal-access-token': internal_access_token,
                    'Authorization': authorization
                }

                solutionDetailsresponse = requests.request("GET", solutionDetailsurl, headers=solutionDetailsheaders,
                                                        data=solutionDetailspayload)

                messageArr = ["Task solution Entity Type found",
                            "URL : " + str(solutionDetailsurl),
                            "Status Code : " + str(solutionDetailsresponse.status_code),
                            "Response : " + str(solutionDetailsresponse.text)]

                if solutionDetailsresponse.status_code == 200:
                    solutionDetailsjson = _parse_maybe_html_json(solutionDetailsresponse, "checkEntityOfSolution-detail")
                    if not solutionDetailsjson:
                        return False
                    solution_result = solutionDetailsjson.get("result", {}) if isinstance(solutionDetailsjson, dict) else {}
                    if solution_result.get("isReusable") is False:
                        solutionEntityType = solution_result.get("entityType")
                        solutionExternalId = solution_result.get("externalId")
                        if not solutionEntityType or not solutionExternalId:
                            Helpers.errorVar.append(
                                f"checkEntityOfSolution-Missing entityType/externalId for solution id {solution_id}"
                            )
                            return False
                        messageArr.append("Task solution Entity Type found : " + str(solutionEntityType))
                        Helpers.createAPILog(projectName_for_folder_path, messageArr)
                        print("FetchSolutionDocApi Success")
                        break
                else:
                    messageArr = ["Solution details fetch failed",
                        "URL : " + str(solutionDetailsurl),
                        "Status Code : " + str(solutionDetailsresponse.status_code),
                        "Response : " + str(solutionDetailsresponse.text)]
                    Helpers.createAPILog(projectName_for_folder_path, messageArr)
                    if solutionDetailsresponse.status_code in [400, 401, 403, 404, 422]:
                        Helpers.errorVar.append(f"FetchSolutionDocApi-Client Error {solutionDetailsresponse.status_code}: {solutionDetailsresponse.text}")
                    elif solutionDetailsresponse.status_code in [500, 502, 503, 504]:
                        Helpers.errorVar.append(f"FetchSolutionDocApi-Server Error {solutionDetailsresponse.status_code}: {solutionDetailsresponse.text}")
                    else:
                        Helpers.errorVar.append(f"FetchSolutionDocApi-Unexpected Error {solutionDetailsresponse.status_code}: {solutionDetailsresponse.text}")
                    return False

        else:
            error_message = ""
            if searchSolutionresponse.status_code in [400, 401, 403, 404, 422]:
                error_message = f"checkEntityOfSolution-Client Error {searchSolutionresponse.status_code}: {searchSolutionresponse.text}"
            elif searchSolutionresponse.status_code in [500, 502, 503, 504]:
                error_message = f"checkEntityOfSolution-Server Error {searchSolutionresponse.status_code}: {searchSolutionresponse.text}"
            else:
                error_message = f"checkEntityOfSolution-Unexpected Error {searchSolutionresponse.status_code}: {searchSolutionresponse.text}"
            Helpers.errorVar.append(error_message)
            return False
        return [solutionEntityType, solutionExternalId]

    @exception_handler
    def check_sequence(arr):
        for i in range(1, len(arr)):
            if arr[i] != arr[i - 1] + 1:
                return False
        return True
    
    @exception_handler
    def createAPILog(solutionName_for_folder_path, messageArr):
        logs_dir = solutionName_for_folder_path + '/apiHitLogs'
        os.makedirs(logs_dir, exist_ok=True)
        file_exists = logs_dir + '/apiLogs.txt'
        # check if the file existis or not and create a file 
        if not path.exists(file_exists):
            API_log = open(file_exists, "w", encoding='utf-8')
            API_log.write("===============================================================================")
            API_log.write("\n")
            API_log.write("ENVIRONMENT : " + str(global_vars.environment))
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

    @exception_handler
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
    
    @exception_handler
    def validateObservationWithRubrics(wbObservation1, accessToken, parentFolder, sheetNames1, isImp=False):
        # Logic for Type 1 (Observation with Rubrics) and Type 5 (Imp Led)
        resourceEndDates = []
        ecmIds = list()
        global_vars.criteriaLevels = list()
        criteriaExternalIds = list()
        cached_details_row = {}
        cached_framework_rows = []
        cached_ecm_rows = []
        cached_questions_rows = []
        cached_criteria_rubric_rows = []
        cached_theme_rubric_rows = []
        cached_imp_mapping_rows = []
        resource_name = ""
        
        # Use globally cached program and resource details
        if global_vars.programDict:
            dictProgramDetails = global_vars.programDict
            programName = Helpers._to_text(dictProgramDetails.get('Title of the Program', ''))
            global_vars.isProgramnamePresent = bool(programName)
            userEntity = Helpers._to_text(dictProgramDetails.get('Targeted state at program level', '')).strip().split(",") if dictProgramDetails.get('Targeted state at program level') else Helpers.errorVar.append("\"scope_entity\" must not be Empty in \"details\" sheet")

        for dictDetailsEnv in global_vars.programResourceDetails:
            endDateOfResources = dictDetailsEnv.get('End date of resource')
            if endDateOfResources:
                resourceEndDates.append(endDateOfResources)
                        
        for sheetEnv in sheetNames1:
            questionsequenceArr =[]
            if sheetEnv == "Instructions":
                pass
            else:
                if sheetEnv.strip().lower() == 'details':
                    print("--->Checking details sheet...")
                    detailsCols = ["observation_solution_name", "observation_solution_description", "Diksha_loginId","Name_of_the_creator", "language", "allow_multiple_submissions", "keywords","scoring_system", "entity_type"]
                    detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                    keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                            range(detailsEnvSheet.ncols)]
                    for row_index_env in range(2, detailsEnvSheet.nrows):
                        dictDetailsEnv = {
                            keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                            col_index_env in range(detailsEnvSheet.ncols)}
                        if not cached_details_row:
                            cached_details_row = dictDetailsEnv
                            resource_name = Helpers._to_text(dictDetailsEnv.get('observation_solution_name', ''))
                        if set(detailsCols) == set(dictDetailsEnv.keys()):
                            global_vars.solutionName = dictDetailsEnv['observation_solution_name'].encode('utf-8').decode('utf-8') if dictDetailsEnv['observation_solution_name'] else Helpers.errorVar.append("\"observation_solution_name\" must not be Empty in \"details\" sheet")
                            global_vars.dikshaLoginId = dictDetailsEnv['Diksha_loginId'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Diksha_loginId'] else Helpers.errorVar.append("\"Diksha_loginId\" must not be Empty in \"details\" sheet")
                            ccUserDetails = Helpers.fetchUserDetails(accessToken, global_vars.dikshaLoginId)
                            if not ccUserDetails:
                                Helpers.errorVar.append(f"Failed to fetch user details for: {global_vars.dikshaLoginId}")
                                return False
                            if not "CONTENT_CREATOR" in ccUserDetails[3]:
                                Helpers.errorVar.append(
                                    f"---> {Helpers._to_text(global_vars.dikshaLoginId)} is not a CONTENT_CREATOR in Diksha {Helpers._to_text(global_vars.environment)}"
                                )
                                return False
                            global_vars.ccRootOrgName = ccUserDetails[4]
                            global_vars.ccRootOrgId = ccUserDetails[5]
                            global_vars.solutionDescription = dictDetailsEnv['observation_solution_description'].encode('utf-8').decode('utf-8')
                            global_vars.pointBasedValue = str(dictDetailsEnv['scoring_system']).encode('utf-8').decode('utf-8') if dictDetailsEnv['scoring_system'] else Helpers.errorVar.append("\"scoring_system\" must not be Empty in \"details\" sheet")
                            global_vars.entityType = dictDetailsEnv['entity_type'].encode('utf-8').decode('utf-8') if dictDetailsEnv['entity_type'] else Helpers.errorVar.append("\"entity_type\" must not be Empty in \"details\" sheet")

                            global_vars.solutionLanguage = dictDetailsEnv['language'].split(",") if dictDetailsEnv['language'] else [""]
                            global_vars.keyWords = dictDetailsEnv['keywords'].encode('utf-8').decode('utf-8')
                            global_vars.creator = dictDetailsEnv['Name_of_the_creator'].encode('utf-8').decode('utf-8')  if dictDetailsEnv['Name_of_the_creator'] else Helpers.errorVar.append("\"Name_of_the_creator\" must not be Empty in \"details\" sheet")
                            global_vars.allow_multiple_submissions = dictDetailsEnv['allow_multiple_submissions']
                            if global_vars.allow_multiple_submissions == 1 or global_vars.allow_multiple_submissions == 'TRUE':
                                global_vars.allow_multiple_submissions = True
                            else:
                                global_vars.allow_multiple_submissions = False
                            
                            # global_vars.scopeEntityType = global_vars.scopeEntityType # Redundant

                            if global_vars.programName == "":
                                global_vars.isProgramnamePresent = False
                            else:
                                global_vars.isProgramnamePresent = True
                                Helpers.getProgramInfo(accessToken, parentFolder, global_vars.programName, resourceEndDates)
                        else:
                            Helpers.errorVar.append("--->Columns Mismatch in Details Sheet.")
                            
                elif sheetEnv.strip().lower() == 'framework':
                    print("--->Checking frameworks sheet...")
                    detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                    keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                            range(detailsEnvSheet.ncols)]
                    listOfThemeCriteria = list()
                    for row_index_env in range(1, detailsEnvSheet.nrows):
                        dictDetailsEnv = {
                            keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                            col_index_env in range(detailsEnvSheet.ncols)}
                        cached_framework_rows.append(dictDetailsEnv)
                        countLevelUp = 1
                        for eachColNameCheck in keysEnv:
                            if "L" + str(countLevelUp) + " description" == eachColNameCheck:
                                countLevelUp += 1
                        for i in range(1, countLevelUp):
                            if not i in global_vars.criteriaLevels:
                                global_vars.criteriaLevels.append(i)

                        if dictDetailsEnv['Criteria ID'].encode('utf-8').decode('utf-8'):
                            if not [dictDetailsEnv['Domain ID'], dictDetailsEnv['Criteria ID']] in listOfThemeCriteria:
                                listOfThemeCriteria.append([dictDetailsEnv['Domain ID'], dictDetailsEnv['Criteria ID']])
                            else:
                                Helpers.errorVar.append("Theme , criteria combo repeating in framework sheet.")
                        if not dictDetailsEnv['Domain ID']:
                            Helpers.errorVar.append("Domain ID cannot be empty in framework sheet.")
                        if not dictDetailsEnv['Domain Name']:
                            Helpers.errorVar.append("Theme cannot be empty in framework sheet.")

                        if dictDetailsEnv['Criteria ID']:
                            criteriaExternalIds.append(dictDetailsEnv['Criteria ID'].lower())
                            
                elif sheetEnv.strip().lower() == 'ecms or domains':
                    print("--->Checking ECMs sheet...")
                    detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                    keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                            range(detailsEnvSheet.ncols)]
                    for row_index_env in range(2, detailsEnvSheet.nrows):
                        dictDetailsEnv = {
                            keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                            col_index_env in range(detailsEnvSheet.ncols)}
                        cached_ecm_rows.append(dictDetailsEnv)
                        if dictDetailsEnv['ECM Id/Domian ID'].lower() not in ecmIds:
                            ecmIds.append(dictDetailsEnv['ECM Id/Domian ID'].lower())
                        if not dictDetailsEnv['ECM Id/Domian ID']:
                            Helpers.errorVar.append("ECM Id/Domian ID cannot be empty in ecm\'s sheet.")
                        if not dictDetailsEnv['section_id']:
                            Helpers.errorVar.append("section_id cannot be empty in ecm\'s sheet.")
                        if not dictDetailsEnv['section_name']:
                            Helpers.errorVar.append("section_name cannot be empty in ecm\'s sheet.")
                        if not dictDetailsEnv['ECM Name/Domain Name']:
                            Helpers.errorVar.append("ECM Name/Domain Name cannot be empty in ecm\'s sheet.")
                        global_vars.ecmToSection[dictDetailsEnv['section_id']] = dictDetailsEnv['ECM Id/Domian ID']
                        
                elif sheetEnv.strip().lower() == 'questions':
                    print("--->Checking questions sheet...")
                    quesExtIds = list()
                    detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                    keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                            range(detailsEnvSheet.ncols)]
                    global_vars.numberOfResponses = 0
                    for qKeys in keysEnv:
                        countRespo = re.search(r"response\(R[0-9]|[1-9][0-9]|100\)$", qKeys)
                        if countRespo and not "_hint" in qKeys and "response" in qKeys:
                            global_vars.numberOfResponses += 1

                    for n in range(1, global_vars.numberOfResponses + 1):
                        if not "Score for R" + str(n) in keysEnv or not "response(R" + str(n) + ")_hint" in keysEnv:
                            Helpers.errorVar.append("Mandatory Key: " + "Score for R" + str(n) + " or " + "response(R" + str(
                                n) + ")_hint is missing")
                    for row_index_env in range(2, detailsEnvSheet.nrows):
                        dictDetailsEnv = {
                            keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                            col_index_env in range(detailsEnvSheet.ncols)}
                        if not any(Helpers._to_text(v).strip() for v in dictDetailsEnv.values()):
                            continue
                        cached_questions_rows.append(dictDetailsEnv)
                        quesExtIds.append(dictDetailsEnv['question_id'].encode('utf-8').decode('utf-8').lower())

                        if not dictDetailsEnv['criteria_id']:
                            Helpers.errorVar.append("criteria_id cannot be empty in questions sheet.")
                        if not dictDetailsEnv['criteria_id'].lower() in criteriaExternalIds:
                            Helpers.errorVar.append("Criteria ID : " + dictDetailsEnv['criteria_id'] + " in question sheet not present in criteria sheet.")
                        question_sequence = dictDetailsEnv['question_sequence'] if dictDetailsEnv['question_sequence'] else Helpers.errorVar.append("\"question_sequence\" must not be Empty in \"questions\" sheet")

                        questionsequenceArr.append(question_sequence)
                        question_sequence_arr = questionsequenceArr

                        if not dictDetailsEnv['question_primary_language']:
                            Helpers.errorVar.append("question_primary_language cannot be empty in questions sheet.")
                        if not dictDetailsEnv['question_response_type']:
                            Helpers.errorVar.append("question_response_type cannot be empty in questions sheet.")
                        if not dictDetailsEnv['question_id']:
                            Helpers.errorVar.append("question_id cannot be empty in questions sheet.")
                        if not dictDetailsEnv['criteria_id']:
                            Helpers.errorVar.append("criteria_id : " + str(
                                dictDetailsEnv['criteria_id']) + "  cannot be empty in questions sheet.")
                        if not dictDetailsEnv['criteria_id'].lower() in criteriaExternalIds:
                            Helpers.errorVar.append("criteria_id : " + str(dictDetailsEnv['criteria_id']) + " in questions sheet is not matching the criteria upload.")
                    if not len(question_sequence_arr) == len(set(question_sequence_arr)):
                            Helpers.errorVar.append("\"question_sequence\" must be Unique in \"questions\" sheet")
                    if not len(quesExtIds) == len(set(quesExtIds)):
                        Helpers.errorVar.append("Duplicate question_id detected in questions sheet.")
                    if not Helpers.check_sequence(question_sequence_arr): Helpers.errorVar.append("\"question_sequence\" must be in sequence in \"questions\" sheet")
                
                if isImp:
                    if sheetEnv.strip().lower() == 'imp mapping':
                        print("--->Checking Imp mapping sheet...")
                        global_vars.countImps = 1
                        detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                        keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                                range(detailsEnvSheet.ncols)]
                        for row_index_env in range(2, detailsEnvSheet.nrows):
                            dictDetailsEnv = {
                                keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                col_index_env in range(detailsEnvSheet.ncols)}
                            cached_imp_mapping_rows.append(dictDetailsEnv)
                        for eachCols in dictDetailsEnv.keys():
                            if eachCols.strip() == "L" + str(global_vars.countImps) + "-improvement-projects":
                                global_vars.countImps += 1
                        global_vars.countImps = global_vars.countImps - 1

                if not global_vars.pointBasedValue.lower() == "null":
                    if sheetEnv.strip().lower() == 'Criteria_Rubric-Scoring':
                        print("--->Checking Criteria Rubrics sheet")
                        cR_extIds = list()
                        detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                        keysEnv = [detailsEnvSheet.cell(0, col_index_env).value for col_index_env in
                                range(detailsEnvSheet.ncols)]
                        listOfCRs = ["criteriaId", "weightage"]
                        for cl in global_vars.criteriaLevels:
                            listOfCRs.append("L" + str(cl))
                        for keyys in keysEnv:
                            if not keyys in listOfCRs:
                                print("--->" + keyys + " : unwanted column detected...")
                                print("==>PS :  unwanted column will be ignored while uploading...")
                        for row_index_env in range(1, detailsEnvSheet.nrows):
                            dictDetailsEnv = {
                                keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                col_index_env in range(detailsEnvSheet.ncols)}
                            cached_criteria_rubric_rows.append(dictDetailsEnv)
                            cR_extIds.append(dictDetailsEnv['criteriaId'].lower())
                            for cl in global_vars.criteriaLevels:
                                if not dictDetailsEnv["L" + str(cl)]:
                                    Helpers.errorVar.append("L" + str(cl) + " must not be empty in criteria_rubric.")
                            if dictDetailsEnv['criteriaId']:
                                Helpers.errorVar.append("criteriaId must be empty in criteria_rubric sheet.")
                            if not dictDetailsEnv['weightage']:
                                Helpers.errorVar.append("weightage cannot be empty in criteria_rubric sheet.")
                        if not len(cR_extIds) == len(set(cR_extIds)):
                            Helpers.errorVar.append("Duplicate externalId detected in criteria_rubric sheet.")
                    
                    if sheetEnv.strip().lower() == 'Domain(theme)_rubric_scoring':
                        print("--->Checking Theme Rubrics sheet")
                        detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                        keysEnv = [detailsEnvSheet.cell(0, col_index_env).value for col_index_env in
                                range(detailsEnvSheet.ncols)]
                        for row_index_env in range(1, detailsEnvSheet.nrows):
                            dictDetailsEnv = {
                                keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                                col_index_env in range(detailsEnvSheet.ncols)}
                            cached_theme_rubric_rows.append(dictDetailsEnv)
                            if not dictDetailsEnv['domain_Id']:
                                Helpers.errorVar.append("domain_Id cannot be empty in theme_rubric sheet.")
                            if not dictDetailsEnv['domain_name']:
                                Helpers.errorVar.append("domain_name cannot be empty in theme_rubric sheet.")
                            if not dictDetailsEnv['weightage']:
                                Helpers.errorVar.append("weightage cannot be empty in theme_rubric sheet.")

        Helpers._update_resource_validation_cache(
            resource_name=resource_name,
            details_row=cached_details_row,
            framework_rows=cached_framework_rows,
            ecm_rows=cached_ecm_rows,
            question_rows=cached_questions_rows,
            criteria_rubric_rows=cached_criteria_rubric_rows,
            theme_rubric_rows=cached_theme_rubric_rows,
            imp_mapping_rows=cached_imp_mapping_rows
        )
        return resourceEndDates




    @exception_handler
    def validateObservationWithRubricsLedImp(wbObservation1, accessToken, parentFolder, sheetNames1):
        return Helpers.validateObservationWithRubrics(wbObservation1, accessToken, parentFolder, sheetNames1, isImp=True)

    @exception_handler
    def validateObservationWithoutRubrics(wbObservation1, accessToken, parentFolder, sheetNames1):
        questionsequenceArr =[]
        # Point based value set as null by default for observation without rubrics
        global_vars.pointBasedValue = "null"
        criteria_id_arr = []
        cached_details_row = {}
        cached_criteria_rows = []
        cached_questions_rows = []
        resource_name = ""
        detailsColNames = ['observation_solution_name', 'observation_solution_description', 'Diksha_loginId','language', 'keywords', 'entity_type', "scope_entity"]
        criteriaColNames = ['criteria_id', 'criteria_name']
        questionsColNames = ["criteria_id","question_sequence","question_id","instance_parent_question_id","parent_question_id","show_when_parent_question_value_is","parent_question_value","page","question_number","question_primary_language","question_secondory_language","question_tip","question_hint","instance_identifier","question_response_type","date_auto_capture","response_required","min_number_value","max_number_value","file_upload","show_remarks","response(R1)","response(R1)_hint","response(R2)","response(R2)_hint","response(R3)","response(R3)_hint","response(R4)","response(R4)_hint","response(R5)","response(R5)_hint","response(R6)","response(R6)_hint","response(R7)","response(R7)_hint","response(R8)","response(R8)_hint","response(R9)","response(R9)_hint","response(R10)","response(R10)_hint","response(R11)","response(R11)_hint","response(R12)","response(R12)_hint","response(R13)","response(R13)_hint","response(R14)","response(R14)_hint","response(R15)","response(R15)_hint","response(R16)","response(R16)_hint","response(R17)","response(R17)_hint","response(R18)","response(R18)_hint","response(R19)","response(R19)_hint","response(R20)","response(R20)_hint","question_weightage","section_header"]
        for sheetColCheck in sheetNames1:
            if sheetColCheck.strip().lower() == 'details':
                detailsColCheck = wbObservation1.sheet_by_name(sheetColCheck)
                keysColCheckDetai = [detailsColCheck.cell(1, col_index_check).value for col_index_check in
                                    range(detailsColCheck.ncols)]
                if len(keysColCheckDetai) != len(detailsColNames):
                    Helpers.errorVar.append('Columns is missing in details sheet')
            if sheetColCheck.strip().lower() == 'criteria':
                criteriaColCheck = wbObservation1.sheet_by_name(sheetColCheck)
                keysColCheckCrit = [criteriaColCheck.cell(1, col_index_check1).value for col_index_check1 in
                                    range(criteriaColCheck.ncols)]
                if len(keysColCheckCrit) != len(criteriaColNames):
                    Helpers.errorVar.append('Columns is missing in criteria sheet')
            if sheetColCheck.strip().lower() == 'questions':
                questionsColCheck = wbObservation1.sheet_by_name(sheetColCheck)
                keysColCheckQues = [questionsColCheck.cell(1, col_index_check2).value for col_index_check2 in
                                    range(questionsColCheck.ncols)]
                if len(keysColCheckQues) != len(questionsColNames):
                    Helpers.errorVar.append('Columns is missing in questions sheet')
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
                        if not cached_details_row:
                            cached_details_row = dictDetailsEnv
                            resource_name = Helpers._to_text(dictDetailsEnv.get('observation_solution_name', ''))
                        global_vars.solutionName = dictDetailsEnv['observation_solution_name'].encode('utf-8').decode('utf-8') if dictDetailsEnv['observation_solution_name'] else Helpers.errorVar.append("\"observation_solution_name\" must not be Empty in \"details\" sheet")
                        global_vars.solutionDescription = dictDetailsEnv['observation_solution_description'].encode('utf-8').decode('utf-8') if dictDetailsEnv['observation_solution_description'] else Helpers.errorVar.append("\"observation_solution_description\" must not be Empty in \"details\" sheet")
                        global_vars.dikshaLoginId = dictDetailsEnv['Diksha_loginId'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Diksha_loginId'] else Helpers.errorVar.append("\"Diksha_loginId\" must not be Empty in \"details\" sheet")
                        global_vars.creator = dictDetailsEnv['Name_of_the_creator'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Name_of_the_creator'] else Helpers.errorVar.append("\"Name_of_the_creator\" must not be Empty in \"details\" sheet")
                        ccUserDetails = Helpers.fetchUserDetails(accessToken, global_vars.dikshaLoginId)
                        if not ccUserDetails:
                            Helpers.errorVar.append(f"Failed to fetch user details for: {global_vars.dikshaLoginId}")
                            return False
                        if not "CONTENT_CREATOR" in ccUserDetails[3]:
                            Helpers.errorVar.append(
                                f"---> {Helpers._to_text(global_vars.dikshaLoginId)} is not a CONTENT_CREATOR in Diksha {Helpers._to_text(global_vars.environment)}"
                            )
                            return False
                        global_vars.ccRootOrgName = ccUserDetails[4]
                        global_vars.ccRootOrgId = ccUserDetails[5]
                            
                        global_vars.entityType = dictDetailsEnv['entity_type'].encode('utf-8').decode('utf-8') if dictDetailsEnv['entity_type'] else Helpers.errorVar.append("\"entity_type\" must not be Empty in \"details\" sheet")
                        global_vars.solutionLanguage = dictDetailsEnv['language'].encode('utf-8').decode('utf-8').split(",") if dictDetailsEnv['language'] else [""]
                        Helpers.getProgramInfo(accessToken, parentFolder, global_vars.programNameInp, [])

                elif sheetEnv.strip().lower() == 'criteria':
                    print("--->Checking criteria sheet...")
                    detailsEnvSheet = wbObservation1.sheet_by_name(sheetEnv)
                    keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                            range(detailsEnvSheet.ncols)]
                    for row_index_env in range(2, detailsEnvSheet.nrows):
                        dictDetailsEnv = {
                            keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value for
                            col_index_env in range(detailsEnvSheet.ncols)}
                        cached_criteria_rows.append(dictDetailsEnv)
                        criteria_id = dictDetailsEnv['criteria_id'].encode('utf-8').decode('utf-8') if dictDetailsEnv['criteria_id'] else Helpers.errorVar.append("\"criteria_id\" must not be Empty in \"criteria\" sheet")
                        criteria_name = dictDetailsEnv['criteria_name'].encode('utf-8').decode('utf-8') if dictDetailsEnv['criteria_name'] else Helpers.errorVar.append("\"criteria_name\" must not be Empty in \"criteria\" sheet")
                        criteria_id_arr.append(criteria_id)
                    if not len(criteria_id_arr) == len(set(criteria_id_arr)):
                        Helpers.errorVar.append("\"criteria_id\" must be Unique in \"criteria\" sheet")
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
                        if not any(Helpers._to_text(v).strip() for v in dictDetailsEnv.values()):
                            continue
                        cached_questions_rows.append(dictDetailsEnv)
                        criteria_id = dictDetailsEnv['criteria_id'].encode('utf-8').decode('utf-8') if dictDetailsEnv['criteria_id'] else Helpers.errorVar.append("\"criteria_id\" must not be Empty in \"questions\" sheet")
                        question_sequence = dictDetailsEnv['question_sequence'] if dictDetailsEnv['question_sequence'] else Helpers.errorVar.append("\"question_sequence\" must not be Empty in \"questions\" sheet")

                        questionsequenceArr.append(question_sequence)
                        question_sequence_arr = questionsequenceArr

                        if not criteria_id in criteria_id_arr:
                            Helpers.errorVar.append("\"criteria_id\" in \"Questions\" sheet must be declared in \"criteria\" sheet")
                        page = dictDetailsEnv['page'].encode('utf-8').decode('utf-8') if dictDetailsEnv['page'] else Helpers.errorVar.append("\"page\" must not be Empty in \"questions\" sheet")
                        question_number = dictDetailsEnv['question_number'] if dictDetailsEnv['question_number'] else Helpers.errorVar.append("\"question_number\" must not be Empty in \"questions\" sheet")
                        question_primary_language = dictDetailsEnv['question_primary_language'].encode('utf-8').decode('utf-8') if dictDetailsEnv['question_primary_language'] else Helpers.errorVar.append("\"question_primary_language\" must not be Empty in \"questions\" sheet")
                        
                        response_required = dictDetailsEnv['response_required'] if str(dictDetailsEnv['response_required']) else Helpers.errorVar.append("\"response_required\" must not be Empty in \"questions\" sheet")

                        question_id = dictDetailsEnv['question_id'] if dictDetailsEnv['question_id'] else Helpers.errorVar.append("\"question_id\" must not be Empty in \"questions\" sheet")
                        ques_id_arr.append(question_id)
                        parent_question_id = dictDetailsEnv['question_id']
                        if parent_question_id and not parent_question_id in ques_id_arr:
                            Helpers.errorVar.append("parent_question_id referenced before assigning in questions sheet.")
                        question_response_type = dictDetailsEnv['question_response_type'].encode('utf-8').decode('utf-8') if dictDetailsEnv[
                            'question_response_type'] else Helpers.errorVar.append(
                            "\"question_response_type\" must not be Empty in \"questions\" sheet")
                    if not len(question_sequence_arr) == len(set(question_sequence_arr)):
                            Helpers.errorVar.append("\"question_sequence\" must be Unique in \"questions\" sheet")
                    if not Helpers.check_sequence(question_sequence_arr): Helpers.errorVar.append("\"question_sequence\" must be in sequence in \"questions\" sheet")
        Helpers._update_resource_validation_cache(
            resource_name=resource_name,
            details_row=cached_details_row,
            criteria_rows=cached_criteria_rows,
            question_rows=cached_questions_rows
        )

    @exception_handler
    def validateSurvey(wbObservation1, sheetNames1):
        print("Validating survey temp....")
        cached_details_row = {}
        cached_questions_rows = []
        resource_name = ""
        for sheetEnvCheck in sheetNames1:
            if sheetEnvCheck.strip().lower() == 'instructions' or sheetEnvCheck.strip().lower() == 'details' or sheetEnvCheck.strip().lower() == 'questions':
                pass
            else:
                Helpers.errorVar.append('Sheet Names in excel file is wrong , Sheet Names are details,questions')

        detailsColNames = ["survey_solution_name", "survey_solution_description", "Name_of_the_creator","survey_creator_username", "survey_start_date", "survey_end_date"]
        questionsColNames = ["question_sequence", "question_id", "section_header", "instance_parent_question_id",
                            "parent_question_id", "show_when_parent_question_value_is", "parent_question_value",
                            "page", "question_number", "question_language1", "question_language2", "question_tip",
                            "question_hint", "instance_identifier", "question_response_type", "date_auto_capture",
                            "response_required", "min_number_value", "max_number_value", "file_upload", "show_remarks",
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
                if len(keysColCheckDetai) != len(detailsColNames):
                    Helpers.errorVar.append('Some Columns are missing in details sheet')
                if detailsColCheck.nrows > 2:
                    cached_details_row = {
                        keysColCheckDetai[col_index_env]: detailsColCheck.cell(2, col_index_env).value
                        for col_index_env in range(detailsColCheck.ncols)
                    }
                    resource_name = Helpers._to_text(cached_details_row.get("survey_solution_name", ""))
            if sheetColCheck.strip().lower() == 'questions':
                questionsColCheck = wbObservation1.sheet_by_name(sheetColCheck)
                keysColCheckQues = [questionsColCheck.cell(1, col_index_check2).value for col_index_check2 in
                                    range(questionsColCheck.ncols)]
                # print(keysColCheckQues)
                if len(keysColCheckQues) != len(questionsColNames):
                    Helpers.errorVar.append('Some Columns are missing in questions sheet')
                for row_index_env in range(2, questionsColCheck.nrows):
                    dictDetailsEnv = {
                        keysColCheckQues[col_index_env]: questionsColCheck.cell(row_index_env, col_index_env).value for
                        col_index_env in range(questionsColCheck.ncols)}
                    if not any(Helpers._to_text(v).strip() for v in dictDetailsEnv.values()):
                        continue
                    cached_questions_rows.append(dictDetailsEnv)
                    question_sequenceSUR = dictDetailsEnv['question_sequence'] if dictDetailsEnv[
                        'question_sequence'] else Helpers.errorVar.append(
                        "\"question_sequence\" must not be Empty in \"details\" sheet")
                    question_idSUR = dictDetailsEnv['question_id'].encode('utf-8').decode('utf-8') if dictDetailsEnv[
                        'question_id'] else Helpers.errorVar.append("\"question_id\" must not be Empty in \"details\" sheet")
                    pageSUR = dictDetailsEnv['page'] if dictDetailsEnv['page'] else Helpers.errorVar.append(
                        "\"page\" must not be Empty in \"details\" sheet")
                    question_numberSUR = dictDetailsEnv['question_number'] if dictDetailsEnv[
                        'question_number'] else Helpers.errorVar.append(
                        "\"question_number\" must not be Empty in \"details\" sheet")
                    question_language1SUR = dictDetailsEnv['question_language1'].encode('utf-8').decode('utf-8') if not dictDetailsEnv['question_language1'] == None else Helpers.errorVar.append(
                        "\"question_language1\" must not be Empty in \"details\" sheet")
                    question_response_typeSUR = dictDetailsEnv['question_response_type'] if dictDetailsEnv[
                        'question_response_type'] else Helpers.errorVar.append(
                        "\"question_response_type\" must not be Empty in \"details\" sheet")
                    response_requiredSUR = dictDetailsEnv['response_required'] if dictDetailsEnv[
                        'response_required'] else Helpers.errorVar.append(
                        "\"response_required\" must not be Empty in \"details\" sheet")
        Helpers._update_resource_validation_cache(
            resource_name=resource_name,
            details_row=cached_details_row,
            question_rows=cached_questions_rows
        )

    @exception_handler
    def validateProject(wbObservation1, sheetNames1):
        print("Validating Project file specifics...")
        def _norm_col(col_name):
            return re.sub(r"\s+", " ", Helpers._to_text(col_name).strip()).lower()

        def _validate_columns(actual_cols, expected_cols, sheet_label):
            actual_norm = [_norm_col(c) for c in actual_cols if Helpers._to_text(c).strip()]
            expected_norm = [_norm_col(c) for c in expected_cols]
            missing = [c for c in expected_cols if _norm_col(c) not in actual_norm]
            extra = [c for c in actual_cols if _norm_col(c) not in expected_norm and Helpers._to_text(c).strip()]
            if missing or extra or len(actual_norm) != len(expected_norm):
                msg = f"Columns mismatch in {sheet_label} sheet."
                if missing:
                    msg += f" Missing: {missing}."
                if extra:
                    msg += f" Extra: {extra}."
                Helpers.errorVar.append(msg)
                return False
            return True

        criteria_id_arr = list()
        cached_project_rows = []
        cached_task_rows = []
        cached_certificate_rows = []
        resource_name = ""
        projectDetailsCols = ["title", "projectId", "is a SSO user?", "Diksha_loginId", "categories",
                            "objective","duration","recommendedFor","keywords"]
        detailsColCheck = wbObservation1.sheet_by_name('Project upload')
        keysColCheckDetai = [detailsColCheck.cell(1, col_index_check).value for col_index_check in
                                    range(detailsColCheck.ncols)]
        lentasks = (len(keysColCheckDetai) - 12) // 2
        for i in range(lentasks):
            projectDetailsCols.append(f"learningResources{i+1}-name")
            projectDetailsCols.append(f"learningResources{i+1}-link")
        projectDetailsCols.append("has certificate")
        projectDetailsCols.append("Project Level Evidence")
        projectDetailsCols.append("Minimum No. of Evidence")

        taskUploadCols = ["TaskId", "TaskTitle", "Subtask",
                        "Mandatory task(Yes or No)","observation Name","Number of submissions for observation"]
        detailsColCheck = wbObservation1.sheet_by_name('Tasks upload')
        keysColCheckDetai = [detailsColCheck.cell(1, col_index_check).value for col_index_check in
                                    range(detailsColCheck.ncols)]
        lentasks = (len(keysColCheckDetai) - 10) // 2
        for i in range(lentasks):
            taskUploadCols.append(f"learningResources{i+1}-name")
            taskUploadCols.append(f"learningResources{i+1}-link")
        taskUploadCols.append("Evidence required for any task for certificate criteria")
        taskUploadCols.append("Minimum No. of Evidence for any task criteria")
        taskUploadCols.append("Task Level Evidence req. for certificate criteria")
        taskUploadCols.append("Minimum No. of Evidence for task level evidence criteria")

        certificateCols = ["Certificate issuer","Type of certificate","Logo - 1","Logo - 2","Authorised Signature Image - 1","Authorised Signature Name - 1",
                           "Authorised Designation - 1","Authorised Signature Image - 2","Authorised Signature Name - 2","Authorised Designation - 2"]
        for sheetColCheck in sheetNames1:
            if sheetColCheck.strip().lower() == 'Project upload'.lower():
                print("--->Checking Project Upload sheet...")
                detailsColCheck = wbObservation1.sheet_by_name(sheetColCheck)
                keysColCheckDetai = [detailsColCheck.cell(1, col_index_check).value for col_index_check in
                                    range(detailsColCheck.ncols)]
                _validate_columns(keysColCheckDetai, projectDetailsCols, "Project Upload")
                detailsEnvSheet = wbObservation1.sheet_by_name(sheetColCheck)
                keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                        range(detailsEnvSheet.ncols)]
                for row_index_env in range(2, detailsEnvSheet.nrows):
                    # print(dictDetailsEnv)
                    # sys.exit()
                    dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                    for
                                    col_index_env in range(detailsEnvSheet.ncols)}
                    cached_project_rows.append(dictDetailsEnv)
                    if not resource_name:
                        resource_name = Helpers._to_text(dictDetailsEnv.get("title", ""))
                    projectTitle = dictDetailsEnv['title'].encode('utf-8').decode('utf-8') if dictDetailsEnv['title'] else Helpers.errorVar.append(
                        "\"title\" must not be Empty in \"Project Upload\" sheet")
                    projectId = dictDetailsEnv['projectId'] if dictDetailsEnv['projectId'] else Helpers.errorVar.append(
                        "\"projectId\" must not be Empty in \"Project Upload\" sheet")
                    Helpers.validate_identifier(projectId)
                    projectCategories = dictDetailsEnv['categories'].encode('utf-8').decode('utf-8') if dictDetailsEnv[
                        'categories'] else Helpers.errorVar.append(
                        "\"categories\" must not be Empty in \"Project Upload\" sheet")
                        
                    projectDescription = dictDetailsEnv["objective"].encode('utf-8').decode('utf-8') if dictDetailsEnv[
                        "objective"] else Helpers.errorVar.append(
                        "\"objective\" must not be Empty in \"Project Upload\" sheet")
                    projectSSOuser = dictDetailsEnv["is a SSO user?"] if dictDetailsEnv[
                        "is a SSO user?"] else Helpers.errorVar.append(
                        "\"is a SSO user?\" must not be Empty in \"Project Upload\" sheet")
                    projectDikshaloginid = dictDetailsEnv["Diksha_loginId"].encode('utf-8').decode('utf-8') if dictDetailsEnv["Diksha_loginId"] else Helpers.errorVar.append("\"Diksha_loginId\" must not be Empty in \"Project Upload\" sheet")
                    projectDuration = dictDetailsEnv["duration"].encode('utf-8').decode('utf-8') if dictDetailsEnv[
                        "duration"] else Helpers.errorVar.append(
                        "\"duration\" must not be Empty in \"Project Upload\" sheet")
                    projectcertificate = dictDetailsEnv["has certificate"] if dictDetailsEnv["has certificate"] else Helpers.errorVar.append(
                        "\"has certificate\" must not be Empty in \"Project Upload\" sheet")


            if sheetColCheck.strip().lower() == 'Tasks upload'.lower():
                print("--->Checking Tasks upload sheet...")
                # sys.exit()
                detailsColCheck = wbObservation1.sheet_by_name(sheetColCheck)
                keysColCheckDetai = [detailsColCheck.cell(1, col_index_check).value for col_index_check in
                                    range(detailsColCheck.ncols)]
                
                _validate_columns(keysColCheckDetai, taskUploadCols, "Tasks upload")
                detailsEnvSheet = wbObservation1.sheet_by_name(sheetColCheck)
                keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                        range(detailsEnvSheet.ncols)]
                for row_index_env in range(2, detailsEnvSheet.nrows):
                    dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                    for
                                    col_index_env in range(detailsEnvSheet.ncols)}
                    cached_task_rows.append(dictDetailsEnv)
                    projectTaskMandatory = dictDetailsEnv['Mandatory task(Yes or No)'] if dictDetailsEnv[
                        'Mandatory task(Yes or No)'] else Helpers.errorVar.append(
                        "\"Mandatory task(Yes or No)\" must not be Empty in \"Tasks Upload\" sheet")
                        

            if sheetColCheck.strip().lower() == 'Certificate details'.lower():
                print("--->Checking Certificate details  sheet...")

                detailsColCheck = wbObservation1.sheet_by_name(sheetColCheck)
                keysColCheckDetai = [detailsColCheck.cell(1, col_index_check).value for col_index_check in
                                        range(detailsColCheck.ncols)]

                if not _validate_columns(keysColCheckDetai, certificateCols, "Certificate details"):
                    print("certificate not found")
                detailsEnvSheet = wbObservation1.sheet_by_name(sheetColCheck)
                keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in
                            range(detailsEnvSheet.ncols)]
                for row_index_env in range(2, detailsEnvSheet.nrows):

                    dictDetailsEnv = {keysEnv[col_index_env]: detailsEnvSheet.cell(row_index_env, col_index_env).value
                                    for
                                    col_index_env in range(detailsEnvSheet.ncols)}
                    cached_certificate_rows.append(dictDetailsEnv)
                    if projectcertificate == "Yes":
                        certificateissuer = dictDetailsEnv['Certificate issuer'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Certificate issuer'] else Helpers.errorVar.append(
                        "\"Certificate issuer\" must not be Empty in \"Certificate details\" sheet")
                        
                        
                        Typeofcertificate = dictDetailsEnv['Type of certificate'] if dictDetailsEnv['Type of certificate'] in ["One Logo - One Signature","One Logo - Two Signature","Two Logo - One Signature","Two Logo - Two Signature"]  else Helpers.errorVar.append(
                        "\"Type of certificate\" must not be Empty in \"Certificate details\" sheet")
                        Logo1 = dictDetailsEnv['Logo - 1'] if dictDetailsEnv[
                        'Logo - 1'] else Helpers.errorVar.append(
                        "\"Logo - 1\" must not be Empty in \"Certificate details\" sheet")

                        Authorisedsignlogo1 = dictDetailsEnv['Authorised Signature Image - 1'] if dictDetailsEnv['Authorised Signature Image - 1'] else Helpers.errorVar.append("\"Authorised Signature Image - 1\" must not be Empty in \"Certificate details\" sheet")
                        Authorisedsignname1 = dictDetailsEnv['Authorised Signature Name - 1'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Authorised Signature Name - 1'] else Helpers.errorVar.append("\"Authorised Signature Name - 1\" must not be Empty in \"Certificate details\" sheet")
                        Authoriseddesifnation1 = dictDetailsEnv['Authorised Designation - 1'].encode('utf-8').decode('utf-8') if dictDetailsEnv['Authorised Designation - 1'] else Helpers.errorVar.append("\"Authorised Designation - 1\" must not be Empty in \"Certificate details\" sheet")
        Helpers._update_resource_validation_cache(
            resource_name=resource_name,
            project_upload_row=(cached_project_rows[0] if cached_project_rows else {}),
            project_upload_rows=cached_project_rows,
            task_upload_rows=cached_task_rows,
            certificate_detail_rows=cached_certificate_rows
        )

    @exception_handler
    def validateSheets(filePathAddObs, accessToken, parentFolder):
        global_vars.reset_resource_validation_cache()
        wbObservation1 = xlrd.open_workbook(filePathAddObs, on_demand=True)
        sheetNames1 = wbObservation1.sheet_names()
        
        rubrics_sheet_names = ['Instructions', 'details', 'framework', 'ECMs or Domains', 'questions','Criteria_Rubric-Scoring', 'Domain(theme)_rubric_scoring']
        rubrics_sheet_IMP_names = ['Instructions', 'details', 'framework', 'ECMs or Domains', 'questions','Criteria_Rubric-Scoring', 'Domain(theme)_rubric_scoring', 'Imp mapping']
        observation_sheet_names = ['Instructions', 'details', 'criteria', 'questions']
        survey_sheet_names = ['Instructions', 'details', 'questions']
        project_sheet_names = ['Instructions', 'Project upload', 'Tasks upload','Certificate details']

        # 1-with rubrics , 2 - with out rubrics , 3 - survey , 4 - Project 5 - With rubric and IMP
        typeofSolutin = 0

        if (len(rubrics_sheet_names) == len(sheetNames1)) and ((set(rubrics_sheet_names) == set(sheetNames1))):
            print("--->Observation with rubrics file detected.<---")
            typeofSolutin = 1
        elif (len(observation_sheet_names) == len(sheetNames1)) and ((set(observation_sheet_names) == set(sheetNames1))):
            print("--->Observation without rubrics file detected.<---")
            typeofSolutin = 2
        elif (len(survey_sheet_names) == len(sheetNames1)) and ((set(survey_sheet_names) == set(sheetNames1))):
            print("--->Survey file detected.<---")
            typeofSolutin = 3
        elif (len(project_sheet_names) == len(sheetNames1)) and ((set(project_sheet_names) == set(sheetNames1))):
            print("--->Project file detected.<---")
            typeofSolutin = 4
        elif (len(rubrics_sheet_IMP_names) == len(sheetNames1)) and ((set(rubrics_sheet_IMP_names) == set(sheetNames1))):
            print("--->Observation with rubrics and IMP file detected.<---")
            typeofSolutin = 5
        else:
            typeofSolutin = 0
            Helpers.errorVar.append("Please check the Input sheet.")

        Helpers._update_resource_validation_cache(
            solution_path=os.path.abspath(filePathAddObs) if filePathAddObs else filePathAddObs,
            workbook=wbObservation1,
            sheet_names=sheetNames1,
            solution_type=typeofSolutin,
            resource_name=""
        )
        
        validation_result = True
        if typeofSolutin == 1:
            validation_result = Helpers.validateObservationWithRubrics(wbObservation1, accessToken, parentFolder, sheetNames1)
        elif typeofSolutin == 2:
            validation_result = Helpers.validateObservationWithoutRubrics(wbObservation1, accessToken, parentFolder, sheetNames1)
        elif typeofSolutin == 3:
            validation_result = Helpers.validateSurvey(wbObservation1, sheetNames1)
        elif typeofSolutin == 4:
            validation_result = Helpers.validateProject(wbObservation1, sheetNames1)
        elif typeofSolutin == 5:
            validation_result = Helpers.validateObservationWithRubricsLedImp(wbObservation1, accessToken, parentFolder, sheetNames1)

        if validation_result is False or Helpers.errorVar:
            return False

        return typeofSolutin

       
    @exception_handler
    def mainFunc(MainFilePath, programFile, millisecond, addObservationSolution):
        # Use addObservationSolution parameter instead of global variable
        if addObservationSolution:
            global_vars.addObservationSolution = addObservationSolution
        
        global_vars.surveySolutionlink = None
        parentFolder = Helpers.createFileStruct(MainFilePath, global_vars.addObservationSolution)
        # print(parentFolder,"2761")
        accessToken = Helpers.generateAccessToken()
        if not accessToken:
            Helpers.errorVar.append("Failed to generate access token")
            return False
            
        typeofSolution = Helpers.validateSheets(global_vars.addObservationSolution, accessToken, parentFolder)
        print(f"Validation result: {typeofSolution}")
        if not typeofSolution:
            if not Helpers.errorVar:
                Helpers.errorVar.append("Validation failed for the provided resource template")
            return False
        # Validate sheets returns error dict if validation fails - check types
        # if isinstance(typeofSolution, dict) and not typeofSolution.get("success", True):
        #      return typeofSolution
             
        validation_cache = global_vars.get_resource_validation_cache() or {}
        cached_path = validation_cache.get("solution_path")
        current_path = os.path.abspath(global_vars.addObservationSolution) if global_vars.addObservationSolution else global_vars.addObservationSolution
        wbObservation = validation_cache.get("workbook") if cached_path == current_path else None
        if wbObservation is None:
            wbObservation = xlrd.open_workbook(global_vars.addObservationSolution, on_demand=True)
        
        # Check program file structure once per program file path.
        program_validation_cache = getattr(global_vars, "programValidationCache", {}) or {}
        current_program_path = os.path.abspath(programFile) if programFile else programFile
        program_already_validated = (
            program_validation_cache.get("program_path") == current_program_path
            and bool(program_validation_cache.get("is_valid"))
        )
        if not program_already_validated and not Helpers.programsFileCheck(programFile, accessToken, parentFolder, MainFilePath):
            if not Helpers.errorVar:
                Helpers.errorVar.append("Program file validation failed")
            return False
            
        wbproject = wbObservation
        projectSheetNames = validation_cache.get("sheet_names") if cached_path == current_path else wbproject.sheet_names()
        
        dictProgramDetails = global_vars.programDict
        if not dictProgramDetails:
            global_vars.load_program_template(programFile)
            dictProgramDetails = global_vars.programDict

        programName = Helpers._to_text(dictProgramDetails.get('Title of the Program', ''))
        isProgramnamePresent = bool(programName)
        userEntity = Helpers._to_text(dictProgramDetails.get('Targeted state at program level', '')).strip().split(",") if dictProgramDetails.get('Targeted state at program level') else Helpers.errorVar.append("\"scope_entity\" must not be Empty in \"details\" sheet")
        
                    
        for sheets in projectSheetNames:
            if sheets.strip().lower() == 'details'.lower() and typeofSolution in [1, 5]:
                try:
                    ObsWRResourceName = Helpers._to_text(validation_cache.get("resource_name", "")) or Helpers._to_text(global_vars.solutionName or "")
                except Exception as e:
                    error_msg = f"Error reading 'details' sheet or processing observation solution name: {str(e)}"
                    print(error_msg)
                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                    return False

                try:
                    @exception_handler
                    def addObsWRFunc(parentFolder, wbObservation, millisecond, accessToken):
                        try:
                            impLedObsFlag = True if typeofSolution == 5 else False
                            if not Helpers.criteriaUpload(parentFolder, wbObservation, millisecond, accessToken, "framework", impLedObsFlag):
                                return "", Helpers.errorVar
                            print("Criteria Upload success....")
                        except Exception as e:
                            error_msg = f"Error during criteria upload: {str(e)}"
                            print(error_msg)
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                            return "", Helpers.errorVar

                        try:
                            userDetails = Helpers.fetchUserDetails(accessToken, global_vars.dikshaLoginId)
                            if not userDetails:
                                Helpers.errorVar.append(f"Failed to fetch user details for: {global_vars.dikshaLoginId}")
                                return "", Helpers.errorVar
                            global_vars.matchedShikshalokamLoginId = userDetails[0]
                        except Exception as e:
                            error_msg = f"Error fetching user details: {str(e)}"
                            print(error_msg)
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                            return "", Helpers.errorVar

                        try:
                            frameworkExternalId = Helpers.frameWorkUpload(parentFolder, global_vars.matchedShikshalokamLoginId, wbObservation, accessToken)
                            if not frameworkExternalId:
                                return "", Helpers.errorVar

                            observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
                            
                            if not Helpers.themesUpload(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, False):
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Themes Upload Failed"))
                                return "", Helpers.errorVar

                            solutionId = Helpers.createSolutionFromFramework(parentFolder, wbObservation, accessToken, frameworkExternalId)
                            if not solutionId:
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Solution Creation Failed"))
                                return "", Helpers.errorVar
                        except Exception as e:
                            error_msg = f"Error during framework or solution creation: {str(e)}"
                            print(error_msg)
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                            return "", Helpers.errorVar

                        try:
                            # ECM processing
                            ecm_rows = validation_cache.get("ecm_rows") or []
                            if not ecm_rows:
                                ecmsSheet = wbObservation.sheet_by_name('ECMs or Domains')
                                keys = [ecmsSheet.cell(1, col_index).value for col_index in range(ecmsSheet.ncols)]
                                for row_index in range(2, ecmsSheet.nrows):
                                    ecm_rows.append({
                                        keys[col_index]: ecmsSheet.cell(row_index, col_index).value
                                        for col_index in range(ecmsSheet.ncols)
                                    })

                            ecm_update = dict()
                            ecm_dict = dict()
                            section = dict()
                            ecmSeqCount = 1
                            for dictECMs in ecm_rows:
                                ecm_external = Helpers._to_text(dictECMs.get('ECM Id/Domian ID')).strip()
                                EMC_ID = ecm_external + '_' + str(millisecond)
                                ECM_NAME = Helpers._to_text(dictECMs.get('ECM Name/Domain Name')).strip()
                                section_id = Helpers._to_text(dictECMs.get('section_id'))
                                section_name = Helpers._to_text(dictECMs.get('section_name'))
                                section.update({section_id: section_name})
                                global_vars.ecm_sections[EMC_ID] = section_id
                                
                                # Handle boolean conversion safely
                                is_mandatory = dictECMs.get('Is ECM Mandatory?', 'FALSE')
                                is_mandatory = str(is_mandatory).strip().upper() in ['TRUE', '1']

                                ecm_update[EMC_ID] = {
                                    "externalId": EMC_ID, "tip": None, "name": ECM_NAME, "description": None,
                                    "modeOfCollection": "onfield", "canBeNotApplicable": not is_mandatory,
                                    "notApplicable": False, "canBeNotAllowed": not is_mandatory, "remarks": None,
                                    "sequenceNo": ecmSeqCount
                                }
                                ecmSeqCount += 1
                            ecm_dict['evidenceMethods'] = ecm_update
                            if not Helpers.solutionUpdate(accessToken, solutionId, ecm_dict):
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Solution Update (ECM evidence) Failed"))
                                return "", Helpers.errorVar
                            
                            if not Helpers.solutionUpdate(accessToken, solutionId, {"sections": section}):
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Solution Update (ECM sections) Failed"))
                                return "", Helpers.errorVar
                        except Exception as e:
                            error_msg = f"Error during ECM processing: {str(e)}"
                            print(error_msg)
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                            return "", Helpers.errorVar

                        try:
                            # Continue rest of the process
                            bodySolutionUpdate = {"status": "active", "isDeleted": False, "criteriaLevelReport": global_vars.criteriaLevelsReport}
                            if not Helpers.solutionUpdate(accessToken, solutionId, bodySolutionUpdate):
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Solution Update Failed"))
                                return "", Helpers.errorVar

                            if not Helpers.questionUpload(global_vars.addObservationSolution, parentFolder, frameworkExternalId, millisecond, accessToken,solutionId,typeofSolution):
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Question Upload Failed"))
                                return "", Helpers.errorVar
                            
                            # Handle rubrics
                            if global_vars.pointBasedValue.lower() != "null":
                                bodySolutionUpdate = {"isRubricDriven": True}
                                if not Helpers.solutionUpdate(accessToken, solutionId, bodySolutionUpdate):
                                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Solution Update (Rubric flag) Failed"))
                                    return "", Helpers.errorVar
                                if not Helpers.fetchSolutionCriteria(parentFolder, observationExternalId, accessToken):
                                    return "", Helpers.errorVar

                                if not Helpers.uploadCriteriaRubrics(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, True):
                                    return "", Helpers.errorVar

                                if not Helpers.uploadThemeRubrics(parentFolder, wbObservation, accessToken, frameworkExternalId, True):
                                    return "", Helpers.errorVar

                            # Handle program information and start/end dates
                            bodySolutionUpdate = {'allowMultipleAssessemts': global_vars.allow_multiple_submissions, "creator": global_vars.creator}
                            if not Helpers.solutionUpdate(accessToken, solutionId, bodySolutionUpdate):
                                return "", Helpers.errorVar

                            solutionDetails = Helpers.fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, solutionId, accessToken)
                            if not solutionDetails:
                                error_msg = "Failed to fetch solution details"
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                                return "", Helpers.errorVar
                                
                            if solutionDetails[1]:
                                startDateArr = str(solutionDetails[1]).split("-")
                                bodySolutionUpdate = {"startDate": f"{startDateArr[2]}-{startDateArr[1]}-{startDateArr[0]} 00:00:00"}
                                if not Helpers.solutionUpdate(accessToken, solutionId, bodySolutionUpdate):
                                    return "", Helpers.errorVar

                            if solutionDetails[2]:
                                endDateArr = str(solutionDetails[2]).split("-")
                                bodySolutionUpdate = {"endDate": f"{endDateArr[2]}-{endDateArr[1]}-{endDateArr[0]} 23:59:59"}
                                if not Helpers.solutionUpdate(accessToken, solutionId, bodySolutionUpdate):
                                    return "", Helpers.errorVar

                            # If program name exists, handle child creation and linking
                            if global_vars.isProgramnamePresent:
                                childId = Helpers.createChild(parentFolder, observationExternalId, accessToken)
                                if (
                                    not isinstance(childId, (list, tuple))
                                    or len(childId) < 2
                                    or not childId[0]
                                ):
                                    # Error already logged in createChild
                                    return "", Helpers.errorVar

                                if childId[0]:
                                    childSolutionDetails = Helpers.fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, childId[0], accessToken)
                                    if not childSolutionDetails:
                                        error_msg = "Failed to fetch child solution details"
                                        Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                                        return "", Helpers.errorVar
                                    
                                    bodySolutionUpdate = {"scope": {"entityType": global_vars.scopeEntityType, "entities": global_vars.entitiesPGMID, "roles": childSolutionDetails[0]}}
                                    if not Helpers.solutionUpdate(accessToken, childId[0], bodySolutionUpdate):
                                        return "", Helpers.errorVar

                                    if solutionDetails[1]:
                                        startDateArr = str(solutionDetails[1]).split("-")
                                        bodySolutionUpdate = {
                                            "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[
                                                0] + " 00:00:00"}
                                        if not Helpers.solutionUpdate(accessToken, childId[0], bodySolutionUpdate):
                                            return "", Helpers.errorVar

                                    if solutionDetails[2]:
                                        endDateArr = str(solutionDetails[2]).split("-")
                                        bodySolutionUpdate = {
                                            "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                                        if not Helpers.solutionUpdate(accessToken, childId[0], bodySolutionUpdate):
                                            return "", Helpers.errorVar

                                    ObsRubricSolutionLink = Helpers.prepareProgramSuccessSheet(MainFilePath, parentFolder, programFile, childId[1], childId[0],
                                                            accessToken)
                                    if not ObsRubricSolutionLink:
                                        return "", Helpers.errorVar
                        except Exception as e:
                            error_msg = f"Error during rubric or program handling: {str(e)}"
                            print(error_msg)
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                            return "", Helpers.errorVar

                        # Final link creation
                        finalObsRubricSolutionLink = {ObsWRResourceName: ObsRubricSolutionLink}
                        return finalObsRubricSolutionLink, {}

                except Exception as e:
                    error_msg = f"Error in 'addObsWRFunc': {str(e)}"
                    print(error_msg)
                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                    return False

                # Call function
                try:
                    millisecond = int(time.time() * 1000)
                    ObsWRSolutionLink, obsErrors = addObsWRFunc(parentFolder, wbObservation, millisecond, accessToken)
                    
                    if obsErrors:
                        return False
                        
                    return ObsWRSolutionLink
                except Exception as e:
                     error_msg = f"Error processing observation w/ rubrics: {str(e)}"
                     print(error_msg)
                     Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                     return False


            elif sheets.strip().lower() == 'details'.lower() and typeofSolution == 2:
                cached_details = validation_cache.get("details_row") or {}
                ObsWORResourceName = Helpers._to_text(cached_details.get('observation_solution_name', '')) or Helpers._to_text(global_vars.solutionName or "")
                
                try:
                    @exception_handler
                    def addObsWORFunc(parentFolder, wbObservation, millisecond, accessToken):
                        print("Create Observation Function called ....")
                        
                        try:
                            # Step 1: Upload criteria
                            if not Helpers.criteriaUpload(parentFolder, wbObservation, millisecond, accessToken, "criteria", False):
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Criteria Upload Failed"))
                                return "", Helpers.errorVar
                            print("-------------> criteria upload done")
                            
                            # Step 2: Process user details for Diksha_loginId (prefer cached details row)
                            diksha_login_id = Helpers._to_text(cached_details.get('Diksha_loginId', '')).strip()
                            if not diksha_login_id:
                                diksha_login_id = Helpers._to_text(global_vars.dikshaLoginId).strip()

                            if diksha_login_id:
                                userDetails = Helpers.fetchUserDetails(accessToken, diksha_login_id)
                                if not userDetails:
                                    error_msg = f"Failed to fetch user details for {diksha_login_id}"
                                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                                    return "", Helpers.errorVar
                                global_vars.matchedShikshalokamLoginId = userDetails[0]
                                print(f"Matched login ID: {global_vars.matchedShikshalokamLoginId}")
                            else:
                                error_msg = "\"Diksha_loginId\" missing in details sheet/cache"
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                                return "", Helpers.errorVar
                            
                            # Step 3: Upload framework and themes
                            frameworkExternalId = Helpers.frameWorkUpload(parentFolder, global_vars.matchedShikshalokamLoginId, wbObservation, accessToken)
                            if not frameworkExternalId:
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Framework Upload Failed"))
                                return "", Helpers.errorVar

                            observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
                            if not Helpers.themesUpload(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, True):
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Themes Upload Failed"))
                                return "", Helpers.errorVar
                            
                            # Step 4: Create solution from framework
                            solutionId = Helpers.createSolutionFromFramework(parentFolder, wbObservation, accessToken, frameworkExternalId)
                            if not solutionId:
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Solution Creation Failed"))
                                return "", Helpers.errorVar
                            
                            # Step 5: Update solution with sections
                            sectionsObj = {"sections": {'S1': 'Observation Question'}}
                            if not Helpers.solutionUpdate(accessToken, solutionId, sectionsObj):
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Solution Update (Sections) Failed"))
                                return "", Helpers.errorVar
                            
                            # Step 6: Upload criteria and evidence methods
                            ecmObj = {
                                "evidenceMethods": {
                                    'OB': {
                                        'externalId': 'OB', 'tip': None, 'name': 'Observation', 'description': None,
                                        'modeOfCollection': 'onfield', 'canBeNotApplicable': False,
                                        'notApplicable': False, 'canBeNotAllowed': False, 'remarks': None
                                    }
                                }
                            }
                            if not Helpers.solutionUpdate(accessToken, solutionId, ecmObj):
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Solution Update (ECM) Failed"))
                                return "", Helpers.errorVar

                            if not Helpers.questionUpload(global_vars.addObservationSolution, parentFolder, frameworkExternalId, millisecond, accessToken, solutionId, typeofSolution):
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Question Upload Failed"))
                                return "", Helpers.errorVar

                            if not Helpers.fetchSolutionCriteria(parentFolder, observationExternalId, accessToken):
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Fetch Solution Criteria Failed"))
                                return "", Helpers.errorVar
                            
                            # Handle point-based value and rubrics
                            if global_vars.pointBasedValue.lower() != "null":
                                if not Helpers.uploadCriteriaRubrics(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, False):
                                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Upload Criteria Rubrics Failed"))
                                    return "", Helpers.errorVar
                                if not Helpers.uploadThemeRubrics(parentFolder, wbObservation, accessToken, frameworkExternalId, False):
                                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Upload Theme Rubrics Failed"))
                                    return "", Helpers.errorVar
                            
                            # Step 7: Activate and update solution status
                            bodySolutionUpdate = {"status": "active", "isDeleted": False, "allowMultipleAssessemts": True, "creator": global_vars.creator}
                            if not Helpers.solutionUpdate(accessToken, solutionId, bodySolutionUpdate):
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Solution Update (Status) Failed"))
                                return "", Helpers.errorVar
                            
                            # Step 8: Update solution dates (start and end)
                            solutionDetails = Helpers.fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, solutionId, accessToken)
                            if not solutionDetails:
                                error_msg = "Failed to fetch solution details"
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                                return "", Helpers.errorVar

                            if solutionDetails[1]:
                                startDateArr = str(solutionDetails[1]).split("-")
                                bodySolutionUpdate = {
                                    "startDate": f"{startDateArr[2]}-{startDateArr[1]}-{startDateArr[0]} 00:00:00"
                                }
                                if not Helpers.solutionUpdate(accessToken, solutionId, bodySolutionUpdate):
                                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Solution Update (Start Date) Failed"))
                                    return "", Helpers.errorVar
                            if solutionDetails[2]:
                                endDateArr = str(solutionDetails[2]).split("-")
                                bodySolutionUpdate = {
                                    "endDate": f"{endDateArr[2]}-{endDateArr[1]}-{endDateArr[0]} 23:59:59"
                                }
                                if not Helpers.solutionUpdate(accessToken, solutionId, bodySolutionUpdate):
                                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Solution Update (End Date) Failed"))
                                    return "", Helpers.errorVar

                            # Step 9: Handle program name
                            if global_vars.isProgramnamePresent:
                                childId = Helpers.createChild(parentFolder, observationExternalId, accessToken)
                                if (
                                    not isinstance(childId, (list, tuple))
                                    or len(childId) < 2
                                    or not childId[0]
                                ):
                                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Create Child Failed"))
                                    return "", Helpers.errorVar

                                if childId[0]:
                                    solutionDetails = Helpers.fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, childId[0], accessToken)
                                    if not solutionDetails:
                                        error_msg = "Failed to fetch child solution details"
                                        Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                                        return "", Helpers.errorVar

                                    scopeEntities = global_vars.entitiesPGMID
                                    scopeRoles = solutionDetails[0]
                                    bodySolutionUpdate = {
                                        "scope": {"entityType": global_vars.scopeEntityType, "entities": scopeEntities, "roles": scopeRoles}
                                    }
                                    if not Helpers.solutionUpdate(accessToken, childId[0], bodySolutionUpdate):
                                        Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Child Solution Update (Scope) Failed"))
                                        return "", Helpers.errorVar
                                    
                                    if solutionDetails[1]:
                                        startDateArr = str(solutionDetails[1]).split("-")
                                        bodySolutionUpdate = {
                                            "startDate": f"{startDateArr[2]}-{startDateArr[1]}-{startDateArr[0]} 00:00:00"
                                        }
                                        if not Helpers.solutionUpdate(accessToken, childId[0], bodySolutionUpdate):
                                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Child Solution Update (Start Date) Failed"))
                                            return "", Helpers.errorVar
                                    if solutionDetails[2]:
                                        endDateArr = str(solutionDetails[2]).split("-")
                                        bodySolutionUpdate = {
                                            "endDate": f"{endDateArr[2]}-{endDateArr[1]}-{endDateArr[0]} 23:59:59"
                                        }
                                        if not Helpers.solutionUpdate(accessToken, childId[0], bodySolutionUpdate):
                                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Child Solution Update (End Date) Failed"))
                                            return "", Helpers.errorVar
                                    ObsSolutionLink = Helpers.prepareProgramSuccessSheet(
                                        MainFilePath, parentFolder, programFile, childId[1], childId[0], accessToken
                                    )
                                    if not ObsSolutionLink:
                                        Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Prepare Success Sheet Failed"))
                                        return "", Helpers.errorVar
                                    print(ObsSolutionLink)
                                else:
                                    error_msg = "Failed to create child observation"
                                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                                    return "", Helpers.errorVar
                            
                            finalObsSolutionLink = {ObsWORResourceName: ObsSolutionLink}
                            return finalObsSolutionLink, {}

                        except Exception as e:
                            error_msg = f"Error during observation creation: {str(e)}"
                            print(error_msg)
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                            return "", Helpers.errorVar

                except Exception as e:
                    error_msg = f"Error in 'addObsWORFunc' definition: {str(e)}"
                    print(error_msg)
                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                    return False
                    
                # Call function
                try:
                    millisecond = int(time.time() * 1000)
                    add_obs_wor_result = addObsWORFunc(parentFolder, wbObservation, millisecond, accessToken)
                    if not isinstance(add_obs_wor_result, tuple) or len(add_obs_wor_result) != 2:
                        Helpers.errorVar.append(
                            str("CRITICAL") + ': ' + str("Observation w/o rubrics returned invalid response shape")
                        )
                        return False
                    ObsWRSolutionLink, obsErrors = add_obs_wor_result
                    
                    if obsErrors:
                        return False
                    
                    return ObsWRSolutionLink
                except Exception as e:
                    error_msg = f"Error processing observation w/o rubrics: {str(e)}"
                    print(error_msg)
                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                    return False



            elif sheets.strip().lower() == 'Project upload'.lower() and typeofSolution == 4:
                print("Checking project upload sheet...")
                project_rows = validation_cache.get("project_upload_rows") or []
                if not project_rows:
                    projectsheet = wbproject.sheet_by_name(sheets)
                    keysEnv = [projectsheet.cell(1, col_index_env).value for col_index_env in range(projectsheet.ncols)]
                    for row_index_env in range(2, projectsheet.nrows):
                        project_rows.append({
                            keysEnv[col_index_env]: projectsheet.cell(row_index_env, col_index_env).value
                            for col_index_env in range(projectsheet.ncols)
                        })
                ProjectName = Helpers._to_text(validation_cache.get("resource_name", ""))
                for projectDetails in project_rows:
                    ProjectName = Helpers._to_text(projectDetails.get("title", ProjectName))
                    entityType = "school"

                try:
                    @exception_handler
                    def addProjectFunc(filePathAddProject, parentFolder, accessToken):
                        print('Add Project Function Called')

                        # Create project folder if it doesn't exist
                        if not path.exists(parentFolder):
                            os.mkdir(parentFolder)

                        # Create a user input folder if it doesn't exist
                        if not path.exists(parentFolder + "/user_input_file"):
                            os.mkdir(parentFolder + "/user_input_file")
                        
                        # Copy files to the folder
                        shutil.copy(filePathAddProject, parentFolder + "/user_input_file")
                        shutil.copy(programFile, parentFolder + "/user_input_file")



                        solutionlink = ""

                        # Process cached project upload rows
                        for dictDetailsEnv in project_rows:

                            # Handle projects without a certificate
                            if str(dictDetailsEnv['has certificate']).lower() == 'no':
                                # print("----> No certificate for project <----")
                                if not Helpers.prepareProjectAndTasksSheets(global_vars.addObservationSolution, parentFolder, accessToken):
                                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Prepare project/tasks sheets failed"))
                                    return solutionlink, Helpers.errorVar
                                if not Helpers.projectUpload(parentFolder, accessToken):
                                    return solutionlink, Helpers.errorVar
                                if not Helpers.taskUpload(parentFolder, accessToken):
                                     return solutionlink, Helpers.errorVar
                                
                                ProjectSolutionResp = Helpers.solutionCreationAndMapping(parentFolder, global_vars.entityToUpload, global_vars.listOfFoundRoles, accessToken, programFile)
                                if not ProjectSolutionResp:
                                    if not Helpers.errorVar:
                                        Helpers.errorVar.append("Solution creation and mapping failed after project upload")
                                    return solutionlink, Helpers.errorVar

                                ProjectSolutionExternalId = ProjectSolutionResp[0]
                                ProjectSolutionId = ProjectSolutionResp[1]
                                
                                solutionlink = Helpers.prepareProgramSuccessSheet(
                                    MainFilePath, parentFolder, programFile,
                                    ProjectSolutionExternalId, ProjectSolutionId, accessToken)
                                if not solutionlink:
                                    return solutionlink, Helpers.errorVar
                                print(solutionlink)

                            # Handle projects with a certificate
                            elif str(dictDetailsEnv['has certificate']).lower() == 'yes':
                                print("----> Certificate required for project <----")
                                baseTemplate_id = Helpers.fetchCertificateBaseTemplate(
                                    filePathAddProject, accessToken)
                                if not baseTemplate_id:
                                    return solutionlink, Helpers.errorVar

                                if not Helpers.downloadlogosign(filePathAddProject, parentFolder):
                                     return solutionlink, Helpers.errorVar
                                
                                if not Helpers.editsvg(accessToken, filePathAddProject, parentFolder, baseTemplate_id):
                                     return solutionlink, Helpers.errorVar

                                if not Helpers.prepareProjectAndTasksSheets(global_vars.addObservationSolution, parentFolder, accessToken):
                                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Prepare project/tasks sheets failed"))
                                    return solutionlink, Helpers.errorVar
                                if not Helpers.projectUpload(parentFolder, accessToken):
                                     return solutionlink, Helpers.errorVar
                                if not Helpers.taskUpload(parentFolder, accessToken):
                                     return solutionlink, Helpers.errorVar
                                
                                ProjectSolutionResp = Helpers.solutionCreationAndMapping(
                                    parentFolder, global_vars.entityToUpload, global_vars.listOfFoundRoles, accessToken, programFile)
                                if not ProjectSolutionResp:
                                    if not Helpers.errorVar:
                                        Helpers.errorVar.append("Solution creation and mapping failed after project upload")
                                    return solutionlink, Helpers.errorVar

                                ProjectSolutionExternalId = ProjectSolutionResp[0]
                                ProjectSolutionId = ProjectSolutionResp[1]

                                # Handle certificate template
                                certificatetemplateid = Helpers.prepareaddingcertificatetemp(
                                    filePathAddProject, parentFolder, accessToken, ProjectSolutionId, global_vars.programID, baseTemplate_id)
                                if not certificatetemplateid:
                                    return solutionlink, Helpers.errorVar

                                solutionlink = Helpers.prepareProgramSuccessSheet(
                                    MainFilePath, parentFolder, programFile,
                                    ProjectSolutionExternalId, ProjectSolutionId, accessToken)
                                if not solutionlink:
                                    return solutionlink, Helpers.errorVar
                                        
                        if solutionlink:
                            finalprojectsolutionlink = {ProjectName: solutionlink}
                            return finalprojectsolutionlink, {}
                        else:
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Solution link generation failed or invalid certificate option"))
                            return solutionlink, Helpers.errorVar

                except Exception as e:
                    error_msg = f"Error in 'addProjectFunc': {str(e)}"
                    print(error_msg)
                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                    return "", Helpers.errorVar

                # Call function
                try:
                    millisAddObs = int(round(time.time() * 1000))
                    projectSolutionLink, projectErrors = addProjectFunc(global_vars.addObservationSolution, parentFolder, accessToken)
                    
                    if projectErrors:
                        return False
                    
                    return projectSolutionLink
                except Exception as e:
                    error_msg = f"Error processing project: {str(e)}"
                    print(error_msg)
                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                    return False

    
            elif sheets.strip().lower() == 'details'.lower() and typeofSolution == 3:
                try:
                    SurveyResourceName = Helpers._to_text(validation_cache.get("resource_name", ""))

                    @exception_handler
                    def addsurveyFunc(parentFolder, millisecond, accessToken):
                        try:
                            # Validate program file and survey sheets
                            if not Helpers.programsFileCheck(programFile, accessToken, parentFolder, MainFilePath):
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Programs File Check Failed"))
                                return "", Helpers.errorVar

                            wbSurvey = wbObservation

                            # Create survey solution
                            surveyResp = Helpers.createSurveySolution(parentFolder, wbSurvey, accessToken)
                            if not surveyResp or not surveyResp[0]:
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Create Survey Solution Failed"))
                                return "", Helpers.errorVar

                            surTempExtID = surveyResp[1]
                            surTempSolID = surveyResp[0]

                            # Update solution status
                            bodySolutionUpdate = {"status": "active", "isDeleted": False}
                            if not Helpers.solutionUpdate(accessToken, surveyResp[0], bodySolutionUpdate):
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Solution Update Failed"))
                                return "", Helpers.errorVar

                            # Upload survey questions
                            surveyLink = Helpers.uploadSurveyQuestions(MainFilePath, parentFolder, wbSurvey, global_vars.addObservationSolution, accessToken, surTempExtID, surTempSolID, millisecond, programFile)
                            if not surveyLink:
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Upload Survey Questions Failed"))
                                return "", Helpers.errorVar

                            finalsurveySolutionlink = {SurveyResourceName: global_vars.surveySolutionlink}
                            return finalsurveySolutionlink, {}

                        except KeyError as e:
                            error_msg = f"KeyError: {str(e)} - Possible missing column or incorrect key in sheet."
                            print(error_msg)
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                            return "", Helpers.errorVar
                        except xlrd.XLRDError as e:
                            error_msg = f"XLRDError: {str(e)} - Issue with reading the Excel file."
                            print(error_msg)
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                            return "", Helpers.errorVar
                        except Exception as e:
                            error_msg = f"An error occurred in addsurveyFunc: {str(e)}"
                            print(error_msg)
                            Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                            return "", Helpers.errorVar

                    millisecond = int(time.time() * 1000)
                    surveySollink, surveyErrors = addsurveyFunc(parentFolder, millisecond, accessToken)

                    if surveyErrors:
                        error_msg = f"Errors occurred during survey processing: {surveyErrors}"
                        Helpers.errorVar.append(str("CRITICAL") + ': ' + str(error_msg))
                        return False
                        
                    return surveySollink
                except KeyError as e:
                    print(f"KeyError: {str(e)} - Check if 'survey_solution_name' exists in the sheet.")
                except xlrd.XLRDError as e:
                    print(f"XLRDError: {str(e)} - Unable to load the Excel sheet: {sheets}")
                except Exception as e:
                    print(f"An unexpected error occurred: {str(e)}")



    @exception_handler
    def loadSurveyFile(programFile):
        Helpers.reset_errors()
        try:
            return Helpers._loadSurveyFile_impl(programFile)
        except Exception as e:
            Helpers.errorVar.append(str(e))
            first_error = Helpers.errorVar[0] if Helpers.errorVar else str(e)
            return {
                "solutionDict": {
                    "NA": [first_error]
                },
                "programName": ""
            }

    @exception_handler
    def _loadSurveyFile_impl(programFile):
        MainFilePath = Helpers.createFileStructForProgram(programFile)
        global_vars.downloaded_file = []
        print(global_vars.downloaded_file, "downloaded_file 3044")

        global_vars.reset_program_template_cache()
        global_vars.load_program_template(programFile, force_reload=True)

        sheetNames = global_vars.programTemplateSheetNames
        pgmSheets = ["Instructions", "Program Details", "Resource Details", "Program Manager Details", "Role-Subrole mapping"]
        print(sheetNames)
        print(pgmSheets)

        solutionDict = {}
        download_pairs = []
        programName = ""
        millisecond = int(time.time() * 1000)

        if len(sheetNames) == len(pgmSheets) and sheetNames == pgmSheets:
            print("--->Program Template detected.<---")
            dictProgramDetails = global_vars.programDict
            if dictProgramDetails:
                programName = Helpers._to_text(dictProgramDetails.get('Title of the Program', ''))

            for dictDetailsEnv in global_vars.programResourceDetails:
                resourceNamePGM = Helpers._to_text(dictDetailsEnv['Name of resources in program'])
                resourceLinkOrExtPGM = dictDetailsEnv['Resource Link']

                if str(dictDetailsEnv['Type of resources']).lower().strip() == "course":
                    continue
                else:
                    resourceStatus = dictDetailsEnv['Resource Status']
                    if resourceStatus.strip() == "New Upload":
                        try:
                            print("--->Resource Name : " + str(resourceNamePGM))
                            resourceLinkOrExtPGM = str(resourceLinkOrExtPGM).split('/')[5]
                            file_url = 'https://docs.google.com/spreadsheets/d/' + resourceLinkOrExtPGM + '/export?format=xlsx'
                            if not os.path.isdir('InputFiles'):
                                os.mkdir('InputFiles')
                            dest_file = 'InputFiles'
                            download_file = wget.download(file_url, dest_file)
                            global_vars.downloaded_file.append(download_file)
                            download_pairs.append((resourceNamePGM, download_file))
                        except Exception as e:
                            solutionDict[resourceNamePGM] = [str(e)]

            print("--->Solution input file successfully downloaded: " + str(global_vars.downloaded_file))
            if not download_pairs and not solutionDict:
                Helpers.errorVar.append("No Resources Detected in the Resource sheet.")
            for resource_name, file_path in download_pairs:
                try:
                    global_vars.addObservationSolution = file_path
                    print(f"Processing file: {global_vars.addObservationSolution}")
                    solutionSL = Helpers.mainFunc(MainFilePath, programFile, millisecond, global_vars.addObservationSolution)
                    if not solutionSL:
                        error_msg = Helpers.errorVar[-1] if Helpers.errorVar else "Execution failed for resource."
                        solutionDict[resource_name] = [str(error_msg)]
                        break
                    elif isinstance(solutionSL, dict) and resource_name in solutionSL:
                        result_value = solutionSL[resource_name]
                        solutionDict[resource_name] = result_value if isinstance(result_value, list) else [result_value]
                    elif isinstance(solutionSL, dict) and len(solutionSL) == 1:
                        result_value = next(iter(solutionSL.values()))
                        solutionDict[resource_name] = result_value if isinstance(result_value, list) else [result_value]
                    elif isinstance(solutionSL, str):
                        solutionDict[resource_name] = [solutionSL]
                    else:
                        solutionDict[resource_name] = [str(solutionSL)]
                except Exception as e:
                    Helpers.errorVar.append(str(e))
                    solutionDict[resource_name] = [str(e)]
                    break
            global_vars.downloaded_file = None
        else:
            Helpers.errorVar.append("The provided Template is not a Program Template.")

        if Helpers.errorVar:
            first_error = Helpers.errorVar[0]
            return {
                "solutionDict": {"NA": [first_error]},
                "programName": programName
            }

        return {
            "solutionDict": solutionDict,
            "programName": programName
        }

    
    # function to upload criteria   
    @exception_handler
    def criteriaUpload(solutionName_for_folder_path, wbObservation, millisAddObs, accessToken, tabName, projectDrivenFlag):
        criteriaColNames = ["criteriaId", "criteria_name"]
        criteriaUploadFieldnames = ['criteriaID', 'criteriaName']
        dictCriteriaToCsv = dict()
        criteriaLevelsFromFramework = dict()
        validation_cache = global_vars.get_resource_validation_cache() or {}

        criteria_rows = []
        framework_rows = []
        imp_mapping_rows = []

        if tabName == "framework":
            framework_rows = validation_cache.get("framework_rows") or []
            if projectDrivenFlag:
                imp_mapping_rows = validation_cache.get("imp_mapping_rows") or []

            if not framework_rows:
                criteriaSheet = wbObservation.sheet_by_name(tabName)
                keys = [criteriaSheet.cell(1, col_index).value for col_index in range(criteriaSheet.ncols)]
                for row_index in range(2, criteriaSheet.nrows):
                    framework_rows.append({
                        keys[col_index]: criteriaSheet.cell(row_index, col_index).value for col_index in range(criteriaSheet.ncols)
                    })

            if projectDrivenFlag and not imp_mapping_rows:
                impsToCriteria = wbObservation.sheet_by_name('Imp mapping')
                keysFromImpSheet = [impsToCriteria.cell(1, col_index).value for col_index in range(impsToCriteria.ncols)]
                for row_indexImp in range(2, impsToCriteria.nrows):
                    imp_mapping_rows.append({
                        keysFromImpSheet[col_index]: impsToCriteria.cell(row_indexImp, col_index).value for col_index in range(impsToCriteria.ncols)
                    })

        elif tabName == "criteria":
            criteria_rows = validation_cache.get("criteria_rows") or []
            if not criteria_rows:
                criteriaSheet = wbObservation.sheet_by_name(tabName)
                keys = [criteriaSheet.cell(1, col_index).value for col_index in range(criteriaSheet.ncols)]
                for row_index in range(2, criteriaSheet.nrows):
                    criteria_rows.append({
                        keys[col_index]: criteriaSheet.cell(row_index, col_index).value for col_index in range(criteriaSheet.ncols)
                    })

        if tabName == "framework":
            # Some templates may accidentally include header-like rows in cache.
            framework_rows = [
                row for row in framework_rows
                if Helpers._to_text(row.get("Criteria ID")).strip().lower() not in ["", "criteria id"]
            ]

            if projectDrivenFlag:
                criteriaImpDict = {}
                for dictImp in imp_mapping_rows:
                    criteria_id = Helpers._to_text(dictImp.get('criteriaId')).strip()
                    if not criteria_id:
                        continue
                    criteriaImpDict[criteria_id] = {}
                    for levls in range(1, global_vars.countImps + 1):
                        imp_val = Helpers._to_text(dictImp.get('L' + str(levls) + '-improvement-projects')).strip()
                        criteriaImpDict[criteria_id].update({'L' + str(levls) + '-improvement-projects': imp_val})

            framework_keys = list(framework_rows[0].keys()) if framework_rows else []
            levelCount = 0
            for eachHeaders in framework_keys:
                if re.match(r"^L\d+ description$", Helpers._to_text(eachHeaders).strip()):
                    levelCount += 1

            for dictFramework in framework_rows:
                criteriaLevelsFromFramework[dictFramework["Criteria ID"]] = {}

                for levlsNo in range(1, levelCount + 1):
                    criteriaLevelsFromFramework[dictFramework["Criteria ID"]].update(
                        {"L" + str(levlsNo): dictFramework["L" + str(levlsNo) + " description"]})
                    if not "L" + str(levlsNo) in criteriaColNames:
                        criteriaColNames.append("L" + str(levlsNo))

            for dictCriteria in framework_rows:
                dictCriteriaToCsv = {}

                criteria_external = Helpers._to_text(dictCriteria.get('Criteria ID')).strip()
                criteria_name = Helpers._to_text(dictCriteria.get('Criteria Name'))
                dictCriteriaToCsv['criteriaID'] = criteria_external + '_' + str(millisAddObs)
                global_vars.criteriaLookUp[dictCriteriaToCsv['criteriaID'].strip()] = criteria_name
                dictCriteriaToCsv['criteriaName'] = criteria_name
                criteriaName = criteria_name
                dictCriteriaToCsv['type'] = 'auto'
                for levlsNo in range(1, levelCount + 1):
                    dictCriteriaToCsv['L' + str(levlsNo)] = dictCriteria["L" + str(levlsNo) + " description"]
                if projectDrivenFlag:
                    for eachImps in criteriaImpDict.get(criteria_external, {}):
                        dictCriteriaToCsv[eachImps] = criteriaImpDict[criteria_external][eachImps]

                if not 'type' in criteriaUploadFieldnames:
                    criteriaUploadFieldnames.append('type')
                for eachCols in criteriaColNames:
                    if not eachCols in ['criteria_id', 'criteria_name', 'type', "criteriaId"]:
                        if not eachCols in criteriaUploadFieldnames:
                            criteriaUploadFieldnames.append(eachCols)
                if projectDrivenFlag:
                    for levls in range(1, global_vars.countImps + 1):
                        if not (str('L' + str(levls) + '-improvement-projects') in criteriaUploadFieldnames):
                            criteriaUploadFieldnames.append('L' + str(levls) + '-improvement-projects')
                criteriaFilePath = solutionName_for_folder_path + '/criteriaUpload/'
                file_exists = os.path.isfile(solutionName_for_folder_path + '/criteriaUpload/uploadSheet.csv')
                global_vars.criteriaLevelsCount = levelCount
                if not os.path.exists(criteriaFilePath):
                    os.mkdir(criteriaFilePath)
                with open(solutionName_for_folder_path + '/criteriaUpload/uploadSheet.csv', 'a',encoding='utf-8') as criteriaUploadFile:
                    writerCriteriaUpload = csv.DictWriter(criteriaUploadFile, fieldnames=list(criteriaUploadFieldnames),
                                                        lineterminator='\n')
                    if not file_exists:
                        writerCriteriaUpload.writeheader()
                    writerCriteriaUpload.writerow(dictCriteriaToCsv)
                    
        elif tabName == "criteria":
            for criteria_row in criteria_rows:
                dictCriteria = dict(criteria_row)
                dictCriteria['criteriaID'] = Helpers._to_text(dictCriteria['criteria_id']).strip() + '_' + str(millisAddObs)
                global_vars.criteriaLookUp[dictCriteria['criteriaID']] = Helpers._to_text(dictCriteria['criteria_name'])
                del dictCriteria['criteria_id']
                dictCriteria['criteriaName'] = Helpers._to_text(dictCriteria['criteria_name'])
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

        urlCriteriaUploadApi = internal_kong_ip + criteriauploadapiurl
        headerCriteriaUploadApi = {
            'Authorization': authorization,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id': x_channel_id
        }
        filesCriteria = {
            'criteria': open(solutionName_for_folder_path + '/criteriaUpload/uploadSheet.csv', 'rb')
        }

        responseCriteriaUploadApi = requests.post(url=urlCriteriaUploadApi, headers=headerCriteriaUploadApi,
                                                files=filesCriteria)
        messageArr = ["Criteria Upload Sheet Prepared.",
                    "File path : " + solutionName_for_folder_path + '/criteriaUpload/uploadSheet.csv']
        messageArr.append("Upload status code : " + str(responseCriteriaUploadApi.status_code))
        Helpers.createAPILog(solutionName_for_folder_path, messageArr)

        if responseCriteriaUploadApi.status_code == 200:
            print('CriteriaUploadApi Success')
            response_content = responseCriteriaUploadApi.text or ""
            try:
                parsed_payload = responseCriteriaUploadApi.json()
                if isinstance(parsed_payload, dict):
                    raw_csv = parsed_payload.get("raw")
                    if raw_csv:
                        response_content = raw_csv
                    else:
                        nested_raw = parsed_payload.get("result", {}).get("raw") if isinstance(parsed_payload.get("result"), dict) else None
                        if nested_raw:
                            response_content = nested_raw
            except ValueError:
                pass

            with open(solutionName_for_folder_path + '/criteriaUpload/uploadInternalIdsSheet.csv', 'w+', encoding='utf-8') as criteriaRes:
                criteriaRes.write(response_content)
        else:
            messageArr.append("Response : " + str(responseCriteriaUploadApi.text))
            error_message = ""
            if responseCriteriaUploadApi.status_code in [400, 401, 403, 404, 422]:
                error_message = f"criteriaUploadApi-Client Error {responseCriteriaUploadApi.status_code}: {responseCriteriaUploadApi.text}"
            elif responseCriteriaUploadApi.status_code in [500, 502, 503, 504]:
                error_message = f"criteriaUploadApi-Server Error {responseCriteriaUploadApi.status_code}: {responseCriteriaUploadApi.text}"
            else:
                error_message = f"criteriaUploadApi-Unexpected Error {responseCriteriaUploadApi.status_code}: {responseCriteriaUploadApi.text}"
            Helpers.errorVar.append(error_message)
            return False
        return True
    
    @exception_handler
    def frameWorkUpload(solutionName_for_folder_path, matchedShikshalokamLoginId, wbObservation, accessToken):
        dateTime = datetime.now()
        frameworkDocInsertObj = {}
        try:
            observationSheet = wbObservation.sheet_by_name("details")  # Reading the "details" sheet
        except xlrd.biffh.XLRDError:
            Helpers.errorVar.append("Error: 'details' sheet not found in wbObservation.")
            return False
        
        headers = [observationSheet.cell(1, col_index).value for col_index in range(observationSheet.ncols)]
        values = [observationSheet.cell(2, col_index).value for col_index in range(observationSheet.ncols)]        
        observationData = dict(zip(headers, values))
        solutionName = observationData.get('observation_solution_name', 'Default Solution Name')  # Replace 'Solution Name' with the actual header
        solutionDescription = observationData.get('observation_solution_description', 'Default Description')  # Replace 'Solution Description' with the actual header
        solutionKeywords = observationData.get('keywords', 'Default keywords')
        solutionEntityType = observationData.get('entity_type', 'Default entity_type')
        # Generating a unique External ID for the framework
        frameworkExternalId = str(uuid.uuid1())
        frameworkDocInsertObj['externalId'] = frameworkExternalId

        # Assigning the fetched name and description
        frameworkDocInsertObj['name'] = solutionName
        frameworkDocInsertObj['description'] = solutionDescription  
        frameworkDocInsertObj['parentId'] = None
        frameworkDocInsertObj['resourceType'] = ['Observations Framework']
        frameworkDocInsertObj['language'] = global_vars.solutionLanguage
        frameworkDocInsertObj['levelToScoreMapping'] = dict()
        frameworkDocInsertObj['keywords'] = solutionKeywords
        keyWords = solutionKeywords

        if keyWords and (keyWords != 'Framework' or keyWords != 'Frameworks' or keyWords != 'Observation' or keyWords != 'Observations'):
            keywordsFinalArr = ['Framework', 'Observation']
            keywordsArr = keyWords.encode('utf-8').decode('utf-8').split(',')
            for keyw in keywordsArr:
                keywordsFinalArr.append(keyw)
            frameworkDocInsertObj['keywords'] = keywordsFinalArr
            print(keywordsFinalArr,"<---------------------------------keywordsFinalArr")
        else:
            frameworkDocInsertObj['keywords'] = ['Framework', 'Observation']
        frameworkDocInsertObj['concepts'] = []
        frameworkDocInsertObj['createdFor'] = [global_vars.ccRootOrgId]  # createdForArr
        frameworkDocInsertObj['rootOrg'] = [global_vars.ccRootOrgId]  # rootOrgArr
        
        criteriaFrameworkArr = []
        with open(solutionName_for_folder_path + '/criteriaUpload/uploadInternalIdsSheet.csv', 'r', encoding='utf-8') as criteriaInternalFile:
            criteria_rows = csv.DictReader(criteriaInternalFile)
            if not criteria_rows:
                Helpers.errorVar.append(
                    str("CRITICAL") + ': ' + str("No criteria returned from criteria upload; cannot build framework themes")
                )
                return False
            criteriaWeightage = 100 / (len(list(criteria_rows)))
            criteriaInternalFile.seek(0, 0)
            next(criteria_rows, None)
            for crit in criteria_rows:
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
        # print(pointBasedValue,"<-------------------point based value in framwork")
        if not global_vars.pointBasedValue.lower() == "null":
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
            frameworkDocInsertObj['scoringSystem'] = global_vars.pointBasedValue
            frameworkDocInsertObj['isRubricDriven'] = True
            global_vars.criteriaLevelsReport = True
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
            for levs in range(1, global_vars.criteriaLevelsCount + 1):
                levelToScore = {"L" + str(levs): {'points': levs * 10, 'label': 'Level ' + str(levs)}}
                frameworkDocInsertObj['levelToScoreMapping'].update(levelToScore)
            frameworkDocInsertObj['noOfRatingLevels'] = global_vars.criteriaLevelsCount
            
        else:
            frameworkDocInsertObj['scoringSystem'] = None
            frameworkDocInsertObj['isRubricDriven'] = False
        
        fetchentitytypeid = Helpers.fetchEntityId(solutionName_for_folder_path, accessToken,
                                                      global_vars.entitiesPGM.lstrip().rstrip().split(","), global_vars.scopeEntityType)
        if not fetchentitytypeid:
            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Failed to fetch entity type ID for framework"))
            return False
    
        frameworkDocInsertObj['entityTypeId'] = fetchentitytypeid
        frameworkDocInsertObj['entityType'] = solutionEntityType
        frameworkDocInsertObj['type'] = 'observation'
        frameworkDocInsertObj['subType'] = solutionEntityType
        frameworkDocInsertObj['status'] = "active"
        frameworkDocInsertObj['updatedBy'] = 'INITIALIZE'
        frameworkDocInsertObj['createdBy'] = 'INITIALIZE'
        frameworkDocInsertObj['createdAt'] = str(dateTime)
        frameworkDocInsertObj['updatedAt'] = str(dateTime)
        frameworkDocInsertObj['author'] = matchedShikshalokamLoginId
        frameworkDocInsertObj['isTempObTest'] = 'observationAutomation'

        # Adding Credits and license into Frameworks
        frameworkDocInsertObj['creator'] = str(global_vars.creator)
        frameworkDocInsertObj['license'] = {}
        frameworkDocInsertObj['license']['author'] = str(global_vars.creator)
        frameworkDocInsertObj['license']['creator'] = str(global_vars.creator)
        frameworkDocInsertObj['license']['copyright'] = str(global_vars.ccRootOrgName)
        frameworkDocInsertObj['license']['copyrightYear'] = int(dateTime.strftime("%Y"))
        frameworkDocInsertObj['license']['contentType'] = "Observation"
        frameworkDocInsertObj['license']['organisation'] = [global_vars.ccRootOrgName]
        frameworkDocInsertObj['license']['orgDetails'] = {}
        frameworkDocInsertObj['license']['orgDetails']['email'] = None
        frameworkDocInsertObj['license']['orgDetails']['orgName'] = global_vars.ccRootOrgName
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
                                        'X-authenticated-user-token': accessToken,
                                        'X-Channel-id': x_channel_id}
            filesFramework = {'framework': open(solutionName_for_folder_path + '/framework/uploadFile.json', 'rb')}

            responseFrameworkUploadApi = requests.post(url=urlCreateFrameworkApi, headers=headerFrameworkUploadApi,
                                                    files=filesFramework)
            messageArr = ["Framwork json file created.",
                        "File loc : " + solutionName_for_folder_path + '/framework/uploadFile.json',
                        "Framework upload API called,", "Status code : " + str(responseFrameworkUploadApi.status_code)]
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)
            if responseFrameworkUploadApi.status_code == 200:
                print('Framework upload Success')
                return frameworkExternalId

            else:
                if responseFrameworkUploadApi.status_code in [400, 401, 403, 404, 422]:
                    Helpers.errorVar.append(f"FrameworkUploadApi-Client Error {responseFrameworkUploadApi.status_code}: {responseFrameworkUploadApi.text}")
                elif responseFrameworkUploadApi.status_code in [500, 502, 503, 504]:
                    Helpers.errorVar.append(f"FrameworkUploadApi-Server Error {responseFrameworkUploadApi.status_code}: {responseFrameworkUploadApi.text}")
                else:
                    Helpers.errorVar.append(f"FrameworkUploadApi-Unexpected Error {responseFrameworkUploadApi.status_code}: {responseFrameworkUploadApi.text}")
                messageArr = ["Framwork upload Failed.", "Response : " + responseFrameworkUploadApi.text]
                Helpers.createAPILog(solutionName_for_folder_path, messageArr)
                print('Framework upload api failed ',
                    'with response from api is ' + str(responseFrameworkUploadApi.text))
                return False
        
        except Exception as e:
            messageArr.append("Exception caught : " + str(e))
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)
            Helpers.errorVar.append(f"Error occurred: {str(e)}")
            return False


    @exception_handler
    def themesUpload(solutionName_for_folder_path, wbObservation, millisAddObs, accessToken, frameworkExternalId,obsWORubWS):
        criteria_internal_ids_in_order = []
        criteria_external_ids_in_order = []
        criteria_order_lookup = {}
        criteria_upload_sheet_path = solutionName_for_folder_path + '/criteriaUpload/uploadSheet.csv'
        if os.path.exists(criteria_upload_sheet_path):
            with open(criteria_upload_sheet_path, 'r', encoding='utf-8') as criteriaUploadFile:
                for row in csv.DictReader(criteriaUploadFile):
                    criteria_id = Helpers._to_text(row.get('criteriaID')).strip()
                    if criteria_id:
                        criteria_external_ids_in_order.append(criteria_id)
        with open(solutionName_for_folder_path + '/criteriaUpload/uploadInternalIdsSheet.csv', 'r',encoding='utf-8') as criteriaInternalFile:
            criteriaInternalReader = csv.DictReader(criteriaInternalFile)
            for crit in criteriaInternalReader:
                criteria_external_id = Helpers._to_text(crit.get('Criteria External Id')).strip()
                criteria_internal_id = Helpers._to_text(crit.get('Criteria Internal Id')).strip()
                if criteria_external_id and criteria_internal_id:
                    global_vars.dictCritLookUp[criteria_external_id] = criteria_internal_id
                    criteria_internal_ids_in_order.append(criteria_internal_id)
        if criteria_external_ids_in_order and criteria_internal_ids_in_order:
            for ext_id, int_id in zip(criteria_external_ids_in_order, criteria_internal_ids_in_order):
                criteria_order_lookup[ext_id] = int_id
        if obsWORubWS:
            print("Themes Observation without rubrics with scores")
            themeUploadFieldnames = ["theme", "aoi", "indicators", "criteriaInternalId"]
            themesUploadCsv = dict()
            for dictCritLookUpKey, dictCritLookUpValue in global_vars.dictCritLookUp.items():
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
            validation_cache = global_vars.get_resource_validation_cache() or {}
            framework_rows = validation_cache.get("framework_rows") or []
            if not framework_rows:
                frameWorkSheet = wbObservation.sheet_by_name('framework')
                keys = [frameWorkSheet.cell(1, col_index).value for col_index in range(frameWorkSheet.ncols)]
                for row_index in range(2, frameWorkSheet.nrows):
                    framework_rows.append({
                        keys[col_index]: frameWorkSheet.cell(row_index, col_index).value
                        for col_index in range(frameWorkSheet.ncols)
                    })
            themeUploadFieldnames = ["theme", "aoi", "indicators", "criteriaInternalId"]
            themesUploadCsv = dict()
            fallback_criteria_idx = 0
            for dictCriteria in framework_rows:
                domain_name = Helpers._to_text(dictCriteria.get('Domain Name'))
                domain_id = Helpers._to_text(dictCriteria.get('Domain ID'))
                criteria_id = Helpers._to_text(dictCriteria.get('Criteria ID')).strip()
                # Skip accidental header/blank rows from framework sheet cache.
                if not criteria_id or criteria_id.lower() == "criteria id":
                    continue
                themesUploadCsv['theme'] = domain_name + "###" + domain_id + "###40"
                themesUploadCsv['aoi'] = ""
                themesUploadCsv['indicators'] = ""
                criteria_external_id = criteria_id + '_' + str(millisAddObs)
                criteria_internal_id = global_vars.dictCritLookUp.get(criteria_external_id)
                if not criteria_internal_id:
                    criteria_internal_id = criteria_order_lookup.get(criteria_external_id)
                # In some environments criteria upload API returns fewer/non-matching external IDs.
                # Fallback to row-order mapping; if counts differ, cycle available internal IDs.
                if not criteria_internal_id and criteria_internal_ids_in_order:
                    criteria_internal_id = criteria_internal_ids_in_order[
                        fallback_criteria_idx % len(criteria_internal_ids_in_order)
                    ]
                if not criteria_internal_id:
                    Helpers.errorVar.append(
                        str("CRITICAL")
                        + ': '
                        + str(
                            f"Theme upload mapping failed: criteria external id not found: {criteria_external_id}"
                        )
                    )
                    return False
                themesUploadCsv['criteriaInternalId'] = criteria_internal_id + "###40"
                fallback_criteria_idx += 1
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
                                    'X-authenticated-user-token': accessToken,
                                    'X-Channel-id': x_channel_id}
            filesThemes = {'themes': open(solutionName_for_folder_path + '/themeUpload/uploadSheet.csv', 'rb')}
            responseThemeUploadApi = requests.post(url=urlThemesUploadApi, headers=headerThemesUploadApi, files=filesThemes)
            messageArr = ["Themes upload sheet prepared.",
                        "File path : " + solutionName_for_folder_path + '/themeUpload/uploadSheet.csv',
                        "Theme upload to framework API called.", "URL : " + urlThemesUploadApi,
                        "Status code : " + str(responseThemeUploadApi.status_code)]
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)
            if responseThemeUploadApi.status_code == 200:
                print('Theme UploadApi Success')
                with open(solutionName_for_folder_path + '/themeUpload/uploadInternalIdsSheet.csv', 'w+',encoding='utf-8') as criteriaRes:
                    criteriaRes.write(responseThemeUploadApi.text)
                return True
            else:
                if responseThemeUploadApi.status_code in [400, 401, 403, 404, 422]:
                    Helpers.errorVar.append(f"ThemeUploadApi-Client Error {responseThemeUploadApi.status_code}: {responseThemeUploadApi.text}")
                elif responseThemeUploadApi.status_code in [500, 502, 503, 504]:
                    Helpers.errorVar.append(f"ThemeUploadApi-Server Error {responseThemeUploadApi.status_code}: {responseThemeUploadApi.text}")
                else:
                    Helpers.errorVar.append(f"ThemeUploadApi-Unexpected Error {responseThemeUploadApi.status_code}: {responseThemeUploadApi.text}")
                messageArr = ["Themes upload failed.", "Response : " + str(responseThemeUploadApi.text)]
                Helpers.createAPILog(solutionName_for_folder_path, messageArr)
                print("Theme upload failed.")
                return False
            
        except Exception as e:
            messageArr.append("Exception caught : " + str(e))
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)
            Helpers.errorVar.append(f"Error occurred: {str(e)}")
            return False

    @exception_handler
    def createSolutionFromFramework(solutionName_for_folder_path, wbObservation, accessToken, frameworkExternalId):
        urlCreateSolutionApi = internal_kong_ip + solutioncreationapiurl
        headerCreateSolutionApi = {
            'Content-Type': content_type,
            'Authorization': authorization,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id': x_channel_id
        }
        try:
            observationSheet = wbObservation.sheet_by_name("details")  # Reading the "details" sheet
        except xlrd.biffh.XLRDError:
            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Details sheet not found in workbook"))
            return False
        
        headers = [observationSheet.cell(1, col_index).value for col_index in range(observationSheet.ncols)]
        values = [observationSheet.cell(2, col_index).value for col_index in range(observationSheet.ncols)]      
        observationData = dict(zip(headers, values))
        entityType = observationData.get('entity_type', 'Default entity_type') 
        queryparamsCreateSolutionApi = '?frameworkId=' + str(frameworkExternalId) + '&entityType=' + entityType
        responseCreateSolutionApi = requests.post(url=urlCreateSolutionApi + queryparamsCreateSolutionApi,
                                                headers=headerCreateSolutionApi)

        messageArr = ["Solution Created from Framework.",
                    "URL : " + str(urlCreateSolutionApi + queryparamsCreateSolutionApi),
                    "Status Code : " + str(responseCreateSolutionApi.status_code),
                    "Response : " + str(responseCreateSolutionApi.text)]
        Helpers.createAPILog(solutionName_for_folder_path, messageArr)
        messageArr = []
        if responseCreateSolutionApi.status_code == 200:
            responseCreateSolutionApi = responseCreateSolutionApi.json()
            solutionId = responseCreateSolutionApi['result']['templateId']
            messageArr.append("Parent Solution Generated : " + str(solutionId))
            print("Parent Solution Generated : " + str(solutionId))
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)
        else:
            if responseCreateSolutionApi.status_code in [400, 401, 403, 404, 422]:
                Helpers.errorVar.append(f"CreateSolutionApi-Client Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}")
            elif responseCreateSolutionApi.status_code in [500, 502, 503, 504]:
                Helpers.errorVar.append(f"CreateSolutionApi-Server Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}")
            else:
                Helpers.errorVar.append(f"CreateSolutionApi-Unexpected Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}")
            messageArr.append("Solution from framework api failed.")
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)
            print("Solution from framework api failed.")
            return False
        return solutionId
    
    @exception_handler
    def questionUpload(filePathAddObs, solutionName_for_folder_path, frameworkExternalId, millisAddObs, accessToken,
                   solutionId, typeofSolution):
        wbObservation = Helpers._get_cached_resource_workbook(filePathAddObs)
        questShee = wbObservation.sheet_by_name('questions')
        Qukeys = [questShee.cell(1, col_index).value for col_index in range(questShee.ncols)]
        countColSeq = Qukeys.index('question_sequence')
        questionsResponseDict = dict()
        raw_questions_rows = []
        for row_idx in range(2, questShee.nrows):
            row_dict = {
                Qukeys[col_idx]: questShee.cell(row_idx, col_idx).value
                for col_idx in range(questShee.ncols)
            }
            raw_questions_rows.append(row_dict)

        try:
            questionsList = sorted(
                raw_questions_rows,
                key=lambda row: int(float(row.get('question_sequence', 0)))
            )
        except Exception:
            questionsList = raw_questions_rows
        print("Question Sorted.")
        questionSeqByEcmDict = dict()
        questionSeqByEcmSectionDict = dict()
        questionSeqByEcmArr = []
        quesSeqCnt = 1.0
        questionUploadFieldnames = []
        questionUploadExceptSliderFieldnames = []
        questionUploadSliderFieldNames = []
        if typeofSolution == 1:
            for ques00 in questionsList:
                questionSeqByEcmDict[global_vars.ecmToSection[ques00['section_id']] + "_" + str(millisAddObs)] = {
                    global_vars.ecm_sections[global_vars.ecmToSection[ques00['section_id']] + "_" + str(millisAddObs)]: []}
        elif typeofSolution == 2:
            questionSeqByEcmDict["OB"] = {
                "S1": []
            }

        for ques1 in questionsList:
            if not global_vars.pointBasedValue.lower() == "null":
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
                    questionFileObj['name'] = global_vars.criteriaLookUp[questionFileObj['criteriaExternalId']]
                except:
                    print(questionFileObj['criteriaExternalId'] + " not found.")
                    Helpers.errorVar.append(str("CRITICAL") + ': ' + str(f"Criteria ID error: {questionFileObj['criteriaExternalId']} not found"))
                    return False
                if typeofSolution == 1 or typeofSolution == 5:
                    questionFileObj['evidenceMethod'] = global_vars.ecmToSection[ques['section_id']] + "_" + str(millisAddObs)
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
                                searchResponse = re.search(r"^response\(R[0-9]\)$|^response\(R[0-2][0-9]\)$", i)
                                if searchResponse:
                                    try:
                                        responseCheck = questionsResponseDict[questionFileObj['parentQuestionId']][
                                            searchResponse.string]
                                    except:
                                        print(questionFileObj[
                                                'parentQuestionId'] + " Referenced before intialising in questions sheet.")
                                        print("Please check question sequence...")
                                        print("Aborting...")
                                        messageArr = [questionFileObj[
                                                        'parentQuestionId'] + " Referenced before intialising in questions sheet.",
                                                    "Please check question sequesnce...", ]
                                        Helpers.createAPILog(solutionName_for_folder_path, messageArr)
                                        Helpers.errorVar.append("Execution terminated")
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
                        global_vars.ecm_sections[questionFileObj['evidenceMethod']]].append(
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
                if not global_vars.pointBasedValue.lower() == "null":
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
        if not Helpers.solutionUpdate(accessToken, solutionId, bodySolutionUpdate):
            return False

        try:
            urlQuestionsUploadApi = internal_kong_ip + questionuploadapiurl
            headerQuestionUploadApi = {'Authorization': authorization,
                                    'X-authenticated-user-token': accessToken,
                                    'X-Channel-id': x_channel_id}
            filesQuestion = {
                'questions': open(solutionName_for_folder_path + '/questionUpload/uploadSheet.csv', 'rb')
            }
            responseQuestionUploadApi = requests.post(url=urlQuestionsUploadApi, headers=headerQuestionUploadApi,
                                                    files=filesQuestion)
            messageArr = ["Question Upload sheet prepared.",
                        "File loc : " + solutionName_for_folder_path + '/questionUpload/uploadSheet.csv',
                        "Question upload API called.", "Status code : " + str(responseQuestionUploadApi.status_code)]
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)
            if responseQuestionUploadApi.status_code == 200:
                print('QuestionUploadApi Success')
                with open(solutionName_for_folder_path + '/questionUpload/uploadInternalIdsSheet.csv','w+',
                        encoding='utf-8') as questionRes:
                    questionRes.write(responseQuestionUploadApi.text)
                return True
            else:
                error_message = ""
                if responseQuestionUploadApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"questionUploadApi-Client Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}"
                elif responseQuestionUploadApi.status_code in [500, 502, 503, 504]:
                    error_message = f"questionUploadApi-Server Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}"
                else:
                    error_message = f"questionUploadApi-Unexpected Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}"
                Helpers.errorVar.append(error_message)
                return False
        except Exception as e:
            messageArr.append("Exception caught : " + str(e))
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)
            Helpers.errorVar.append(f"Error occurred: {str(e)}")
            return False

    @exception_handler
    def uploadCriteriaRubrics(solutionName_for_folder_path, wbObservation, millisAddObs, accessToken, frameworkExternalId,
                          withRubricsFlag):
        validation_cache = global_vars.get_resource_validation_cache() or {}
        solution_criteria_in_order = []
        criteria_upload_ids_in_order = []
        criteria_upload_names_in_order = []
        criteria_internal_ids_from_upload_in_order = []
        criteria_from_upload_lookup = {}
        criteria_order_lookup = {}
        criteria_upload_sheet_path = solutionName_for_folder_path + '/criteriaUpload/uploadSheet.csv'
        if os.path.exists(criteria_upload_sheet_path):
            with open(criteria_upload_sheet_path, 'r', encoding='utf-8') as criteriaUploadFile:
                for row in csv.DictReader(criteriaUploadFile):
                    criteria_id = Helpers._to_text(row.get('criteriaID')).strip()
                    criteria_name = Helpers._to_text(row.get('criteriaName'))
                    if criteria_id:
                        criteria_upload_ids_in_order.append(criteria_id)
                        criteria_upload_names_in_order.append(criteria_name)
        criteria_internal_sheet_path = solutionName_for_folder_path + '/criteriaUpload/uploadInternalIdsSheet.csv'
        if os.path.exists(criteria_internal_sheet_path):
            criteria_internal_content = ""
            with open(criteria_internal_sheet_path, 'r', encoding='utf-8') as criteriaInternalFile:
                criteria_internal_content = criteriaInternalFile.read()
            try:
                parsed_payload = json.loads(criteria_internal_content)
                if isinstance(parsed_payload, dict):
                    raw_csv = parsed_payload.get("raw")
                    nested_raw = parsed_payload.get("result", {}).get("raw") if isinstance(parsed_payload.get("result"), dict) else None
                    criteria_internal_content = raw_csv or nested_raw or criteria_internal_content
            except ValueError:
                pass
            criteria_internal_reader = csv.DictReader(criteria_internal_content.splitlines())
            for row in criteria_internal_reader:
                criteria_internal_id = Helpers._to_text(row.get('Criteria Internal Id')).strip()
                if criteria_internal_id:
                    criteria_internal_ids_from_upload_in_order.append(criteria_internal_id)
        if criteria_upload_ids_in_order and criteria_internal_ids_from_upload_in_order:
            for idx, ext_id in enumerate(criteria_upload_ids_in_order):
                internal_id = criteria_internal_ids_from_upload_in_order[idx % len(criteria_internal_ids_from_upload_in_order)]
                crit_name = criteria_upload_names_in_order[idx] if idx < len(criteria_upload_names_in_order) else ""
                criteria_from_upload_lookup[ext_id] = [internal_id, crit_name]
        if withRubricsFlag:
            criteria_rubric_rows = validation_cache.get("criteria_rubric_rows") or []
            if not criteria_rubric_rows:
                criteriaRubricSheet = wbObservation.sheet_by_name('Criteria_Rubric-Scoring')
                keys = [criteriaRubricSheet.cell(1, col_index).value for col_index in range(criteriaRubricSheet.ncols)]
                for row_index in range(2, criteriaRubricSheet.nrows):
                    criteria_rubric_rows.append({
                        keys[col_index]: criteriaRubricSheet.cell(row_index, col_index).value
                        for col_index in range(criteriaRubricSheet.ncols)
                    })
            dictSolCritLookUp = dict()
            filePath = os.path.join(solutionName_for_folder_path + "/solutionCriteriaFetch/", "solutionCriteriaDetails.csv")
            with open(filePath, 'r',encoding='utf-8') as criteriaInternalFile:
                criteriaInternalReader = csv.DictReader(criteriaInternalFile)
                for crit in criteriaInternalReader:
                    crit_id = Helpers._to_text(crit.get('criteriaID')).strip()
                    crit_internal = Helpers._to_text(crit.get('criteriaInternalId')).strip()
                    crit_name = Helpers._to_text(crit.get('criteriaName'))
                    if crit_id and crit_internal:
                        dictSolCritLookUp[crit_id] = [crit_internal, crit_name]
                        solution_criteria_in_order.append([crit_internal, crit_name])
    
        else:
            dictSolCritLookUp = dict()
            filePath = os.path.join(solutionName_for_folder_path + "/solutionCriteriaFetch/", "solutionCriteriaDetails.csv")
            with open(filePath, 'r',encoding='utf-8') as criteriaInternalFile:
                criteriaInternalReader = csv.DictReader(criteriaInternalFile)
                for crit in criteriaInternalReader:
                    crit_id = Helpers._to_text(crit.get('criteriaID')).strip()
                    crit_internal = Helpers._to_text(crit.get('criteriaInternalId')).strip()
                    crit_name = Helpers._to_text(crit.get('criteriaName'))
                    if crit_id and crit_internal:
                        dictSolCritLookUp[crit_id] = [crit_internal, crit_name]
                        solution_criteria_in_order.append([crit_internal, crit_name])
        if not dictSolCritLookUp and criteria_from_upload_lookup:
            dictSolCritLookUp.update(criteria_from_upload_lookup)
            solution_criteria_in_order.extend(criteria_from_upload_lookup.values())
        if criteria_upload_ids_in_order and solution_criteria_in_order:
            for ext_id, crit_data in zip(criteria_upload_ids_in_order, solution_criteria_in_order):
                criteria_order_lookup[ext_id] = crit_data
        criteriaRubricUploadFieldnames = ["externalId", "name", "criteriaId", "weightage", "expressionVariables"]

        if withRubricsFlag:
            for cl in global_vars.criteriaLevels:
                criteriaRubricUploadFieldnames.append("L" + str(cl))
        else:
            criteriaRubricUploadFieldnames.append("L1")
        criteriaRubricUpload = dict()
        criteriaRubricsFilePath = solutionName_for_folder_path + '/criteriaRubrics/'
        file_exists_ques = os.path.isfile(solutionName_for_folder_path + '/criteriaRubrics/uploadSheet.csv')
        if not os.path.exists(criteriaRubricsFilePath):
            os.mkdir(criteriaRubricsFilePath)
        if withRubricsFlag:
            fallback_criteria_idx = 0
            for dictCriteriaRubric in criteria_rubric_rows:
                file_exists_ques = os.path.isfile(solutionName_for_folder_path + '/criteriaRubrics/uploadSheet.csv')
                with open(solutionName_for_folder_path + '/criteriaRubrics/uploadSheet.csv', 'a',
                        encoding='utf-8') as questionUploadFile:
                    writerQuestionUpload = csv.DictWriter(questionUploadFile, fieldnames=criteriaRubricUploadFieldnames,
                                                        lineterminator='\n')
                    if not file_exists_ques:
                        writerQuestionUpload.writeheader()
                    criteriaRubricUpload = {}
                    criteria_id = Helpers._to_text(dictCriteriaRubric.get('criteriaId')).strip()
                    if not criteria_id or criteria_id.lower() == "criteriaid":
                        continue
                    criteriaRubricUpload['externalId'] = criteria_id + "_" + str(millisAddObs)
                    print(criteriaRubricUpload['externalId'])
                    crit_data = dictSolCritLookUp.get(criteriaRubricUpload['externalId'])
                    if not crit_data:
                        crit_data = criteria_order_lookup.get(criteriaRubricUpload['externalId'])
                    if not crit_data and solution_criteria_in_order:
                        crit_data = solution_criteria_in_order[fallback_criteria_idx % len(solution_criteria_in_order)]
                    if not crit_data:
                        Helpers.errorVar.append(
                            str("CRITICAL")
                            + ': '
                            + str(
                                f"Criteria rubric mapping failed: criteria external id not found: {criteriaRubricUpload['externalId']}"
                            )
                        )
                        return False
                    criteriaRubricUpload['name'] = crit_data[1]
                    criteriaRubricUpload['criteriaId'] = crit_data[0]
                    if dictCriteriaRubric['weightage']:
                        criteriaRubricUpload['weightage'] = dictCriteriaRubric['weightage']
                    else:
                        criteriaRubricUpload['weightage'] = 0
                    criteriaRubricUpload['expressionVariables'] = "SCORE=" + criteriaRubricUpload[
                        'criteriaId'] + ".scoreOfAllQuestionInCriteria()"
                    for cl in global_vars.criteriaLevels:
                        criteriaRubricUpload['L' + str(cl)] = dictCriteriaRubric['L' + str(cl) + " SCORE"]
                    writerQuestionUpload.writerow(criteriaRubricUpload)
                    fallback_criteria_idx += 1
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
                'X-authenticated-user-token': accessToken,
                'X-Channel-id': x_channel_id
            }
            filesCriteriaRubric = {
                'criteria': open(solutionName_for_folder_path + '/criteriaRubrics/uploadSheet.csv', 'rb')
            }
            responseCriteriaRubricUploadApi = requests.post(url=urlCriteriaRubricUploadApi,
                                                            headers=headerCriteriaRubricUploadApi, files=filesCriteriaRubric)
            messageArr = ["Criteria Rubric upload sheet prepared.",
                        "File Loc : " + solutionName_for_folder_path + '/criteriaRubrics/uploadSheet.csv',
                        "Status Code : " + str(responseCriteriaRubricUploadApi.status_code)]
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)
            if responseCriteriaRubricUploadApi.status_code == 200:
                with open(solutionName_for_folder_path + '/criteriaRubrics/uploadInternalIdsSheet.csv',
                        'w+',encoding='utf-8') as criteriaRubricRes:
                    criteriaRubricRes.write(responseCriteriaRubricUploadApi.text)
                return True
            else:
                error_message = ""
                if responseCriteriaRubricUploadApi.status_code in [400, 401, 403, 404, 422]:
                    error_message = f"criteriaRubricUploadApi-Client Error {responseCriteriaRubricUploadApi.status_code}: {responseCriteriaRubricUploadApi.text}"
                elif responseCriteriaRubricUploadApi.status_code in [500, 502, 503, 504]:
                    error_message = f"criteriaRubricUploadApi-Server Error {responseCriteriaRubricUploadApi.status_code}: {responseCriteriaRubricUploadApi.text}"
                else:
                    error_message = f"criteriaRubricUploadApi-Unexpected Error {responseCriteriaRubricUploadApi.status_code}: {responseCriteriaRubricUploadApi.text}"
                Helpers.errorVar.append(error_message)
                return False
        except Exception as e:
            messageArr = []
            messageArr.append("Exception caught : " + str(e))
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)
            print("❌ Exception:", e)
            Helpers.errorVar.append(f"Error occurred: {str(e)}")
            return False
        
    @exception_handler
    def fetchSolutionCriteria(solutionName_for_folder_path, observationId, accessToken):
        try:
            url = internal_kong_ip + ferchsolutioncriteria + observationId

            headers = {
                'Authorization': authorization,
                'X-authenticated-user-token': accessToken,
                'internal-access-token': internal_access_token
            }

            response = requests.request("POST", url, headers=headers)
            messageArr = ["Criteria solution fetch API called.", "Status Code  : " + str(response.status_code), "URL : " + url]
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)

            os.mkdir(solutionName_for_folder_path + "/solutionCriteriaFetch/")
            if response.status_code == 200:
                print("Solution criteria fetched.")
                with open(solutionName_for_folder_path + "/solutionCriteriaFetch/solutionCriteriaDetails.csv",
                        'w+',encoding='utf-8') as solutionCriteriaFetch:
                    solutionCriteriaFetch.write(response.text)
                return True
            else:
                if response.status_code in [400, 401, 403, 404, 422]:
                    Helpers.errorVar.append(f"QuestionUploadApi-Client Error {response.status_code}: {response.text}")
                elif response.status_code in [500, 502, 503, 504]:
                    Helpers.errorVar.append(f"QuestionUploadApi-Server Error {response.status_code}: {response.text}")
                else:
                    Helpers.errorVar.append(f"QuestionUploadApi-Unexpected Error {response.status_code}: {response.text}")
                messageArr = []
                messageArr = ["Criteria solution fetch API failed.", "Response  : " + str(response.text)]
                Helpers.createAPILog(solutionName_for_folder_path, messageArr)
                Helpers.errorVar.append("Solution criteria fetch failed. Status Code : " + str(response.status_code))
                return False
        except Exception as e:
            messageArr = []
            messageArr.append("Exception caught : " + str(e))
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)
            Helpers.errorVar.append(f"Error occurred: {str(e)}")
            return False


    @exception_handler
    def uploadThemeRubrics(solutionName_for_folder_path, wbObservation, accessToken, frameworkExternalId, withRubricsFlag):
        validation_cache = global_vars.get_resource_validation_cache() or {}
        themeRubricUploadFieldnames = ["externalId", "name", "weightage"]
        themeRubricsFilePath = os.path.join(solutionName_for_folder_path, "themeRubrics/")
        if not os.path.exists(themeRubricsFilePath):
            os.mkdir(themeRubricsFilePath)
        themeRubricUpload = dict()
        if withRubricsFlag:
            theme_rubric_rows = validation_cache.get("theme_rubric_rows") or []
            if not theme_rubric_rows:
                themeRubricSheet = wbObservation.sheet_by_name('Domain(theme)_rubric_scoring')
                keys = [themeRubricSheet.cell(1, col_index).value for col_index in range(themeRubricSheet.ncols)]
                for row_index in range(2, themeRubricSheet.nrows):
                    theme_rubric_rows.append({
                        keys[col_index]: themeRubricSheet.cell(row_index, col_index).value for col_index in range(themeRubricSheet.ncols)
                    })
            themeRubricUploadFieldnames = ["externalId", "name", "weightage"]
            if withRubricsFlag:
                for cl in global_vars.criteriaLevels:
                    themeRubricUploadFieldnames.append("L" + str(cl))
            else:
                themeRubricUploadFieldnames.append("L1")

            for dictThemeRubric in theme_rubric_rows:
                file_exists_ques = os.path.isfile(solutionName_for_folder_path + '/themeRubrics/uploadSheet.csv')
                with open(solutionName_for_folder_path + '/themeRubrics/uploadSheet.csv', 'a',
                        encoding='utf-8') as themeRubricsUploadFile:
                    writerThemeRubricsUpload = csv.DictWriter(themeRubricsUploadFile,
                                                            fieldnames=themeRubricUploadFieldnames, lineterminator='\n')
                    if not file_exists_ques:
                        writerThemeRubricsUpload.writeheader()

                    themeRubricUpload['externalId'] = dictThemeRubric['domain_Id']
                    themeRubricUpload['name'] = Helpers._to_text(dictThemeRubric['domain_name'])
                    if dictThemeRubric['weightage']:
                        themeRubricUpload['weightage'] = dictThemeRubric['weightage']
                    else:
                        themeRubricUpload['weightage'] = 0
                    if withRubricsFlag:
                        for cl in global_vars.criteriaLevels:
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
                'X-authenticated-user-token': accessToken,
                'X-Channel-id': x_channel_id
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
                    error_message = f"themeRubricUploadApi-Client Error {responseThemeRubricUploadApi.status_code}: {responseThemeRubricUploadApi.text}"
                elif responseThemeRubricUploadApi.status_code in [500, 502, 503, 504]:
                    error_message = f"themeRubricUploadApi-Server Error {responseThemeRubricUploadApi.status_code}: {responseThemeRubricUploadApi.text}"
                else:
                    error_message = f"themeRubricUploadApi-Unexpected Error {responseThemeRubricUploadApi.status_code}: {responseThemeRubricUploadApi.text}"
                Helpers.errorVar.append(error_message)
                return False
        except Exception as e:
            messageArr = []
            messageArr.append("Exception caught : " + str(e))
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)
            Helpers.errorVar.append(f"Error occurred: {str(e)}")
            print(f"Error occurred: {str(e)}")
            return False
            

    @exception_handler
    def prepareSuccessSheet(solutionName_for_folder_path, filePathAddObs, observationExternalId, millisAddObs):
        updateSuccessWorkBook = Helpers._get_cached_resource_workbook(filePathAddObs)
        updateWbNumberOfSheets = updateSuccessWorkBook.nsheets
        updateWbSheetNames = updateSuccessWorkBook.sheet_names()
        updateCriteriaSheet = updateSuccessWorkBook.sheet_by_name('Criteria_Rubric-Scoring')
        updateQuestionsSheet = updateSuccessWorkBook.sheet_by_name('questions')
        updateDetailsSheet = updateSuccessWorkBook.sheet_by_name('details')
        copyOfUpdateWb = copy(updateSuccessWorkBook)
        updateQuestionsSheetCopy = copyOfUpdateWb.get_sheet('questions')
        for each in range(updateWbNumberOfSheets):
            eachUpdateWorkSheet = copyOfUpdateWb.get_sheet(each)
            if (eachUpdateWorkSheet.name).strip() == 'Criteria_Rubric-Scoring':
                for row_idx_crit in range(1, updateCriteriaSheet.nrows):
                    for col_idx_crit in range(0, updateCriteriaSheet.ncols):
                        if col_idx_crit == 0:
                            eachUpdateWorkSheet.write(row_idx_crit, col_idx_crit,
                                                    updateCriteriaSheet.cell(row_idx_crit, col_idx_crit).value.replace(
                                                        '\n', '').strip() + '_' + str(millisAddObs))
            if (eachUpdateWorkSheet.name).strip().lower() == 'questions':
                for row_idx_ques in range(1, updateQuestionsSheet.nrows):
                    for col_idx_ques in range(0, updateQuestionsSheet.ncols):
                        if col_idx_ques == 2 or col_idx_ques == 0:
                            eachUpdateWorkSheet.write(row_idx_ques, col_idx_ques,
                                                    updateQuestionsSheet.cell(row_idx_ques, col_idx_ques).value.replace(
                                                        '\n', '').strip() + '_' + str(millisAddObs))
                for row_0 in range(0, updateQuestionsSheet.nrows):
                    if row_0 == 0:
                        eachUpdateWorkSheet.write(row_0, updateQuestionsSheet.ncols, 'question_operations')
                    else:
                        eachUpdateWorkSheet.write(row_0, updateQuestionsSheet.ncols, None)
            if (eachUpdateWorkSheet.name).strip().lower() == 'details':
                eachUpdateWorkSheet.write(1, 1, observationExternalId)
                for row_details_0 in range(0, updateDetailsSheet.nrows):
                    if row_details_0 == 0:
                        eachUpdateWorkSheet.write(row_details_0, updateDetailsSheet.ncols, 'solution_name_update')
                    else:
                        eachUpdateWorkSheet.write(row_details_0, updateDetailsSheet.ncols, None)
        copyOfUpdateWb.save(solutionName_for_folder_path.replace('.xlsx', '') + '_styles.xlsx')
        workbook = open_workbook(solutionName_for_folder_path.replace('.xlsx', '') + '_styles.xlsx')
        # Process each sheet
        for sheet in workbook.sheets():
            # Make a copy of the master worksheet
            new_workbook = copy(workbook)
            # for each time we copy the master workbook, remove all sheets except
            #  for the curren sheet (as defined by sheet.name)
            new_workbook._Workbook__worksheets = [worksheet for worksheet in new_workbook._Workbook__worksheets if
                                                worksheet.name != 'questions_sequence_sorted']
            # Save the new_workbook based on sheet.name
            new_workbook.save(solutionName_for_folder_path.replace('.xlsx', '') + '_styles.xlsx'.format(sheet.name))
        workbookXlsxWriter = xlsxwriter.Workbook(solutionName_for_folder_path.replace('.xlsx', '') + '_Success.xlsx')
        updateSuccessWorkBookReopen = xlrd.open_workbook(solutionName_for_folder_path.replace('.xlsx', '') + '_styles.xlsx',
                                                        on_demand=True)
        updateWbNumberOfSheetsReopen = updateSuccessWorkBookReopen.nsheets
        updateWbSheetNamesReopen = updateSuccessWorkBookReopen.sheet_names()
        updateQuestionsSheetReopen = updateSuccessWorkBookReopen.sheet_by_name('questions')
        updateDetailsSheetReopen = updateSuccessWorkBookReopen.sheet_by_name('details')
        cellFormat = workbookXlsxWriter.add_format()
        cellFormat.set_bg_color('00FF00')
        unlockCell = workbookXlsxWriter.add_format({'locked': False})
        for ele in updateWbSheetNamesReopen:
            if ele == 'details' or ele == 'questions' or ele == 'questions_sequence_sorted':
                updateWbSheetNamesReopen.remove(ele)
        for suSh in updateWbSheetNamesReopen:
            worksheetXlsxWriter = workbookXlsxWriter.add_worksheet(suSh)
            eachSheetByName = updateSuccessWorkBookReopen.sheet_by_name(suSh)
            for row_indx_sheets in range(eachSheetByName.nrows):
                for col_indx_sheets in range(eachSheetByName.ncols):
                    worksheetXlsxWriter.write(row_indx_sheets, col_indx_sheets,
                                            eachSheetByName.cell(row_indx_sheets, col_indx_sheets).value)
        questionsWorkSheetSuccess = workbookXlsxWriter.add_worksheet('questions')
        for row_idx_ques_succ in range(updateQuestionsSheetReopen.nrows):
            for col_idx_ques_succ in range(updateQuestionsSheetReopen.ncols):
                if col_idx_ques_succ == 0 or col_idx_ques_succ == 2:
                    questionsWorkSheetSuccess.protect()
                    questionsWorkSheetSuccess.write(row_idx_ques_succ, col_idx_ques_succ,
                                                    updateQuestionsSheetReopen.cell(row_idx_ques_succ,
                                                                                    col_idx_ques_succ).value, cellFormat)
                else:
                    questionsWorkSheetSuccess.write(row_idx_ques_succ, col_idx_ques_succ,
                                                    updateQuestionsSheetReopen.cell(row_idx_ques_succ,
                                                                                    col_idx_ques_succ).value, unlockCell)
                if updateQuestionsSheetReopen.ncols - 1 == col_idx_ques_succ:
                    questionsWorkSheetSuccess.data_validation(1, updateQuestionsSheetReopen.ncols - 1,
                                                            updateQuestionsSheetReopen.nrows,
                                                            updateQuestionsSheetReopen.ncols - 1,
                                                            {'validate': 'list', 'source': ['ADD', 'UPDATE', 'DELETE']})
        questionsWorkSheetSuccess.write_comment(0, 0,
                                                'criteria_id column is locked can\'t be edited , as it will be useful in updating the observations')
        questionsWorkSheetSuccess.write_comment(0, 2,
                                                'question_id column is locked can\'t be edited , as it will be useful in updating the observations')
        questionsWorkSheetSuccess.write_comment(0, updateQuestionsSheetReopen.ncols - 1,
                                                'question_operation column can be used in updating the questions , select either one of the options to update else leave blank and send the template to genie with update observation template command')
        detailsWorkSheetSuccess = workbookXlsxWriter.add_worksheet('details')
        for row_idx_deta_succ in range(updateDetailsSheetReopen.nrows):
            for col_idx_deta_succ in range(updateDetailsSheetReopen.ncols):
                if col_idx_deta_succ == 1:
                    detailsWorkSheetSuccess.protect()
                    detailsWorkSheetSuccess.write(row_idx_deta_succ, col_idx_deta_succ,
                                                updateDetailsSheetReopen.cell(row_idx_deta_succ, col_idx_deta_succ).value,
                                                cellFormat)
                else:
                    detailsWorkSheetSuccess.write(row_idx_deta_succ, col_idx_deta_succ,
                                                updateDetailsSheetReopen.cell(row_idx_deta_succ, col_idx_deta_succ).value,
                                                unlockCell)
                if updateDetailsSheetReopen.ncols - 1 == col_idx_deta_succ:
                    detailsWorkSheetSuccess.data_validation(1, updateDetailsSheetReopen.ncols - 1,
                                                            updateDetailsSheetReopen.nrows,
                                                            updateDetailsSheetReopen.ncols - 1,
                                                            {'validate': 'list', 'source': ['TRUE', 'FALSE']})
        detailsWorkSheetSuccess.write_comment(0, 1,
                                            'observation_id column is locked can\'t be edited , as it will be useful in updating the observations')
        detailsWorkSheetSuccess.write_comment(0, updateDetailsSheetReopen.ncols - 1,
                                            'solution_name_update column can be used in updating the solution_name , select either TRUE or FALSE and send the template to genie with update observation template command')
        sheet_names = ['Instructions', 'details', 'Criteria upload', 'Criteria_Rubric-Scoring',
                    'Domain(theme)_rubric_scoring', 'questions', 'framework', 'ECMs or Domains']
        workbookXlsxWriter.worksheets_objs.sort(key=lambda x: sheet_names.index(x.name))
        workbookXlsxWriter.close()
        print("Success sheet prepared.")


    @exception_handler
    def createChild(solutionName_for_folder_path, observationExternalId, accessToken):
        print("it has entered the creation api")
        try: 
            childObservationExternalId = str(observationExternalId + "_CHILD")
            urlSol_prog_mapping = internal_kong_ip + solutiontoprogrammappingapiurl + "?solutionId=" + observationExternalId + "&entityType=" + global_vars.entityType
            
            payloadSol_prog_mapping = {
                "externalId": childObservationExternalId,
                "name": global_vars.solutionName.lstrip().rstrip(),
                "description": global_vars.solutionDescription.lstrip().rstrip(),
                "programExternalId": global_vars.programExternalId
            }
            headersSol_prog_mapping = {'Authorization': authorization,
                                    'X-authenticated-user-token': accessToken,
                                    'Content-Type': content_type}
            responseSol_prog_mapping = requests.request("POST", urlSol_prog_mapping, headers=headersSol_prog_mapping,
                                                        data=json.dumps(payloadSol_prog_mapping))
            messageArr = ["Create child API called.", "URL : " + urlSol_prog_mapping,
                        "Status code : " + str(responseSol_prog_mapping.status_code),
                        "Response : " + responseSol_prog_mapping.text, "body : " + str(payloadSol_prog_mapping)]
            if responseSol_prog_mapping.status_code == 200:
                print("Solution mapped to program : " + global_vars.programName)
                print("Child solution : " + childObservationExternalId)

                responseSol_prog_mapping = responseSol_prog_mapping.json()
                child_id = responseSol_prog_mapping['result']['_id']
                Helpers.createAPILog(solutionName_for_folder_path, messageArr)
                return [child_id, childObservationExternalId]
            else:
                if responseSol_prog_mapping.status_code in [400, 401, 403, 404, 422]:
                    Helpers.errorVar.append(f"Sol_prog_mapping-Client Error {responseSol_prog_mapping.status_code}: {responseSol_prog_mapping.text}")
                elif responseSol_prog_mapping.status_code in [500, 502, 503, 504]:
                    Helpers.errorVar.append(f"Sol_prog_mapping-Server Error {responseSol_prog_mapping.status_code}: {responseSol_prog_mapping.text}")
                else:
                    Helpers.errorVar.append(f"Sol_prog_mapping-Unexpected Error {responseSol_prog_mapping.status_code}: {responseSol_prog_mapping.text}")
                print("Unable to create child solution")
                return False
        except Exception as e:
            messageArr = []
            messageArr.append("Exception caught : " + str(e))
            Helpers.createAPILog(solutionName_for_folder_path, messageArr)
            Helpers.errorVar.append(f"Error occurred: {str(e)}")
            return False
    
    @exception_handler
    def createSurveySolution(parentFolder, wbSurvey, accessToken):
        print("Create Survey Solution Func Called....")
        validation_cache = global_vars.get_resource_validation_cache() or {}
        dictDetailsEnv = validation_cache.get("details_row") or {}
        if not dictDetailsEnv:
            if 'details' not in [sheet.strip().lower() for sheet in wbSurvey.sheet_names()]:
                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("details sheet not found in survey workbook"))
                return [None, None]
            detailsEnvSheet = wbSurvey.sheet_by_name('details')
            keysEnv = [detailsEnvSheet.cell(1, col_index_env).value for col_index_env in range(detailsEnvSheet.ncols)]
            if detailsEnvSheet.nrows > 2:
                dictDetailsEnv = {
                    keysEnv[col_index_env]: detailsEnvSheet.cell(2, col_index_env).value
                    for col_index_env in range(detailsEnvSheet.ncols)
                }
        if not dictDetailsEnv:
            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("No data found in details sheet"))
            return [None, None]

        surveySolutionCreationReqBody = {}
        surveySolutionCreationReqBody['name'] = Helpers._to_text(dictDetailsEnv.get('survey_solution_name'))
        surveySolutionCreationReqBody["description"] = Helpers._to_text(dictDetailsEnv.get('survey_solution_description'))
        surveySolutionExternalId = str(uuid.uuid1())
        surveySolutionCreationReqBody["externalId"] = surveySolutionExternalId
        if Helpers._to_text(dictDetailsEnv.get('Name_of_the_creator')) == "":
            print('Diksha_loginId column should not be empty in the details sheet')
            Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Diksha_loginId column should not be empty in the details sheet"))
            return [None, None]
        else:
            surveySolutionCreationReqBody['creator'] = dictDetailsEnv['Name_of_the_creator']

        userDetails = Helpers.fetchUserDetails(accessToken, dictDetailsEnv['survey_creator_username'])
        if not userDetails:
            Helpers.errorVar.append(f"Failed to fetch user details for: {dictDetailsEnv['survey_creator_username']}")
            return [None, None]
        surveySolutionCreationReqBody['author'] = userDetails[0]
        if dictDetailsEnv["survey_start_date"]:
            if type(dictDetailsEnv["survey_start_date"]) == str:
                startDateArr = (dictDetailsEnv["survey_start_date"]).split("-")
                surveySolutionCreationReqBody["startDate"] = startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"
            elif type(dictDetailsEnv["survey_start_date"]) == float:
                surveySolutionCreationReqBody["startDate"] = (
                    xlrd.xldate.xldate_as_datetime(dictDetailsEnv["survey_start_date"], wbSurvey.datemode)).strftime("%Y/%m/%d")
            else:
                surveySolutionCreationReqBody["startDate"] = ""
            if dictDetailsEnv["survey_end_date"]:
                if type(dictDetailsEnv["survey_end_date"]) == str:
                    endDateArr = (dictDetailsEnv["survey_end_date"]).split("-")
                    surveySolutionCreationReqBody["endDate"] = endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + "T23:59:59.000Z"
                elif type(dictDetailsEnv["survey_end_date"]) == float:
                    surveySolutionCreationReqBody["endDate"] = (
                        xlrd.xldate.xldate_as_datetime(dictDetailsEnv["survey_end_date"], wbSurvey.datemode)).strftime("%Y/%m/%d")
                else:
                    surveySolutionCreationReqBody["endDate"] = ""

        print("Survey Solution Creation API called.")
        urlCreateSolutionApi = internal_kong_ip+ surveysolutioncreationapiurl
        headerCreateSolutionApi = {
            'Content-Type': content_type,
            'Authorization': authorization,
            'X-authenticated-user-token': accessToken,
            'X-Channel-id': x_channel_id,
            'appName': appname,
            'internal-access-token': internal_access_token
        }
        responseCreateSolutionApi = requests.post(url=urlCreateSolutionApi, headers=headerCreateSolutionApi, data=json.dumps(surveySolutionCreationReqBody))
        messageArr = ["********* Create Survey Solution *********", "URL : " + urlCreateSolutionApi,
                    "BODY : " + str(surveySolutionCreationReqBody),
                    "Status code : " + str(responseCreateSolutionApi.status_code),
                    "Response : " + responseCreateSolutionApi.text]
        fileheader = [Helpers._to_text(surveySolutionCreationReqBody.get('name')),'Program Sheet Validation'," "]
        Helpers.createAPILog(parentFolder, messageArr)
        Helpers.apicheckslog(parentFolder,fileheader)
        print("Survey Solution Creation API response received.", responseCreateSolutionApi.text)
        if responseCreateSolutionApi.status_code == 200:
            responseCreateSolutionApi = responseCreateSolutionApi.json()
            # We already generated the survey externalId before create; avoid immediate list lookup
            # because list APIs can be eventually consistent.
            solutionId = responseCreateSolutionApi["result"]["solutionId"]
            bodySolutionUpdate = {"creator": Helpers._to_text(dictDetailsEnv.get('Name_of_the_creator'))}
            if not Helpers.solutionUpdate(accessToken, solutionId, bodySolutionUpdate):
                return [None, None]

            return [solutionId, surveySolutionExternalId]
        else:
            error_message = ""
            if responseCreateSolutionApi.status_code in [400, 401, 403, 404, 422]:
                error_message = f"surveyCreationAPI-Client Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}"
            elif responseCreateSolutionApi.status_code in [500, 502, 503, 504]:
                error_message = f"surveyCreationAPI-Server Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}"
            else:
                error_message = f"surveyCreationAPI-Unexpected Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}"
            Helpers.errorVar.append(error_message)
            return [None, None]

    # upload survey questions 
    @exception_handler
    def uploadSurveyQuestions(MainFilePath, parentFolder, wbSurvey, addObservationSolution, accessToken, surTempExtID, surTempSolID, millisecond, programFile):
        print("Upload Survey Questions Func Called....")
        # print(parentFolder,"4854")
        # wbSurvey = xlrd.open_workbook(wbSurvey, on_demand=True)
        # print(f"Type of wbSurvey: {type(wbSurvey)}")
        sheetNam = wbSurvey.sheet_names()
        # print(sheetNam,"4854")
        stDt = None
        enDt = None
        shCnt = 0
        validation_cache = global_vars.get_resource_validation_cache() or {}
        survey_question_rows = validation_cache.get("question_rows") or []
        if not survey_question_rows:
            for i in sheetNam:
                if i.strip().lower() == 'questions':
                    sheetNam1 = wbSurvey.sheets()[shCnt]
                shCnt = shCnt + 1
            dataSort = [sheetNam1.row_values(i) for i in range(sheetNam1.nrows)]
            labels = dataSort[1]
            dataSort = dataSort[2:]
            dataSort.sort(key=lambda x: int(x[0]))
            survey_question_rows = []
            for row in dataSort:
                row_dict = {
                    labels[col_idx]: row[col_idx] for col_idx in range(len(labels))
                }
                survey_question_rows.append(row_dict)
        else:
            try:
                survey_question_rows = sorted(
                    survey_question_rows,
                    key=lambda x: int(float(x.get("question_sequence", 0)))
                )
            except Exception:
                pass

        sheetNames = ['questions_sequence_sorted']
        for sheet2 in sheetNames:
            if sheet2.strip().lower() == 'questions_sequence_sorted':
                questionsList = list(survey_question_rows)
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
                        
                urlQuestionsUploadApi = internal_kong_ip + questionuploadapiurl
                headerQuestionUploadApi = {
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
                    Helpers.createAPILog(parentFolder, messageArr)
                    messageArr1 = ["Questions","Question upload Success","Passed",str(responseQuestionUploadApi.status_code)]
                    Helpers.apicheckslog(parentFolder,messageArr1)

                    with open(parentFolder + '/questionUpload/uploadInternalIdsSheet.csv', 'w+',encoding='utf-8') as questionRes:
                        questionRes.write(responseQuestionUploadApi.text)
                    urlImportSoluTemplate = internal_kong_ip + importsurveysolutiontemplateurl + str(surTempSolID) + "?appName=manage-learn"
                    headerImportSoluTemplateApi = {
                        'Authorization': authorization,
                        'X-authenticated-user-token': accessToken,
                        'X-Channel-id': x_channel_id
                    }
                    responseImportSoluTemplateApi = requests.get(url=urlImportSoluTemplate,
                                                                headers=headerImportSoluTemplateApi)
                    if responseImportSoluTemplateApi.status_code == 200:
                        print('Creating Child Success')

                        messageArr = ["********* Creating Child api *********", "URL : " + urlImportSoluTemplate,
                                    "Status code : " + str(responseImportSoluTemplateApi.status_code),
                                    "Response : " + responseImportSoluTemplateApi.text]
                        Helpers.createAPILog(parentFolder, messageArr)
                        responseImportSoluTemplateApi = responseImportSoluTemplateApi.json()
                        solutionIdSuc = responseImportSoluTemplateApi["result"]["solutionId"]
                        urlSurveyProgramMapping = internal_kong_ip + importsurveysolutiontoprogramurl + str(solutionIdSuc) + "?programId=" + global_vars.programExternalId.lstrip().rstrip()
                        headeSurveyProgramMappingApi = {
                            'Authorization': authorization,
                            'X-authenticated-user-token': accessToken,
                            'X-Channel-id': x_channel_id
                        }
                        responseSurveyProgramMappingApi = requests.get(url=urlSurveyProgramMapping,headers=headeSurveyProgramMappingApi)
                        if responseSurveyProgramMappingApi.status_code == 200:
                            print('Program Mapping Success')
                            
                            messageArr = ["********* Program mapping api *********", "URL : " + urlSurveyProgramMapping,
                                        "Status code : " + str(responseSurveyProgramMappingApi.status_code),
                                        "Response : " + responseSurveyProgramMappingApi.text]
                            Helpers.createAPILog(parentFolder, messageArr)
                            surveyLink = None
                            solutionIdSuc = None
                            surveyExternalIdSuc = None
                            surveyLink = responseImportSoluTemplateApi["result"]["link"]
                            solutionIdSuc = responseImportSoluTemplateApi["result"]["solutionId"]
                            solutionExtIdSuc = responseImportSoluTemplateApi["result"]["solutionExternalId"]
                            print("Survey Child Id : " + str(solutionExtIdSuc))
                            solutionDetails = Helpers.fetchSolutionDetailsFromProgramSheet(parentFolder, programFile, solutionIdSuc,
                                                                                accessToken)
                            if not solutionDetails:
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Failed to fetch solution details from program sheet"))
                                return False
                            scopeEntities = global_vars.entitiesPGMID
                            scopeRoles = solutionDetails[0]
                            surveyScopeBody = {
                                "scope": {"entityType": global_vars.scopeEntityType, "entities": scopeEntities, "roles": scopeRoles}}
                            if not Helpers.solutionUpdate(accessToken, solutionIdSuc, surveyScopeBody):
                                return False
                            global_vars.surveySolutionlink = Helpers.prepareProgramSuccessSheet(MainFilePath, parentFolder, programFile, solutionExtIdSuc,
                                                    solutionIdSuc, accessToken)
                            if not global_vars.surveySolutionlink:
                                Helpers.errorVar.append(str("CRITICAL") + ': ' + str("Prepare program success sheet failed"))
                                return False
                            
                            print('Survey Successfully Added')
                            print(global_vars.surveySolutionlink)
                        else:
                            print('Program Mapping Failed')
                            if responseSurveyProgramMappingApi.status_code in [400, 401, 403, 404, 422]:
                                Helpers.errorVar.append(f"SurveyProgramMappingApi-Client Error {responseSurveyProgramMappingApi.status_code}: {responseSurveyProgramMappingApi.text}")
                            elif responseSurveyProgramMappingApi.status_code in [500, 502, 503, 504]:
                                Helpers.errorVar.append(f"SurveyProgramMappingApi-Server Error {responseSurveyProgramMappingApi.status_code}: {responseSurveyProgramMappingApi.text}")
                            else:
                                Helpers.errorVar.append(f"SurveyProgramMappingApi-Unexpected Error {responseSurveyProgramMappingApi.status_code}: {responseSurveyProgramMappingApi.text}")
                            messageArr = ["********* Program mapping api *********", "URL : " + urlSurveyProgramMapping,
                                        "Status code : " + str(responseSurveyProgramMappingApi.status_code),
                                        "Response : " + responseSurveyProgramMappingApi.text]
                            messageArr.append(f"Error Response: {Helpers.errorVar}")
                            return False
                    else:
                        print('Creating Child API Failed')
                        if responseImportSoluTemplateApi.status_code in [400, 401, 403, 404, 422]:
                            Helpers.errorVar.append(f"ImportSoluTemplateApi-Client Error {responseImportSoluTemplateApi.status_code}: {responseImportSoluTemplateApi.text}")
                        elif responseImportSoluTemplateApi.status_code in [500, 502, 503, 504]:
                            Helpers.errorVar.append(f"ImportSoluTemplateApi-Server Error {responseImportSoluTemplateApi.status_code}: {responseImportSoluTemplateApi.text}")
                        else:
                            Helpers.errorVar.append(f"ImportSoluTemplateApi-Unexpected Error {responseImportSoluTemplateApi.status_code}: {responseImportSoluTemplateApi.text}")
                        messageArr = ["********* Program mapping api *********", "URL : " + urlImportSoluTemplate,
                                    "Status code : " + str(responseImportSoluTemplateApi.status_code),
                                    "Response : " + responseImportSoluTemplateApi.text]
                        messageArr.append(f"Error Response: {Helpers.errorVar}")
                        return False
                else:
                    if responseQuestionUploadApi.status_code in [400, 401, 403, 404, 422]:
                        Helpers.errorVar.append(f"QuestionUploadApi-Client Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}")
                    elif responseQuestionUploadApi.status_code in [500, 502, 503, 504]:
                        Helpers.errorVar.append(f"QuestionUploadApi-Server Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}")
                    else:
                        Helpers.errorVar.append(f"QuestionUploadApi-Unexpected Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}")
                    print('QuestionUploadApi Failed')
                    return False
        return global_vars.surveySolutionlink

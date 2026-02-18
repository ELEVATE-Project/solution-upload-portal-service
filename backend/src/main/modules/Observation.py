from backend.src.main.modules.common_config import *
from dotenv import load_dotenv
from pathlib import Path
import os, json, requests, sys, csv, time, uuid, re, openpyxl
from openpyxl.styles import Color, PatternFill
from datetime import datetime
from bson import ObjectId
import pandas as pd
import backend.src.main.modules.headers as apiHeader


env_path = Path(__file__).resolve().parents[1] / "apiServices" / "src" / "main" / ".env"

# Load the .env file
load_dotenv(dotenv_path=env_path)
internal_access_token = os.getenv("internal_access_token")
# adminTokenHeaderName = os.getenv("adminTokenHeaderName")
jwtTokenSecret = os.getenv("jwtTokenSecret")
# projAdminAccessToken = os.getenv("projAdminAccessToken")
# adminAccessToken = os.getenv("adminAccessToken")
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


class CreateObservation:
    def __init__(self):
        self.errorVar = []

    def createAPILog(self, solutionName_for_folder_path, messageArr):
        file_exists = os.path.join(solutionName_for_folder_path, 'apiHitLogs', 'apiLogs.txt')
        os.makedirs(os.path.dirname(file_exists), exist_ok=True)
        # Create file with header if it doesn't exist
        if not os.path.exists(file_exists):
            with open(file_exists, "w", encoding='utf-8') as API_log:
                API_log.write("===============================================================================\n")
                API_log.write("ENVIRONMENT LOGS\n")
                API_log.write("===============================================================================\n")
        # Append logs
        with open(file_exists, "a", encoding='utf-8') as API_log:
            API_log.write("\n")
            for msg in messageArr:
                API_log.write(str(msg))
                API_log.write("\n")

    def apicheckslog(self, solutionName_for_folder_path, messageArr):
        file_exists = os.path.join(solutionName_for_folder_path, 'apiHitLogs', 'apiLogs.csv')
        fileheader = ["Resource", "Process", "Status", "Remark"]
        os.makedirs(os.path.dirname(file_exists), exist_ok=True)

        # ✅ Create file with header if it doesn't exist
        if not os.path.exists(file_exists):
            with open(file_exists, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',', lineterminator='\n')
                writer.writerow(fileheader)  
                
        # ✅ Append a new log entry
        with open(file_exists, 'a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',', lineterminator='\n')
            writer.writerow(messageArr)

    def FetchTempExternalID(self, parentFolder, accessToken, solutionName, programdetails, userRole):
        try:
            urldbFindPT = elevateprojecthost + dbfindapi_projectTemplate
            headers = apiHeader.headers().headerFetchEntitytype(accessToken)
            searchSolutionpayloadPT = {
                "query": {
                    "title": solutionName,
                    "isReusable": True
                    },
                "projection": ["externalId"],
                "mongoIdKeys": ["_id", "solutionId", "metaInformation.solutionId"],
                "limit": 10000
            }
            responsePT = requests.post(urldbFindPT, headers=headers, json=searchSolutionpayloadPT)

            if responsePT.status_code != 200:
                self.errorVar.append(f"DBFindPT-Error {responsePT.status_code}: {responsePT.text}")
                self.createAPILog(parentFolder, responsePT.text)
                print("Unable to fetch Project Template External ID.")
                return False

            resultsPT = responsePT.json().get("result", [])
            if not resultsPT:
                self.errorVar.append( "No projectTemplate found for: " + solutionName)
                return False

            ProjectTempExternalID = resultsPT[0].get("externalId")
            print(ProjectTempExternalID,"ProjectTempExternalID")
            urldbFind = elevateprojecthost + dbfindapi_url
            headers = apiHeader.headers().headerFetchEntitytype(accessToken)
            searchSolutionpayload = {
                "query": {"name": solutionName},
                "projection": ["status", "name"],
                "mongoIdKeys": ["_id"],
                "limit": 10000
            }
            response = requests.post(urldbFind, headers=headers, json=searchSolutionpayload)

            if response.status_code != 200:
                self.errorVar.append(f"DBFind-Error {response.status_code}: {response.text}")
                self.createAPILog(parentFolder, response.text)
                print("Unable to fetch Solution...")
                return False

            results = response.json().get("result", [])
            if not results:
                self.errorVar.append("No solutions found for name: " + solutionName)
                return False

            projectSolutionID = results[0].get("_id")
            if not projectSolutionID:
                self.errorVar.append("Solution found but no projectTemplateId for: " + solutionName)
                return False
            
            urldbFindPT = elevateprojecthost + solutionupdateapi + projectSolutionID
            headerUpdateSolutionApi = apiHeader.headers().headersObservationsolutionUpdate(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
            
            searchSolutionpayloadPT = {
                "status": "inactive",
                "isDeleted": False
            }
            responseSolutionUpdate = requests.post(urldbFindPT, headers=headerUpdateSolutionApi, json=searchSolutionpayloadPT)

            if responseSolutionUpdate.status_code == 200:
                print("Solution Update Success.")
            else:
                self.errorVar.append("Solution Update Failed.")
                return False
            return ProjectTempExternalID

        except Exception as e:
            self.errorVar.append(f"Exception in FetchTempExternalID: {str(e)}")
            return False

        except Exception as e:
            self.errorVar.append(f"Exception in FetchTempExternalID: {str(e)}")
            return False
        
    def criteriaUpload(self, parentFolder, wbObservation, millisAddObs, accessToken, tabName, impLedObsFlag, programdetails, userRole):
        criteriaLookUp = dict()
        criteriaColNames = ["criteriaId", "criteria_name"]
        criteriaUploadFieldnames = ['criteriaID', 'criteriaName']
        dictCriteriaToCsv = dict()
        criteriaLevelsFromFramework = dict()
        countImps = 0

        # === CASE 1: FRAMEWORK TAB ===
        if tabName == "framework":
            # --- fetch framework sheet (list of dicts) ---
            fetchLevelsFromFramework = wbObservation.get("framework")
            if not fetchLevelsFromFramework or not isinstance(fetchLevelsFromFramework, list):
                self.errorVar.append("Framework sheet missing or in unexpected format.")
                return False

            # --- Handle Improvement Mapping (normalize both possible shapes) ---
            criteriaImpDict = {}
            TcountImps = []
            if impLedObsFlag:
                impsToCriteria = wbObservation.get("imp mapping")  # expected: list of dicts

                if impsToCriteria:
                    # take first sample row (fall back to empty dict)
                    sample_row = next(iter(impsToCriteria), {})

                    # regex to match keys like "L1-improvement-projects" and capture the number
                    pattern = re.compile(r'^\s*L(\d+)-improvement-projects\s*$', re.IGNORECASE)

                    levels = set()  # use a set to avoid duplicates
                    for key in sample_row.keys():
                        if not isinstance(key, str):
                            continue
                        m = pattern.match(key)
                        if m:
                            level_num = int(m.group(1))
                            levels.add(level_num)

                    TcountImps = sorted(levels)

                countImps = len(TcountImps)
                if not impsToCriteria:
                    self.errorVar.append("Imp mapping expected but missing.")
                    return False

                # Handle both dict and list formats
                if isinstance(impsToCriteria[0], dict):
                    # Dict format
                    print(impsToCriteria,"impsToCriteria")
                    for dictImp in impsToCriteria:
                        print(dictImp,"dictImp")
                        crit_key = str(dictImp.get('criteriaId', '')).strip()
                        print(crit_key,"crit_key")  
                        if not crit_key:
                            continue
                        criteriaImpDict[crit_key] = {}
                        for levls in range(1, countImps + 1):
                            colname = f"L{levls}-improvement-projects"
                            print(colname,"colname")
                            solutionName = str(dictImp.get(colname, "") or "").strip()
                            print(solutionName,"solutionName")
                            if solutionName!= "":
                                ProjectTempExternalID = self.FetchTempExternalID(parentFolder, accessToken, solutionName, programdetails, userRole)
                                if not ProjectTempExternalID:
                                    return False
                            else:
                                ProjectTempExternalID = ""
                            criteriaImpDict[crit_key][colname] = ProjectTempExternalID
                else:
                    # List format
                    keysFromImpSheet = impsToCriteria[0]
                    for row in impsToCriteria[1:]:
                        dictImp = {keysFromImpSheet[i]: row[i] for i in range(len(keysFromImpSheet))}
                        crit_key = str(dictImp.get('criteriaId', '')).strip()
                        if not crit_key:
                            continue
                        criteriaImpDict[crit_key] = {}
                        for levls in range(1, countImps + 1):
                            colname = f"L{levls}-improvement-projects"
                            solutionName = str(dictImp.get(colname, "") or "").strip()
                            ProjectTempExternalID = self.FetchTempExternalID(parentFolder, accessToken, solutionName, programdetails, userRole)
                            if not ProjectTempExternalID:
                                return False
                            criteriaImpDict[crit_key][colname] = ProjectTempExternalID

            # --- Determine levelCount from framework headers ---
            first_row_keys = list(fetchLevelsFromFramework[0].keys())
            levelCount = 0
            for k in first_row_keys:
                if (
                    isinstance(k, str)
                    and k.strip().lower().startswith("l")
                    and "description" in k.lower()
                    and not k.strip().lower().startswith("ln")  # 👈 Ignore 'Ln' cases
                ):
                    levelCount += 1

            # --- Build criteriaLevelsFromFramework ---
            for dictFramework in fetchLevelsFromFramework:
                criteria_id = str(dictFramework.get("Criteria ID", "")).strip()
                if not criteria_id:
                    continue
                criteriaLevelsFromFramework[criteria_id] = {}
                for levlsNo in range(1, levelCount + 1):
                    key_name = f"L{levlsNo} description"
                    criteriaLevelsFromFramework[criteria_id][f"L{levlsNo}"] = dictFramework.get(key_name, "") or ""
                    if f"L{levlsNo}" not in criteriaColNames:
                        criteriaColNames.append(f"L{levlsNo}")

            # --- Create CSV ---
            criteriaFilePath = os.path.join(parentFolder, 'criteriaUpload')
            os.makedirs(criteriaFilePath, exist_ok=True)
            csv_path = os.path.join(criteriaFilePath, 'uploadSheet.csv')
            file_exists = os.path.isfile(csv_path)
            criteriaLevelsCount = levelCount
            wbObservation['criteriaLevelsCount'] = criteriaLevelsCount

            # --- Write rows ---
            for dictCriteria in fetchLevelsFromFramework:
                criteria_id_raw = str(dictCriteria.get('Criteria ID', '')).strip()
                if not criteria_id_raw:
                    continue

                dictCriteriaToCsv = {}
                dictCriteriaToCsv['criteriaID'] = criteria_id_raw + '_' + str(millisAddObs)
                criteriaLookUp[dictCriteriaToCsv['criteriaID'].strip()] = dictCriteria.get('Criteria Name', '')
                dictCriteriaToCsv['criteriaName'] = dictCriteria.get('Criteria Name', '')
                dictCriteriaToCsv['type'] = 'auto'

                # Add levels
                for levlsNo in range(1, levelCount + 1):
                    level_col = f"L{levlsNo} description"
                    dictCriteriaToCsv[f"L{levlsNo}"] = dictCriteria.get(level_col, "") or ""

                # Add improvement mappings (if any)
                if impLedObsFlag:
                    crit_key = criteria_id_raw
                    for eachImps, val in criteriaImpDict.get(crit_key, {}).items():
                        dictCriteriaToCsv[eachImps] = val

                # Ensure headers include everything
                if 'type' not in criteriaUploadFieldnames:
                    criteriaUploadFieldnames.append('type')
                for eachCols in criteriaColNames:
                    if eachCols not in ['criteria_id', 'criteria_name', 'type', "criteriaId"]:
                        if eachCols not in criteriaUploadFieldnames:
                            criteriaUploadFieldnames.append(eachCols)
                if impLedObsFlag:
                    for levls in range(1, countImps + 1):
                        imp_col = f"L{levls}-improvement-projects"
                        if imp_col not in criteriaUploadFieldnames:
                            criteriaUploadFieldnames.append(imp_col)

                # Write to CSV
                with open(csv_path, 'a', encoding='utf-8', newline='') as criteriaUploadFile:
                    writerCriteriaUpload = csv.DictWriter(criteriaUploadFile, fieldnames=list(criteriaUploadFieldnames), lineterminator='\n')
                    if not file_exists:
                        writerCriteriaUpload.writeheader()
                        file_exists = True
                    writerCriteriaUpload.writerow(dictCriteriaToCsv)
        # === CASE 2: CRITERIA TAB ===
        elif tabName == "criteria":
            print(wbObservation,"wbObservation")
            criteriaList = wbObservation.get("criteria")  # This is already a list of dicts
            if not criteriaList or not isinstance(criteriaList, list):
                self.errorVar.append("Criteria sheet missing or not in expected dict format.")
                return False

            criteriaFilePath = os.path.join(parentFolder, 'criteriaUpload')
            os.makedirs(criteriaFilePath, exist_ok=True)
            csv_path = os.path.join(criteriaFilePath, 'uploadSheet.csv')
            file_exists = os.path.isfile(csv_path)
            criteriaUploadFieldnames = ['criteriaID', 'criteriaName', 'L1', 'L2', 'L3', 'type']

            for row in criteriaList:  # ✅ directly loop dicts
                criteria_id_raw = str(row.get('criteria_id', '')).strip()
                criteria_name_raw = str(row.get('criteria_name', '')).encode('utf-8').decode('utf-8')

                if not criteria_id_raw or not criteria_name_raw:
                    continue

                data = {
                    'criteriaID': criteria_id_raw + '_' + str(millisAddObs),
                    'criteriaName': criteria_name_raw,
                    'L1': 'NA',
                    'L2': 'NA',
                    'L3': 'NA',
                    'type': 'auto'
                }

                with open(csv_path, 'a', encoding='utf-8', newline='') as file:
                    writer = csv.DictWriter(file, fieldnames=criteriaUploadFieldnames, lineterminator='\n')
                    if not file_exists:
                        writer.writeheader()
                        file_exists = True
                    writer.writerow(data)

            print(f"✅ Criteria upload CSV created successfully at: {csv_path}")
        else:
            self.errorVar.append("Invalid tabName provided. Expected 'framework' or 'criteria'.")

        # === Upload to API (common for both) ===
        try:
            urlCriteriaUploadApi = internal_kong_ip + criteriauploadapiurl
            headerCriteriaUploadApi = apiHeader.headers().headersCriteriaUpload(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
            filesCriteria = {'criteria': open(csv_path, 'rb')}
            responseCriteriaUploadApi = requests.post(url=urlCriteriaUploadApi, headers=headerCriteriaUploadApi, files=filesCriteria)

            messageArr = [
                "Criteria Upload Sheet Prepared.",
                f"File path : {csv_path}",
                f"Upload status code : {responseCriteriaUploadApi.status_code}"
            ]
            self.createAPILog(parentFolder, messageArr)
            if responseCriteriaUploadApi.status_code == 200:
                print('✅ CriteriaUploadApi Success')
                with open(os.path.join(criteriaFilePath, 'uploadInternalIdsSheet.csv'), 'w+', encoding='utf-8') as criteriaRes:
                    criteriaRes.write(responseCriteriaUploadApi.text)
                return True
            else:
                messageArr.append("Response : " + str(responseCriteriaUploadApi.text))
                self.createAPILog(parentFolder, messageArr)
                self.errorVar.append(f"CriteriaUploadApi Error {responseCriteriaUploadApi.status_code}: {responseCriteriaUploadApi.text}")
                print("❌ Criteria Upload failed.")
                return False
        except Exception as e:
            messageArr.append("Exception caught : " + str(e))
            self.createAPILog(parentFolder, messageArr)
            self.errorVar.append(f"Error occurred during Criteria Upload API: {str(e)}")
            return False

    def frameWorkUpload(self, solutionName_for_folder_path, wbObservation, millisAddObs, accessToken, programdetails, typeofsolution, userRole):
        detailsSheet = wbObservation['details']
        if typeofsolution == 2:
            detailsSheet["scoring_system"] = "null"
        dateTime = datetime.now()
        frameworkDocInsertObj = {}
        frameworkExternalId = None
        frameworkExternalId = uuid.uuid1()
        frameworkExternalId = str(frameworkExternalId)
        frameworkDocInsertObj['externalId'] = frameworkExternalId
        frameworkDocInsertObj['name'] = detailsSheet.get("observation_solution_name").strip()
        frameworkDocInsertObj['description'] = detailsSheet.get("observation_solution_description").strip()
        frameworkDocInsertObj['parentId'] = None
        frameworkDocInsertObj['resourceType'] = ['Observations Framework']
        frameworkDocInsertObj['language'] = detailsSheet.get("language")
        frameworkDocInsertObj['levelToScoreMapping'] = dict()
        keyWords = detailsSheet.get("language")
        if keyWords and (keyWords != 'Framework' or keyWords != 'Frameworks' or keyWords != 'Observation' or keyWords != 'Observations'):
            keywordsFinalArr = ['Framework', 'Observation']
            keywordsArr = keyWords.encode('utf-8').decode('utf-8').split(',')
            for keyw in keywordsArr:
                keywordsFinalArr.append(keyw)
            frameworkDocInsertObj['keywords'] = keywordsFinalArr
        else:
            frameworkDocInsertObj['keywords'] = ['Framework', 'Observation']
        frameworkDocInsertObj['concepts'] = []
        frameworkDocInsertObj['createdFor'] = [detailsSheet.get("Username/user id/email id/phone no. of the Content creator") or ""]  # createdForArr
        frameworkDocInsertObj['rootOrg'] = [detailsSheet.get("Username/user id/email id/phone no. of the Content creator") or ""]     # rootOrgArr
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
        if not detailsSheet.get("scoring_system").lower() == "null":
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
            frameworkDocInsertObj['scoringSystem'] = detailsSheet.get("scoring_system")
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
            criteriaLevelsCount = wbObservation.get('criteriaLevelsCount')
            for levs in range(1, criteriaLevelsCount + 1):
                levelToScore = {"L" + str(levs): {'points': levs * 10, 'label': 'Level ' + str(levs)}}
                frameworkDocInsertObj['levelToScoreMapping'].update(levelToScore)
            frameworkDocInsertObj['noOfRatingLevels'] = criteriaLevelsCount
            
        else:
            frameworkDocInsertObj['scoringSystem'] = None
            frameworkDocInsertObj['isRubricDriven'] = False
        entitydetails = programdetails.get("entitiesType")
        entity_type_id = entitydetails[1]
        if isinstance(entity_type_id, list) and len(entity_type_id) > 0:
            entity_type_id = entity_type_id[0]
        entity_type = detailsSheet.get("entity_type")
        frameworkDocInsertObj['entityType'] = str(entity_type)
        frameworkDocInsertObj['entityTypeId'] = None
        frameworkDocInsertObj['type'] = 'observation'
        frameworkDocInsertObj['subType'] = str(entity_type)
        frameworkDocInsertObj['status'] = "active"
        frameworkDocInsertObj['updatedBy'] = 'INITIALIZE'
        frameworkDocInsertObj['createdBy'] = 'INITIALIZE'
        frameworkDocInsertObj['createdAt'] = str(dateTime)
        frameworkDocInsertObj['updatedAt'] = str(dateTime)
        frameworkDocInsertObj['author'] = detailsSheet.get("Name_of_the_creator")
        frameworkDocInsertObj['isTempObTest'] = 'observationAutomation'

        # Adding Credits and license into Frameworks
        frameworkDocInsertObj['creator'] = str(detailsSheet.get("Username/user id/email id/phone no. of the Content creator"))
        frameworkDocInsertObj['license'] = {}
        frameworkDocInsertObj['license']['author'] = str(detailsSheet.get("Username/user id/email id/phone no. of the Content creator"))
        frameworkDocInsertObj['license']['creator'] = str(detailsSheet.get("Username/user id/email id/phone no. of the Content creator"))
        frameworkDocInsertObj['license']['copyright'] = str(detailsSheet.get("Name_of_the_creator"))
        frameworkDocInsertObj['license']['copyrightYear'] = int(dateTime.strftime("%Y"))
        frameworkDocInsertObj['license']['contentType'] = "Observation"
        frameworkDocInsertObj['license']['organisation'] = [detailsSheet.get("Username/user id/email id/phone no. of the Content creator")]
        frameworkDocInsertObj['license']['orgDetails'] = {}
        frameworkDocInsertObj['license']['orgDetails']['email'] = None
        frameworkDocInsertObj['license']['orgDetails']['orgName'] = None
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
            
            headerFrameworkUploadApi = apiHeader.headers().headersFrameworkUpload(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
            
            filesFramework = {'framework': open(solutionName_for_folder_path + '/framework/uploadFile.json', 'rb')}

            responseFrameworkUploadApi = requests.post(url=urlCreateFrameworkApi, headers=headerFrameworkUploadApi,
                                                    files=filesFramework)
            messageArr = ["Framwork json file created.",
                        "File loc : " + solutionName_for_folder_path + '/framework/uploadFile.json',
                        "Framework upload API called,", "Status code : " + str(responseFrameworkUploadApi.status_code)]
            self.createAPILog(solutionName_for_folder_path, messageArr)
            if responseFrameworkUploadApi.status_code == 200:
                print('Framework upload Success')
                return frameworkExternalId

            else:
                if responseFrameworkUploadApi.status_code in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"FrameworkUploadApi-Client Error {responseFrameworkUploadApi.status_code}: {responseFrameworkUploadApi.text}")
                elif responseFrameworkUploadApi.status_code in [500, 502, 503, 504]:
                    self.errorVar.append(f"FrameworkUploadApi-Server Error {responseFrameworkUploadApi.status_code}: {responseFrameworkUploadApi.text}")
                else:
                    self.errorVar.append(f"FrameworkUploadApi-Unexpected Error {responseFrameworkUploadApi.status_code}: {responseFrameworkUploadApi.text}")
                messageArr = ["Framwork upload Failed.", "Response : " + responseFrameworkUploadApi.text]
                self.createAPILog(solutionName_for_folder_path, messageArr)
                print('Framework upload api failed ',
                    'with response from api is ' + str(responseFrameworkUploadApi.text))
                return False
                
        except Exception as e:
            messageArr.append("Exception caught : " + str(e))
            self.createAPILog(solutionName_for_folder_path, messageArr)
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False
            
    def themesUpload(self, solutionName_for_folder_path, wbObservation, millisAddObs, accessToken, frameworkExternalId,obsWORubWS,programdetails, userRole):
        dictCritLookUp = {}
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
            # Extract the framework list directly from your dict
            framework_list = wbObservation['framework']  # This is the list of dicts you showed

            # CSV fieldnames
            themeUploadFieldnames = ["theme", "aoi", "indicators", "criteriaInternalId"]

            # List to hold all rows
            themesUploadList = []

            # Build dicts for CSV
            for dictCriteria in framework_list:
                criteria_id = dictCriteria['Criteria ID'].strip()
                themeRow = {
                    'theme': f"{dictCriteria['Domain Name']}###{dictCriteria['Domain ID']}###40",
                    'aoi': "",
                    'indicators': "",
                    'criteriaInternalId': dictCritLookUp[criteria_id + f"_{millisAddObs}"] + "###40"
                }
                themesUploadList.append(themeRow)

            # Ensure folder exists
            themeFilePath = os.path.join(solutionName_for_folder_path, 'themeUpload')
            os.makedirs(themeFilePath, exist_ok=True)

            # Write CSV
            uploadCsvPath = os.path.join(themeFilePath, 'uploadSheet.csv')
            file_exists = os.path.isfile(uploadCsvPath)

            with open(uploadCsvPath, 'a', encoding='utf-8', newline='') as themeUploadFile:
                writer = csv.DictWriter(themeUploadFile, fieldnames=themeUploadFieldnames, lineterminator='\n')
                if not file_exists:
                    writer.writeheader()
                writer.writerows(themesUploadList)
        try:
            urlThemesUploadApi = internal_kong_ip + themeuploadapiurl + frameworkExternalId
            headerThemesUploadApi = apiHeader.headers().headersFrameworkUpload(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
            filesThemes = {'themes': open(solutionName_for_folder_path + '/themeUpload/uploadSheet.csv', 'rb')}
            responseThemeUploadApi = requests.post(url=urlThemesUploadApi, headers=headerThemesUploadApi, files=filesThemes)
            messageArr = ["Themes upload sheet prepared.",
                        "File path : " + solutionName_for_folder_path + '/themeUpload/uploadSheet.csv',
                        "Theme upload to framework API called.", "URL : " + urlThemesUploadApi,
                        "Status code : " + str(responseThemeUploadApi.status_code)]
            self.createAPILog(solutionName_for_folder_path, messageArr)
            if responseThemeUploadApi.status_code == 200:
                print('Theme UploadApi Success')
                with open(solutionName_for_folder_path + '/themeUpload/uploadInternalIdsSheet.csv', 'w+',encoding='utf-8') as criteriaRes:
                    criteriaRes.write(responseThemeUploadApi.text)
                return True
            else:
                if responseThemeUploadApi.status_code in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"ThemeUploadApi-Client Error {responseThemeUploadApi.status_code}: {responseThemeUploadApi.text}")
                elif responseThemeUploadApi.status_code in [500, 502, 503, 504]:
                    self.errorVar.append(f"ThemeUploadApi-Server Error {responseThemeUploadApi.status_code}: {responseThemeUploadApi.text}")
                else:
                    self.errorVar.append(f"ThemeUploadApi-Unexpected Error {responseThemeUploadApi.status_code}: {responseThemeUploadApi.text}")
                messageArr = ["Themes upload failed.", "Response : " + str(responseThemeUploadApi.text)]
                self.createAPILog(solutionName_for_folder_path, messageArr)
                print("Theme upload failed.")
                return False
                # sys.exit()
        except Exception as e:
            messageArr.append("Exception caught : " + str(e))
            self.createAPILog(solutionName_for_folder_path, messageArr)
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False
        
    def createSolutionFromFramework(self, solutionName_for_folder_path, accessToken, frameworkExternalId,programdetails, userRole):
        try:
            entitydetails = programdetails.get("entitiesType")
            entity_type = entitydetails[0] if isinstance(entitydetails[0], str) else str(entitydetails[0][0])
            entity_type_id = entitydetails[1]
            urlCreateSolutionApi = internal_kong_ip + solutioncreationapiurl
            headerCreateSolutionApi = apiHeader.headers().headersCreateSolutionFromFramework(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
            isExternalProgram = "true"
            queryparamsCreateSolutionApi = '?frameworkId=' + str(frameworkExternalId) + '&entityType=' + entity_type + '&isExternalProgram=' + isExternalProgram
            responseCreateSolutionApi = requests.post(url=urlCreateSolutionApi + queryparamsCreateSolutionApi,
                                                    headers=headerCreateSolutionApi)

            messageArr = ["Solution Created from Framework.",
                        "URL : " + str(urlCreateSolutionApi + queryparamsCreateSolutionApi),
                        "Status Code : " + str(responseCreateSolutionApi.status_code),
                        "Response : " + str(responseCreateSolutionApi.text)]
            self.createAPILog(solutionName_for_folder_path, messageArr)
            messageArr = []
            if responseCreateSolutionApi.status_code == 200:
                responseCreateSolutionApi = responseCreateSolutionApi.json()
                solutionId = responseCreateSolutionApi['result']['templateId']
                messageArr.append("Parent Solution Generated : " + str(solutionId))
                print("Parent Solution Generated : " + str(solutionId))
                self.createAPILog(solutionName_for_folder_path, messageArr)
                return solutionId
            else:
                if responseCreateSolutionApi.status_code in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"CreateSolutionApi-Client Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}")
                elif responseCreateSolutionApi.status_code in [500, 502, 503, 504]:
                    self.errorVar.append(f"CreateSolutionApi-Server Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}")
                else:
                    self.errorVar.append(f"CreateSolutionApi-Unexpected Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}")
                messageArr.append("Solution from framework api failed.")
                self.createAPILog(solutionName_for_folder_path, messageArr)
                print("Solution from framework api failed.")
                return False
        except Exception as e:
            messageArr.append("Exception caught : " + str(e))
            self.createAPILog(solutionName_for_folder_path, messageArr)
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False
        
    def ECMUpdatebody(self,ECMSheet,millisecond):
        ecm_update = {}
        ecm_dict = {}
        section = {}
        ecm_sections = {}
        ecmSeqCount = 1

        for dictECMs in ECMSheet:
            EMC_ID = dictECMs['ECM Id/Domian ID'].strip() + '_' + str(millisecond)
            ECM_NAME = dictECMs['ECM Name/Domain Name'].strip()

            # Update section mapping
            section.update({dictECMs['section_id']: dictECMs['section_name']})
            ecm_sections[EMC_ID] = dictECMs['section_id']

            # Normalize 'Is ECM Mandatory?' field
            is_mandatory = dictECMs.get('Is ECM Mandatory?')
            if is_mandatory in ["TRUE", 1]:
                mandatory_flag = False
            elif is_mandatory in ["FALSE", 0]:
                mandatory_flag = True
            else:
                mandatory_flag = False

            ecm_update[EMC_ID] = {
                "externalId": EMC_ID,
                "tip": None,
                "name": ECM_NAME,
                "description": None,
                "modeOfCollection": "onfield",
                "canBeNotApplicable": mandatory_flag,
                "notApplicable": False,
                "canBeNotAllowed": mandatory_flag,
                "remarks": None,
                "sequenceNo": ecmSeqCount
            }

            print(ecm_update[EMC_ID])
            ecmSeqCount += 1

        ecm_dict['evidenceMethods'] = ecm_update
        bodySolutionUpdate = [ecm_dict,{"sections": section}]
        return bodySolutionUpdate

    def safe_decode(self, val):
        """Safely decode and strip UTF-8 strings."""
        if val is None:
            return None
        if isinstance(val, str):
            return val.strip()
        try:
            return str(val).encode('utf-8').decode('utf-8').strip()
        except Exception:
            return str(val).strip()

    def get_numeric_or_str(self,val):
        """Return int if numeric, else decoded string."""
        if val is None or val == "":
            return None
        try:
            if float(val).is_integer():
                return int(float(val))
            return float(val)
        except Exception:
            return self.safe_decode(val)

    def process_responses(self, ques, prefix='response', max_n=20):
        """Build R1–R20 response and hint pairs dynamically."""
        responses = {}
        for i in range(1, max_n + 1):
            key = f"{prefix}(R{i})"
            hint_key = f"{prefix}(R{i})_hint"
            responses[f"R{i}"] = self.get_numeric_or_str(ques.get(key))
            responses[f"R{i}-hint"] = self.get_numeric_or_str(ques.get(hint_key))
        return responses

    def process_scores(self, ques, max_n=20):
        """Build R1–R20 score mapping dynamically."""
        scores = {}
        for i in range(1, max_n + 1):
            scores[f"R{i}-score"] = ques.get(f"Score for R{i}")
        return scores

    def questionUpload(self, wbObservation, parentFolder, frameworkExternalId, millisAddObs, accessToken,
                   solutionId, typeofSolution, programdetails, userRole):

        # Prepare ECM and section mappings for solution types other than 2
        ecm_sections = {}
        ecmToSection = {}
        if typeofSolution != 2:
            if 'ecms or domains' not in wbObservation:
                self.errorVar.append("ECM/Domain sheet missing for this solution type.")
                return False

            ECMSheet = wbObservation['ecms or domains']
            section = {}
            for dictECMs in ECMSheet:
                EMC_ID = dictECMs['ECM Id/Domian ID'].strip() + '_' + str(millisAddObs)
                section[dictECMs['section_id']] = dictECMs['section_name']
                ecm_sections[EMC_ID] = dictECMs['section_id']

            # Prepare ECM to Section mapping
            ecmToSection = {ecm['section_id']: ecm['ECM Id/Domian ID'] for ecm in wbObservation['ecms or domains']}

        # Prepare criteria lookup
        criteriaLookUp = {}
        if typeofSolution == 2:
            # For type 2, build criteria lookup directly from questions
            for ques in wbObservation['questions']:
                criteriaKey = ques['criteria_id'].strip() + '_' + str(millisAddObs)
                criteriaLookUp[criteriaKey] = ques.get('criteria_name', 'Unknown Criteria')
        else:
            if 'framework' not in wbObservation:
                self.errorVar.append("Framework sheet missing for this solution type.")
                return False
            criteriaLookUp = {criteria['Criteria ID'].strip() + '_' + str(millisAddObs):
                            criteria['Criteria Name'] for criteria in wbObservation['framework']}

        # Sort questions by sequence
        questionsList = wbObservation['questions']
        questionsList.sort(key=lambda x: float(x.get('question_sequence') or 0))

        # Prepare CSV folder
        questionFilePath = os.path.join(parentFolder, 'questionUpload')
        os.makedirs(questionFilePath, exist_ok=True)
        uploadCSV = os.path.join(questionFilePath, 'uploadSheet.csv')
        file_exists_ques = os.path.isfile(uploadCSV)

        # CSV headers
        questionUploadFieldnames = [
            'solutionId','criteriaExternalId','name','evidenceMethod','section','instanceParentQuestionId',
            'hasAParentQuestion','parentQuestionOperator','parentQuestionValue','parentQuestionId','externalId',
            'question0','question1','tip','hint','instanceIdentifier','responseType','dateFormat','autoCapture',
            'validation','validationIsNumber','validationRegex','validationMax','validationMin','file','fileIsRequired',
            'fileUploadType','minFileCount','maxFileCount','allowAudioRecording','caption','questionGroup',
            'modeOfCollection','accessibility','showRemarks','rubricLevel','isAGeneralQuestion'
        ] + [f'R{i}' for i in range(1,21)] + [f'R{i}-hint' for i in range(1,21)] + [f'R{i}-score' for i in range(1,21)] + [
            'weightage','sectionHeader','page','questionNumber','_arrayFields','prefillFromEntityProfile',
            'isEditable','entityFieldName'
        ]

        questionSeqByEcmDict = {}

        # Open CSV to write
        with open(uploadCSV, 'a', encoding='utf-8', newline='') as questionUploadFile:
            writerQuestionUpload = csv.DictWriter(questionUploadFile, fieldnames=questionUploadFieldnames, lineterminator='\n')
            if not file_exists_ques:
                writerQuestionUpload.writeheader()

            for ques in questionsList:
                questionFileObj = {key: None for key in questionUploadFieldnames}

                observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
                questionFileObj['solutionId'] = observationExternalId
                questionFileObj['criteriaExternalId'] = ques['criteria_id'].strip() + '_' + str(millisAddObs)

                # Set Criteria Name
                try:
                    questionFileObj['name'] = criteriaLookUp[questionFileObj['criteriaExternalId']]
                except KeyError:
                    self.errorVar.append("criteria Id error....")
                    print(questionFileObj['criteriaExternalId'] + " not found.")
                    return False

                # Evidence method & section
                if typeofSolution in [1,5]:
                    evidence_key = ecmToSection.get(ques['section_id'], '') + "_" + str(millisAddObs)
                    section_key = ecm_sections.get(evidence_key, '')
                    questionFileObj['evidenceMethod'] = evidence_key
                    questionFileObj['section'] = ques['section_id']
                    questionSeqByEcmDict.setdefault(evidence_key, {}).setdefault(section_key, [])
                    questionSeqByEcmDict[evidence_key][section_key].append(
                        ques['question_id'].strip() + '_' + str(millisAddObs)
                    )
                elif typeofSolution == 2:
                    questionFileObj['evidenceMethod'] = "OB"
                    questionFileObj['section'] = "S1"
                    questionSeqByEcmDict.setdefault("OB", {}).setdefault("S1", [])
                    questionSeqByEcmDict["OB"]["S1"].append(ques['question_id'].strip() + '_' + str(millisAddObs))

                # Responses R1-R20
                for i in range(1,21):
                    questionFileObj[f'R{i}'] = ques.get(f'response(R{i})')
                    questionFileObj[f'R{i}-hint'] = ques.get(f'response(R{i})_hint')
                    questionFileObj[f'R{i}-score'] = ques.get(f'Score for R{i}')

                # Parent question handling
                if ques.get('instance_parent_question_id'):
                    questionFileObj['instanceParentQuestionId'] = ques['instance_parent_question_id'].strip() + '_' + str(millisAddObs)
                    questionFileObj['hasAParentQuestion'] = 'NO'
                else:
                    questionFileObj['instanceParentQuestionId'] = 'NA'

                if ques.get('parent_question_id'):
                    questionFileObj['hasAParentQuestion'] = 'YES'
                    op = ques.get('show_when_parent_question_value_is','').strip().upper()
                    if op in ['OR', '||']:
                        questionFileObj['parentQuestionOperator'] = '||'
                        questionFileObj['parentQuestionValue'] = ques.get('parent_question_value','').replace(" ","")
                    elif op == 'EQUALS':
                        questionFileObj['parentQuestionOperator'] = 'EQUALS'
                        questionFileObj['parentQuestionValue'] = ques.get('parent_question_value','').replace(" ","")
                    elif op == 'NOT_EQUALS_TO':
                        questionFileObj['parentQuestionOperator'] = '||'
                    else:
                        questionFileObj['parentQuestionOperator'] = ''
                    questionFileObj['parentQuestionId'] = ques['parent_question_id'].strip() + '_' + str(millisAddObs)

                # Basic question details
                questionFileObj['externalId'] = ques['question_id'].strip() + '_' + str(millisAddObs)
                questionFileObj['question0'] = ques.get('question_primary_language')
                questionFileObj['question1'] = ques.get('question_secondory_language')
                questionFileObj['tip'] = ques.get('question_tip')
                questionFileObj['hint'] = ques.get('question_hint')
                questionFileObj['instanceIdentifier'] = ques.get('instance_identifier')
                questionFileObj['responseType'] = ques.get('question_response_type','').strip().lower()
                questionFileObj['weightage'] = ques.get('question_weightage',0)
                questionFileObj['sectionHeader'] = ques.get('section_header')
                questionFileObj['page'] = ques.get('page')
                questionFileObj['questionNumber'] = int(ques['question_number']) if ques.get('question_number') else None

                # AutoCapture & date
                if questionFileObj['responseType'] == 'date':
                    questionFileObj['dateFormat'] = 'DD-MM-YYYY'
                    questionFileObj['autoCapture'] = 'TRUE' if ques.get('date_auto_capture') in [1,'true','True'] else 'FALSE'
                else:
                    questionFileObj['dateFormat'] = ''
                    questionFileObj['autoCapture'] = None

                # Validation
                questionFileObj['validation'] = 'TRUE' if ques.get('response_required') in [1,'true','True'] else 'FALSE'
                if questionFileObj['responseType'] in ['number','slider']:
                    questionFileObj['validationIsNumber'] = 'TRUE'
                    questionFileObj['validationRegex'] = 'isNumber'
                    questionFileObj['validationMax'] = ques.get('max_number_value') or (5 if questionFileObj['responseType']=='slider' else 10000)
                    questionFileObj['validationMin'] = ques.get('min_number_value') or 0

                # File upload
                if ques.get('file_upload') in [1,'TRUE','true']:
                    questionFileObj['file'] = 'Snapshot'
                    questionFileObj['fileIsRequired'] = 'TRUE'
                    questionFileObj['fileUploadType'] = 'image/jpeg,docx,pdf,ppt'
                    questionFileObj['minFileCount'] = 0
                    questionFileObj['maxFileCount'] = 10
                else:
                    questionFileObj['file'] = 'NA'
                    questionFileObj['fileIsRequired'] = 'FALSE'

                questionFileObj['_arrayFields'] = 'parentQuestionValue'
                questionFileObj['prefillFromEntityProfile'] = None
                questionFileObj['isEditable'] = 'TRUE'
                questionFileObj['entityFieldName'] = None
                questionFileObj['allowAudioRecording'] = False
                questionFileObj['caption'] = 'FALSE'
                questionFileObj['questionGroup'] = 'A1'
                questionFileObj['modeOfCollection'] = 'onfield'
                questionFileObj['accessibility'] = 'No'
                questionFileObj['showRemarks'] = 'TRUE' if ques.get('show_remarks') in [1,'TRUE','true'] else 'FALSE'
                questionFileObj['rubricLevel'] = None
                questionFileObj['isAGeneralQuestion'] = None

                # Write row
                writerQuestionUpload.writerow(questionFileObj)

        # Update solution sequence
        bodySolutionUpdate = {"questionSequenceByEcm": questionSeqByEcmDict}
        if not self.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate, programdetails, userRole):
            return False

        # Call Question Upload API
        try:
            urlQuestionsUploadApi = internal_kong_ip + questionuploadapiurl
            headerQuestionUploadApi = apiHeader.headers().headersQuestionUpload(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
            filesQuestion = {
                'questions': open(parentFolder + '/questionUpload/uploadSheet.csv', 'rb')
            }
            responseQuestionUploadApi = requests.post(url=urlQuestionsUploadApi, headers=headerQuestionUploadApi,
                                                    files=filesQuestion)
            messageArr = ["Solution Created from Framework.",
                        "URL : " + str(urlQuestionsUploadApi),
                        "Status Code : " + str(responseQuestionUploadApi.status_code),
                        "Response : " + str(responseQuestionUploadApi.text)]
            self.createAPILog(parentFolder, messageArr)
            messageArr = []
            if responseQuestionUploadApi.status_code == 200:
                print('QuestionUploadApi Success')
                with open(parentFolder + '/questionUpload/uploadInternalIdsSheet.csv','w+',
                        encoding='utf-8') as questionRes:
                    questionRes.write(responseQuestionUploadApi.text)
                return True
            else:
                self.errorVar.append(f"QuestionUploadApi Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}")
                return False
        except Exception as e:
            messageArr.append("Exception caught : " + str(e))
            self.createAPILog(parentFolder, messageArr)
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False

    def fetchSolutionCriteria(self, solutionName_for_folder_path, observationId, accessToken, programdetails, userRole):
        try:
            url = internal_kong_ip + ferchsolutioncriteria + observationId

            headers = apiHeader.headers().headersFetchSolutionCriteria(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)

            response = requests.request("POST", url, headers=headers)
            messageArr = ["Solution Created from Framework.",
                        "URL : " + str(url),
                        "Headers: "+ str(headers),
                        "Status Code : " + str(response.status_code),
                        "Response : " + str(response.text)]
            self.createAPILog(solutionName_for_folder_path, messageArr)

            os.mkdir(solutionName_for_folder_path + "/solutionCriteriaFetch/")
            if response.status_code == 200:
                print("Solution criteria fetched.")
                with open(solutionName_for_folder_path + "/solutionCriteriaFetch/solutionCriteriaDetails.csv",
                        'w+',encoding='utf-8') as solutionCriteriaFetch:
                    solutionCriteriaFetch.write(response.text)
                return True
            else:
                if response.status_code in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"QuestionUploadApi-Client Error {response.status_code}: {response.text}")
                elif response.status_code in [500, 502, 503, 504]:
                    self.errorVar.append(f"QuestionUploadApi-Server Error {response.status_code}: {response.text}")
                else:
                    self.errorVar.append(f"QuestionUploadApi-Unexpected Error {response.status_code}: {response.text}")
                messageArr = []
                messageArr = ["Criteria solution fetch API failed.", "Response  : " + str(response.text)]
                self.createAPILog(solutionName_for_folder_path, messageArr)
                self.errorVar.append("Solution criteria fetch failed. Status Code : " + str(response.status_code))
                return False
        except Exception as e:
            messageArr = []
            messageArr.append("Exception caught : " + str(e))
            self.createAPILog(solutionName_for_folder_path, messageArr)
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False

    def uploadCriteriaRubrics(self,parentFolder,wbObservation,millisecond,accessToken,frameworkExternalId,withRubricsFlag,programdetails, userRole):
        # --- Load appropriate sheet ---
        criteriaRubricSheet = (
            wbObservation['criteria_rubric-scoring']
            if withRubricsFlag
            else wbObservation['criteria']
        )

        # --- Load internal criteria lookup from CSV ---
        dictSolCritLookUp = {}
        filePath = os.path.join(parentFolder, "solutionCriteriaFetch", "solutionCriteriaDetails.csv")
        with open(filePath, 'r', encoding='utf-8') as criteriaInternalFile:
            reader = csv.DictReader(criteriaInternalFile)
            for crit in reader:
                dictSolCritLookUp[crit['criteriaID']] = [crit['criteriaInternalId'], crit['criteriaName']]

        # --- Detect valid L<number> SCORE columns ---
        criteriaLevelsCount = []
        if withRubricsFlag and criteriaRubricSheet:
            sample_row = next(iter(criteriaRubricSheet), {})
            for key in sample_row.keys():
                key_upper = key.upper().strip()
                if key_upper.startswith("L") and key_upper[1:].split()[0].isdigit() and "SCORE" in key_upper:
                    level_num = int(''.join(filter(str.isdigit, key_upper)))
                    if level_num not in criteriaLevelsCount:
                        criteriaLevelsCount.append(level_num)
            criteriaLevelsCount.sort()

        # --- Set CSV headers ---
        criteriaRubricUploadFieldnames = [
            "externalId",
            "name",
            "criteriaId",
            "weightage",
            "expressionVariables",
        ]
        if withRubricsFlag and criteriaLevelsCount:
            criteriaRubricUploadFieldnames += [f"L{cl}" for cl in criteriaLevelsCount]
        else:
            criteriaRubricUploadFieldnames.append("L1")

        # --- Create folder & file path ---
        criteriaRubricsDir = os.path.join(parentFolder, 'criteriaRubrics')
        os.makedirs(criteriaRubricsDir, exist_ok=True)
        uploadSheetPath = os.path.join(criteriaRubricsDir, 'uploadSheet.csv')

        # --- Write CSV cleanly ---
        with open(uploadSheetPath, 'w', encoding='utf-8', newline='') as csvfile:
            writer = csv.DictWriter(
                csvfile,
                fieldnames=criteriaRubricUploadFieldnames,
                lineterminator='\n'
            )
            writer.writeheader()

            if withRubricsFlag:
                for row in criteriaRubricSheet:
                    criteria_id = row.get('criteriaId', '').strip()
                    if not criteria_id:
                        continue

                    lookup_key = f"{criteria_id}_{millisecond}"
                    criteria_info = dictSolCritLookUp.get(lookup_key)

                    # Skip if no valid lookup
                    if not criteria_info or not criteria_info[0] or not criteria_info[1]:
                        print(f"⚠️ Skipping {criteria_id} — missing lookup info.")
                        continue

                    criteriaRubricUpload = {
                        'externalId': lookup_key,
                        'criteriaId': criteria_info[0].strip(),
                        'name': criteria_info[1].strip(),
                        'weightage': float(row.get('weightage', 1)),
                        'expressionVariables': f"SCORE={criteria_info[0]}.scoreOfAllQuestionInCriteria()"
                    }

                    # Add levels
                    if criteriaLevelsCount:
                        for cl in criteriaLevelsCount:
                            criteriaRubricUpload[f"L{cl}"] = str(row.get(f"L{cl} SCORE", "")).strip()
                    else:
                        criteriaRubricUpload["L1"] = str(row.get("L1 SCORE", "")).strip()

                    # Validate before writing
                    if not criteriaRubricUpload['criteriaId'] or not criteriaRubricUpload['name']:
                        print(f"⚠️ Skipping incomplete record: {criteriaRubricUpload}")
                        continue

                    writer.writerow(criteriaRubricUpload)

            else:
                # Non-rubric mode
                for criteriaIds, criteriaDetails in dictSolCritLookUp.items():
                    if not criteriaDetails[0] or not criteriaDetails[1]:
                        continue

                    criteriaRubricUpload = {
                        'externalId': criteriaIds,
                        'name': criteriaDetails[1],
                        'criteriaId': criteriaDetails[0],
                        'weightage': 1,
                        'expressionVariables': f"SCORE={criteriaDetails[0]}.scoreOfAllQuestionInCriteria()",
                        'L1': '0<=SCORE<=100000'
                    }

                    writer.writerow(criteriaRubricUpload)

        try:
            # --- Prepare API call ---
            urlCriteriaRubricUploadApi = (
                internal_kong_ip + criteriarubricuploadapiurl + frameworkExternalId + "-OBSERVATION-TEMPLATE"
            )
            headerCriteriaRubricUploadApi = apiHeader.headers().headersCriteriaRubricUpload(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)

            # --- Send file ---
            with open(uploadSheetPath, 'rb') as f:
                filesCriteriaRubric = {
                    'criteria': ('uploadSheet.csv', f, 'text/csv')
                }
                response = requests.post(
                    url=urlCriteriaRubricUploadApi,
                    headers=headerCriteriaRubricUploadApi,
                    files=filesCriteriaRubric
                )
            messageArr = []
            messageArr = ["Solution Created from Framework.",
                        "URL : " + str(urlCriteriaRubricUploadApi),
                        "Headers: "+ str(headerCriteriaRubricUploadApi),
                        "Status Code : " + str(response.status_code),
                        "Response : " + str(response.text)]
            self.createAPILog(parentFolder, messageArr)

            # --- Handle response ---
            if response.status_code == 200:
                output_path = os.path.join(criteriaRubricsDir, 'uploadInternalIdsSheet.csv')
                with open(output_path, 'w+', encoding='utf-8') as res_file:
                    res_file.write(response.text)
                print("✅ uploadCriteriaRubrics success.")
                return True
            else:
                msg = f"CriteriaRubricUploadApi Error {response.status_code}: {response.text}"
                self.errorVar.append(msg)
                print("❌", msg)
                return False

        except Exception as e:
            messageArr = []
            messageArr.append("Exception caught : " + str(e))
            self.createAPILog(parentFolder, messageArr)
            print("❌ Exception:", e)
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False
             
    def uploadThemeRubrics(self, parentFolder, wbObservation, accessToken, frameworkExternalId, withRubricsFlag, programdetails, userRole):
        # --- Detect appropriate rubric sheet ---
        if withRubricsFlag:
            themeRubricSheet = wbObservation['theme_rubric_scoring']
        else:
            themeRubricSheet = wbObservation.get('Domain(theme)', [])

        # --- Detect valid L<number> columns (ignore LN etc.) ---
        criteriaLevels = []
        if withRubricsFlag and themeRubricSheet:
            sample_row = next(iter(themeRubricSheet), {})
            for key in sample_row.keys():
                key_upper = key.upper().strip()
                if key_upper.startswith("L") and key_upper[1:].split()[0].isdigit() and "SCORE" not in key_upper:
                    # For safety: also allow just "L1", "L2" style
                    level_num = int(''.join(filter(str.isdigit, key_upper)))
                    if level_num not in criteriaLevels:
                        criteriaLevels.append(level_num)
            criteriaLevels.sort()

        # --- Set up CSV folder and headers ---
        themeRubricUploadFieldnames = ["externalId", "name", "weightage"]
        if withRubricsFlag and criteriaLevels:
            for cl in criteriaLevels:
                themeRubricUploadFieldnames.append(f"L{cl}")
        else:
            themeRubricUploadFieldnames.append("L1")

        themeRubricsFilePath = os.path.join(parentFolder, "themeRubrics")
        os.makedirs(themeRubricsFilePath, exist_ok=True)
        uploadSheetPath = os.path.join(themeRubricsFilePath, 'uploadSheet.csv')

        # --- Write data to CSV ---
        file_exists = os.path.isfile(uploadSheetPath)
        with open(uploadSheetPath, 'a', encoding='utf-8', newline='') as themeRubricsUploadFile:
            writer = csv.DictWriter(themeRubricsUploadFile, fieldnames=themeRubricUploadFieldnames, lineterminator='\n')
            if not file_exists:
                writer.writeheader()

            # --- Rubric mode ---
            if withRubricsFlag:
                for row in themeRubricSheet:
                    themeRubricUpload = {}
                    domain_id = row.get('domain_Id', '').strip()
                    domain_name = row.get('domain_name', '').strip()
                    if not domain_id:
                        continue

                    themeRubricUpload['externalId'] = domain_id
                    themeRubricUpload['name'] = domain_name
                    themeRubricUpload['weightage'] = row.get('weightage', 0)

                    if criteriaLevels:
                        for cl in criteriaLevels:
                            themeRubricUpload[f"L{cl}"] = row.get(f"L{cl}", "")
                    else:
                        themeRubricUpload["L1"] = row.get("L1", "0<=SCORE<=100000")

                    writer.writerow(themeRubricUpload)

            # --- Non-rubric mode ---
            else:
                themeRubricUpload = {
                    "externalId": "OB",
                    "name": "Observation Theme",
                    "weightage": 1,
                    "L1": "0<=SCORE<=100000"
                }
                writer.writerow(themeRubricUpload)

        # --- Upload to API ---
        try:
            urlThemeRubricUploadApi = internal_kong_ip + themerubricuploadapiurl + frameworkExternalId + "-OBSERVATION-TEMPLATE"
            headerThemeRubricUploadApi = apiHeader.headers().headersThemeRubricUpload(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
            
            filesThemeRubric = {
                'themes': open(uploadSheetPath, 'rb')
            }

            response = requests.post(url=urlThemeRubricUploadApi, headers=headerThemeRubricUploadApi, files=filesThemeRubric)
            messageArr = []
            messageArr = ["Solution Created from Framework.",
                        "URL : " + str(urlThemeRubricUploadApi),
                        "Headers: "+ str(headerThemeRubricUploadApi),
                        "Status Code : " + str(response.status_code),
                        "Response : " + str(response.text)]
            self.createAPILog(parentFolder, messageArr)

            if response.status_code == 200:
                print('ThemeRubricUploadApi Success')
                with open(os.path.join(themeRubricsFilePath, 'uploadInternalIdsSheet.csv'), 'w+', encoding='utf-8') as themeRubricRes:
                    themeRubricRes.write(response.text)
                return True
            else:
                self.errorVar.append(f"ThemeRubricUploadApi Error {response.status_code}: {response.text}")
                return False

        except Exception as e:
            messageArr = []
            messageArr.append("Exception caught : " + str(e))
            self.createAPILog(parentFolder, messageArr)
            self.errorVar.append(f"Error occurred: {str(e)}")
            print(f"Error occurred: {str(e)}")
            return False

    def solutionUpdate(self, solutionName_for_folder_path, accessToken, solutionId, bodySolutionUpdate,programdetails, userRole):
        try:
            solutionUpdateApi = internal_kong_ip + solutionupdateapi + str(solutionId)
            headerUpdateSolutionApi = apiHeader.headers().headersObservationsolutionUpdate(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
            responseUpdateSolutionApi = requests.post(url=solutionUpdateApi, headers=headerUpdateSolutionApi,data=json.dumps(bodySolutionUpdate))
            messageArr = []
            messageArr = ["Solution Created from Framework.",
                        "URL : " + str(solutionUpdateApi),
                        "Headers: "+ str(headerUpdateSolutionApi),
                        "bodySolutionUpdate: "+str(bodySolutionUpdate),
                        "Status Code : " + str(responseUpdateSolutionApi.status_code),
                        "Response : " + str(responseUpdateSolutionApi.text)]
            self.createAPILog(solutionName_for_folder_path, messageArr)
            if responseUpdateSolutionApi.status_code == 200:
                print("Solution Update Success.")
                return True
            else:
                if responseUpdateSolutionApi.status_code in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"UpdateSolutionApi-Client Error {responseUpdateSolutionApi.status_code}: {responseUpdateSolutionApi.text}")
                elif responseUpdateSolutionApi.status_code in [500, 502, 503, 504]:
                    self.errorVar.append(f"UpdateSolutionApi-Server Error {responseUpdateSolutionApi.status_code}: {responseUpdateSolutionApi.text}")
                else:
                    self.errorVar.append(f"UpdateSolutionApi-Unexpected Error {responseUpdateSolutionApi.status_code}: {responseUpdateSolutionApi.text}")
                # ElevateObservation.createAPILog(solutionName_for_folder_path, errorVar)
                return False
            
        except Exception as e:
            messageArr = []
            messageArr.append("Exception caught : " + str(e))
            self.createAPILog(solutionName_for_folder_path, messageArr)
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False
        
    def UpdateCertForSolution(self, solutionName_for_folder_path, childTemplateId, childSolutionId, accessToken,programdetails, userRole):
        try:
            urldbFind = elevateprojecthost + dbfindapi_url
            headers = {
                'X-auth-token': accessToken,
                'Content-Type': content_type
            }
            searchSolutionpayload = {
                "query": {"_id": childSolutionId},
                "projection": ["status", "name"],
                "mongoIdKeys": ["_id"],
                "limit": 10000
            }
            response = requests.post(urldbFind, headers=headers, json=searchSolutionpayload)
            results = response.json().get("result", [])
            if not results:
                self.errorVar = "No solutions found for name: "
                return False

            projectSolutionName = results[0].get("name")
            if response.status_code != 200:
                self.errorVar = f"DBFind-Error {response.status_code}: {response.text}"
                self.createAPILog(solutionName_for_folder_path, response.text)
                print("Unable to fetch Solution...")
                return False
            
            urldbFind = elevateprojecthost + dbfindapi_url
            headers = {
                'X-auth-token': accessToken,
                'Content-Type': content_type
            }
            searchSolutionpayload = {
                "query": {"name": projectSolutionName,
                          "status": "inactive"},
                "projection": ["status", "name", "certificateTemplateId"],
                "mongoIdKeys": ["_id"],
                "limit": 10000
            }
            response = requests.post(urldbFind, headers=headers, json=searchSolutionpayload)
            results = response.json().get("result", [])
            if not results:
                self.errorVar = "No solutions found for name: "
                return False
            if results[0].get("certificateTemplateId"):
                certificateTemplateId = results[0].get("certificateTemplateId") 
                print(certificateTemplateId,"certificateTemplateId")
                urldbFindPT = elevateprojecthost + solutionupdateapi + childSolutionId
                headerUpdateSolutionApi = apiHeader.headers().headersObservationsolutionUpdate(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
                searchSolutionpayloadPT = {
                    "certificateTemplateId": certificateTemplateId
                }
                responseSolutionUpdate = requests.post(urldbFindPT, headers=headerUpdateSolutionApi, json=searchSolutionpayloadPT)

                if responseSolutionUpdate.status_code == 200:
                    print("Child Solution Update Success.")
                else:
                    print("Child Solution Update Failed.")
                    return False

                urldbFindCPT = elevateprojecthost + projectTemplateupdateapi + childTemplateId
                headerUpdateSolutionApi = apiHeader.headers().headersObservationsolutionUpdate(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
                searchSolutionpayloadPT = {
                    "certificateTemplateId": certificateTemplateId
                }
                responseSolutionUpdate = requests.post(urldbFindCPT, headers=headerUpdateSolutionApi, json=searchSolutionpayloadPT)

                if responseSolutionUpdate.status_code == 200:
                    print("Child Solution Update Success.")
                else:
                    print("Child Solution Update Failed.")
                    return False
            if response.status_code != 200:
                self.errorVar = f"DBFind-Error {response.status_code}: {response.text}"
                self.createAPILog(solutionName_for_folder_path, response.text)
                print("Unable to fetch Solution...")
                return False
            return True
        except Exception as e:
            self.errorVar = f"Exception in UpdateCertForSolution: {str(e)}"
            print(self.errorVar, "---> API-Error")
            return False

    def createChild(self, parentFolder,wbObservation, observationExternalId, accessToken, programdetails, userRole):
        entitydetails = programdetails.get("entitiesType")
        entity_type = entitydetails[0] if isinstance(entitydetails[0], str) else str(entitydetails[0][0])
        observationChildId = None
        try:
            childObservationExternalId = str(observationExternalId + "_CHILD")
            urlSol_prog_mapping = internal_kong_ip + solutiontoprogrammappingapiurl + "?solutionId=" + observationExternalId + "&entityType=" + entity_type
            ObsDict = wbObservation.get('details')
            # isExternalProgram == 'true'
            payloadSol_prog_mapping = {
                "externalId": childObservationExternalId,
                "name": ObsDict.get("observation_solution_name"),
                "description": ObsDict.get("observation_solution_description"),
                "programExternalId": programdetails.get("_id")
            }
            
            headersSol_prog_mapping = apiHeader.headers().headersSolutionToProgramMapping(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
            
            responseSol_prog_mapping = requests.request("POST", urlSol_prog_mapping, headers=headersSol_prog_mapping,
                                                        data=json.dumps(payloadSol_prog_mapping))
            messageArr = []
            messageArr = ["Create child API called.", "URL : " + urlSol_prog_mapping,
                        "Status code : " + str(responseSol_prog_mapping.status_code),
                        "Response : " + responseSol_prog_mapping.text, "body : " + str(payloadSol_prog_mapping)]
            self.createAPILog(parentFolder, messageArr)

            if responseSol_prog_mapping.status_code == 200:
                if programdetails.get('TitleoftheProgram') :
                    print("Solution mapped to program : " + programdetails.get('TitleoftheProgram'))
                print("Child solution : " + childObservationExternalId)

                responseSol_prog_mapping = responseSol_prog_mapping.json()
                child_id = responseSol_prog_mapping['result']['_id']
                solutionDetails = responseSol_prog_mapping['result']['projectTemplateDetails']
                for sol in solutionDetails:
                    childTemplateId = sol.get('childProjectTemplateId')
                    childSolutionId = sol.get('solutionId')
                    self.UpdateCertForSolution(parentFolder, childTemplateId, childSolutionId, accessToken, programdetails, userRole)

                self.createAPILog(parentFolder, messageArr)
                print("child solutionId: " + child_id)
                return [child_id, childObservationExternalId]
            else:
                if responseSol_prog_mapping.status_code in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"Sol_prog_mapping-Client Error {responseSol_prog_mapping.status_code}: {responseSol_prog_mapping.text}")
                elif responseSol_prog_mapping.status_code in [500, 502, 503, 504]:
                    self.errorVar.append(f"Sol_prog_mapping-Server Error {responseSol_prog_mapping.status_code}: {responseSol_prog_mapping.text}")
                else:
                    self.errorVar.append(f"Sol_prog_mapping-Unexpected Error {responseSol_prog_mapping.status_code}: {responseSol_prog_mapping.text}")
                print("Unable to create child solution")
                return False
        except Exception as e:
            messageArr = []
            messageArr.append("Exception caught : " + str(e))
            self.createAPILog(parentFolder, messageArr)
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False
            
    def fetchSolutionDetailsFromProgramSheet(self, solutionName_for_folder_path, programdetails, solutionId, accessToken, ProgramGlobalDict, userRole):
        try:
            urlFetchSolutionApi = internal_kong_ip + dbfindapi_url
            headerFetchSolutionApi = apiHeader.headers().headersFetchSolutionDetails(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
            payloadFetchSolutionApi = json.dumps({
                    "query": {
                        "_id": solutionId
                    },
                    "mongoIdKeys": [
                        "_id","name", "externalId"
                    ],
                    "limit": 10000
                })
            responseFetchSolutionApiUrl = requests.post(url=urlFetchSolutionApi, headers=headerFetchSolutionApi,
                                                    data=payloadFetchSolutionApi)
            responseFetchSolutionJson = responseFetchSolutionApiUrl.json()
            messageArr = []
            messageArr = ["Solution Fetch Link.",
                        "solution name : " + responseFetchSolutionJson["result"][-1]["name"],
                        "solution ExternalId : " + responseFetchSolutionJson["result"][-1]["externalId"],
                        "Upload status code : " + str(responseFetchSolutionApiUrl.status_code),
                        "Response : " + str(responseFetchSolutionApiUrl.text)]
            self.createAPILog(solutionName_for_folder_path, messageArr)
            if responseFetchSolutionApiUrl.status_code == 200:
                solutionName = responseFetchSolutionJson["result"][-1]["name"]
                print(solutionName,"solutionName")
                resourceDetailsSheet = ProgramGlobalDict.get('Program Resources')
                for solutions in resourceDetailsSheet:
                    if solutionName == solutions.get('Nameofresourcesinprogram'):
                        solutionMainRole = solutions.get('Targetroleattheresourcelevel')
                        solutionRolesArray = solutions.get('Targetedsubroleatresourcelevel')
                        solutionStartDate = solutions.get('Startdateofresource')
                        solutionEndDate = solutions.get('Enddateofresource')
                        return [solutionMainRole,solutionRolesArray, solutionStartDate, solutionEndDate]
            else:
                if responseFetchSolutionApiUrl.status_code in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"FetchSolutionApiUrl-Client Error {responseFetchSolutionApiUrl.status_code}: {responseFetchSolutionApiUrl.text}")
                elif responseFetchSolutionApiUrl.status_code in [500, 502, 503, 504]:
                    self.errorVar.append(f"FetchSolutionApiUrl-Server Error {responseFetchSolutionApiUrl.status_code}: {responseFetchSolutionApiUrl.text}")
                else:
                    self.errorVar.append(f"FetchSolutionApiUrl-Unexpected Error {responseFetchSolutionApiUrl.status_code}: {responseFetchSolutionApiUrl.text}")
                return False
        except Exception as e:
            messageArr = []
            messageArr.append("Exception caught : " + str(e))
            self.createAPILog(solutionName_for_folder_path, messageArr)
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False

    def validate_roles_against_api(self, mainRoles, subRoles, programdetails, parentFolder):
        urlFetchRoleList = userLoginHost + fetchprofessionalRole
        headers = apiHeader.headers().header_validate_roles_against_api(programdetails.get('TenantID'))
        payload = {}

        response = requests.request("GET", urlFetchRoleList, headers=headers, data=payload)
        messageArr = []
        messageArr = ["Solution Fetch Link.",
                        "solution name : " + str(urlFetchRoleList),
                        "headers: " + str(headers),
                        "Upload status code : " + str(response.status_code),
                        "Response : " + str(response.text)]
        self.createAPILog(parentFolder, messageArr)

        if response.status_code != 200:
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
                            messageArr.append(f"Subrole '{sub_external_id}' validated under mainRole '{role}'")
                            print(f"Subrole '{sub_external_id}' validated under mainRole '{role}'")
                else:
                    messageArr.append(f"Failed to fetch subroles for mainRole '{role}'")
                    self.errorVar.append(f"Failed to fetch subroles for mainRole '{role}'")
            else:
                messageArr.append(f"MainRole '{role}' not found in API")
                self.errorVar.append(f"MainRole '{role}' not found in API")

        for s in remaining_subroles:
            messageArr.append(f"Subrole '{s}' not found in any of the provided mainRoles.")
            self.errorVar.append(f"Subrole '{s}' not found in any of the provided mainRoles.")
        # Programs.createAPILog(parentFolder, messageArr)
        for err in self.errorVar:
            print(err)
        return validated_main_role_ids, validated_subrole_ids_list 
    
    def convert_to_date(self, date_str):
        return datetime.strptime(date_str, "%d-%m-%Y")
    
    def prepareProgramSuccessSheet(self, MainFilePath, solutionName_for_folder_path, programFile, solutionExternalId, solutionId,accessToken, programdetails, userRole):
        urlFetchSolutionApi = internal_kong_ip + dbfindapi_url
        headerFetchSolutionApi = apiHeader.headers().headersFetchSolutionDetails(programdetails.get('TenantID'), programdetails.get('Org ID'), accessToken, userRole)
        
        payloadFetchSolutionApi = json.dumps({
            "query": {
                "_id": solutionId
            },
            "mongoIdKeys": [
                "_id",
                "solutionId",
                "metaInformation.solutionId"
            ],
            "limit": 10000
        })
        responseFetchSolutionApi = requests.request("POST", url=urlFetchSolutionApi, headers=headerFetchSolutionApi,
                                                data=payloadFetchSolutionApi)
        print(responseFetchSolutionApi.text, "responseFetchSolutionApi")
        responseFetchSolutionJson = responseFetchSolutionApi.json()
        messageArr = []
        messageArr.append("Solution Fetch Link.")
        messageArr.append("solution name : " + responseFetchSolutionJson["result"][-1]["name"])
        messageArr.append("solution ExternalId : " + responseFetchSolutionJson["result"][-1]["externalId"])
        messageArr.append("Upload status code : " + str(responseFetchSolutionApi.status_code))
        self.createAPILog(solutionName_for_folder_path, messageArr)

        if responseFetchSolutionApi.status_code == 200:
            print('Fetch solution Api Success')
            solutionName = responseFetchSolutionJson["result"][-1]["name"]
        urlFetchSolutionLinkApi = internal_kong_ip + fetchlink + solutionId
        print(urlFetchSolutionLinkApi,"urlFetchSolutionLinkApi")
        headerFetchSolutionLinkApi = apiHeader.headers().headersFetchSolutionLink(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
        
        payloadFetchSolutionLinkApi = {}

        responseFetchSolutionLinkApi = requests.get(url=urlFetchSolutionLinkApi, headers=headerFetchSolutionLinkApi,
                                                    data=payloadFetchSolutionLinkApi)
        
        print(responseFetchSolutionLinkApi.text, "responseFetchSolutionLinkApi")
        messageArr = ["Solution Fetch Link.","solution id : " + solutionId,"solution ExternalId : " + solutionExternalId]
        messageArr.append("Upload status code : " + str(responseFetchSolutionLinkApi.status_code))
        self.createAPILog(solutionName_for_folder_path, messageArr)
        if responseFetchSolutionLinkApi.status_code == 200:
            print(responseFetchSolutionLinkApi.text,"responseFetchSolutionLinkApi")
            print('Fetch solution Link Api Success')
            responseProjectUploadJson = responseFetchSolutionLinkApi.json()
            solutionLink = responseProjectUploadJson["result"]
            solutionLink = ','.join(solutionLink)
            print(solutionLink,"solutionLink")
            messageArr = []
            messageArr.append("Response : " + str(responseFetchSolutionLinkApi.text))
            self.createAPILog(solutionName_for_folder_path, messageArr)
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
            self.solutionUpdate(solutionName_for_folder_path, accessToken, solutionId, {"status": "active", "isDeleted": False}, programdetails, userRole)
            return solutionLink
        
    def ObservationSolutionCreate(self, resource, parentFolder, accessToken, ProgramGlobalDict,programdetails, MainFilePath, programFile, userRole):
        finalObsRubricSolutionLink = ""
        if resource.get('typeofSolution') == 5:
            impLedObsFlag = True
        else:
            impLedObsFlag = False
        wbObservation = resource.get("ResourceCre")
        millisecond = int(time.time() * 1000)
        if resource.get('typeofSolution') == 1 or resource.get('typeofSolution') == 5:
            if not self.criteriaUpload(parentFolder, resource.get("ResourceCre"), millisecond, accessToken, "framework", impLedObsFlag, programdetails, userRole):
                self.errorVar.append("Criteria Upload Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            print("Criteria Upload success....")
            frameworkExternalId = self.frameWorkUpload(parentFolder, resource.get("ResourceCre"), millisecond, accessToken, programdetails, resource.get('typeofSolution'), userRole)
            if not frameworkExternalId:
                self.errorVar.append("Framework Upload Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            print(frameworkExternalId,"frameworkExternalId")
            observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
            if not self.themesUpload(parentFolder, resource.get("ResourceCre"), millisecond, accessToken, frameworkExternalId, False, programdetails, userRole):
                self.errorVar.append("Theme Upload Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            solutionId = self.createSolutionFromFramework(parentFolder, accessToken, frameworkExternalId, programdetails, userRole)
            if not solutionId:
                self.errorVar.append("Unable to create Solution From Framework.")
                return finalObsRubricSolutionLink, self.errorVar
            print(solutionId,"solutionId")
            print("parent solution created....")
            ECMSheet = wbObservation['ecms or domains']
            ECMUpdate = self.ECMUpdatebody(ECMSheet,millisecond)
            if not self.solutionUpdate(parentFolder, accessToken, solutionId, ECMUpdate[0],programdetails, userRole):
                self.errorVar.append("Solution Update Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            if not self.solutionUpdate(parentFolder, accessToken, solutionId, ECMUpdate[1],programdetails, userRole):
                self.errorVar.append("Solution Update Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            entitydetails = programdetails.get("entitiesType")
            Entity_To_Upload = entitydetails[0] if isinstance(entitydetails[0], str) else str(entitydetails[0][0])
            if Entity_To_Upload.strip().lower() in ['state', 'district', 'block', 'cluster', 'school']:
                parentEntityKey = "state"
            else:
                parentEntityKey = None
            detailsSheet = wbObservation['details']
            if not detailsSheet.get("scoring_system").lower() == "null":
                criteriaLevelsReport = True
            else:
                criteriaLevelsReport = False
            bodySolutionUpdate = {"status": "active", "isDeleted": False, "criteriaLevelReport": criteriaLevelsReport,"parentEntityKey": parentEntityKey}
            if not self.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate, programdetails, userRole):
                self.errorVar.append("Solution Update Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            if not self.questionUpload(wbObservation, parentFolder, frameworkExternalId, millisecond, accessToken,solutionId,resource.get('typeofSolution'), programdetails, userRole):
                self.errorVar.append("question Upload Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            if not programdetails.get('scoring_system') == "null":
                bodySolutionUpdate = {"isRubricDriven": True}
                if not self.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate, programdetails, userRole):
                    self.errorVar.append("Solution Update Failed.")
                    return finalObsRubricSolutionLink, self.errorVar
                if not self.fetchSolutionCriteria(parentFolder, observationExternalId, accessToken, programdetails, userRole):
                    self.errorVar.append("fetchSolutionCriteria Failed.")
                    return finalObsRubricSolutionLink, self.errorVar
                if not self.uploadCriteriaRubrics(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, True, programdetails, userRole):
                    self.errorVar.append("upload Criteria Rubrics Failed.")
                    return finalObsRubricSolutionLink, self.errorVar
                if not self.uploadThemeRubrics(parentFolder, wbObservation, accessToken, frameworkExternalId, True, programdetails, userRole):
                    self.errorVar.append("upload Theme Rubrics Failed.")
                    return finalObsRubricSolutionLink, self.errorVar
            else:
                print("Observation with scoring system : null.")
            allow_multiple_submissions = programdetails.get('allow_multiple_submissions')
            if allow_multiple_submissions == 1 or allow_multiple_submissions == 'TRUE':
                allow_multiple_submissions = True
            else:
                allow_multiple_submissions = False
            bodySolutionUpdate = {'allowMultipleAssessemts': allow_multiple_submissions, "creator": programdetails.get('Name_of_the_creator'),"parentEntityKey": parentEntityKey}
            if not self.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate, programdetails, userRole):
                self.errorVar.append("upload Theme Rubrics Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            childId = self.createChild(parentFolder, wbObservation, solutionId, accessToken, programdetails, userRole)
            if not childId:
                self.errorVar.append("upload Theme Rubrics Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            if childId[0]:
                print("Fetching solution details")
                solutionDetails = self.fetchSolutionDetailsFromProgramSheet(
                    parentFolder, programdetails, childId[0], accessToken, ProgramGlobalDict, userRole
                )
                print(solutionDetails, "solutionDetails")
                if not solutionDetails:
                    self.errorVar.append("Fetch solution details API Failed.")
                    return finalObsRubricSolutionLink, self.errorVar
                self.solutionUpdate(parentFolder, accessToken, childId[0], {"status": "inactive", "isDeleted": True}, programdetails, userRole)
                scopeRoles = solutionDetails[0]
                scopeSubRoles = solutionDetails[1]
                # ✅ Fix: ensure roles are lists
                if isinstance(scopeRoles, str):
                    scopeRoles = [scopeRoles]
                if isinstance(scopeSubRoles, str):
                    scopeSubRoles = [scopeSubRoles]
                verifiedRoles = self.validate_roles_against_api(
                    scopeRoles, scopeSubRoles, programdetails, parentFolder
                )
                mainRoleproff = verifiedRoles[0]
                rolesPGMID = verifiedRoles[1]
                print("mainRole", mainRoleproff)
                print("rolesPGMID--------22", rolesPGMID)
                entities = programdetails.get('entitiesType')
                entitiesPGMID = entities[1]
                entitiesType = entities[0]
                scopeEntities = entitiesPGMID
                print("scopeEntities", scopeEntities)
                print("entitiesType", entitiesType)
                scope = {}
                entityHierarchy = programdetails.get('entityHierarchy')
                scope.update(entityHierarchy)
                scope["organizations"] = programdetails.get('OrgID', '')
                scope["professional_subroles"] = rolesPGMID
                scope["professional_role"] = mainRoleproff
                bodySolutionUpdate = {
                "scope": scope
                }
                if not self.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate, programdetails, userRole):
                    self.errorVar.append("Solution Update Failed.")
                    return finalObsRubricSolutionLink, self.errorVar
                print("solution update success.")
                ReffstartDateOfProgram1 = self.convert_to_date(programdetails.get('Startdateofprogram'))
                ReffendDateOfProgram1 = self.convert_to_date(programdetails.get('Enddateofprogram'))
                solutionDetails2 = self.convert_to_date(solutionDetails[2])
                solutionDetails3 = self.convert_to_date(solutionDetails[3])
                if ReffstartDateOfProgram1 <= solutionDetails2 <= ReffendDateOfProgram1 and ReffstartDateOfProgram1 <= solutionDetails3 <= ReffendDateOfProgram1:
                    print("dates validated...")
                    if solutionDetails[2]:
                        startDateArr = str(solutionDetails[2]).split("-")
                        bodySolutionUpdate = {
                            "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                        if not self.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate, programdetails, userRole):
                            self.errorVar.append("Solution Update Failed.")
                            return finalObsRubricSolutionLink, self.errorVar
                    if solutionDetails[3]:
                        endDateArr = str(solutionDetails[3]).split("-")
                        bodySolutionUpdate = {
                            "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                        if not self.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate, programdetails, userRole):
                            self.errorVar.append("Solution Update Failed.")
                            return finalObsRubricSolutionLink, self.errorVar
                else:
                    self.errorVar.append("Date Mismatched! Creation Stopped.")
                    print("Date Mismatched! Creation Stopped")
                    return finalObsRubricSolutionLink, self.errorVar
                ObsRubricSolutionLink = self.prepareProgramSuccessSheet(MainFilePath, parentFolder, programFile, childId[1], childId[0],
                                        accessToken, programdetails, userRole)
                if not ObsRubricSolutionLink:
                    self.errorVar.append("Solution Fetch Link Failed.")
                    return finalObsRubricSolutionLink, self.errorVar
                else:
                    print(ObsRubricSolutionLink,"ObsRubricSolutionLink")
                    finalObsRubricSolutionLink =  ObsRubricSolutionLink
                    return finalObsRubricSolutionLink, self.errorVar   
        elif resource.get('typeofSolution') == 2:
            finalObsSolutionLink = ""
            if not self.criteriaUpload(parentFolder, resource.get("ResourceCre"), millisecond, accessToken, "criteria", False, programdetails, userRole):
                self.errorVar.append("Criteria Upload Failed.")
                return finalObsSolutionLink,self.errorVar
            print("Criteria Upload success....")
            frameworkExternalId = self.frameWorkUpload(parentFolder, resource.get("ResourceCre"), millisecond, accessToken, programdetails, resource.get('typeofSolution'), userRole)
            if not frameworkExternalId:
                self.errorVar.append("Framework Upload Failed.")
                return finalObsSolutionLink,self.errorVar
            observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
            if not self.themesUpload(parentFolder, resource.get("ResourceCre"), millisecond, accessToken, frameworkExternalId, True, programdetails, userRole):
                self.errorVar.append("Theme Upload Failed.")
                return finalObsSolutionLink,self.errorVar
            solutionId = self.createSolutionFromFramework(parentFolder, accessToken, frameworkExternalId, programdetails, userRole)
            if not solutionId:
                self.errorVar.append("Unable to create Solution From Framework.")
                return finalObsSolutionLink,self.errorVar
            sectionsObj = {"sections": {'S1': 'Observation Question'}}
            if not self.solutionUpdate(parentFolder, accessToken, solutionId, sectionsObj, programdetails, userRole):
                self.errorVar.append("Solution Update Failed.")
                return finalObsSolutionLink,self.errorVar
            ecmObj = {}
            ecmExternalId = None
            ecmObj = {
                "evidenceMethods": {'OB': {'externalId': 'OB', 'tip': None, 'name': 'Observation', 'description': None,
                                        'modeOfCollection': 'onfield', 'canBeNotApplicable': False,
                                        'notApplicable': False, 'canBeNotAllowed': False, 'remarks': None}}}
            if not self.solutionUpdate(parentFolder, accessToken, solutionId, ecmObj, programdetails, userRole):
                self.errorVar.append("Solution Update Failed.")
                return finalObsSolutionLink,self.errorVar
            if not self.questionUpload(wbObservation, parentFolder, frameworkExternalId, millisecond, accessToken,solutionId,resource.get('typeofSolution'), programdetails, userRole):
                self.errorVar.append("question Upload Failed.")
                return finalObsSolutionLink,self.errorVar
            if not programdetails.get('scoring_system') == None:
                if not self.uploadCriteriaRubrics(parentFolder, wbObservation, millisecond, accessToken, frameworkExternalId, False, programdetails, userRole):
                    self.errorVar.append("upload Criteria Rubrics Failed.")
                    return finalObsSolutionLink,self.errorVar
                if not self.uploadThemeRubrics(parentFolder, wbObservation, accessToken, frameworkExternalId, False, programdetails, userRole):
                    self.errorVar.append("upload Theme Rubrics Failed.")
                    return finalObsSolutionLink,self.errorVar
            else:
                print("Observation with scoring system : null.")
            entitydetails = programdetails.get("entitiesType")
            Entity_To_Upload = entitydetails[0] if isinstance(entitydetails[0], str) else str(entitydetails[0][0])
            if Entity_To_Upload.strip().lower() in ['state', 'district', 'block', 'cluster', 'school']:
                parentEntityKey = "state"
            else:
                parentEntityKey = None
            detailsSheet = wbObservation['details']
            bodySolutionUpdate = {"status": "active", "isDeleted": False, "allowMultipleAssessemts": True,
                                "creator": programdetails.get('Name_of_the_creator'),"parentEntityKey": parentEntityKey}
            if not self.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate, programdetails, userRole):
                self.errorVar.append("Solution Update Failed.")
                return finalObsSolutionLink,self.errorVar
            childId = self.createChild(parentFolder, wbObservation, solutionId, accessToken, programdetails, userRole)
            if not childId:
                self.errorVar.append("upload Theme Rubrics Failed.")
                return finalObsSolutionLink,self.errorVar
            if childId[0]:
                print("Fetching solutions details")
                solutionDetails = self.fetchSolutionDetailsFromProgramSheet(
                    parentFolder, programdetails, childId[0], accessToken, ProgramGlobalDict, userRole
                )
                print(solutionDetails, "solutionDetails")
                if not solutionDetails:
                    self.errorVar.append("Fetch solution details API Failed.")
                    return finalObsSolutionLink,self.errorVar
                self.solutionUpdate(parentFolder, accessToken, childId[0], {"status":"inactive", "isDeleted": True}, programdetails, userRole)
                scopeRoles = solutionDetails[0]
                scopeSubRoles = solutionDetails[1]
                # ✅ Fix: ensure roles are lists
                if isinstance(scopeRoles, str):
                    scopeRoles = [scopeRoles]
                if isinstance(scopeSubRoles, str):
                    scopeSubRoles = [scopeSubRoles]
                verifiedRoles = self.validate_roles_against_api(
                    scopeRoles, scopeSubRoles, programdetails, parentFolder
                )
                mainRoleproff = verifiedRoles[0]
                rolesPGMID = verifiedRoles[1]
                print("mainRole", mainRoleproff)
                print("rolesPGMID--------22", rolesPGMID)
                entities = programdetails.get('entitiesType')
                entitiesPGMID = entities[1]
                entitiesType = entities[0]
                scopeEntities = entitiesPGMID
                print("scopeEntities", scopeEntities)
                print("entitiesType", entitiesType)
                scope = {}
                entityHierarchy = programdetails.get('entityHierarchy')
                scope.update(entityHierarchy)
                scope["organizations"] = programdetails.get('OrgID', '')
                scope["professional_subroles"] = rolesPGMID
                scope["professional_role"] = mainRoleproff
                bodySolutionUpdate = {
                "scope": scope
                }
                if not self.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate, programdetails, userRole):
                    self.errorVar.append("Solution Update Failed.")
                    return finalObsSolutionLink,self.errorVar
                print("solution update success.")
                ReffstartDateOfProgram1 = self.convert_to_date(programdetails.get('Startdateofprogram'))
                ReffendDateOfProgram1 = self.convert_to_date(programdetails.get('Enddateofprogram'))
                solutionDetails2 = self.convert_to_date(solutionDetails[2])
                solutionDetails3 = self.convert_to_date(solutionDetails[3])
                if ReffstartDateOfProgram1 <= solutionDetails2 <= ReffendDateOfProgram1 and ReffstartDateOfProgram1 <= solutionDetails3 <= ReffendDateOfProgram1:
                    print("dates validated...")
                    if solutionDetails[2]:
                        startDateArr = str(solutionDetails[2]).split("-")
                        bodySolutionUpdate = {
                            "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                        if not self.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate, programdetails, userRole):
                            self.errorVar.append("Solution Update Failed.")
                            return finalObsSolutionLink, self.errorVar
                    if solutionDetails[3]:
                        endDateArr = str(solutionDetails[3]).split("-")
                        bodySolutionUpdate = {
                            "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                        if not self.solutionUpdate(parentFolder, accessToken, childId[0], bodySolutionUpdate, programdetails, userRole):
                            self.errorVar.append("Solution Update Failed.")
                            return finalObsSolutionLink,self.errorVar
                else:
                    self.errorVar.append("Date Mismatched! Creation Stopped.")
                    print("Date Mismatched! Creation Stopped")
                    return finalObsSolutionLink,self.errorVar
                ObsSolutionLink = self.prepareProgramSuccessSheet(MainFilePath, parentFolder, programFile, childId[1], childId[0],
                                        accessToken, programdetails, userRole)
                print(ObsSolutionLink,"ObsSolutionLink")
                if not ObsSolutionLink:
                    self.errorVar.append("Solution Fetch Link Failed.")
                    return finalObsSolutionLink, self.errorVar
                else:
                    finalObsSolutionLink = ObsSolutionLink
                    return finalObsSolutionLink, self.errorVar 
                            

from backend.src.main.modules.common_config import *
from dotenv import load_dotenv
from pathlib import Path
import os, json, requests, sys, csv, time, uuid, re, openpyxl
from openpyxl.styles import Color, PatternFill
from datetime import datetime
from bson import ObjectId
import pandas as pd
import backend.src.main.modules.headers as apiHeader
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

env_path = Path(__file__).resolve().parents[1] / "apiServices" / "src" / "main" / ".env"

load_dotenv(dotenv_path=env_path)
internal_access_token = os.getenv("internal_access_token")
jwtTokenSecret = os.getenv("jwtTokenSecret")
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

# ─── Tuneable constants ──────────────────────────────────────────────────────
DEFAULT_TIMEOUT = (10, 60)   # (connect_timeout_sec, read_timeout_sec)
MAX_RETRIES     = 3
BACKOFF_FACTOR  = 2          # wait 2, 4, 8 … seconds between retries
# Status codes that are worth retrying (server-side / transient)
RETRY_STATUS_CODES = {429, 500, 502, 503, 504}
# How long to pause between improvement-project API calls (seconds)
IMP_PROJECT_SLEEP = 2
# ─────────────────────────────────────────────────────────────────────────────


def _build_session() -> requests.Session:
    """
    Return a requests.Session with:
      • automatic retries (with exponential back-off) for transient errors
      • a shared connection pool (avoids opening a new TCP socket every call)
    """
    session = requests.Session()
    retry = Retry(
        total=MAX_RETRIES,
        backoff_factor=BACKOFF_FACTOR,
        status_forcelist=RETRY_STATUS_CODES,
        allowed_methods=["GET", "POST"],   # retry POST as well
        raise_on_status=False,
    )
    adapter = HTTPAdapter(
        max_retries=retry,
        pool_connections=10,
        pool_maxsize=20,
    )
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session


class CreateObservation:
    def __init__(self):
        self.errorVar = []
        # One session shared for the lifetime of this object
        self._session = _build_session()

    # ── thin wrapper so every call gets a timeout automatically ──────────────
    def _post(self, url, *, headers=None, json=None, data=None,
              files=None, params=None, timeout=DEFAULT_TIMEOUT):
        """
        Wrapper around session.post that:
          • always sets a timeout
          • raises a clear RuntimeError on requests.exceptions.Timeout
            instead of hanging silently
        """
        try:
            return self._session.post(
                url, headers=headers, json=json, data=data,
                files=files, params=params, timeout=timeout,
            )
        except requests.exceptions.Timeout:
            raise RuntimeError(
                f"POST {url} timed out after {timeout[1]}s read / "
                f"{timeout[0]}s connect"
            )
        except requests.exceptions.ConnectionError as exc:
            raise RuntimeError(f"POST {url} connection error: {exc}") from exc

    def _get(self, url, *, headers=None, params=None,
             timeout=DEFAULT_TIMEOUT):
        try:
            return self._session.get(
                url, headers=headers, params=params, timeout=timeout,
            )
        except requests.exceptions.Timeout:
            raise RuntimeError(
                f"GET {url} timed out after {timeout[1]}s read / "
                f"{timeout[0]}s connect"
            )
        except requests.exceptions.ConnectionError as exc:
            raise RuntimeError(f"GET {url} connection error: {exc}") from exc

    # ─────────────────────────────────────────────────────────────────────────

    def createAPILog(self, solutionName_for_folder_path, messageArr):
        file_exists = os.path.join(solutionName_for_folder_path, 'apiHitLogs', 'apiLogs.txt')
        os.makedirs(os.path.dirname(file_exists), exist_ok=True)
        if not os.path.exists(file_exists):
            with open(file_exists, "w", encoding='utf-8') as API_log:
                API_log.write("=" * 79 + "\nENVIRONMENT LOGS\n" + "=" * 79 + "\n")
        with open(file_exists, "a", encoding='utf-8') as API_log:
            API_log.write("\n")
            for msg in messageArr:
                API_log.write(str(msg) + "\n")

    def apicheckslog(self, solutionName_for_folder_path, messageArr):
        file_exists = os.path.join(solutionName_for_folder_path, 'apiHitLogs', 'apiLogs.csv')
        fileheader = ["Resource", "Process", "Status", "Remark"]
        os.makedirs(os.path.dirname(file_exists), exist_ok=True)
        if not os.path.exists(file_exists):
            with open(file_exists, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',', lineterminator='\n')
                writer.writerow(fileheader)
        with open(file_exists, 'a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',', lineterminator='\n')
            writer.writerow(messageArr)

    def FetchBulkTempExternalIDs(self, parentFolder, accessToken, solutionNames, programdetails, userRole):
        """
        Consolidates multiple FetchTempExternalID calls into bulk searches.
        Reduces 2N dbFind calls to 2 total.
        """
        if not solutionNames:
            return {}

        # Filter unique non-empty names
        unique_names = list(set(name.strip() for name in solutionNames if name and name.strip()))
        if not unique_names:
            return {}

        headers = apiHeader.headers().headerFetchEntitytype(accessToken)
        mapping = {}  # {solution_name: external_id}

        try:
            # ── 1. Bulk Fetch project templates ──────────────────────────────
            urldbFindPT = elevateprojecthost + dbfindapi_projectTemplate
            payloadPT = {
                "query": {"title": {"$in": unique_names}, "isReusable": True},
                "projection": ["externalId", "title"],
                "mongoIdKeys": ["_id"],
                "limit": 10000,
            }
            respPT = self._post(urldbFindPT, headers=headers, json=payloadPT)

            if respPT.status_code == 200:
                for res in respPT.json().get("result", []):
                    mapping[res['title']] = {"externalId": res['externalId']}
            else:
                self.errorVar.append(f"BulkDBFindPT-Error {respPT.status_code}")

            # ── 2. Bulk Fetch solutions to get IDs for status update ─────────
            urldbFindSol = elevateprojecthost + dbfindapi_url
            payloadSol = {
                "query": {"name": {"$in": unique_names}},
                "projection": ["name"],
                "mongoIdKeys": ["_id"],
                "limit": 10000,
            }
            respSol = self._post(urldbFindSol, headers=headers, json=payloadSol)

            if respSol.status_code == 200:
                headerUpdate = apiHeader.headers().headersObservationsolutionUpdate(
                    programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole
                )
                for res in respSol.json().get("result", []):
                    solName = res['name']
                    solId = res['_id']
                    if solName in mapping:
                        # Sequential updates still required as most APIs update single resources
                        self._post(
                            elevateprojecthost + solutionupdateapi + solId,
                            headers=headerUpdate,
                            json={"status": "inactive", "isDeleted": False},
                        )
            else:
                self.errorVar.append(f"BulkDBFindSol-Error {respSol.status_code}")

            # Construct lookup mapping {name: externalId}
            return {name: info.get("externalId") for name, info in mapping.items() if "externalId" in info}

        except RuntimeError as e:
            self.errorVar.append(f"Network error in FetchBulkTempExternalIDs: {e}")
            return {}
        except Exception as e:
            self.errorVar.append(f"Exception in FetchBulkTempExternalIDs: {str(e)}")
            return {}

    def criteriaUpload(self, parentFolder, wbObservation, millisAddObs, accessToken, tabName,
                       impLedObsFlag, programdetails, userRole):
        criteriaLookUp = dict()
        criteriaColNames = ["criteriaId", "criteria_name"]
        criteriaUploadFieldnames = ['criteriaID', 'criteriaName']
        dictCriteriaToCsv = dict()
        criteriaLevelsFromFramework = dict()
        all_imp_cols = set()

        # ── CASE 1: FRAMEWORK TAB ─────────────────────────────────────────────
        if tabName == "framework":
            fetchLevelsFromFramework = wbObservation.get("framework")
            if not fetchLevelsFromFramework or not isinstance(fetchLevelsFromFramework, list):
                self.errorVar.append("Framework sheet missing or in unexpected format.")
                return False

            criteriaImpDict = {}
            TcountImps = []
            if impLedObsFlag:
                impsToCriteria = wbObservation.get("imp mapping")
                if not impsToCriteria:
                    self.errorVar.append("Imp mapping expected but missing.")
                    return False
                
                # ── CONSOLIDATION: Collect all unique project names first ──
                all_names = []
                if isinstance(impsToCriteria[0], dict):
                    for row in impsToCriteria:
                        levels = [k for k in row.keys() if "improvement-projects" in k.lower()]
                        all_imp_cols.update(levels)
                        for k, v in row.items():
                            if "improvement-projects" in k.lower() and v:
                                all_names.append(str(v).strip())
                else:
                    keys = impsToCriteria[0]
                    levels = [k for k in keys if "improvement-projects" in k.lower()]
                    all_imp_cols.update(levels)
                    for row in impsToCriteria[1:]:
                        for i, v in enumerate(row):
                            if "improvement-projects" in keys[i].lower() and v:
                                all_names.append(str(v).strip())
                
                # Fetch all metadata in one batch
                lookup_table = self.FetchBulkTempExternalIDs(
                    parentFolder, accessToken, all_names, programdetails, userRole
                )

                if isinstance(impsToCriteria[0], dict):
                    for dictImp in impsToCriteria:
                        crit_key = str(dictImp.get('criteriaId', '')).strip()
                        if not crit_key:
                            continue
                        criteriaImpDict[crit_key] = {}
                        # Identify how many levels exist in this sheet
                        levels = [k for k in dictImp.keys() if "improvement-projects" in k.lower()]
                        for colname in levels:
                            solutionName = str(dictImp.get(colname, "") or "").strip()
                            criteriaImpDict[crit_key][colname] = lookup_table.get(solutionName, "")
                else:
                    keysFromImpSheet = impsToCriteria[0]
                    for row in impsToCriteria[1:]:
                        dictImp = {keysFromImpSheet[i]: row[i] for i in range(len(keysFromImpSheet))}
                        crit_key = str(dictImp.get('criteriaId', '')).strip()
                        if not crit_key:
                            continue
                        criteriaImpDict[crit_key] = {}
                        levels = [k for k in keysFromImpSheet if "improvement-projects" in k.lower()]
                        for colname in levels:
                            solutionName = str(dictImp.get(colname, "") or "").strip()
                            criteriaImpDict[crit_key][colname] = lookup_table.get(solutionName, "")

            first_row_keys = list(fetchLevelsFromFramework[0].keys())
            levelCount = 0
            for k in first_row_keys:
                if (
                    isinstance(k, str)
                    and k.strip().lower().startswith("l")
                    and "description" in k.lower()
                    and not k.strip().lower().startswith("ln")
                ):
                    levelCount += 1

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

            criteriaFilePath = os.path.join(parentFolder, 'criteriaUpload')
            os.makedirs(criteriaFilePath, exist_ok=True)
            csv_path = os.path.join(criteriaFilePath, 'uploadSheet.csv')
            file_exists = os.path.isfile(csv_path)
            criteriaLevelsCount = levelCount
            wbObservation['criteriaLevelsCount'] = criteriaLevelsCount

            for dictCriteria in fetchLevelsFromFramework:
                criteria_id_raw = str(dictCriteria.get('Criteria ID', '')).strip()
                if not criteria_id_raw:
                    continue

                dictCriteriaToCsv = {}
                dictCriteriaToCsv['criteriaID'] = criteria_id_raw + '_' + str(millisAddObs)
                criteriaLookUp[dictCriteriaToCsv['criteriaID'].strip()] = dictCriteria.get('Criteria Name', '')
                dictCriteriaToCsv['criteriaName'] = dictCriteria.get('Criteria Name', '')
                dictCriteriaToCsv['type'] = 'auto'

                for levlsNo in range(1, levelCount + 1):
                    level_col = f"L{levlsNo} description"
                    dictCriteriaToCsv[f"L{levlsNo}"] = dictCriteria.get(level_col, "") or ""

                if impLedObsFlag:
                    crit_key = criteria_id_raw
                    for eachImps, val in criteriaImpDict.get(crit_key, {}).items():
                        dictCriteriaToCsv[eachImps] = val

                if 'type' not in criteriaUploadFieldnames:
                    criteriaUploadFieldnames.append('type')
                for eachCols in criteriaColNames:
                    if eachCols not in ['criteria_id', 'criteria_name', 'type', "criteriaId"]:
                        if eachCols not in criteriaUploadFieldnames:
                            criteriaUploadFieldnames.append(eachCols)
                if impLedObsFlag:
                    for imp_col in sorted(list(all_imp_cols)):
                        if imp_col not in criteriaUploadFieldnames:
                            criteriaUploadFieldnames.append(imp_col)

                with open(csv_path, 'a', encoding='utf-8', newline='') as criteriaUploadFile:
                    writerCriteriaUpload = csv.DictWriter(
                        criteriaUploadFile, fieldnames=list(criteriaUploadFieldnames), lineterminator='\n'
                    )
                    if not file_exists:
                        writerCriteriaUpload.writeheader()
                        file_exists = True
                    writerCriteriaUpload.writerow(dictCriteriaToCsv)

        # ── CASE 2: CRITERIA TAB ──────────────────────────────────────────────
        elif tabName == "criteria":
            print(wbObservation, "wbObservation")
            criteriaList = wbObservation.get("criteria")
            if not criteriaList or not isinstance(criteriaList, list):
                self.errorVar.append("Criteria sheet missing or not in expected dict format.")
                return False

            criteriaFilePath = os.path.join(parentFolder, 'criteriaUpload')
            os.makedirs(criteriaFilePath, exist_ok=True)
            csv_path = os.path.join(criteriaFilePath, 'uploadSheet.csv')
            file_exists = os.path.isfile(csv_path)
            criteriaUploadFieldnames = ['criteriaID', 'criteriaName', 'L1', 'L2', 'L3', 'type']

            for row in criteriaList:
                criteria_id_raw = str(row.get('criteria_id', '')).strip()
                criteria_name_raw = str(row.get('criteria_name', '')).encode('utf-8').decode('utf-8')
                if not criteria_id_raw or not criteria_name_raw:
                    continue

                data = {
                    'criteriaID': criteria_id_raw + '_' + str(millisAddObs),
                    'criteriaName': criteria_name_raw,
                    'L1': 'NA', 'L2': 'NA', 'L3': 'NA',
                    'type': 'auto',
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

        # ── Upload to API (common for both tabs) ──────────────────────────────
        try:
            urlCriteriaUploadApi = internal_kong_ip + criteriauploadapiurl
            headerCriteriaUploadApi = apiHeader.headers().headersCriteriaUpload(
                programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole
            )
            # Use a context manager so the file handle is always closed
            with open(csv_path, 'rb') as criteria_file:
                responseCriteriaUploadApi = self._post(
                    urlCriteriaUploadApi,
                    headers=headerCriteriaUploadApi,
                    files={'criteria': criteria_file},
                    # Criteria upload may be large – give it more time
                    timeout=(10, 120),
                )

            messageArr = [
                "Criteria Upload Sheet Prepared.",
                f"File path : {csv_path}",
                f"Upload status code : {responseCriteriaUploadApi.status_code}",
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
                self.errorVar.append(
                    f"CriteriaUploadApi Error {responseCriteriaUploadApi.status_code}: {responseCriteriaUploadApi.text}"
                )
                print("❌ Criteria Upload failed.")
                return False

        except RuntimeError as e:
            self.createAPILog(parentFolder, [str(e)])
            self.errorVar.append(str(e))
            return False
        except Exception as e:
            self.createAPILog(parentFolder, [f"Exception caught: {e}"])
            self.errorVar.append(f"Error occurred during Criteria Upload API: {str(e)}")
            return False

    def frameWorkUpload(self, solutionName_for_folder_path, wbObservation, millisAddObs, accessToken,
                        programdetails, typeofsolution, userRole):
        detailsSheet = wbObservation['details']
        if typeofsolution == 2:
            detailsSheet["scoring_system"] = "null"
        dateTime = datetime.now()
        frameworkDocInsertObj = {}
        frameworkExternalId = str(uuid.uuid1())
        frameworkDocInsertObj['externalId'] = frameworkExternalId
        frameworkDocInsertObj['name'] = detailsSheet.get("observation_solution_name").strip()
        frameworkDocInsertObj['description'] = detailsSheet.get("observation_solution_description").strip()
        frameworkDocInsertObj['parentId'] = None
        frameworkDocInsertObj['resourceType'] = ['Observations Framework']
        frameworkDocInsertObj['language'] = detailsSheet.get("language")
        frameworkDocInsertObj['levelToScoreMapping'] = dict()
        keyWords = detailsSheet.get("language")
        if keyWords and (keyWords != 'Framework' or keyWords != 'Frameworks'
                         or keyWords != 'Observation' or keyWords != 'Observations'):
            keywordsFinalArr = ['Framework', 'Observation']
            keywordsArr = keyWords.encode('utf-8').decode('utf-8').split(',')
            for keyw in keywordsArr:
                keywordsFinalArr.append(keyw)
            frameworkDocInsertObj['keywords'] = keywordsFinalArr
        else:
            frameworkDocInsertObj['keywords'] = ['Framework', 'Observation']
        frameworkDocInsertObj['concepts'] = []
        frameworkDocInsertObj['createdFor'] = [
            detailsSheet.get("Username/user id/email id/phone no. of the Content creator") or ""
        ]
        frameworkDocInsertObj['rootOrg'] = [
            detailsSheet.get("Username/user id/email id/phone no. of the Content creator") or ""
        ]
        criteriaFrameworkArr = []
        with open(solutionName_for_folder_path + '/criteriaUpload/uploadInternalIdsSheet.csv', 'r', encoding='utf-8') as criteriaInternalFile:
            criteriaInternalReader = csv.DictReader(criteriaInternalFile)
            criteriaWeightage = 100 / (len(list(criteriaInternalReader)))
            criteriaInternalFile.seek(0, 0)
            next(criteriaInternalReader, None)
            for crit in criteriaInternalReader:
                dictCritInter = dict(crit)
                criteriaFrameworkArr.append({
                    'criteriaId': str(ObjectId(dictCritInter['Criteria Internal Id'])),
                    'weightage': criteriaWeightage,
                })

        frameworkDocInsertObj['themes'] = [{
            'type': 'theme', 'label': 'theme', 'name': 'Observation Theme',
            'externalId': 'OB', 'weightage': 100, 'criteria': criteriaFrameworkArr,
        }]
        if not detailsSheet.get("scoring_system").lower() == "null":
            frameworkDocInsertObj['flattenedThemes'] = {
                "type": "theme", "label": "theme", "name": "Observation Theme",
                "externalId": "OB", "weightage": 1, "criteria": criteriaFrameworkArr,
                "rubric": {
                    "expressionVariables": {"SCORE": "OB.sumOfPointsOfAllChildren()"},
                    "levels": {"L1": {"expression": "(0<=SCORE<=100000)"}},
                },
                "hierarchyLevel": 0, "hierarchyTrack": [],
            }
            frameworkDocInsertObj['scoringSystem'] = detailsSheet.get("scoring_system")
            frameworkDocInsertObj['isRubricDriven'] = True
            frameworkDocInsertObj['themes'] = [{
                'type': 'theme', 'label': 'theme', 'name': 'Observation Theme',
                'externalId': 'OB', 'weightage': 100, 'criteria': criteriaFrameworkArr,
                "rubric": {
                    "expressionVariables": {"SCORE": "OB.sumOfPointsOfAllChildren()"},
                    "levels": {"L1": {"expression": "(0<=SCORE<=100000)"}},
                },
            }]
            criteriaLevelsCount = wbObservation.get('criteriaLevelsCount')
            for levs in range(1, criteriaLevelsCount + 1):
                frameworkDocInsertObj['levelToScoreMapping'].update(
                    {f"L{levs}": {'points': levs * 10, 'label': f'Level {levs}'}}
                )
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

        creator = str(detailsSheet.get("Username/user id/email id/phone no. of the Content creator"))
        frameworkDocInsertObj['creator'] = creator
        frameworkDocInsertObj['license'] = {
            'author': creator, 'creator': creator,
            'copyright': str(detailsSheet.get("Name_of_the_creator")),
            'copyrightYear': int(dateTime.strftime("%Y")),
            'contentType': "Observation",
            'organisation': [creator],
            'orgDetails': {'email': None, 'orgName': None},
            'licenseDetails': {
                'name': "CC BY 4.0",
                'url': "https://creativecommons.org/licenses/by/4.0/legalcode",
                'description': "For details see below:",
            },
        }
        try:
            urlCreateFrameworkApi = internal_kong_ip + frameworkcreationapi
            frameworkFilePath = solutionName_for_folder_path + '/framework/'
            if not os.path.exists(frameworkFilePath):
                os.mkdir(frameworkFilePath)

            fw_json_path = frameworkFilePath + "uploadFile.json"
            with open(fw_json_path, "w", encoding='utf-8') as outfile:
                json.dump(frameworkDocInsertObj, outfile)

            headerFrameworkUploadApi = apiHeader.headers().headersFrameworkUpload(
                programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole
            )
            # Use context manager – file handle is closed after the call
            with open(fw_json_path, 'rb') as fw_file:
                responseFrameworkUploadApi = self._post(
                    urlCreateFrameworkApi,
                    headers=headerFrameworkUploadApi,
                    files={'framework': fw_file},
                    timeout=(10, 120),
                )

            messageArr = [
                "Framework json file created.",
                f"File loc : {fw_json_path}",
                "Framework upload API called.",
                f"Status code : {responseFrameworkUploadApi.status_code}",
            ]
            self.createAPILog(solutionName_for_folder_path, messageArr)

            if responseFrameworkUploadApi.status_code == 200:
                print('Framework upload Success')
                return frameworkExternalId
            else:
                status = responseFrameworkUploadApi.status_code
                if status in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"FrameworkUploadApi-Client Error {status}: {responseFrameworkUploadApi.text}")
                elif status in [500, 502, 503, 504]:
                    self.errorVar.append(f"FrameworkUploadApi-Server Error {status}: {responseFrameworkUploadApi.text}")
                else:
                    self.errorVar.append(f"FrameworkUploadApi-Unexpected Error {status}: {responseFrameworkUploadApi.text}")
                self.createAPILog(solutionName_for_folder_path, ["Framework upload Failed.", "Response : " + responseFrameworkUploadApi.text])
                return False

        except RuntimeError as e:
            self.createAPILog(solutionName_for_folder_path, [str(e)])
            self.errorVar.append(str(e))
            return False
        except Exception as e:
            self.createAPILog(solutionName_for_folder_path, [f"Exception caught : {e}"])
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False

    def themesUpload(self, solutionName_for_folder_path, wbObservation, millisAddObs, accessToken,
                     frameworkExternalId, obsWORubWS, programdetails, userRole):
        dictCritLookUp = {}
        with open(solutionName_for_folder_path + '/criteriaUpload/uploadInternalIdsSheet.csv', 'r', encoding='utf-8') as f:
            for crit in csv.DictReader(f):
                dictCritLookUp[crit['Criteria External Id']] = crit['Criteria Internal Id']

        themeUploadFieldnames = ["theme", "aoi", "indicators", "criteriaInternalId"]
        themeFilePath = os.path.join(solutionName_for_folder_path, 'themeUpload')
        os.makedirs(themeFilePath, exist_ok=True)
        uploadCsvPath = os.path.join(themeFilePath, 'uploadSheet.csv')
        file_exists = os.path.isfile(uploadCsvPath)

        if obsWORubWS:
            print("Themes Observation without rubrics with scores")
            with open(uploadCsvPath, 'a', encoding='utf-8', newline='') as themeUploadFile:
                writer = csv.DictWriter(themeUploadFile, fieldnames=themeUploadFieldnames, lineterminator='\n')
                if not file_exists:
                    writer.writeheader()
                for dictCritLookUpValue in dictCritLookUp.values():
                    writer.writerow({
                        'theme': "Observation Theme###OB###40",
                        'aoi': "", 'indicators': "",
                        'criteriaInternalId': dictCritLookUpValue + "###40",
                    })
        else:
            framework_list = wbObservation['framework']
            themesUploadList = []
            for dictCriteria in framework_list:
                criteria_id = dictCriteria['Criteria ID'].strip()
                themesUploadList.append({
                    'theme': f"{dictCriteria['Domain Name']}###{dictCriteria['Domain ID']}###40",
                    'aoi': "", 'indicators': "",
                    'criteriaInternalId': dictCritLookUp[criteria_id + f"_{millisAddObs}"] + "###40",
                })
            with open(uploadCsvPath, 'a', encoding='utf-8', newline='') as themeUploadFile:
                writer = csv.DictWriter(themeUploadFile, fieldnames=themeUploadFieldnames, lineterminator='\n')
                if not file_exists:
                    writer.writeheader()
                writer.writerows(themesUploadList)

        try:
            urlThemesUploadApi = internal_kong_ip + themeuploadapiurl + frameworkExternalId
            headerThemesUploadApi = apiHeader.headers().headersFrameworkUpload(
                programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole
            )
            with open(uploadCsvPath, 'rb') as theme_file:
                responseThemeUploadApi = self._post(
                    urlThemesUploadApi,
                    headers=headerThemesUploadApi,
                    files={'themes': theme_file},
                    timeout=(10, 120),
                )

            messageArr = [
                "Themes upload sheet prepared.",
                f"File path : {uploadCsvPath}",
                "Theme upload to framework API called.",
                f"URL : {urlThemesUploadApi}",
                f"Status code : {responseThemeUploadApi.status_code}",
            ]
            self.createAPILog(solutionName_for_folder_path, messageArr)

            if responseThemeUploadApi.status_code == 200:
                print('Theme UploadApi Success')
                with open(solutionName_for_folder_path + '/themeUpload/uploadInternalIdsSheet.csv', 'w+', encoding='utf-8') as criteriaRes:
                    criteriaRes.write(responseThemeUploadApi.text)
                return True
            else:
                status = responseThemeUploadApi.status_code
                if status in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"ThemeUploadApi-Client Error {status}: {responseThemeUploadApi.text}")
                elif status in [500, 502, 503, 504]:
                    self.errorVar.append(f"ThemeUploadApi-Server Error {status}: {responseThemeUploadApi.text}")
                else:
                    self.errorVar.append(f"ThemeUploadApi-Unexpected Error {status}: {responseThemeUploadApi.text}")
                self.createAPILog(solutionName_for_folder_path, ["Themes upload failed.", "Response : " + str(responseThemeUploadApi.text)])
                return False

        except RuntimeError as e:
            self.createAPILog(solutionName_for_folder_path, [str(e)])
            self.errorVar.append(str(e))
            return False
        except Exception as e:
            self.createAPILog(solutionName_for_folder_path, [f"Exception caught : {e}"])
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False

    def createSolutionFromFramework(self, solutionName_for_folder_path, accessToken,
                                    frameworkExternalId, programdetails, userRole):
        try:
            entitydetails = programdetails.get("entitiesType")
            entity_type = entitydetails[0] if isinstance(entitydetails[0], str) else str(entitydetails[0][0])
            urlCreateSolutionApi = internal_kong_ip + solutioncreationapiurl
            headerCreateSolutionApi = apiHeader.headers().headersCreateSolutionFromFramework(
                programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole
            )
            queryparams = (f"?frameworkId={frameworkExternalId}"
                           f"&entityType={entity_type}&isExternalProgram=true")
            responseCreateSolutionApi = self._post(
                urlCreateSolutionApi + queryparams,
                headers=headerCreateSolutionApi,
                timeout=(10, 60),
            )

            messageArr = [
                "Solution Created from Framework.",
                f"URL : {urlCreateSolutionApi + queryparams}",
                f"Status Code : {responseCreateSolutionApi.status_code}",
                f"Response : {responseCreateSolutionApi.text}",
            ]
            self.createAPILog(solutionName_for_folder_path, messageArr)

            if responseCreateSolutionApi.status_code == 200:
                solutionId = responseCreateSolutionApi.json()['result']['templateId']
                print(f"Parent Solution Generated : {solutionId}")
                self.createAPILog(solutionName_for_folder_path, [f"Parent Solution Generated : {solutionId}"])
                return solutionId
            else:
                status = responseCreateSolutionApi.status_code
                if status in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"CreateSolutionApi-Client Error {status}: {responseCreateSolutionApi.text}")
                elif status in [500, 502, 503, 504]:
                    self.errorVar.append(f"CreateSolutionApi-Server Error {status}: {responseCreateSolutionApi.text}")
                else:
                    self.errorVar.append(f"CreateSolutionApi-Unexpected Error {status}: {responseCreateSolutionApi.text}")
                self.createAPILog(solutionName_for_folder_path, ["Solution from framework api failed."])
                return False

        except RuntimeError as e:
            self.createAPILog(solutionName_for_folder_path, [str(e)])
            self.errorVar.append(str(e))
            return False
        except Exception as e:
            self.createAPILog(solutionName_for_folder_path, [f"Exception caught : {e}"])
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False

    def ECMUpdatebody(self, ECMSheet, millisecond):
        ecm_update = {}
        section = {}
        ecm_sections = {}
        ecmSeqCount = 1

        for dictECMs in ECMSheet:
            EMC_ID = dictECMs['ECM Id/Domian ID'].strip() + '_' + str(millisecond)
            ECM_NAME = dictECMs['ECM Name/Domain Name'].strip()
            section.update({dictECMs['section_id']: dictECMs['section_name']})
            ecm_sections[EMC_ID] = dictECMs['section_id']

            is_mandatory = dictECMs.get('Is ECM Mandatory?')
            mandatory_flag = is_mandatory not in ["TRUE", 1]

            ecm_update[EMC_ID] = {
                "externalId": EMC_ID, "tip": None, "name": ECM_NAME,
                "description": None, "modeOfCollection": "onfield",
                "canBeNotApplicable": mandatory_flag, "notApplicable": False,
                "canBeNotAllowed": mandatory_flag, "remarks": None,
                "sequenceNo": ecmSeqCount,
            }
            print(ecm_update[EMC_ID])
            ecmSeqCount += 1

        return [{"evidenceMethods": ecm_update}, {"sections": section}]

    def safe_decode(self, val):
        if val is None:
            return None
        if isinstance(val, str):
            return val.strip()
        try:
            return str(val).encode('utf-8').decode('utf-8').strip()
        except Exception:
            return str(val).strip()

    def get_numeric_or_str(self, val):
        if val is None or val == "":
            return None
        try:
            f = float(val)
            return int(f) if f.is_integer() else f
        except Exception:
            return self.safe_decode(val)

    def process_responses(self, ques, prefix='response', max_n=20):
        responses = {}
        for i in range(1, max_n + 1):
            responses[f"R{i}"] = self.get_numeric_or_str(ques.get(f"{prefix}(R{i})"))
            responses[f"R{i}-hint"] = self.get_numeric_or_str(ques.get(f"{prefix}(R{i})_hint"))
        return responses

    def process_scores(self, ques, max_n=20):
        return {f"R{i}-score": ques.get(f"Score for R{i}") for i in range(1, max_n + 1)}

    def questionUpload(self, wbObservation, parentFolder, frameworkExternalId, millisAddObs,
                       accessToken, solutionId, typeofSolution, programdetails, userRole):
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
            ecmToSection = {ecm['section_id']: ecm['ECM Id/Domian ID'] for ecm in wbObservation['ecms or domains']}

        criteriaLookUp = {}
        if typeofSolution == 2:
            for ques in wbObservation['questions']:
                criteriaKey = ques['criteria_id'].strip() + '_' + str(millisAddObs)
                criteriaLookUp[criteriaKey] = ques.get('criteria_name', 'Unknown Criteria')
        else:
            if 'framework' not in wbObservation:
                self.errorVar.append("Framework sheet missing for this solution type.")
                return False
            criteriaLookUp = {
                criteria['Criteria ID'].strip() + '_' + str(millisAddObs): criteria['Criteria Name']
                for criteria in wbObservation['framework']
            }

        questionsList = sorted(wbObservation['questions'], key=lambda x: float(x.get('question_sequence') or 0))

        questionFilePath = os.path.join(parentFolder, 'questionUpload')
        os.makedirs(questionFilePath, exist_ok=True)
        uploadCSV = os.path.join(questionFilePath, 'uploadSheet.csv')
        file_exists_ques = os.path.isfile(uploadCSV)

        questionUploadFieldnames = [
            'solutionId', 'criteriaExternalId', 'name', 'evidenceMethod', 'section',
            'instanceParentQuestionId', 'hasAParentQuestion', 'parentQuestionOperator',
            'parentQuestionValue', 'parentQuestionId', 'externalId', 'question0', 'question1',
            'tip', 'hint', 'instanceIdentifier', 'responseType', 'dateFormat', 'autoCapture',
            'validation', 'validationIsNumber', 'validationRegex', 'validationMax', 'validationMin',
            'file', 'fileIsRequired', 'fileUploadType', 'minFileCount', 'maxFileCount',
            'allowAudioRecording', 'caption', 'questionGroup', 'modeOfCollection', 'accessibility',
            'showRemarks', 'rubricLevel', 'isAGeneralQuestion',
        ] + [f'R{i}' for i in range(1, 21)] \
          + [f'R{i}-hint' for i in range(1, 21)] \
          + [f'R{i}-score' for i in range(1, 21)] \
          + ['weightage', 'sectionHeader', 'page', 'questionNumber', '_arrayFields',
             'prefillFromEntityProfile', 'isEditable', 'entityFieldName']

        questionSeqByEcmDict = {}

        with open(uploadCSV, 'a', encoding='utf-8', newline='') as questionUploadFile:
            writerQuestionUpload = csv.DictWriter(
                questionUploadFile, fieldnames=questionUploadFieldnames, lineterminator='\n'
            )
            if not file_exists_ques:
                writerQuestionUpload.writeheader()

            for ques in questionsList:
                questionFileObj = {key: None for key in questionUploadFieldnames}
                observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
                questionFileObj['solutionId'] = observationExternalId
                questionFileObj['criteriaExternalId'] = ques['criteria_id'].strip() + '_' + str(millisAddObs)

                try:
                    questionFileObj['name'] = criteriaLookUp[questionFileObj['criteriaExternalId']]
                except KeyError:
                    self.errorVar.append("criteria Id error....")
                    print(questionFileObj['criteriaExternalId'] + " not found.")
                    return False

                if typeofSolution in [1, 5]:
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
                    questionSeqByEcmDict["OB"]["S1"].append(
                        ques['question_id'].strip() + '_' + str(millisAddObs)
                    )

                for i in range(1, 21):
                    questionFileObj[f'R{i}'] = ques.get(f'response(R{i})')
                    questionFileObj[f'R{i}-hint'] = ques.get(f'response(R{i})_hint')
                    questionFileObj[f'R{i}-score'] = ques.get(f'Score for R{i}')

                if ques.get('instance_parent_question_id'):
                    questionFileObj['instanceParentQuestionId'] = (
                        ques['instance_parent_question_id'].strip() + '_' + str(millisAddObs)
                    )
                    questionFileObj['hasAParentQuestion'] = 'NO'
                else:
                    questionFileObj['instanceParentQuestionId'] = 'NA'

                if ques.get('parent_question_id'):
                    questionFileObj['hasAParentQuestion'] = 'YES'
                    op = ques.get('show_when_parent_question_value_is', '').strip().upper()
                    if op in ['OR', '||']:
                        questionFileObj['parentQuestionOperator'] = '||'
                        questionFileObj['parentQuestionValue'] = ques.get('parent_question_value', '').replace(" ", "")
                    elif op == 'EQUALS':
                        questionFileObj['parentQuestionOperator'] = 'EQUALS'
                        questionFileObj['parentQuestionValue'] = ques.get('parent_question_value', '').replace(" ", "")
                    elif op == 'NOT_EQUALS_TO':
                        questionFileObj['parentQuestionOperator'] = '||'
                    else:
                        questionFileObj['parentQuestionOperator'] = ''
                    questionFileObj['parentQuestionId'] = ques['parent_question_id'].strip() + '_' + str(millisAddObs)

                questionFileObj['externalId'] = ques['question_id'].strip() + '_' + str(millisAddObs)
                questionFileObj['question0'] = ques.get('question_primary_language')
                questionFileObj['question1'] = ques.get('question_secondory_language')
                questionFileObj['tip'] = ques.get('question_tip')
                questionFileObj['hint'] = ques.get('question_hint')
                questionFileObj['instanceIdentifier'] = ques.get('instance_identifier')
                questionFileObj['responseType'] = ques.get('question_response_type', '').strip().lower()
                questionFileObj['weightage'] = ques.get('question_weightage', 0)
                questionFileObj['sectionHeader'] = ques.get('section_header')
                questionFileObj['page'] = ques.get('page')
                questionFileObj['questionNumber'] = ques['question_number'] if ques.get('question_number') else None

                if questionFileObj['responseType'] == 'date':
                    questionFileObj['dateFormat'] = 'DD-MM-YYYY'
                    questionFileObj['autoCapture'] = (
                        'TRUE' if ques.get('date_auto_capture') in [1, 'true', 'True'] else 'FALSE'
                    )
                else:
                    questionFileObj['dateFormat'] = ''
                    questionFileObj['autoCapture'] = None

                questionFileObj['validation'] = (
                    'TRUE' if ques.get('response_required') in [1, 'true', 'True'] else 'FALSE'
                )
                if questionFileObj['responseType'] in ['number', 'slider']:
                    questionFileObj['validationIsNumber'] = 'TRUE'
                    questionFileObj['validationRegex'] = 'isNumber'
                    questionFileObj['validationMax'] = (
                        ques.get('max_number_value') or (5 if questionFileObj['responseType'] == 'slider' else 10000)
                    )
                    questionFileObj['validationMin'] = ques.get('min_number_value') or 0

                if ques.get('file_upload') in [1, 'TRUE', 'true']:
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
                questionFileObj['showRemarks'] = (
                    'TRUE' if ques.get('show_remarks') in [1, 'TRUE', 'true'] else 'FALSE'
                )
                questionFileObj['rubricLevel'] = None
                questionFileObj['isAGeneralQuestion'] = None
                writerQuestionUpload.writerow(questionFileObj)

        bodySolutionUpdate = {"questionSequenceByEcm": questionSeqByEcmDict}
        if not self.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate, programdetails, userRole):
            return False

        try:
            urlQuestionsUploadApi = internal_kong_ip + questionuploadapiurl
            headerQuestionUploadApi = apiHeader.headers().headersQuestionUpload(
                programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole
            )
            with open(parentFolder + '/questionUpload/uploadSheet.csv', 'rb') as q_file:
                responseQuestionUploadApi = self._post(
                    urlQuestionsUploadApi,
                    headers=headerQuestionUploadApi,
                    files={'questions': q_file},
                    # Question uploads can be large – generous timeout
                    timeout=(10, 180),
                )

            messageArr = [
                "Question upload API called.",
                f"URL : {urlQuestionsUploadApi}",
                f"Status Code : {responseQuestionUploadApi.status_code}",
                f"Response : {responseQuestionUploadApi.text}",
            ]
            self.createAPILog(parentFolder, messageArr)

            if responseQuestionUploadApi.status_code == 200:
                print('QuestionUploadApi Success')
                with open(parentFolder + '/questionUpload/uploadInternalIdsSheet.csv', 'w+', encoding='utf-8') as questionRes:
                    questionRes.write(responseQuestionUploadApi.text)
                return True
            else:
                self.errorVar.append(
                    f"QuestionUploadApi Error {responseQuestionUploadApi.status_code}: {responseQuestionUploadApi.text}"
                )
                return False

        except RuntimeError as e:
            self.createAPILog(parentFolder, [str(e)])
            self.errorVar.append(str(e))
            return False
        except Exception as e:
            self.createAPILog(parentFolder, [f"Exception caught : {e}"])
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False

    def fetchSolutionCriteria(self, solutionName_for_folder_path, observationId, accessToken, programdetails, userRole):
        try:
            url = internal_kong_ip + ferchsolutioncriteria + observationId
            headers = apiHeader.headers().headersFetchSolutionCriteria(
                programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole
            )
            response = self._post(url, headers=headers, timeout=(10, 60))

            messageArr = [
                "fetchSolutionCriteria called.",
                f"URL : {url}",
                f"Status Code : {response.status_code}",
                f"Response : {response.text}",
            ]
            self.createAPILog(solutionName_for_folder_path, messageArr)

            os.makedirs(solutionName_for_folder_path + "/solutionCriteriaFetch/", exist_ok=True)
            if response.status_code == 200:
                print("Solution criteria fetched.")
                with open(
                    solutionName_for_folder_path + "/solutionCriteriaFetch/solutionCriteriaDetails.csv",
                    'w+', encoding='utf-8',
                ) as f:
                    f.write(response.text)
                return True
            else:
                status = response.status_code
                if status in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"FetchSolutionCriteria-Client Error {status}: {response.text}")
                elif status in [500, 502, 503, 504]:
                    self.errorVar.append(f"FetchSolutionCriteria-Server Error {status}: {response.text}")
                else:
                    self.errorVar.append(f"FetchSolutionCriteria-Unexpected Error {status}: {response.text}")
                self.createAPILog(solutionName_for_folder_path,
                                  ["Criteria solution fetch API failed.", f"Response : {response.text}"])
                return False

        except RuntimeError as e:
            self.createAPILog(solutionName_for_folder_path, [str(e)])
            self.errorVar.append(str(e))
            return False
        except Exception as e:
            self.createAPILog(solutionName_for_folder_path, [f"Exception caught : {e}"])
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False

    def uploadCriteriaRubrics(self, parentFolder, wbObservation, millisecond, accessToken,
                               frameworkExternalId, withRubricsFlag, programdetails, userRole):
        criteriaRubricSheet = (
            wbObservation['criteria_rubric-scoring'] if withRubricsFlag else wbObservation['criteria']
        )
        dictSolCritLookUp = {}
        filePath = os.path.join(parentFolder, "solutionCriteriaFetch", "solutionCriteriaDetails.csv")
        with open(filePath, 'r', encoding='utf-8') as f:
            for crit in csv.DictReader(f):
                dictSolCritLookUp[crit['criteriaID']] = [crit['criteriaInternalId'], crit['criteriaName']]

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

        criteriaRubricUploadFieldnames = ["externalId", "name", "criteriaId", "weightage", "expressionVariables"]
        if withRubricsFlag and criteriaLevelsCount:
            criteriaRubricUploadFieldnames += [f"L{cl}" for cl in criteriaLevelsCount]
        else:
            criteriaRubricUploadFieldnames.append("L1")

        criteriaRubricsDir = os.path.join(parentFolder, 'criteriaRubrics')
        os.makedirs(criteriaRubricsDir, exist_ok=True)
        uploadSheetPath = os.path.join(criteriaRubricsDir, 'uploadSheet.csv')

        with open(uploadSheetPath, 'w', encoding='utf-8', newline='') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=criteriaRubricUploadFieldnames, lineterminator='\n')
            writer.writeheader()

            if withRubricsFlag:
                for row in criteriaRubricSheet:
                    criteria_id = row.get('criteriaId', '').strip()
                    if not criteria_id:
                        continue
                    lookup_key = f"{criteria_id}_{millisecond}"
                    criteria_info = dictSolCritLookUp.get(lookup_key)
                    if not criteria_info or not criteria_info[0] or not criteria_info[1]:
                        print(f"⚠️ Skipping {criteria_id} — missing lookup info.")
                        continue
                    record = {
                        'externalId': lookup_key,
                        'criteriaId': criteria_info[0].strip(),
                        'name': criteria_info[1].strip(),
                        'weightage': float(row.get('weightage', 1)),
                        'expressionVariables': f"SCORE={criteria_info[0]}.scoreOfAllQuestionInCriteria()",
                    }
                    for cl in (criteriaLevelsCount or [1]):
                        record[f"L{cl}"] = str(row.get(f"L{cl} SCORE", "")).strip()
                    if not record['criteriaId'] or not record['name']:
                        print(f"⚠️ Skipping incomplete record: {record}")
                        continue
                    writer.writerow(record)
            else:
                for criteriaIds, criteriaDetails in dictSolCritLookUp.items():
                    if not criteriaDetails[0] or not criteriaDetails[1]:
                        continue
                    writer.writerow({
                        'externalId': criteriaIds,
                        'name': criteriaDetails[1],
                        'criteriaId': criteriaDetails[0],
                        'weightage': 1,
                        'expressionVariables': f"SCORE={criteriaDetails[0]}.scoreOfAllQuestionInCriteria()",
                        'L1': '0<=SCORE<=100000',
                    })

        try:
            urlCriteriaRubricUploadApi = (
                internal_kong_ip + criteriarubricuploadapiurl + frameworkExternalId + "-OBSERVATION-TEMPLATE"
            )
            headerCriteriaRubricUploadApi = apiHeader.headers().headersCriteriaRubricUpload(
                programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole
            )
            with open(uploadSheetPath, 'rb') as f:
                response = self._post(
                    urlCriteriaRubricUploadApi,
                    headers=headerCriteriaRubricUploadApi,
                    files={'criteria': ('uploadSheet.csv', f, 'text/csv')},
                    timeout=(10, 120),
                )

            messageArr = [
                "CriteriaRubric upload called.",
                f"URL : {urlCriteriaRubricUploadApi}",
                f"Status Code : {response.status_code}",
                f"Response : {response.text}",
            ]
            self.createAPILog(parentFolder, messageArr)

            if response.status_code == 200:
                with open(os.path.join(criteriaRubricsDir, 'uploadInternalIdsSheet.csv'), 'w+', encoding='utf-8') as res_file:
                    res_file.write(response.text)
                print("✅ uploadCriteriaRubrics success.")
                return True
            else:
                msg = f"CriteriaRubricUploadApi Error {response.status_code}: {response.text}"
                self.errorVar.append(msg)
                print("❌", msg)
                return False

        except RuntimeError as e:
            self.createAPILog(parentFolder, [str(e)])
            self.errorVar.append(str(e))
            return False
        except Exception as e:
            self.createAPILog(parentFolder, [f"Exception caught : {e}"])
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False

    def uploadThemeRubrics(self, parentFolder, wbObservation, accessToken, frameworkExternalId,
                           withRubricsFlag, programdetails, userRole):
        themeRubricSheet = (
            wbObservation['theme_rubric_scoring'] if withRubricsFlag else wbObservation.get('Domain(theme)', [])
        )
        criteriaLevels = []
        if withRubricsFlag and themeRubricSheet:
            sample_row = next(iter(themeRubricSheet), {})
            for key in sample_row.keys():
                key_upper = key.upper().strip()
                if key_upper.startswith("L") and key_upper[1:].split()[0].isdigit() and "SCORE" not in key_upper:
                    level_num = int(''.join(filter(str.isdigit, key_upper)))
                    if level_num not in criteriaLevels:
                        criteriaLevels.append(level_num)
            criteriaLevels.sort()

        themeRubricUploadFieldnames = ["externalId", "name", "weightage"]
        if withRubricsFlag and criteriaLevels:
            themeRubricUploadFieldnames += [f"L{cl}" for cl in criteriaLevels]
        else:
            themeRubricUploadFieldnames.append("L1")

        themeRubricsFilePath = os.path.join(parentFolder, "themeRubrics")
        os.makedirs(themeRubricsFilePath, exist_ok=True)
        uploadSheetPath = os.path.join(themeRubricsFilePath, 'uploadSheet.csv')
        file_exists = os.path.isfile(uploadSheetPath)

        with open(uploadSheetPath, 'a', encoding='utf-8', newline='') as themeRubricsUploadFile:
            writer = csv.DictWriter(themeRubricsUploadFile, fieldnames=themeRubricUploadFieldnames, lineterminator='\n')
            if not file_exists:
                writer.writeheader()

            if withRubricsFlag:
                for row in themeRubricSheet:
                    domain_id = row.get('domain_Id', '').strip()
                    if not domain_id:
                        continue
                    record = {
                        'externalId': domain_id,
                        'name': row.get('domain_name', '').strip(),
                        'weightage': row.get('weightage', 0),
                    }
                    if criteriaLevels:
                        for cl in criteriaLevels:
                            record[f"L{cl}"] = row.get(f"L{cl}", "")
                    else:
                        record["L1"] = row.get("L1", "0<=SCORE<=100000")
                    writer.writerow(record)
            else:
                writer.writerow({
                    "externalId": "OB", "name": "Observation Theme",
                    "weightage": 1, "L1": "0<=SCORE<=100000",
                })

        try:
            urlThemeRubricUploadApi = (
                internal_kong_ip + themerubricuploadapiurl + frameworkExternalId + "-OBSERVATION-TEMPLATE"
            )
            headerThemeRubricUploadApi = apiHeader.headers().headersThemeRubricUpload(
                programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole
            )
            with open(uploadSheetPath, 'rb') as f:
                response = self._post(
                    urlThemeRubricUploadApi,
                    headers=headerThemeRubricUploadApi,
                    files={'themes': f},
                    timeout=(10, 120),
                )

            messageArr = [
                "ThemeRubric upload called.",
                f"URL : {urlThemeRubricUploadApi}",
                f"Status Code : {response.status_code}",
                f"Response : {response.text}",
            ]
            self.createAPILog(parentFolder, messageArr)

            if response.status_code == 200:
                print('ThemeRubricUploadApi Success')
                with open(os.path.join(themeRubricsFilePath, 'uploadInternalIdsSheet.csv'), 'w+', encoding='utf-8') as f:
                    f.write(response.text)
                return True
            else:
                self.errorVar.append(f"ThemeRubricUploadApi Error {response.status_code}: {response.text}")
                return False

        except RuntimeError as e:
            self.createAPILog(parentFolder, [str(e)])
            self.errorVar.append(str(e))
            return False
        except Exception as e:
            self.createAPILog(parentFolder, [f"Exception caught : {e}"])
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False

    def solutionUpdate(self, solutionName_for_folder_path, accessToken, solutionId,
                       bodySolutionUpdate, programdetails, userRole):
        try:
            solutionUpdateApi = internal_kong_ip + solutionupdateapi + str(solutionId)
            headerUpdateSolutionApi = apiHeader.headers().headersObservationsolutionUpdate(
                programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole
            )
            responseUpdateSolutionApi = self._post(
                solutionUpdateApi,
                headers=headerUpdateSolutionApi,
                data=json.dumps(bodySolutionUpdate),
                timeout=(10, 60),
            )

            messageArr = [
                "solutionUpdate called.",
                f"URL : {solutionUpdateApi}",
                f"bodySolutionUpdate: {bodySolutionUpdate}",
                f"Status Code : {responseUpdateSolutionApi.status_code}",
                f"Response : {responseUpdateSolutionApi.text}",
            ]
            self.createAPILog(solutionName_for_folder_path, messageArr)

            if responseUpdateSolutionApi.status_code == 200:
                print("Solution Update Success.")
                return True
            else:
                status = responseUpdateSolutionApi.status_code
                if status in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"UpdateSolutionApi-Client Error {status}: {responseUpdateSolutionApi.text}")
                elif status in [500, 502, 503, 504]:
                    self.errorVar.append(f"UpdateSolutionApi-Server Error {status}: {responseUpdateSolutionApi.text}")
                else:
                    self.errorVar.append(f"UpdateSolutionApi-Unexpected Error {status}: {responseUpdateSolutionApi.text}")
                return False

        except RuntimeError as e:
            self.createAPILog(solutionName_for_folder_path, [str(e)])
            self.errorVar.append(str(e))
            return False
        except Exception as e:
            self.createAPILog(solutionName_for_folder_path, [f"Exception caught : {e}"])
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False

    def UpdateCertForSolution(self, solutionName_for_folder_path, childTemplateId, childSolutionId,
                               accessToken, programdetails, userRole):
        try:
            headers = {'X-auth-token': accessToken, 'Content-Type': content_type}

            # ── 1. Fetch child solution name ──────────────────────────────────
            response = self._post(
                elevateprojecthost + dbfindapi_url, headers=headers,
                json={"query": {"_id": childSolutionId}, "projection": ["status", "name"],
                      "mongoIdKeys": ["_id"], "limit": 10000},
                timeout=(10, 60),
            )
            results = response.json().get("result", [])
            if not results or response.status_code != 200:
                self.errorVar = f"DBFind-Error {response.status_code}: {response.text}"
                return False
            projectSolutionName = results[0].get("name")

            # ── 2. Find inactive solution with matching name ───────────────────
            response = self._post(
                elevateprojecthost + dbfindapi_url, headers=headers,
                json={"query": {"name": projectSolutionName, "status": "inactive"},
                      "projection": ["status", "name", "certificateTemplateId"],
                      "mongoIdKeys": ["_id"], "limit": 10000},
                timeout=(10, 60),
            )
            results = response.json().get("result", [])
            if not results or response.status_code != 200:
                self.errorVar = f"DBFind-Error {response.status_code}: {response.text}"
                return False

            if results[0].get("certificateTemplateId"):
                certificateTemplateId = results[0]["certificateTemplateId"]
                print(certificateTemplateId, "certificateTemplateId")
                headerUpdateSolutionApi = apiHeader.headers().headersObservationsolutionUpdate(
                    programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole
                )
                payload = {"certificateTemplateId": certificateTemplateId}

                r1 = self._post(
                    elevateprojecthost + solutionupdateapi + childSolutionId,
                    headers=headerUpdateSolutionApi, json=payload, timeout=(10, 60),
                )
                if r1.status_code != 200:
                    print("Child Solution Update Failed.")
                    return False
                print("Child Solution Update Success.")

                r2 = self._post(
                    elevateprojecthost + projectTemplateupdateapi + childTemplateId,
                    headers=headerUpdateSolutionApi, json=payload, timeout=(10, 60),
                )
                if r2.status_code != 200:
                    print("Child Template Update Failed.")
                    return False
                print("Child Template Update Success.")

            return True

        except RuntimeError as e:
            self.errorVar = str(e)
            print(self.errorVar, "---> API-Error")
            return False
        except Exception as e:
            self.errorVar = f"Exception in UpdateCertForSolution: {str(e)}"
            print(self.errorVar, "---> API-Error")
            return False

    def createChild(self, parentFolder, wbObservation, observationExternalId,
                    accessToken, programdetails, userRole):
        entitydetails = programdetails.get("entitiesType")
        entity_type = entitydetails[0] if isinstance(entitydetails[0], str) else str(entitydetails[0][0])
        try:
            childObservationExternalId = str(observationExternalId + "_CHILD")
            urlSol_prog_mapping = (
                internal_kong_ip + solutiontoprogrammappingapiurl
                + "?solutionId=" + observationExternalId
                + "&entityType=" + entity_type
            )
            ObsDict = wbObservation.get('details')
            payload = {
                "externalId": childObservationExternalId,
                "name": ObsDict.get("observation_solution_name"),
                "description": ObsDict.get("observation_solution_description"),
                "programExternalId": programdetails.get("_id"),
            }
            headersSol_prog_mapping = apiHeader.headers().headersSolutionToProgramMapping(
                programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole
            )
            responseSol_prog_mapping = self._post(
                urlSol_prog_mapping,
                headers=headersSol_prog_mapping,
                data=json.dumps(payload),
                timeout=(10, 60),
            )

            messageArr = [
                "Create child API called.", f"URL : {urlSol_prog_mapping}",
                f"Status code : {responseSol_prog_mapping.status_code}",
                f"Response : {responseSol_prog_mapping.text}",
                f"body : {payload}",
            ]
            self.createAPILog(parentFolder, messageArr)

            if responseSol_prog_mapping.status_code == 200:
                if programdetails.get('TitleoftheProgram'):
                    print("Solution mapped to program : " + programdetails.get('TitleoftheProgram'))
                print("Child solution : " + childObservationExternalId)
                resp_json = responseSol_prog_mapping.json()
                child_id = resp_json['result']['_id']
                solutionDetails = resp_json['result']['projectTemplateDetails']
                for sol in solutionDetails:
                    self.UpdateCertForSolution(
                        parentFolder, sol.get('childProjectTemplateId'),
                        sol.get('solutionId'), accessToken, programdetails, userRole,
                    )
                    time.sleep(1)
                print("child solutionId: " + child_id)
                return [child_id, childObservationExternalId]
            else:
                status = responseSol_prog_mapping.status_code
                if status in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"Sol_prog_mapping-Client Error {status}: {responseSol_prog_mapping.text}")
                elif status in [500, 502, 503, 504]:
                    self.errorVar.append(f"Sol_prog_mapping-Server Error {status}: {responseSol_prog_mapping.text}")
                else:
                    self.errorVar.append(f"Sol_prog_mapping-Unexpected Error {status}: {responseSol_prog_mapping.text}")
                print("Unable to create child solution")
                return False

        except RuntimeError as e:
            self.createAPILog(parentFolder, [str(e)])
            self.errorVar.append(str(e))
            return False
        except Exception as e:
            self.createAPILog(parentFolder, [f"Exception caught : {e}"])
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False

    def fetchSolutionDetailsFromProgramSheet(self, solutionName_for_folder_path, programdetails,
                                             solutionId, accessToken, ProgramGlobalDict, userRole):
        try:
            urlFetchSolutionApi = internal_kong_ip + dbfindapi_url
            headerFetchSolutionApi = apiHeader.headers().headersFetchSolutionDetails(
                programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole
            )
            payload = json.dumps({
                "query": {"_id": solutionId},
                "mongoIdKeys": ["_id", "name", "externalId"],
                "limit": 10000,
            })
            responseFetchSolution = self._post(
                urlFetchSolutionApi, headers=headerFetchSolutionApi,
                data=payload, timeout=(10, 60),
            )
            responseFetchSolutionJson = responseFetchSolution.json()

            self.createAPILog(solutionName_for_folder_path, [
                "Solution Fetch.",
                "solution name : " + responseFetchSolutionJson["result"][-1]["name"],
                "solution ExternalId : " + responseFetchSolutionJson["result"][-1]["externalId"],
                f"Upload status code : {responseFetchSolution.status_code}",
            ])

            if responseFetchSolution.status_code == 200:
                solutionName = responseFetchSolutionJson["result"][-1]["name"]
                print(solutionName, "solutionName")
                for solutions in ProgramGlobalDict.get('Program Resources', []):
                    if solutionName == solutions.get('Nameofresourcesinprogram'):
                        return [
                            solutions.get('Targetroleattheresourcelevel'),
                            solutions.get('Targetedsubroleatresourcelevel'),
                            solutions.get('Startdateofresource'),
                            solutions.get('Enddateofresource'),
                        ]
            else:
                status = responseFetchSolution.status_code
                if status in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"FetchSolutionApiUrl-Client Error {status}: {responseFetchSolution.text}")
                elif status in [500, 502, 503, 504]:
                    self.errorVar.append(f"FetchSolutionApiUrl-Server Error {status}: {responseFetchSolution.text}")
                else:
                    self.errorVar.append(f"FetchSolutionApiUrl-Unexpected Error {status}: {responseFetchSolution.text}")
                return False

        except RuntimeError as e:
            self.createAPILog(solutionName_for_folder_path, [str(e)])
            self.errorVar.append(str(e))
            return False
        except Exception as e:
            self.createAPILog(solutionName_for_folder_path, [f"Exception caught : {e}"])
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False

    def validate_roles_against_api(self, mainRoles, subRoles, programdetails, parentFolder):
        urlFetchRoleList = userLoginHost + fetchprofessionalRole
        headers = apiHeader.headers().header_validate_roles_against_api(programdetails.get('TenantID'))
        try:
            response = self._get(urlFetchRoleList, headers=headers, timeout=(10, 30))
        except RuntimeError as e:
            self.errorVar.append(str(e))
            return [], []

        self.createAPILog(parentFolder, [
            "validate_roles_against_api called.",
            f"URL : {urlFetchRoleList}",
            f"Upload status code : {response.status_code}",
        ])

        if response.status_code != 200:
            print("Error fetching roles.")
            return [], []

        role_data = response.json()
        validated_main_role_ids = []
        validated_subrole_ids_list = []
        remaining_subroles = [s.strip() for s in subRoles]

        for role in mainRoles:
            matched = next(
                (r for r in role_data['result']
                 if r.get('externalId', '').strip() == role.strip()
                 or r.get('name', '').strip() == role.strip()),
                None,
            )
            if matched:
                main_role_id = matched['_id']
                validated_main_role_ids.append(main_role_id)
                subrole_url = (
                    f"{userLoginHost}entity-management/v1/entities/subEntityList/"
                    f"{main_role_id}?type=professional_subroles"
                )
                try:
                    subrole_resp = self._get(subrole_url, headers=headers, timeout=(10, 30))
                except RuntimeError as e:
                    self.errorVar.append(str(e))
                    continue

                if subrole_resp.status_code == 200:
                    for item in subrole_resp.json()['result']['data']:
                        sub_external_id = item.get('externalId', '').strip()
                        sub_name = item.get('name', '').strip()
                        if sub_external_id in remaining_subroles or sub_name in remaining_subroles:
                            validated_subrole_ids_list.append(item['_id'])
                            remaining_subroles = [
                                s for s in remaining_subroles
                                if s not in (sub_external_id, sub_name)
                            ]
                            print(f"Subrole '{sub_external_id}' validated under mainRole '{role}'")
                else:
                    self.errorVar.append(f"Failed to fetch subroles for mainRole '{role}'")
            else:
                self.errorVar.append(f"MainRole '{role}' not found in API")

        for s in remaining_subroles:
            self.errorVar.append(f"Subrole '{s}' not found in any of the provided mainRoles.")

        for err in self.errorVar:
            print(err)
        return validated_main_role_ids, validated_subrole_ids_list

    def convert_to_date(self, date_str):
        return datetime.strptime(date_str, "%d-%m-%Y")

    def prepareProgramSuccessSheet(self, MainFilePath, solutionName_for_folder_path, programFile,
                                   solutionExternalId, solutionId, accessToken, programdetails, userRole):
        urlFetchSolutionApi = internal_kong_ip + dbfindapi_url
        headerFetchSolutionApi = apiHeader.headers().headersFetchSolutionDetails(
            programdetails.get('TenantID'), programdetails.get('Org ID'), accessToken, userRole
        )
        payload = json.dumps({
            "query": {"_id": solutionId},
            "mongoIdKeys": ["_id", "solutionId", "metaInformation.solutionId"],
            "limit": 10000,
        })
        try:
            responseFetchSolution = self._post(
                urlFetchSolutionApi, headers=headerFetchSolutionApi,
                data=payload, timeout=(10, 60),
            )
        except RuntimeError as e:
            self.errorVar.append(str(e))
            return None

        print(responseFetchSolution.text, "responseFetchSolutionApi")
        responseFetchSolutionJson = responseFetchSolution.json()
        self.createAPILog(solutionName_for_folder_path, [
            "Solution Fetch.",
            "solution name : " + responseFetchSolutionJson["result"][-1]["name"],
            "solution ExternalId : " + responseFetchSolutionJson["result"][-1]["externalId"],
            f"Upload status code : {responseFetchSolution.status_code}",
        ])

        if responseFetchSolution.status_code == 200:
            solutionName = responseFetchSolutionJson["result"][-1]["name"]

        urlFetchSolutionLinkApi = internal_kong_ip + fetchlink + solutionId
        headerFetchSolutionLinkApi = apiHeader.headers().headersFetchSolutionLink(
            programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole
        )
        try:
            responseFetchSolutionLink = self._get(
                urlFetchSolutionLinkApi, headers=headerFetchSolutionLinkApi, timeout=(10, 60),
            )
        except RuntimeError as e:
            self.errorVar.append(str(e))
            return None

        print(responseFetchSolutionLink.text, "responseFetchSolutionLinkApi")
        self.createAPILog(solutionName_for_folder_path, [
            "Solution Fetch Link.", f"solution id : {solutionId}",
            f"solution ExternalId : {solutionExternalId}",
            f"Upload status code : {responseFetchSolutionLink.status_code}",
        ])

        if responseFetchSolutionLink.status_code == 200:
            print('Fetch solution Link Api Success')
            responseProjectUploadJson = responseFetchSolutionLink.json()
            solutionLink = ','.join(responseProjectUploadJson["result"])
            print(solutionLink, "solutionLink")
            self.createAPILog(solutionName_for_folder_path, [f"Response : {responseFetchSolutionLink.text}"])

            programFileBase = str(programFile).replace(".xlsx", "")
            success_file_path = os.path.join(MainFilePath, programFileBase + '-SuccessSheet.xlsx')
            os.makedirs(os.path.dirname(success_file_path), exist_ok=True)

            xfile = openpyxl.load_workbook(success_file_path if os.path.exists(success_file_path) else programFile)
            resourceDetailsSheet = xfile["Resource Details"]
            greenFill = PatternFill(start_color='0000FF00', end_color='0000FF00', fill_type='solid')
            rowCountRD = resourceDetailsSheet.max_row

            for row in range(3, rowCountRD + 1):
                cell_b = str(resourceDetailsSheet["B" + str(row)].value).strip().lower()
                cell_a = str(resourceDetailsSheet["A" + str(row)].value).strip()

                if cell_b == "course":
                    resourceDetailsSheet["D1"].value = ""
                    resourceDetailsSheet["E1"].value = ""
                    resourceDetailsSheet['I2'].value = "External id of the resource"
                    resourceDetailsSheet['J2'].value = "link to access the resource/Response"
                    for col in ['I2', 'J2']:
                        resourceDetailsSheet[col].fill = greenFill
                    resourceDetailsSheet['I' + str(row)].value = solutionExternalId
                    resourceDetailsSheet['J' + str(row)].value = (
                        "The course has been successfully mapped to the program"
                    )
                    for col in ['I', 'J']:
                        resourceDetailsSheet[col + str(row)].fill = greenFill

                elif cell_a == solutionName:
                    resourceDetailsSheet["D1"].value = ""
                    resourceDetailsSheet["E1"].value = ""
                    resourceDetailsSheet['I2'].value = "External id of the resource"
                    resourceDetailsSheet['J2'].value = "link to access the resource/Response"
                    for col in ['I2', 'J2']:
                        resourceDetailsSheet[col].fill = greenFill
                    resourceDetailsSheet['I' + str(row)].value = solutionExternalId
                    resourceDetailsSheet['J' + str(row)].value = solutionLink
                    for col in ['I', 'J']:
                        resourceDetailsSheet[col + str(row)].fill = greenFill

            xfile.save(success_file_path)
            print("Program success sheet is created")
            self.solutionUpdate(
                solutionName_for_folder_path, accessToken, solutionId,
                {"status": "active", "isDeleted": False}, programdetails, userRole,
            )
            return solutionLink

    # ── The large orchestration methods are unchanged in logic; they already ──
    # ── call the above helpers, which now carry timeouts + retries.          ──

    def ObservationSolutionCreate(self, resource, parentFolder, accessToken, ProgramGlobalDict,
                                  programdetails, MainFilePath, programFile, userRole):
        finalObsRubricSolutionLink = ""
        impLedObsFlag = resource.get('typeofSolution') == 5
        wbObservation = resource.get("ResourceCre")
        millisecond = int(time.time() * 1000)

        if resource.get('typeofSolution') in [1, 5]:
            if not self.criteriaUpload(parentFolder, resource.get("ResourceCre"), millisecond, accessToken,
                                       "framework", impLedObsFlag, programdetails, userRole):
                self.errorVar.append("Criteria Upload Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            print("Criteria Upload success....")
            frameworkExternalId = self.frameWorkUpload(parentFolder, resource.get("ResourceCre"), millisecond,
                                                       accessToken, programdetails,
                                                       resource.get('typeofSolution'), userRole)
            if not frameworkExternalId:
                self.errorVar.append("Framework Upload Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            print(frameworkExternalId, "frameworkExternalId")
            observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
            if not self.themesUpload(parentFolder, resource.get("ResourceCre"), millisecond, accessToken,
                                     frameworkExternalId, False, programdetails, userRole):
                self.errorVar.append("Theme Upload Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            solutionId = self.createSolutionFromFramework(parentFolder, accessToken, frameworkExternalId,
                                                          programdetails, userRole)
            if not solutionId:
                self.errorVar.append("Unable to create Solution From Framework.")
                return finalObsRubricSolutionLink, self.errorVar
            print(solutionId, "solutionId")
            print("parent solution created....")
            ECMSheet = wbObservation['ecms or domains']
            ECMUpdate = self.ECMUpdatebody(ECMSheet, millisecond)
            if not self.solutionUpdate(parentFolder, accessToken, solutionId, ECMUpdate[0], programdetails, userRole):
                self.errorVar.append("Solution Update Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            if not self.solutionUpdate(parentFolder, accessToken, solutionId, ECMUpdate[1], programdetails, userRole):
                self.errorVar.append("Solution Update Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            Resourcedet = wbObservation.get("details")
            ObsEntityType = Resourcedet.get("entity_type")
            parentEntityKey = "state" if ObsEntityType.strip().lower() in ['state', 'district', 'block', 'cluster', 'school'] else None
            detailsSheet = wbObservation['details']
            criteriaLevelsReport = not detailsSheet.get("scoring_system").lower() == "null"
            bodySolutionUpdate = {
                "status": "active", "isDeleted": False,
                "criteriaLevelReport": criteriaLevelsReport, "parentEntityKey": parentEntityKey,
            }
            if not self.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate, programdetails, userRole):
                self.errorVar.append("Solution Update Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            if not self.questionUpload(wbObservation, parentFolder, frameworkExternalId, millisecond,
                                       accessToken, solutionId, resource.get('typeofSolution'),
                                       programdetails, userRole):
                self.errorVar.append("question Upload Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            if not programdetails.get('scoring_system') == "null":
                if not self.solutionUpdate(parentFolder, accessToken, solutionId,
                                           {"isRubricDriven": True}, programdetails, userRole):
                    self.errorVar.append("Solution Update Failed.")
                    return finalObsRubricSolutionLink, self.errorVar
                if not self.fetchSolutionCriteria(parentFolder, observationExternalId, accessToken,
                                                  programdetails, userRole):
                    self.errorVar.append("fetchSolutionCriteria Failed.")
                    return finalObsRubricSolutionLink, self.errorVar
                if not self.uploadCriteriaRubrics(parentFolder, wbObservation, millisecond, accessToken,
                                                  frameworkExternalId, True, programdetails, userRole):
                    self.errorVar.append("upload Criteria Rubrics Failed.")
                    return finalObsRubricSolutionLink, self.errorVar
                if not self.uploadThemeRubrics(parentFolder, wbObservation, accessToken, frameworkExternalId,
                                               True, programdetails, userRole):
                    self.errorVar.append("upload Theme Rubrics Failed.")
                    return finalObsRubricSolutionLink, self.errorVar
            else:
                print("Observation with scoring system : null.")
            ObsDictVal = resource.get("ResourceCre")
            ObsDet = ObsDictVal.get("details")
            allow_multiple = ObsDet.get("allow_multiple_submissions") in [1, 'TRUE', True]
            bodySolutionUpdate = {
                'allowMultipleAssessemts': allow_multiple,
                "creator": programdetails.get('Name_of_the_creator'),
                "parentEntityKey": parentEntityKey,
            }
            if not self.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate, programdetails, userRole):
                self.errorVar.append("Solution Update Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            childId = self.createChild(parentFolder, wbObservation, solutionId, accessToken, programdetails, userRole)
            if not childId:
                self.errorVar.append("createChild Failed.")
                return finalObsRubricSolutionLink, self.errorVar
            if childId[0]:
                print("Fetching solution details")
                solutionDetails = self.fetchSolutionDetailsFromProgramSheet(
                    parentFolder, programdetails, childId[0], accessToken, ProgramGlobalDict, userRole
                )
                if not solutionDetails:
                    self.errorVar.append("Fetch solution details API Failed.")
                    return finalObsRubricSolutionLink, self.errorVar
                self.solutionUpdate(parentFolder, accessToken, childId[0],
                                    {"status": "inactive", "isDeleted": True}, programdetails, userRole)
                scopeRoles = [r.strip() for r in solutionDetails[0].split(",") if r.strip()]
                scopeSubRoles = [r.strip() for r in solutionDetails[1].split(",") if r.strip()]
                verifiedRoles = self.validate_roles_against_api(scopeRoles, scopeSubRoles, programdetails, parentFolder)
                scope = {}
                scope.update(programdetails.get('entityHierarchy'))
                scope["organizations"] = programdetails.get('OrgID', '')
                scope["professional_subroles"] = verifiedRoles[1]
                scope["professional_role"] = verifiedRoles[0]
                if not self.solutionUpdate(parentFolder, accessToken, childId[0],
                                           {"scope": scope}, programdetails, userRole):
                    self.errorVar.append("Solution Update Failed.")
                    return finalObsRubricSolutionLink, self.errorVar
                ReffStart = self.convert_to_date(programdetails.get('Startdateofprogram'))
                ReffEnd = self.convert_to_date(programdetails.get('Enddateofprogram'))
                solStart = self.convert_to_date(solutionDetails[2])
                solEnd = self.convert_to_date(solutionDetails[3])
                if ReffStart <= solStart <= ReffEnd and ReffStart <= solEnd <= ReffEnd:
                    print("dates validated...")
                    if solutionDetails[2]:
                        s = str(solutionDetails[2]).split("-")
                        if not self.solutionUpdate(parentFolder, accessToken, childId[0],
                                                   {"startDate": f"{s[2]}-{s[1]}-{s[0]} 00:00:00"},
                                                   programdetails, userRole):
                            self.errorVar.append("Solution Update Failed.")
                            return finalObsRubricSolutionLink, self.errorVar
                    if solutionDetails[3]:
                        e = str(solutionDetails[3]).split("-")
                        if not self.solutionUpdate(parentFolder, accessToken, childId[0],
                                                   {"endDate": f"{e[2]}-{e[1]}-{e[0]} 23:59:59"},
                                                   programdetails, userRole):
                            self.errorVar.append("Solution Update Failed.")
                            return finalObsRubricSolutionLink, self.errorVar
                else:
                    self.errorVar.append("Date Mismatched! Creation Stopped.")
                    print("Date Mismatched! Creation Stopped")
                    return finalObsRubricSolutionLink, self.errorVar
                ObsRubricSolutionLink = self.prepareProgramSuccessSheet(
                    MainFilePath, parentFolder, programFile, childId[1], childId[0],
                    accessToken, programdetails, userRole,
                )
                if not ObsRubricSolutionLink:
                    self.errorVar.append("Solution Fetch Link Failed.")
                    return finalObsRubricSolutionLink, self.errorVar
                finalObsRubricSolutionLink = ObsRubricSolutionLink
                return finalObsRubricSolutionLink, self.errorVar

        elif resource.get('typeofSolution') == 2:
            finalObsSolutionLink = ""
            if not self.criteriaUpload(parentFolder, resource.get("ResourceCre"), millisecond, accessToken,
                                       "criteria", False, programdetails, userRole):
                self.errorVar.append("Criteria Upload Failed.")
                return finalObsSolutionLink, self.errorVar
            print("Criteria Upload success....")
            frameworkExternalId = self.frameWorkUpload(parentFolder, resource.get("ResourceCre"), millisecond,
                                                       accessToken, programdetails,
                                                       resource.get('typeofSolution'), userRole)
            if not frameworkExternalId:
                self.errorVar.append("Framework Upload Failed.")
                return finalObsSolutionLink, self.errorVar
            observationExternalId = frameworkExternalId + "-OBSERVATION-TEMPLATE"
            if not self.themesUpload(parentFolder, resource.get("ResourceCre"), millisecond, accessToken,
                                     frameworkExternalId, True, programdetails, userRole):
                self.errorVar.append("Theme Upload Failed.")
                return finalObsSolutionLink, self.errorVar
            solutionId = self.createSolutionFromFramework(parentFolder, accessToken, frameworkExternalId,
                                                          programdetails, userRole)
            if not solutionId:
                self.errorVar.append("Unable to create Solution From Framework.")
                return finalObsSolutionLink, self.errorVar
            if not self.solutionUpdate(parentFolder, accessToken, solutionId,
                                       {"sections": {'S1': 'Observation Question'}}, programdetails, userRole):
                self.errorVar.append("Solution Update Failed.")
                return finalObsSolutionLink, self.errorVar
            ecmObj = {
                "evidenceMethods": {
                    'OB': {
                        'externalId': 'OB', 'tip': None, 'name': 'Observation',
                        'description': None, 'modeOfCollection': 'onfield',
                        'canBeNotApplicable': False, 'notApplicable': False,
                        'canBeNotAllowed': False, 'remarks': None,
                    }
                }
            }
            if not self.solutionUpdate(parentFolder, accessToken, solutionId, ecmObj, programdetails, userRole):
                self.errorVar.append("Solution Update Failed.")
                return finalObsSolutionLink, self.errorVar
            if not self.questionUpload(wbObservation, parentFolder, frameworkExternalId, millisecond,
                                       accessToken, solutionId, resource.get('typeofSolution'),
                                       programdetails, userRole):
                self.errorVar.append("question Upload Failed.")
                return finalObsSolutionLink, self.errorVar
            if programdetails.get('scoring_system') is not None:
                if not self.uploadCriteriaRubrics(parentFolder, wbObservation, millisecond, accessToken,
                                                  frameworkExternalId, False, programdetails, userRole):
                    self.errorVar.append("upload Criteria Rubrics Failed.")
                    return finalObsSolutionLink, self.errorVar
                if not self.uploadThemeRubrics(parentFolder, wbObservation, accessToken, frameworkExternalId,
                                               False, programdetails, userRole):
                    self.errorVar.append("upload Theme Rubrics Failed.")
                    return finalObsSolutionLink, self.errorVar
            else:
                print("Observation with scoring system : null.")
            Resourcedet = wbObservation.get("details")
            ObsEntityType = Resourcedet.get("entity_type")
            parentEntityKey = "state" if ObsEntityType.strip().lower() in ['state', 'district', 'block', 'cluster', 'school'] else None
            bodySolutionUpdate = {
                "status": "active", "isDeleted": False, "allowMultipleAssessemts": True,
                "creator": programdetails.get('Name_of_the_creator'), "parentEntityKey": parentEntityKey,
            }
            if not self.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate, programdetails, userRole):
                self.errorVar.append("Solution Update Failed.")
                return finalObsSolutionLink, self.errorVar
            childId = self.createChild(parentFolder, wbObservation, solutionId, accessToken, programdetails, userRole)
            if not childId:
                self.errorVar.append("createChild Failed.")
                return finalObsSolutionLink, self.errorVar
            if childId[0]:
                print("Fetching solution details")
                solutionDetails = self.fetchSolutionDetailsFromProgramSheet(
                    parentFolder, programdetails, childId[0], accessToken, ProgramGlobalDict, userRole
                )
                if not solutionDetails:
                    self.errorVar.append("Fetch solution details API Failed.")
                    return finalObsSolutionLink, self.errorVar
                self.solutionUpdate(parentFolder, accessToken, childId[0],
                                    {"status": "inactive", "isDeleted": True}, programdetails, userRole)
                scopeRoles = [r.strip() for r in solutionDetails[0].split(",") if r.strip()]
                scopeSubRoles = [r.strip() for r in solutionDetails[1].split(",") if r.strip()]
                verifiedRoles = self.validate_roles_against_api(scopeRoles, scopeSubRoles, programdetails, parentFolder)
                scope = {}
                scope.update(programdetails.get('entityHierarchy'))
                scope["organizations"] = programdetails.get('OrgID', '')
                scope["professional_subroles"] = verifiedRoles[1]
                scope["professional_role"] = verifiedRoles[0]
                if not self.solutionUpdate(parentFolder, accessToken, childId[0],
                                           {"scope": scope}, programdetails, userRole):
                    self.errorVar.append("Solution Update Failed.")
                    return finalObsSolutionLink, self.errorVar
                ReffStart = self.convert_to_date(programdetails.get('Startdateofprogram'))
                ReffEnd = self.convert_to_date(programdetails.get('Enddateofprogram'))
                solStart = self.convert_to_date(solutionDetails[2])
                solEnd = self.convert_to_date(solutionDetails[3])
                if ReffStart <= solStart <= ReffEnd and ReffStart <= solEnd <= ReffEnd:
                    print("dates validated...")
                    if solutionDetails[2]:
                        s = str(solutionDetails[2]).split("-")
                        if not self.solutionUpdate(parentFolder, accessToken, childId[0],
                                                   {"startDate": f"{s[2]}-{s[1]}-{s[0]} 00:00:00"},
                                                   programdetails, userRole):
                            self.errorVar.append("Solution Update Failed.")
                            return finalObsSolutionLink, self.errorVar
                    if solutionDetails[3]:
                        e = str(solutionDetails[3]).split("-")
                        if not self.solutionUpdate(parentFolder, accessToken, childId[0],
                                                   {"endDate": f"{e[2]}-{e[1]}-{e[0]} 23:59:59"},
                                                   programdetails, userRole):
                            self.errorVar.append("Solution Update Failed.")
                            return finalObsSolutionLink, self.errorVar
                else:
                    self.errorVar.append("Date Mismatched! Creation Stopped.")
                    print("Date Mismatched! Creation Stopped")
                    return finalObsSolutionLink, self.errorVar
                ObsSolutionLink = self.prepareProgramSuccessSheet(
                    MainFilePath, parentFolder, programFile, childId[1], childId[0],
                    accessToken, programdetails, userRole,
                )
                print(ObsSolutionLink, "ObsSolutionLink")
                if not ObsSolutionLink:
                    self.errorVar.append("Solution Fetch Link Failed.")
                    return finalObsSolutionLink, self.errorVar
                finalObsSolutionLink = ObsSolutionLink
                return finalObsSolutionLink, self.errorVar
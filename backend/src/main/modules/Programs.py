from backend.src.main.modules.common_config import *
from backend.src.main.modules.Observation import *
from backend.src.main.modules.Survey import *
from backend.src.main.modules.Project import *
from dotenv import load_dotenv
from pathlib import Path
import os, json, requests, sys, csv, shutil, wget
from datetime import datetime
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

class Programs:

    def __init__(self):
        self.errorVar = []

    def createAPILog(solutionName_for_folder_path, messageArr):
        print(solutionName_for_folder_path, "solutionName_for_folder_path")
        file_exists = os.path.join(solutionName_for_folder_path, 'apiHitLogs', 'apiLogs.txt')

        # ✅ Ensure folder structure exists
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

    def apicheckslog(solutionName_for_folder_path, messageArr):
        file_exists = os.path.join(solutionName_for_folder_path, 'apiHitLogs', 'apiLogs.csv')
        fileheader = ["Resource", "Process", "Status", "Remark"]

        # ✅ Ensure folder exists
        os.makedirs(os.path.dirname(file_exists), exist_ok=True)

        # ✅ Create file with header if it doesn't exist
        if not os.path.exists(file_exists):
            with open(file_exists, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',', lineterminator='\n')
                writer.writerow(fileheader)  # <-- use writerow, not writerows for a single header row

        # ✅ Append a new log entry
        with open(file_exists, 'a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',', lineterminator='\n')
            writer.writerow(messageArr)  # <-- use writerow for a single record

    def CheckProgramExistance(self, programdetails, accessToken, parentFolder):
        mainRoleStr = programdetails.get('Targetedroleatprogramlevel', "")
        mainRoles = [r.strip() for r in mainRoleStr.split(",") if r.strip()]

        rolesPGM = programdetails.get('Targetedsubroleatprogramlevel', "")
        subRoles = [r.strip() for r in str(rolesPGM).split(",") if r.strip()]

        # Validate roles against API
        verifiedRoles = self.validate_roles_against_api(mainRoles, subRoles, programdetails, parentFolder)
        mainRoleproff = verifiedRoles[0]
        rolesPGMID = verifiedRoles[1]

        scope={}
        scope["organizations"] = programdetails.get('OrgID')
        scope["professional_subroles"] = rolesPGMID
        scope["professional_role"] = mainRoleproff
        entityHierarchy = self.fetchEntityParentChilds(accessToken, programdetails, parentFolder)
        scope.update(entityHierarchy)
        programdetails['entityHierarchy'] = entityHierarchy

        program_name = (programdetails.get('TitleoftheProgram') or "").strip()
        tenant_id = programdetails.get('TenantID')

        programUrl = elevateprojecthost + fetchprograminfoapiurl
        payload = json.dumps({
            "query": {
                "name": program_name,
                "isAPrivateProgram": False,
                "status": "active",
                "tenantId": tenant_id
            },
            "mongoIdKeys": []
        })
        headersProgramSearch = apiHeader.headers().programSearchHeaders(accessToken)

        print(f"Checking program existence for: '{program_name}' (Tenant: {tenant_id})")

        try:
            responseProgramSearch = requests.post(programUrl, headers=headersProgramSearch, data=payload)
            messageArr = []
            messageArr.append("Program Search API")
            messageArr.append("URL : " + programUrl)
            messageArr.append("Status Code : " + str(responseProgramSearch.status_code))
            messageArr.append("Response : " + str(responseProgramSearch.text))
            if responseProgramSearch.status_code != 200:
                print("Program search API failed...")
                messageArr.append("Program search API failed...")
                Programs.createAPILog(parentFolder, messageArr)
                print("Response Code:", responseProgramSearch.status_code)
                self.errorVar.append(str(responseProgramSearch.text))
                return False

            print("---> Program fetch API Success")
            data = responseProgramSearch.json()
            result = data.get("result", [])
            countOfPrograms = len(result)

            if countOfPrograms == 0:
                messageArr.append("No program found with the name : " + str(program_name.lstrip().rstrip()))
                messageArr.append("******************** Preparing for program Upload **********************")
                print(f"No program found with the name: {program_name}")
                print("******************** Preparing for program Upload **********************")
                Programs.createAPILog(parentFolder, messageArr)
                fileheader = ["Program name fetch","Successfully fetched program name","Passed"]
                Programs.apicheckslog(parentFolder,fileheader)
                return False

            getProgramDetails = []
            for eachPgm in result:
                if not eachPgm.get("isAPrivateProgram", True):
                    # Populate programdetails once
                    if not programdetails.get("_id"):
                        programdetails["_id"] = eachPgm.get("_id")
                        # programdetails["externalId"] = eachPgm.get("externalId")
                        # programdetails["isAPrivateProgram"] = eachPgm.get("isAPrivateProgram")
                    
                    # Append a dictionary instead of list
                    getProgramDetails.append({
                        "_id": eachPgm.get("_id"),
                        "externalId": eachPgm.get("externalId"),
                        "description": eachPgm.get("description"),
                        "isAPrivateProgram": eachPgm.get("isAPrivateProgram")
                    })

            
            messageArr = []
            if len(getProgramDetails) == 0:
                print(f"Total 0 backend programs found with the name: {program_name}")
                messageArr.append("Total " + str(len(getProgramDetails)) + " backend programs found with the name : " + program_name.lstrip().rstrip())
                Programs.createAPILog(parentFolder, messageArr)
                fileheader = ["program find api is running","found"+str(len(
                    getProgramDetails))+"programs in backend","Failed","found"+str(len(
                    getProgramDetails))+"programs ,check logs"]
                Programs.apicheckslog(parentFolder,fileheader)
                Programs.createAPILog(parentFolder, messageArr)
            elif len(getProgramDetails) > 1:
                print(f"Total {len(getProgramDetails)} backend programs found with the name: {program_name}")
                messageArr.append("Total " + str(len(getProgramDetails)) + " backend programs found with the name : " + program_name.lstrip().rstrip())
                Programs.createAPILog(parentFolder, messageArr)
            else:
                print(f"Program '{program_name}' already exists.")
                messageArr.append(f"Program '{program_name}' already exists.")
                Programs.createAPILog(parentFolder, messageArr)
                # programdetails['ExistingProgramDB'] = getProgramDetails
                return True

        except Exception as e:
            print("Error while checking program existence:", str(e))
            self.errorVar.append(e)
            return False

        return True

    def parse_excel_date(self,date_value, end_of_day=False):
        """
        Converts date strings or Excel serial numbers into a proper datetime string.
        Adds 00:00:00 or 23:59:59 depending on `end_of_day`.
        """
        if not date_value:
            return None

        try:
            # If it's a float (Excel serial)
            if isinstance(date_value, (int, float)):
                from datetime import timedelta
                base_date = datetime(1899, 12, 30)  # Excel date base
                date_obj = base_date + timedelta(days=float(date_value))
            else:
                # Try multiple formats safely
                for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
                    try:
                        date_obj = datetime.strptime(str(date_value).strip(), fmt)
                        break
                    except ValueError:
                        continue
                else:
                    self.errorVar.append(f"Unknown date format: {date_value}")

            time_part = "23:59:59" if end_of_day else "00:00:00"
            return date_obj.strftime(f"%Y-%m-%d {time_part}")

        except Exception as e:
            self.errorVar.append(f"Date parsing error for '{date_value}': {e}")
            return None

    def validate_roles_against_api(self, mainRoles, subRoles, programdetails, parentFolder):
        urlFetchRoleList = userLoginHost + fetchprofessionalRole
        validateRoleHeader = apiHeader.headers().validateRoleHeaders(programdetails.get('TenantID'))
        payload = {}

        response = requests.request("GET", urlFetchRoleList, headers=validateRoleHeader, data=payload)
        # response = requests.get(urlFetchRoleList, headers=headers, data=json.dumps({}))

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
                subrole_url = f"{userLoginHost}entity-management/v1/entities/subEntityList/{main_role_id}?type=professional_subroles"
                subrole_resp = requests.request("GET",subrole_url, headers=validateRoleHeader,data=payload)
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
        Programs.createAPILog(parentFolder, messageArr)
        for err in self.errorVar:
            print(err)

        return validated_main_role_ids, validated_subrole_ids_list   
    
    def ensure_list(entity):
            if isinstance(entity, list):
                return [str(e).strip() for e in entity if str(e).strip()]
            elif isinstance(entity, str):
                return [s.strip() for s in entity.split(",") if s.strip()]
            else:
                return []
            
    def fetchEntityType(self, programdetails, entitiesPGM, scopeEntityType,
                    schoolEntitiesPGM, clusterEntitiesPGM, blockEntitiesPGM,
                    districtEntitiesPGM, stateEntitiesPGM):

        # Ensure all entity inputs are lists
        stateEntitiesPGM = Programs.ensure_list(stateEntitiesPGM)
        districtEntitiesPGM = Programs.ensure_list(districtEntitiesPGM)
        blockEntitiesPGM = Programs.ensure_list(blockEntitiesPGM)
        clusterEntitiesPGM = Programs.ensure_list(clusterEntitiesPGM)
        schoolEntitiesPGM = Programs.ensure_list(schoolEntitiesPGM)
        entitiesPGM = Programs.ensure_list(entitiesPGM)

        urlFetchEntityListApi = elevateentityhost + searchforlocation
        headerFetchEntityListApi = apiHeader.headers().PheaderFetchEntitytype()

        entityTypes = []
        entityTypeID = []

        for entityName in entitiesPGM:
            entityName = entityName.strip()
            payload = {
                "query": {
                    "metaInformation.name": entityName,
                    "tenantId": programdetails.get('TenantID'),
                    "entityType": scopeEntityType[0]
                },
                "projection": ["entityType", "_id", "metaInformation.name"]
            }

            responseFetchEntityListApi = requests.post(url=urlFetchEntityListApi,
                                                    headers=headerFetchEntityListApi,
                                                    data=json.dumps(payload))

            if responseFetchEntityListApi.status_code != 200:
                self.errorVar.append(
                    f"Failed to fetch entity type for '{entityName}'. Status code: {responseFetchEntityListApi.status_code}")
                continue

            responseFetchEntityListApi = responseFetchEntityListApi.json()
            entityToUpload = None

            for listEntities in responseFetchEntityListApi.get('result', []):
                entityToUpload = listEntities.get('entityType')
                EntityName = listEntities.get('metaInformation', {}).get('name', '')
                entityId = listEntities.get('_id')

                DetailsFetchURL = elevateentityhost + fetchDetailsEntity + entityId
                headerEntityDetails = apiHeader.headers().headerFetchEntityDetails(programdetails.get('TenantID'))
                EntityDetailsResponse = requests.get(DetailsFetchURL, headers=headerEntityDetails)

                if EntityDetailsResponse.status_code != 200:
                    self.errorVar.append(
                        f"Failed to fetch entity details for '{entityName}'. Status code: {EntityDetailsResponse.status_code}")
                    continue

                EntityDetailsResponseJson = EntityDetailsResponse.json()
                entities = EntityDetailsResponseJson.get("result", [])

                for entity in entities:
                    entityId = entity.get("_id")
                    entityToUpload = entity.get("entityType")
                    parent_info = entity.get("parentInformation", {})
                    EntityFlag = False

                    # Extract names safely
                    school_name = (parent_info.get("school") or [{}])[0].get("name", "")
                    cluster_name = (parent_info.get("cluster") or [{}])[0].get("name", "")
                    block_name = (parent_info.get("block") or [{}])[0].get("name", "")
                    district_name = (parent_info.get("district") or [{}])[0].get("name", "")
                    state_name = (parent_info.get("state") or [{}])[0].get("name", "")

                    # Check each entity list for a match
                    if schoolEntitiesPGM and school_name in schoolEntitiesPGM:
                        EntityFlag = True
                    elif clusterEntitiesPGM and cluster_name in clusterEntitiesPGM:
                        EntityFlag = True
                    elif blockEntitiesPGM and block_name in blockEntitiesPGM:
                        EntityFlag = True
                    elif districtEntitiesPGM and district_name in districtEntitiesPGM:
                        EntityFlag = True
                    elif stateEntitiesPGM and state_name in stateEntitiesPGM:
                        EntityFlag = True
                    else:
                        EntityFlag = True  # fallback if no lists are provided

                    if EntityFlag:
                        entityTypes.append(entityToUpload)
                        entityTypeID.append(entityId)
                        print(entityToUpload, entityId)
                    else:
                        print(f"Entity '{EntityName}' did not match any provided entity list.")

            if not entityToUpload:
                self.errorVar.append(f"Entity type not found for entity '{entityName}'.")

        return entityTypes, entityTypeID

    def fetchEntityParentChilds(self, accessToken, programdetails, parentFolder):
        stateEntitiesPGM = programdetails.get('Targetedstateatprogramlevel')
        districtEntitiesPGM = programdetails.get('TargetedDistrictatprogramlevel')
        blockEntitiesPGM = programdetails.get('TargetedBlockatprogramlevel')
        clusterEntitiesPGM = programdetails.get('TargetedClusteratprogramlevel')
        schoolEntitiesPGM = programdetails.get('TargetedSchoolatprogramlevel')
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
        scopeEntityType = [EntityType] if isinstance(EntityType, str) else EntityType
        entitiesType = self.fetchEntityType(programdetails, entitiesPGM, scopeEntityType,schoolEntitiesPGM,clusterEntitiesPGM,blockEntitiesPGM,districtEntitiesPGM,stateEntitiesPGM)
        programdetails['entitiesType'] = entitiesType
        programdetails['TargetedEntity'] = entitiesType[0]
        if entitiesPGM:
            entitiesPGM = entitiesPGM
            scopeEntityType = entitiesType[0]
        entitiesPGMID = entitiesType[1]
        try:
            if not isinstance(entitiesPGMID, list):
                entitiesPGMID = [entitiesPGMID]

            hierarchy = ["state", "district", "block", "cluster", "school"]
            merged_output = {level: [] for level in hierarchy}

            for entityId in entitiesPGMID:
                urlFetchEntity = elevateentityhost + fetchDetailsEntity + entityId

                headers = apiHeader.headers().headerFetchDetailsEntity(programdetails.get('TenantID'), accessToken)

                response = requests.get(url=urlFetchEntity, headers=headers)
                messageArr = []
                messageArr.append("Fetched professional roles from: " + urlFetchEntity)
                messageArr.append("Status Code: " + str(response.status_code))
                Programs.createAPILog(parentFolder, messageArr)
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
                    self.errorVar.append(response.text)
                    messageArr = []
                    messageArr.append(f"---> Error in fetching entity details for {entityId}. "
                        f"Status {response.status_code} Response {response.text}")
                    Programs.createAPILog(parentFolder, messageArr)
                    self.errorVar.append(f"---> Error in fetching entity details for {entityId}. "
                        f"Status {response.status_code} Response {response.text}")

            # Final check: ensure each level has at least "ALL" if empty
            for level in hierarchy:
                if not merged_output[level]:
                    merged_output[level].append("ALL")

            print("Structured Entity Hierarchy:", json.dumps(merged_output, indent=2))            
            return merged_output

        except Exception as e:
            self.errorVar.append(str(e))
            return None

    def ProgramCreate(self,programdetails, accessToken, parentFolder, userRole):
        print("-----> Creating a Program...")
        startDateOfProgram = self.parse_excel_date(programdetails.get('Startdateofprogram'))
        endDateOfProgram = self.parse_excel_date(programdetails.get('Enddateofprogram'), end_of_day=True)

        mainRoleStr = programdetails.get('Targetedroleatprogramlevel', "")
        mainRoles = [r.strip() for r in mainRoleStr.split(",") if r.strip()]

        rolesPGM = programdetails.get('Targetedsubroleatprogramlevel', "")
        subRoles = [r.strip() for r in str(rolesPGM).split(",") if r.strip()]

        # Validate roles against API
        verifiedRoles = self.validate_roles_against_api(mainRoles, subRoles, programdetails, parentFolder)
        mainRoleproff = verifiedRoles[0]
        rolesPGMID = verifiedRoles[1]

        scope={}
        scope["organizations"] = programdetails.get('OrgID')
        scope["professional_subroles"] = rolesPGMID
        scope["professional_role"] = mainRoleproff
        entityHierarchy = self.fetchEntityParentChilds(accessToken, programdetails, parentFolder)
        scope.update(entityHierarchy)
        try:
            programCreationurl = elevateprojecthost + programcreationurl
            payload = json.dumps({
                    "externalId": (programdetails.get('ProgramID') or "").strip(),
                    "name": (programdetails.get('TitleoftheProgram') or "").strip(),
                    "description": (programdetails.get('DescriptionoftheProgram') or "").strip(),
                    "isDeleted": False,
                    "resourceType": [
                        "program"
                    ],
                    "language": [
                        "English"
                    ],
                    "metaInformation": {
                    "state": [
                        s.strip()
                        for s in (programdetails.get('Targetedstateatprogramlevel') or "").split(",")
                        if s.strip()
                    ],
                    "recommendedFor" : mainRoles
                    },
                    "keywords": [],
                    "concepts": [],
                    # "userId":userId,
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
            headerProgramCreateApi = apiHeader.headers().headerProgramCreate(accessToken, programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), userRole)

            responsePgmCreate = requests.request("POST", programCreationurl, headers=headerProgramCreateApi, data=(payload))
            messageArr = []
            messageArr.append("++++++++++++ Program Creation ++++++++++++")
            messageArr.append("Program Creation URL : " + programCreationurl)
            messageArr.append("Body : " + str(payload))
            messageArr.append("Program Creation Status Code : " + str(responsePgmCreate.status_code))
            messageArr.append("Program Creation Response : " + str(responsePgmCreate.text))
            messageArr.append("Program body : " + str(payload))
            Programs.createAPILog(parentFolder, messageArr)
            fileheader = [programdetails.get('TitleoftheProgram'), ('Program Sheet Validation'), ('Passed')]
            Programs.createAPILog(parentFolder, messageArr)
            Programs.apicheckslog(parentFolder, fileheader)
            if responsePgmCreate.status_code == 200:
                    responsePgmCreatejson = responsePgmCreate.json()
                    program_data = responsePgmCreatejson.get("result", {})
                    programID = program_data.get("_id")
                    if programID:
                        programdetails['_id'] = programID
                        print("Program created successfully.")
                        print("Program ID:", programID)
                        return True
                    else:
                        if responsePgmCreate.status_code in [400, 401, 403, 404, 422]:
                            self.errorVar.append(f"PgmCreate-Client Error {responsePgmCreate.status_code}: {responsePgmCreate.text}")
                        elif responsePgmCreate.status_code in [500, 502, 503, 504]:
                            self.errorVar.append(f"PgmCreate-Server Error {responsePgmCreate.status_code}: {responsePgmCreate.text}")
                        else:
                            self.errorVar.append(f"PgmCreate-Unexpected Error {responsePgmCreate.status_code}: {responsePgmCreate.text}")
                        print("Program creation API failed. Please check logs.")
                        return False
        except Exception as e:
            self.errorVar.append(str(e))
            return False
    
    def createFileStruct(self, MainFilePath, addSolutionFile):
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
    
    def programCheckCreate(self, programFile, MainFilePath, parentFolder, ProgramGlobalDict, PRName, accessToken,userRole):
        programdetails_list = ProgramGlobalDict.get('Program Details', [])
        programdetails = programdetails_list[0] if programdetails_list else {}
        # accessToken = self.generateAccessToken(parentFolder)
        if not accessToken:
            print("Access token generation failed")
            return False

        IsExistingProgram = self.CheckProgramExistance(programdetails, accessToken, parentFolder)
        SolutionResults = {}
        if not IsExistingProgram:
            CreateProgram = self.ProgramCreate(programdetails, accessToken, parentFolder, userRole)
            if CreateProgram:
                print("Program created successfully ... ")
        for resource in ProgramGlobalDict['Program Resources']:
            dest_dir = "InputFiles"
            os.makedirs(dest_dir, exist_ok=True)
            match = re.search(r"/d/([a-zA-Z0-9-_]+)", resource['ResourceLink'])
            if not match:
                self.errorVar.append(f"Invalid Resource Link format for '{resource['Nameofresourcesinprogram']}'")
                continue
            file_id = match.group(1)
            file_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
            download_file = wget.download(file_url, out=dest_dir)
            solParentFolder = self.createFileStruct(parentFolder, download_file)
            if resource.get('typeofSolution') == 1 or resource.get('typeofSolution') == 5:

                print('Observation with rubrics File detected...')
                SolutionName = resource.get("Nameofresourcesinprogram")

                ObservationInstance = CreateObservation()
                ObservationCreation = ObservationInstance.ObservationSolutionCreate(resource, solParentFolder, accessToken, ProgramGlobalDict, programdetails, MainFilePath, programFile, userRole)
                if not ObservationCreation:
                    self.errorVar.append("Solution creation failed.")
                if ObservationCreation[1] == []:
                    SolutionResults[SolutionName] = ObservationCreation[0]
                else:
                    SolutionResults[SolutionName] = ObservationCreation[1]
            elif resource.get('typeofSolution') == 2:
                print('Observation without rubrics File detected...')
                SolutionName = resource.get("Nameofresourcesinprogram")

                ObservationInstance = CreateObservation()
                ObservationCreation = ObservationInstance.ObservationSolutionCreate(resource, solParentFolder, accessToken, ProgramGlobalDict, programdetails, MainFilePath, programFile, userRole)
                if not ObservationCreation:
                    self.errorVar.append("Solution creation failed.")
                if ObservationCreation[1] == []:
                    SolutionResults[SolutionName] = ObservationCreation[0]
                else:
                    SolutionResults[SolutionName] = ObservationCreation[1]
            elif resource.get('typeofSolution') == 3:
                print('Survey File detected...')
                SolutionName = resource.get("Nameofresourcesinprogram")
                SurveyInstance = CreateSurvey()
                SurveyCreation = SurveyInstance.CreateSurvey(resource,PRName, solParentFolder, accessToken, ProgramGlobalDict, programdetails, MainFilePath, programFile, userRole)
                if not SurveyCreation:
                    self.errorVar.append("Solution creation failed.")
                # self.errorVar.append(SurveyCreation[1])
                if SurveyCreation[1] == []:
                    SolutionResults[SolutionName] = SurveyCreation[0]
                else:
                    SolutionResults[SolutionName] = SurveyCreation[1]
            elif resource.get('typeofSolution') == 4:
                print('Project File detected...')
                SolutionName = resource.get("Nameofresourcesinprogram")
                ProjectInstance = CreateProject()
                ProjectCreation = ProjectInstance.CreateProject(resource, solParentFolder, accessToken, ProgramGlobalDict, programdetails, MainFilePath, programFile, userRole)
                if not ProjectCreation:
                    self.errorVar.append("Solution creation failed.")
                # self.errorVar.append(SurveyCreation[1])
                if ProjectCreation[1] == []:
                    SolutionResults[SolutionName] = ProjectCreation[0]
                else:
                    SolutionResults[SolutionName] = ProjectCreation[1]
        print(SolutionResults,"SolutionResults")
        return SolutionResults


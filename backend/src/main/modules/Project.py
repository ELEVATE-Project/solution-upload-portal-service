import os, uuid, requests, sys, json, time, csv, jwt, openpyxl, datetime, gdown
from openpyxl.styles import Color, PatternFill
from difflib import get_close_matches
from dotenv import load_dotenv
from pathlib import Path
from backend.src.main.modules.common_config import *
from difflib import get_close_matches
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

class CreateProject():

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

    def checkEntityOfSolution(self, projectName_for_folder_path, solutionNameOrId, accessToken, programdetails, userRole):
        urldbFind = internal_kong_ip + dbfindapi_url
        searchSolutionpayload = {}
        headerdbFindApi = apiHeader.headers().headerCheckEntityOfSolution(programdetails.get('TenantID'), programdetails.get('Org ID'), accessToken, userRole)
        searchSolutionpayload = json.dumps({
            "query": {
                "name": solutionNameOrId,
                "isReusable": True
            },
            "mongoIdKeys": [
                "_id",
                "solutionId",
                "metaInformation.solutionId"
            ],
            "limit": 10000
        })
        searchSolutionresponse = requests.request("POST", url=urldbFind, headers=headerdbFindApi,
                                                data=searchSolutionpayload)
        print(searchSolutionresponse.text,"searchSolutionresponse")
        if searchSolutionresponse.status_code == 200:
            searchSolutionjson = searchSolutionresponse.json()
            results = searchSolutionjson.get("result", [])
            print(results[-1].get("isReusable"))
            solutionId_parent = results[-1].get("_id")
            solutionEntityType = results[-1].get("entityType")
            solutionExternalId = results[-1].get("externalId")

            urldbFind = internal_kong_ip + dbfindapi_url
            searchSolutionpayload = {}
            headerdbFindApi = apiHeader.headers().headerCheckEntityOfSolution(programdetails.get('TenantID'), programdetails.get('Org ID'), accessToken, userRole)
            searchSolutionpayload = json.dumps({
                "query": {
                    "name": solutionNameOrId,
                    "isReusable": False
                },
                "mongoIdKeys": [
                    "_id",
                    "solutionId",
                    "metaInformation.solutionId"
                ],
                "limit": 10000
            })
            print(searchSolutionpayload,"searchSolutionpayload")
            searchSolutionresponse = requests.request("POST", url=urldbFind, headers=headerdbFindApi,
                                                    data=searchSolutionpayload)
            print(searchSolutionresponse.text,"searchSolutionresponse")
            if searchSolutionresponse.status_code == 200:
                searchSolutionjson = searchSolutionresponse.json()
                results = searchSolutionjson.get("result", [])
                print(results[-1].get("isReusable"))
                solutionId_child = results[-1].get("_id")
                return [solutionExternalId, solutionId_child, solutionEntityType]

            else:
                messageArr = [
                    "No solution Found..",
                    f"URL : {urldbFind}",
                    f"Status Code : {searchSolutionresponse.status_code}"
                    ]
                self.createAPILog(projectName_for_folder_path, messageArr)
                return False
        else:
            messageArr = [
                "No parent solution Found..",
                f"URL : {urldbFind}",
                f"Status Code : {searchSolutionresponse.status_code}"
                ]
            self.createAPILog(projectName_for_folder_path, messageArr)
            return False


    def ObservationsolutionUpdate(self, solutionName_for_folder_path, accessToken, solutionId, bodySolutionUpdate, programdetails, userRole):
        solutionUpdateApiurl = internal_kong_ip + solutionupdateapi + str(solutionId)
        headerUpdateSolutionApi = apiHeader.headers().headersObservationsolutionUpdate(programdetails.get('TenantID'), programdetails.get('Org ID'), accessToken, userRole)
        responseUpdateSolutionApi = requests.post(url=solutionUpdateApiurl, headers=headerUpdateSolutionApi,data=json.dumps(bodySolutionUpdate))
        messageArr = ["Solution Update API called.", "URL : " + str(solutionUpdateApiurl), "Body : " + str(bodySolutionUpdate),"Response : " + str(responseUpdateSolutionApi.text),"Status Code : " + str(responseUpdateSolutionApi.status_code)]
        self.createAPILog(solutionName_for_folder_path, messageArr)
        if responseUpdateSolutionApi.status_code == 200:
            print("Solution Update Success.")
            return True
        else:
            print("Solution Update Failed.")
            return False

    def prepareProjectAndTasksSheets(self, wbObservation, parentFolder, accessToken, programdetails, userRole):
        print("prepareProjectAndTasksSheets")

        millisecond = int(time.time() * 1000)
        projectFilePath = os.path.join(parentFolder, 'projectUpload')
        taskFilePath = os.path.join(parentFolder, 'taskUpload')

        os.makedirs(projectFilePath, exist_ok=True)
        os.makedirs(taskFilePath, exist_ok=True)

        # --------------------- Project Sheet ---------------------
        projectDetailsSheet = wbObservation.get("Project upload")
        if not projectDetailsSheet:
            self.errorVar.append("Missing 'Project upload' sheet in wbObservation.")

        # Normalize to list-of-rows format
        if isinstance(projectDetailsSheet, dict):
            keysProject = list(projectDetailsSheet.keys())
            projectRows = [list(projectDetailsSheet.values())]
        elif isinstance(projectDetailsSheet, list) and len(projectDetailsSheet) >= 2:
            keysProject = projectDetailsSheet[1]
            projectRows = projectDetailsSheet[2:]
        else:
            self.errorVar.append("Invalid structure for 'Project upload'. Expected dict or list-with-headers.")

        # Prepare project CSV headers
        projectColnames1 = [
            "title", "externalId", "categories", "recommendedFor", "description", 
            "entityType", "goal"
        ]
        learningResource_count = sum(1 for h in keysProject if str(h).startswith('learningResources')) // 2
        for lr_idx in range(1, learningResource_count + 1):
            projectColnames1 += [
                f"learningResources{lr_idx}-name",
                f"learningResources{lr_idx}-link",
                f"learningResources{lr_idx}-app",
                f"learningResources{lr_idx}-id"
            ]
        projectColnames2 = [
            "rationale", "primaryAudience", "taskCreationForm", "duration", "concepts",
            "keywords", "successIndicators", "risks", "approaches", "_arrayFields"
        ]
        projectColnames1 += projectColnames2

        projectCsvPath = os.path.join(projectFilePath, 'projectUpload.csv')
        write_header_project = not os.path.exists(projectCsvPath)

        with open(projectCsvPath, 'a', encoding='utf-8', newline='') as file:
            writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',', lineterminator='\n')
            if write_header_project:
                writer.writerow(projectColnames1)

            categories_list = ["teachers", "students", "infrastructure", "community", 
                            "educationLeader", "schoolProcess", "learner", "facilitator"]

            for row in projectRows:
                dictProjectDetails = {keysProject[i]: row[i] for i in range(len(keysProject))}
                title = str(dictProjectDetails.get("title", "")).strip()
                externalId = f"{dictProjectDetails.get('projectId', '')}-{millisecond}"
                entityType = str(dictProjectDetails.get("entityType", "")).strip()
                categories = str(dictProjectDetails.get("categories", "")).split(",")
                categories_final = ",".join([
                    get_close_matches(cat.strip().lower().replace(" ", ""), categories_list, n=1)[0]
                    for cat in categories if cat.strip()
                ])
                projectGoal = "TEMP"
                recommendedFor = str(dictProjectDetails.get("recommendedFor", "")).strip()
                objective = str(dictProjectDetails.get("objective", "")).strip()

                project_values = [title, externalId, categories_final, recommendedFor, objective, entityType, projectGoal]

                for lr_idx in range(1, learningResource_count + 1):
                    lr_name = str(dictProjectDetails.get(f"learningResources{lr_idx}-name", "")).strip()
                    lr_link = str(dictProjectDetails.get(f"learningResources{lr_idx}-link", "")).strip()
                    if lr_name == "" and lr_link == "":
                        project_values += ["", "", "", ""]
                    else:
                        lr_link_id = lr_link.split("/")[-1] if lr_link else ""
                        project_values += [lr_name, lr_link, "projectService", lr_link_id]

                for col in projectColnames2:
                    if col == "_arrayFields":
                        project_values.append("categories,primaryAudience,successIndicators,risks,approaches,recommendedFor")
                    else:
                        project_values.append(str(dictProjectDetails.get(col, "")).strip())

                writer.writerow(project_values)

        print("Project CSV prepared successfully")

        # --------------------- Tasks Sheet ---------------------
        tasksDetailsSheet = wbObservation.get("Tasks upload")
        if not tasksDetailsSheet:
            self.errorVar.append("Missing 'Tasks upload' sheet in wbObservation.")

        # Normalize tasks
        if isinstance(tasksDetailsSheet, dict):
            keysTasks = list(tasksDetailsSheet.keys())
            taskRows = [list(tasksDetailsSheet.values())]
        elif isinstance(tasksDetailsSheet, list) and len(tasksDetailsSheet) >= 1:
            if isinstance(tasksDetailsSheet[0], dict):
                keysTasks = list(tasksDetailsSheet[0].keys())
                taskRows = [list(task.values()) for task in tasksDetailsSheet]
            else:
                self.errorVar.append("Invalid format for 'Tasks upload'.")
        else:
            self.errorVar.append("Invalid structure for 'Tasks upload'.")

        taskColumns1 = [
            "name", "externalId", "description", "type", "hasAParentTask", "parentTaskOperator",
            "parentTaskValue", "parentTaskId", "solutionType", "solutionSubType", 
            "solutionId", "isDeletable", "isAnExternalTask"
        ]
        taskLearningResource_count = sum(1 for h in keysTasks if str(h).startswith('learningResources')) // 2
        for lr_idx in range(1, taskLearningResource_count + 1):
            taskColumns1 += [
                f"learningResources{lr_idx}-name",
                f"learningResources{lr_idx}-link",
                f"learningResources{lr_idx}-app",
                f"learningResources{lr_idx}-id"
            ]
        taskColumns1 += ["minNoOfSubmissionsRequired", "sequenceNumber", "redirectLink", "buttonLabel"]

        taskCsvPath = os.path.join(taskFilePath, 'taskUpload.csv')
        write_header_task = not os.path.exists(taskCsvPath)
        with open(taskCsvPath, 'a', encoding='utf-8', newline='') as file:
            writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',', lineterminator='\n')
            if write_header_task:
                writer.writerow(taskColumns1)

            sequenceNumber = 0
            for row in taskRows:
                dictTasksDetails = {keysTasks[i]: row[i] for i in range(len(keysTasks))}
                taskName = str(dictTasksDetails.get("TaskTitle", "")).strip()
                taskId = f"{dictTasksDetails.get('TaskId', '')}-{millisecond}"
                sequenceNumber += 1
                taskDescription = str(dictTasksDetails.get("description", "")).strip()
                if programdetails.get('TenantID') != 'shikshalokam':
                    Mitra_Link = str(dictTasksDetails["Mitra_Link"]).strip()
                    if Mitra_Link == "":
                        Mitra_Link = Mitra_Link
                if dictTasksDetails.get("solutionType"):
                    taskSolutionType = dictTasksDetails.get("solutionType", "")
                elif dictTasksDetails.get("learningResources1-name") != "" and dictTasksDetails.get("learningResources1-link") != "":
                    taskSolutionType = "content"
                elif programdetails.get('TenantID') != 'shikshalokam':
                    if dictTasksDetails["Mitra_Link"] != "":
                        taskSolutionType = "reflection"
                    else:
                        taskSolutionType = "simple"
                else:
                        taskSolutionType = "simple"

                taskType = taskSolutionType
                hasAParentTask = "YES" if dictTasksDetails.get("parentTaskId", "") else "NO"
                parentTaskId = f"{dictTasksDetails.get('parentTaskId', '')}-{millisecond}" if hasAParentTask == "YES" else ""
                parentTaskOperator = "EQUALS" if hasAParentTask == "YES" else ""
                parentTaskValue = "started" if hasAParentTask == "YES" else ""
                solutionSubType = dictTasksDetails.get("SolutionSubType", "")

                # solutionId = dictTasksDetails.get("SolutionId", "")
                if dictTasksDetails.get("Solution Name"):
                    solutionNameOrId = dictTasksDetails["Solution Name"]
                    print(solutionNameOrId, "solutionNameOrId----")

                    taskSolutionType = taskType
                    solutionDetailsInTask = self.checkEntityOfSolution(
                        parentFolder, solutionNameOrId, accessToken, programdetails, userRole
                    )

                    ObservationChildfrom = solutionDetailsInTask[1]
                    # ObservationChildfrom = solutionDetailsInTask[2] if len(solutionDetailsInTask) > 2 else None
                    print(ObservationChildfrom, "ObservationChild")

                    # Prepare body for update
                    bodysolutionUpdate = {
                        "status": "inactive",
                        "isDeleted": True
                    }

                    # Use whichever ID exists
                    child_id_to_update = ObservationChildfrom
                    if child_id_to_update:
                        self.ObservationsolutionUpdate(
                            parentFolder,
                            accessToken,
                            child_id_to_update,
                            bodysolutionUpdate,
                            programdetails,
                            userRole
                        )
                    print("observation child solution deactivated")
                    solutionSubType = solutionDetailsInTask[2]
                    solutionId = solutionDetailsInTask[0]

                    taskSolutionType = dictTasksDetails["solutionType"]
                else:
                    solutionId = ""
                AnExternalTask = "True" if str(dictTasksDetails.get("isAnExternalTask", "")).lower() == "yes" else "False"
                isDeletable = "TRUE" if str(dictTasksDetails.get("Mandatory task(Yes or No)", "")).lower() == "no" else "FALSE"
                taskminNoOfSubmissionsRequired = str(dictTasksDetails.get("Number of submissions for observation", "")).strip()

                task_values = [
                    taskName, taskId, taskDescription, taskType, hasAParentTask, 
                    parentTaskOperator, parentTaskValue, parentTaskId, 
                    taskSolutionType, solutionSubType, solutionId, isDeletable, AnExternalTask
                ]

                for lr_idx in range(1, taskLearningResource_count + 1):
                    lr_name = str(dictTasksDetails.get(f"learningResources{lr_idx}-name", "")).strip()
                    lr_link = str(dictTasksDetails.get(f"learningResources{lr_idx}-link", "")).strip()
                    if lr_link and not lr_name:
                        self.errorVar.append(f"Name is required for the learning resource with link: '{lr_link}'")
                    if lr_name == "" and lr_link == "":
                        task_values += ["", "", "", ""]
                    else:
                        lr_link_id = lr_link.split("/")[-1] if lr_link else ""
                        task_values += [lr_name, lr_link, "projectService", lr_link_id]
                
                redirectLink = []
                if programdetails.get('TenantID') != 'shikshalokam':
                    redirectLink.append(Mitra_Link)
                    if Mitra_Link.strip() != "":
                        redirectLink.append("Start Reflection")
                    else:
                        redirectLink.append("")
                task_values += [taskminNoOfSubmissionsRequired, sequenceNumber, redirectLink, ""]
                writer.writerow(task_values)
        if self.errorVar == []:
            print("Task CSV prepared successfully")
            return True
        else:
            return False

    def projectUpload(self, parentFolder, accessToken, programdetails, userRole):
        try:
            urlProjectUploadApi = elevateprojecthost + projectuploadapi
            print("Project Upload API URL:", urlProjectUploadApi)
            headerProjectUploadApi = apiHeader.headers().headersprojectUpload(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
            print("Project Upload API Headers:", headerProjectUploadApi)
            project_payload = {}
            filesProject = {
                'projectTemplates': open(parentFolder + '/projectUpload/projectUpload.csv', 'rb')
            }
            print(filesProject,"filesProject")
            responseProjectUploadApi = requests.post(url=urlProjectUploadApi, headers=headerProjectUploadApi,data=project_payload,files=filesProject)
            print(responseProjectUploadApi,"responseProjectUploadApi")
            messageArr = ["program mapping is success.","File path : " + parentFolder + '/projectUpload/projectUpload.csv']
            messageArr.append("Upload status code : " + str(responseProjectUploadApi.status_code))
            self.createAPILog(parentFolder, messageArr)
            if responseProjectUploadApi.status_code == 200:
                print('ProjectUploadApi Success')
                with open(parentFolder + '/projectUpload/projectInternal.csv','w+',encoding='utf-8') as projectRes:
                    projectRes.write(responseProjectUploadApi.text)

                    messageArr=["responnse :" ,responseProjectUploadApi.text ]
                    self.createAPILog(parentFolder,messageArr)
                    return True
            else:
                if responseProjectUploadApi.status_code in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"ProjectUploadApi-Client Error {responseProjectUploadApi.status_code}: {responseProjectUploadApi.text}")
                elif responseProjectUploadApi.status_code in [500, 502, 503, 504]:
                    self.errorVar.append(f"ProjectUploadApi-Server Error {responseProjectUploadApi.status_code}: {responseProjectUploadApi.text}")
                else:
                    self.errorVar.append(f"ProjectUploadApi-Unexpected Error {responseProjectUploadApi.status_code}: {responseProjectUploadApi.text}")

                print(self.errorVar)
                messageArr.append(f"Error Response: {responseProjectUploadApi.text}")
                self.createAPILog(parentFolder, messageArr)
                return False
        except Exception as e:
            self.errorVar.append(f"⚠️ Exception: {str(e)}")
            self.createAPILog(parentFolder, [f"Exception: {str(e)}"])
            return False

    def taskUpload(self, parentFolder, accessToken, programdetails, userRole):
        try:
            projectInternalfile = open(parentFolder + '/projectUpload/projectInternal.csv', mode='r',encoding='utf-8')
            projectInternalfile = csv.DictReader(projectInternalfile)
            for projectInternal in projectInternalfile:
                projectExternalId = projectInternal["externalId"]
                project_id = projectInternal["_SYSTEM_ID"]
                if str(project_id).strip() == "Could not pushed to kafka":
                    fetchProjectIdApi = elevateentityhost + fetchsolutiondetails
                    headersFetchSolAPI =apiHeader.headers().headersFetchSol(accessToken, programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), userRole)
                    fetchProjectIdPayload = {}

                    responseProjectListApi = requests.get(url=fetchProjectIdApi, headers=headersFetchSolAPI,
                                                        data=fetchProjectIdPayload)
                    messageArr = ["Tasks Upload Sheet Prepared.",
                                "File path : " + parentFolder + '/taskUpload/taskUpload.csv']
                    messageArr.append("URL : " + str(fetchProjectIdApi))
                    messageArr.append("Upload status code : " + str(responseProjectListApi.status_code))
                    self.createAPILog(parentFolder, messageArr)

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
                        self.createAPILog(parentFolder, messageArr) 
                        return False 

                urlTasksUploadApi = elevateprojecthost + taskuploadapi + project_id
                headerTasksUploadApi = apiHeader.headers().headersTaskUpload(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
                task_payload = {}
                filesTasks = {
                    'projectTemplateTasks': open(parentFolder + '/taskUpload/taskUpload.csv',
                                                'rb')
                }

                responseTasksUploadApi = requests.post(url=urlTasksUploadApi, headers=headerTasksUploadApi,
                                                    data=task_payload,
                                                    files=filesTasks)
                messageArr = ["Tasks Upload Sheet Prepared.",
                            "File path : " + parentFolder + '/taskUpload/taskUpload.csv']
                messageArr.append("URL : " + str(urlTasksUploadApi))
                messageArr.append("Upload status code : " + str(responseTasksUploadApi.status_code))
                messageArr.append("Response : " + str(responseTasksUploadApi.text))
                self.createAPILog(parentFolder, messageArr)
                print(responseTasksUploadApi.text,"responseTasksUploadApi")
                if responseTasksUploadApi.status_code == 200:
                    print('TaskUploadApi Success')
                    with open(parentFolder + '/taskUpload/taskInternal.csv','w+',encoding='utf-8') as tasksRes:
                        tasksRes.write(responseTasksUploadApi.text)
                    return True
                else:
                    if responseTasksUploadApi.status_code in [400, 401, 403, 404, 422]:
                        self.errorVar.append(f"TasksUploadApi-Client Error {responseTasksUploadApi.status_code}: {responseTasksUploadApi.text}")
                    elif responseTasksUploadApi.status_code in [500, 502, 503, 504]:
                        self.errorVar.append(f"TasksUploadApi-Server Error {responseTasksUploadApi.status_code}: {responseTasksUploadApi.text}")
                    else:
                        self.errorVar.append(f"TasksUploadApi-Unexpected Error {responseTasksUploadApi.status_code}: {responseTasksUploadApi.text}")

                    messageArr.append(f"Error Response: {responseTasksUploadApi.text}")
                    self.createAPILog(parentFolder, messageArr)
                    return False
        except Exception as e:
            self.errorVar.append(f"⚠️ Exception: {str(e)}")
            self.createAPILog(parentFolder, [f"Exception: {str(e)}"])
            return False

    def fetchSolutionDetailsFromProgramSheet(self, parentFolder, solutionId, accessToken, programdetails, ProgramGlobalDict, userRole):
        try:
            print("fetching solution details...")
            urlFetchSolutionApi = elevateprojecthost + dbfindapi_url
            print(urlFetchSolutionApi,"urlFetchSolutionApi")
            headerFetchSolutionApi = apiHeader.headers().headerFetchSolutionDetailFromProgramSheetApi(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
            print(headerFetchSolutionApi,"headerFetchSolutionApi")
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
            print(responseFetchSolutionApiUrl.text,"responseFetchSolutionApiUrl")
            responseFetchSolutionJson = responseFetchSolutionApiUrl.json()
            messageArr = ["Solution Fetch Link.",
                        "solution name : " + responseFetchSolutionJson["result"][-1]["name"],
                        "solution ExternalId : " + responseFetchSolutionJson["result"][-1]["externalId"],
                        "Upload status code : " + str(responseFetchSolutionApiUrl.status_code)]
            self.createAPILog(parentFolder, messageArr)
            if responseFetchSolutionApiUrl.status_code == 200:
                solutionName = responseFetchSolutionJson["result"][-1]["name"]
                print(solutionName,"solutionName")
                # xfile = openpyxl.load_workbook(programdetails)
                # sheet_name = 'Resource Details'.strip()
                # resourceDetailsSheet = xfile[sheet_name]
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
                messageArr = self.errorVar
                self.createAPILog(parentFolder,messageArr)
                return False
        except Exception as e:
            self.errorVar.append(f"Error occurred: {str(e)}")
            messageArr = self.errorVar
            self.createAPILog(parentFolder,messageArr)
            return False
    
    def validate_roles_against_api(self, mainRoles, subRoles, programdetails, parentFolder):
        urlFetchRoleList = userLoginHost + fetchprofessionalRole
        headersvalidateRolesApi = apiHeader.headers().header_validate_roles_against_api(programdetails.get('TenantID'))
        payload = {}

        response = requests.request("GET", urlFetchRoleList, headers=headersvalidateRolesApi, data=payload)
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
                subrole_resp = requests.request("GET",subrole_url, headers=headersvalidateRolesApi,data=payload)
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
        self.createAPILog(parentFolder, messageArr)
        for err in self.errorVar:
            print(err)
        if self.errorVar:
            return False

        return validated_main_role_ids, validated_subrole_ids_list   
    
    def solutionUpdate(self, solutionName_for_folder_path, accessToken, solutionId, bodySolutionUpdate,programdetails,userRole):
        try:
            solutionUpdateApi = elevateprojecthost + solutionupdateapi + str(solutionId)
            headerUpdateSolutionApi = apiHeader.headers().headerSolutionUpdateAPI(programdetails.get('TenantID'),programdetails.get('OrgForAPIs'),accessToken, userRole)
             
            responseUpdateSolutionApi = requests.post(url=solutionUpdateApi, headers=headerUpdateSolutionApi,data=json.dumps(bodySolutionUpdate))
            messageArr = ["Solution Update API called.", "URL : " + str(solutionUpdateApi), "Body : " + str(bodySolutionUpdate),"Response : " + str(responseUpdateSolutionApi.text),"Status Code : " + str(responseUpdateSolutionApi.status_code)]
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
                self.createAPILog(solutionName_for_folder_path, self.errorVar)
                return False
            
        except Exception as e:
            self.errorVar.append(f"Error occurred: {str(e)}")
            messageArr = self.errorVar
            self.createAPILog(solutionName_for_folder_path, messageArr)
            return False
        
    def fetchUserDetails(self, parentFolder, accessToken, projectServiceId):
        try:
            # decoded_token = jwt.decode(accessToken, options={"verify_signature": False}, algorithms=["HS256"])
            # data=decoded_token.get('data')
            # user_id = data.get('id')  
            url = userLoginHost + userinfoapiurl
            messageArr = ["User search API called."]
            headerfetchUserDetails = apiHeader.headers().headerfetchUserDetailsAPI(accessToken)
            responseUserSearch = requests.request("GET", url, headers=headerfetchUserDetails)
            messageArr.append(["Solution Update API called.", "URL : " + str(url),"Response : " + str(responseUserSearch.text),"Status Code : " + str(responseUserSearch.status_code)])
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
                    self.errorVar.append("-->Given username/email is not present in projectService platform<--.")
                    messageArr.append(self.errorVar)
                    self.createAPILog(parentFolder, messageArr)
                    return False
            else:
                if responseUserSearch.status_code in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"FetchSolutionApiUrl-Client Error {responseUserSearch.status_code}: {responseUserSearch.text}")
                elif responseUserSearch.status_code in [500, 502, 503, 504]:
                    self.errorVar.append(f"FetchSolutionApiUrl-Server Error {responseUserSearch.status_code}: {responseUserSearch.text}")
                else:
                    self.errorVar.append(f"FetchSolutionApiUrl-Unexpected Error {responseUserSearch.status_code}: {responseUserSearch.text}")
                    messageArr.append(self.errorVar)
                    self.createAPILog(parentFolder, messageArr)
                return False
        except Exception as e:
            self.errorVar.append(f"Error occurred: {str(e)}")
            messageArr.append(self.errorVar)
            self.createAPILog(parentFolder,messageArr)
            return False

    def solutionCreationAndMapping(self, parentFolder, wbObservation, listOfFoundRoles, accessToken, programFile, programdetails, ProgramGlobalDict, userRole):
        try:
            print("solutionCreationAndMapping....")
            SolutionFilePath = parentFolder + '/solutionDetails/'
            if not os.path.exists(SolutionFilePath):
                os.mkdir(SolutionFilePath)
            with open(parentFolder + '/solutionDetails/solutionDetails.csv', 'w',encoding='utf-8') as file:
                writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                writer.writerows(
                    [["solutionExtId", "solutionName", "solutionDescription", "solution_id", "programExternalId", "entityType",
                    "scopeEntityType", "entityNames", "roles", "duplicateTemplateExtId", "duplicateTemplate_id"]])

            projectInternalfile = open(parentFolder + '/projectUpload/projectInternal.csv', mode='r',encoding='utf-8')
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
                programExternalId = programdetails.get('ProgramID')
                subroles = programdetails.get('Targetedsubroleatprogramlevel', '')  # get string safely
                subroleslst = [item.strip() for item in subroles.split(',') if item.strip()]
                entityToUpload = programdetails.get("entity_type")
                scopeEntityType = [programdetails.get("entitiesType")]

                urlCreateProjectSolutionApi = elevateprojecthost + projectsolutioncreationapi
                print(urlCreateProjectSolutionApi,"urlCreateProjectSolutionApi")
                headerCreateSolutionApi = apiHeader.headers().headerCreateSolutionApi(programdetails.get("TenantID"), programdetails.get('OrgForAPIs'), accessToken, userRole)
                startdate = programdetails.get('Startdateofprogram')
                d, m, y = startdate.split('-')
                ProgramStartDate = f"{y}-{m}-{d} 00:00:00"
                enddate = programdetails.get('Enddateofprogram')
                d, m, y = enddate.split('-')
                ProgramEndDate = f"{y}-{m}-{d} 00:00:00"
                sol_payload = {
                    "createdFor": programdetails.get('OrgID'),
                    "rootOrganisations": programdetails.get('OrgID'),
                    "programExternalId": programExternalId,
                    "entityType": projectEntityType,
                    "externalId": solutionExternalId,
                    "name": project_name,
                    "scope": {
                        "roles": subroleslst,
                        },
                    "description": project_description,
                    "isReusable" : False,
                    "startDate": ProgramStartDate,
                    "endDate": ProgramEndDate,
                }
                print(sol_payload,"sol_payload")
                responseCreateSolutionApi = requests.post(url=urlCreateProjectSolutionApi,headers=headerCreateSolutionApi, data=json.dumps(sol_payload))
                print(responseCreateSolutionApi.text,"responseCreateSolutionApi")
                messageArr = ["Project Solution Created.","URL : " + str(urlCreateProjectSolutionApi),"Status Code : " + str(responseCreateSolutionApi.status_code),"Response : " + str(responseCreateSolutionApi.text)]
                if responseCreateSolutionApi.status_code == 200:
                    responseCreateSolutionApi = responseCreateSolutionApi.json()
                    solutionId = responseCreateSolutionApi['result']['_id']
                    self.solutionUpdate(parentFolder, accessToken, solutionId, {"status": "inactive", "isDeleted": True}, programdetails, userRole)
                    print(solutionId,"solutionId")
                    messageArr.append("Solution Generated : " + str(solutionId))
                    self.createAPILog(parentFolder, messageArr)
                    print("ProjectSolutionCreationApi Success")
                    duplicateTemplateExtId = projectExternalId + '_IMPORTED'
                    queryparamsMapProjectSolutionApi = projectExternalId + '?solutionId='+solutionId
                    urlMapProjectSolutionApi = elevateprojecthost + mapsolutiontoproject + queryparamsMapProjectSolutionApi
                    headerMapSolutionProject = apiHeader.headers().headerMapSolutionProjectAPI(programdetails.get("TenantID"), programdetails.get('OrgForAPIs'), accessToken, userRole)
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
                    if responseMapProjectSolutionApi.status_code == 200:
                        responseMapProjectSolutionApi = responseMapProjectSolutionApi.json()
                        duplicateTemplateId = responseMapProjectSolutionApi['result']['_id']
                        messageArr.append("duplicate TemplateId successfully created: " + str(duplicateTemplateId))
                        self.createAPILog(parentFolder, messageArr)
                        print("MapSolutionToProjectApi Sucsess")
                        with open(parentFolder + '/solutionDetails/solutionDetails.csv', 'a',encoding='utf-8') as file:
                            writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
                            writer.writerows([[solutionExternalId, project_name, project_description, solutionId,
                                            programExternalId, projectEntityType,
                                            scopeEntityType, entityToUpload, listOfFoundRoles, duplicateTemplateExtId,
                                            duplicateTemplateId]])
                        solutionDetails = self.fetchSolutionDetailsFromProgramSheet(parentFolder,
                                                                            solutionId, accessToken, programdetails, ProgramGlobalDict, userRole)
                        print("solutionDetails---",solutionDetails)
                        if solutionDetails:
                            scopeRoles = solutionDetails[0]
                            scopeSubRoles = solutionDetails[1]
                            if isinstance(scopeRoles, str):
                                scopeRoles = [r.strip() for r in scopeRoles.split(',') if r.strip()]
                            if isinstance(scopeSubRoles, str):
                                scopeSubRoles = [r.strip() for r in scopeSubRoles.split(',') if r.strip()]
                            verifiedRoles = self.validate_roles_against_api(scopeRoles, scopeSubRoles, programdetails, parentFolder)
                            if not verifiedRoles:
                                return False
                            mainRoleproff = verifiedRoles[0]
                            rolesPGMID = verifiedRoles[1]
                            print("mainRole", mainRoleproff)
                            print("rolesPGMID", rolesPGMID)
                            scope = {}
                            # s = programdetails.get('OrgID', '')
                            scope["organizations"] = programdetails.get('OrgID', '')      
                            scope["professional_subroles"] = rolesPGMID
                            scope["professional_role"] = mainRoleproff
                            scope.update(programdetails.get('entityHierarchy'))
                            print(scope)
                            bodySolutionUpdate = {
                            "scope": scope
                            }
                            if not self.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate, programdetails, userRole):
                                return False
                            projectuploadsheet = (wbObservation.get("Project upload"))
                            projectAuthor = projectuploadsheet.get('Username/user id/email id/phone no. of content creator')
                            print("projectAuthor",projectAuthor)
                            userDetails = self.fetchUserDetails(parentFolder, accessToken, projectAuthor)
                            if not userDetails:
                                return False
                            print("userDetails",userDetails)
                            matchedShikshalokamLoginId = userDetails[0]
                            projectCreator = userDetails[1]
                            bodySolutionUpdate = {
                                "creator": projectCreator, "author": matchedShikshalokamLoginId}
                            self.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate, programdetails, userRole)
                            if solutionDetails[2]:
                                startDateArr = str(solutionDetails[2]).split("-")
                                bodySolutionUpdate = {
                                    "startDate": startDateArr[2] + "-" + startDateArr[1] + "-" + startDateArr[0] + " 00:00:00"}
                                self.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate, programdetails, userRole)
                            if solutionDetails[3]:
                                endDateArr = str(solutionDetails[3]).split("-")
                                bodySolutionUpdate = {
                                    "endDate": endDateArr[2] + "-" + endDateArr[1] + "-" + endDateArr[0] + " 23:59:59"}
                                self.solutionUpdate(parentFolder, accessToken, solutionId, bodySolutionUpdate, programdetails, userRole)
                    
                                return [solutionExternalId, solutionId]
                            else:
                                self.errorVar.append("Date mismatching!")
                                return False
                        else:
                            if responseMapProjectSolutionApi.status_code in [400, 401, 403, 404, 422]:
                                self.errorVar.append(f"MapProjectSolutionApi-Client Error {responseMapProjectSolutionApi.status_code}: {responseMapProjectSolutionApi.text}")
                            elif responseMapProjectSolutionApi.status_code in [500, 502, 503, 504]:
                                self.errorVar.append(f"MapProjectSolutionApi-Server Error {responseMapProjectSolutionApi.status_code}: {responseMapProjectSolutionApi.text}")
                            else:
                                self.errorVar.append(f"MapProjectSolutionApi-Unexpected Error {responseMapProjectSolutionApi.status_code}: {responseMapProjectSolutionApi.text}")
                            print("Map project to solution api failed.")
                            return False
                    else:
                        self.errorVar.append(str(responseCreateSolutionApi.text))
                        print("Project solution creation api failed.")
                        if responseCreateSolutionApi.status_code in [400, 401, 403, 404, 422]:
                            self.errorVar.append(f"CreateSolutionApi-Client Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}")
                        elif responseCreateSolutionApi.status_code in [500, 502, 503, 504]:
                            self.errorVar.append(f"CreateSolutionApi-Server Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}")
                        else:
                            self.errorVar.append(f"CreateSolutionApi-Unexpected Error {responseCreateSolutionApi.status_code}: {responseCreateSolutionApi.text}")
                        return False
        except Exception as e:
            self.errorVar.append(f"Error occurred: {str(e)}")
            messageArr = self.errorVar
            self.createAPILog(parentFolder, messageArr)
            return False
    
    def prepareProgramSuccessSheet(self, MainFilePath, solutionName_for_folder_path, programFile, solutionExternalId, solutionId,accessToken, programdetails, userRole):
        urlFetchSolutionApi = elevateprojecthost + dbfindapi_url
        headerFetchSolutionApi = apiHeader.headers().headerUrlFetchSolutionApi(programdetails.get('TenantID'),programdetails.get('OrgForAPIs'), accessToken, userRole)
        payloadFetchSolutionApi = json.dumps({
                    "query": {
                        "_id": solutionId
                    },
                    "mongoIdKeys": [
                        "_id","name", "externalId"
                    ],
                    "limit": 10000
                })

        responseFetchSolutionApi = requests.post(url=urlFetchSolutionApi, headers=headerFetchSolutionApi,
                                                data=payloadFetchSolutionApi)
        responseFetchSolutionJson = responseFetchSolutionApi.json()
        messageArr = ["Solution Fetch Link.",
                    "solution name : " + responseFetchSolutionJson["result"][-1]["name"],
                    "solution ExternalId : " + responseFetchSolutionJson["result"][-1]["externalId"]]
        messageArr.append("Upload status code : " + str(responseFetchSolutionApi.status_code))
        self.createAPILog(solutionName_for_folder_path, messageArr)

        if responseFetchSolutionApi.status_code == 200:
            print('Fetch solution Api Success')
            solutionName = responseFetchSolutionJson["result"][-1]["name"]
        urlFetchSolutionLinkApi = elevateprojecthost + fetchlink + solutionId
        print(urlFetchSolutionLinkApi,"urlFetchSolutionLinkApi")
        headerFetchSolutionLinkApi = apiHeader.headers().headerURLFetchSolutionLinkApi(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
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
            messageArr.append("Response : " + str(responseFetchSolutionLinkApi.text))
            self.createAPILog(solutionName_for_folder_path, messageArr)
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
    
    def fetchCertificateBaseTemplate(self, wbObservation, accessToken, parentFolder, programdetails, userRole):
        try:
            # --- Get certificate details dict from wbObservation ---
            certificateDetails = wbObservation.get('Certificate details', {})
            if not certificateDetails:
                self.errorVar.append("No 'Certificate details' sheet found in wbObservation.")
                return False
            
            # --- Extract type of certificate ---
            typeOfCertificate = certificateDetails.get("Type of certificate", "")
            if not typeOfCertificate:
                self.errorVar.append("Type of certificate is missing in Certificate details.")
                return False

            # --- Normalize value ---
            typeOfCertificate = typeOfCertificate.lower().replace(" ", "")

            # --- Call DBFind API to get base templates ---
            urldbFind = elevateprojecthost + dbfindapi
            headerdbFindApi = apiHeader.headers().headerSolutionUpdateAPI(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)

            payload = json.dumps({
                "query": {'tenantId': programdetails.get('TenantID')},
                "mongoIdKeys": []
            })

            print(urldbFind, "2163")
            responsedbFindApi = requests.request("POST", url=urldbFind, headers=headerdbFindApi, data=payload)
            print(responsedbFindApi.text, "responsedbFindApi 2166")
            messageArr = ["DBFind API called.", "URL : " + str(urldbFind), "Body : " + str(payload), "Status Code : " + str(responsedbFindApi.status_code)]
            self.createAPILog(parentFolder, messageArr)
            if responsedbFindApi.status_code == 200:
                responseaddcetificate = responsedbFindApi.json()
                result_list = responseaddcetificate.get('result', [])
                baseTemplateLookup = {i['code']: i['_id'] for i in result_list if 'code' in i and '_id' in i}

                baseTemplateCode = certificatetypeof.get(typeOfCertificate)
                if not baseTemplateCode:
                    self.errorVar.append(f"Certificate type '{typeOfCertificate}' not found in certificatetypeof mapping.")
                    return False

                baseTemplateId = baseTemplateLookup.get(baseTemplateCode)
                if not baseTemplateId:
                    self.errorVar.append(f"Base template code '{baseTemplateCode}' not found in DBFind API result.")
                    return False
                messageArr = [f"Base template ID '{baseTemplateId}' found for certificate type '{typeOfCertificate}'."]
                self.createAPILog(parentFolder, messageArr)
                return baseTemplateId

            else:
                if responsedbFindApi.status_code in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"dbFindApi-Client Error {responsedbFindApi.status_code}: {responsedbFindApi.text}")
                elif responsedbFindApi.status_code in [500, 502, 503, 504]:
                    self.errorVar.append(f"dbFindApi-Server Error {responsedbFindApi.status_code}: {responsedbFindApi.text}")
                else:
                    self.errorVar.append(f"dbFindApi-Unexpected Error {responsedbFindApi.status_code}: {responsedbFindApi.text}")

                messageArr = f"Error Response: {responsedbFindApi.text}"
                self.createAPILog(parentFolder, messageArr)
                print("---> Error in fetching DBFind data; please give proper code value <---")
                return False

        except Exception as e:
            self.errorVar.append(f"Error occurred: {str(e)}")
            return False

    def downloadlogosign(self, wbObservation, projectName_for_folder_path):
        try:
            certificateDetails = wbObservation.get('Certificate details', {})
            if not certificateDetails:
                self.errorVar.append("---> No 'Certificate details' found in wbObservation. <---")
                return False

            print("---> Checking Certificate details sheet (from wbObservation)...")

            certificateissuer = certificateDetails.get('Certificate issuer', '').strip()
            if not certificateissuer:
                self.errorVar.append("\"Certificate issuer\" must not be empty in 'Certificate details'")

            typeOfCertificate = certificateDetails.get('Type of certificate', '').strip()
            if not typeOfCertificate:
                self.errorVar.append("\"Type of certificate\" must not be empty in 'Certificate details'")

            Logofilepath = os.path.join(projectName_for_folder_path, 'Logofile')
            os.makedirs(Logofilepath, exist_ok=True)

            # Helper function to extract Google Drive ID and download
            def download_from_drive(drive_url, dest_name):
                if not drive_url:
                    self.errorVar.append(f"⚠️ Missing URL for {dest_name}")
                    return False
                try:
                    parts = str(drive_url).split('/')
                    if len(parts) < 6:
                        self.errorVar.append(f"⚠️ Invalid Google Drive link for {dest_name}: {drive_url}")
                        return False
                    file_id = parts[5]
                    file_url = f'https://drive.google.com/uc?export=download&id={file_id}'
                    dest_file = os.path.join(Logofilepath, dest_name)
                    print(f"⬇️ Downloading {dest_name}...")
                    return gdown.download(file_url, dest_file, quiet=False)
                except Exception as e:
                    self.errorVar.append(f"Error downloading {dest_name}: {str(e)}")
                    return False

            # --- Handle certificate types ---
            if typeOfCertificate == 'One Logo - One Signature':
                download_from_drive(certificateDetails.get('Logo - 1'), 'logo1.jpg')
                download_from_drive(certificateDetails.get('Authorised Signature Image - 1'), 'signature1.jpg')

            elif typeOfCertificate == 'One Logo - Two Signature':
                download_from_drive(certificateDetails.get('Logo - 1'), 'logo1.jpg')
                download_from_drive(certificateDetails.get('Authorised Signature Image - 1'), 'signature1.jpg')
                download_from_drive(certificateDetails.get('Authorised Signature Image - 2'), 'signature2.jpg')

            elif typeOfCertificate == 'Two Logo - One Signature':
                download_from_drive(certificateDetails.get('Logo - 1'), 'logo1.jpg')
                download_from_drive(certificateDetails.get('Logo - 2'), 'logo2.jpg')
                download_from_drive(certificateDetails.get('Authorised Signature Image - 1'), 'signature1.jpg')

            elif typeOfCertificate == 'Two Logo - Two Signature':
                download_from_drive(certificateDetails.get('Logo - 1'), 'logo1.jpg')
                download_from_drive(certificateDetails.get('Logo - 2'), 'logo2.jpg')
                download_from_drive(certificateDetails.get('Authorised Signature Image - 1'), 'signature1.jpg')
                download_from_drive(certificateDetails.get('Authorised Signature Image - 2'), 'signature2.jpg')

            else:
                self.errorVar.append("---> Logos and signature downloading failed (check 'Type of certificate' or link access). <---")
                return False

            print("---> Logo(s) and signature(s) downloaded successfully. <---")
            return True

        except Exception as e:
            self.errorVar.append(f"Error occurred while downloading logos/signatures: {str(e)}")
            return False
    
    def editsvg(self, accessToken, wbObservation, parentFolder, baseTemplate_id, programdetails, userRole):
        try:
            certificateDetails = wbObservation.get('Certificate details', {})
            # --- Read values from dictionary safely ---
            Typeofcertificate = certificateDetails.get('Type of certificate', '').strip()
            Certificateissuer = str(certificateDetails.get('Certificate issuer', '')).encode('utf-8').decode('utf-8')
            Logo1 = certificateDetails.get('Logo - 1', '')
            Logo2 = certificateDetails.get('Logo - 2', '')
            authsignaturelogo1 = certificateDetails.get('Authorised Signature Image - 1', '')
            authsignaturelogo2 = certificateDetails.get('Authorised Signature Image - 2', '')
            authsignatory1 = str(certificateDetails.get('Authorised Signatory - 1', '')).encode('utf-8').decode('utf-8')
            authsignatory2 = str(certificateDetails.get('Authorised Signatory - 2', '')).encode('utf-8').decode('utf-8')

            print(f"---> Certificate type: {Typeofcertificate}")

            # --- Prepare request payload ---
            payload = {}
            downloadedfiles = []
            baseTemplateId = baseTemplate_id

            # --- Build paths ---
            # print(parentFolder)
            # projectName_for_folder_path = os.path.join(parentFolder, 'CertificateAssets')

            # --- Based on type ---
            if Typeofcertificate == 'One Logo - One Signature':
                print("-->This is One Logo - One Signature<--")
                stateLogo1 = ('stateLogo1', ('logo1.jpg', open(os.path.join(parentFolder, 'Logofile/logo1.jpg'), 'rb'), 'image/jpeg'))
                downloadedfiles.append(stateLogo1)

                payload['stateTitle'] = Certificateissuer
                signatureImg1 = ('signatureImg1', ('signatureImg1.jpg', open(os.path.join(parentFolder, 'Logofile/signature1.jpg'), 'rb'), 'image/jpeg'))
                downloadedfiles.append(signatureImg1)
                payload['signatureTitle1a'] = authsignatory1

            elif Typeofcertificate == 'One Logo - Two Signature':
                print("-->This is One Logo - Two Signature<--")
                stateLogo1 = ('stateLogo1', ('logo1.jpg', open(os.path.join(parentFolder, 'Logofile/logo1.jpg'), 'rb'), 'image/jpeg'))
                downloadedfiles.append(stateLogo1)

                payload['stateTitle'] = Certificateissuer
                signatureImg1 = ('signatureImg1', ('signature1.jpg', open(os.path.join(parentFolder, 'Logofile/signature1.jpg'), 'rb'), 'image/jpeg'))
                downloadedfiles.append(signatureImg1)
                signatureImg2 = ('signatureImg2', ('signature2.jpg', open(os.path.join(parentFolder, 'Logofile/signature2.jpg'), 'rb'), 'image/jpeg'))
                downloadedfiles.append(signatureImg2)

                payload['signatureTitle1a'] = authsignatory1
                payload['signatureTitle2a'] = authsignatory2

            elif Typeofcertificate == 'Two Logo - One Signature':
                print("-->This is Two Logo - One Signature<--")
                stateLogo1 = ('stateLogo1', ('logo1.jpg', open(os.path.join(parentFolder, 'Logofile/logo1.jpg'), 'rb'), 'image/jpeg'))
                stateLogo2 = ('stateLogo2', ('logo2.jpg', open(os.path.join(parentFolder, 'Logofile/logo2.jpg'), 'rb'), 'image/jpeg'))
                downloadedfiles.extend([stateLogo1, stateLogo2])

                payload['stateTitle'] = Certificateissuer
                signatureImg1 = ('signatureImg1', ('signature1.jpg', open(os.path.join(parentFolder, 'Logofile/signature1.jpg'), 'rb'), 'image/jpeg'))
                downloadedfiles.append(signatureImg1)
                payload['signatureTitle1a'] = authsignatory1

            elif Typeofcertificate == 'Two Logo - Two Signature':
                print("-->This is Two Logo - Two Signature<--")
                stateLogo1 = ('stateLogo1', ('logo1.jpg', open(os.path.join(parentFolder, 'Logofile/logo1.jpg'), 'rb'), 'image/jpeg'))
                stateLogo2 = ('stateLogo2', ('logo2.jpg', open(os.path.join(parentFolder, 'Logofile/logo2.jpg'), 'rb'), 'image/jpeg'))
                signatureImg1 = ('signatureImg1', ('signature1.jpg', open(os.path.join(parentFolder, 'Logofile/signature1.jpg'), 'rb'), 'image/jpeg'))
                signatureImg2 = ('signatureImg2', ('signature2.jpg', open(os.path.join(parentFolder, 'Logofile/signature2.jpg'), 'rb'), 'image/jpeg'))
                downloadedfiles.extend([stateLogo1, stateLogo2, signatureImg1, signatureImg2])

                payload['stateTitle'] = Certificateissuer
                payload['signatureTitle1a'] = authsignatory1
                payload['signatureTitle2a'] = authsignatory2

            else:
                self.errorVar.append(f"Unknown certificate type: {Typeofcertificate}")
                return False

            # --- API call setup ---
            urleditnigsvgApi = f"{elevateprojecthost}{editsvgtemp}{baseTemplateId}"
            headereditingsvgApi = apiHeader.headers().headereditingsvgApi(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
            print(urleditnigsvgApi, "urleditnigsvgApi")
            print(headereditingsvgApi, "headereditingsvgApi")

            responseeditsvg = requests.post(url=urleditnigsvgApi, headers=headereditingsvgApi, data=payload, files=downloadedfiles)
            print(responseeditsvg.text, "responseeditsvg")
            messageArr = ["editsvg API called.", "URL : " + str(urleditnigsvgApi), "Payload : " + str(payload), "Response : " + str(responseeditsvg.text), "Status Code : " + str(responseeditsvg.status_code)]
            self.createAPILog(parentFolder, messageArr)
            # --- Handle response ---
            if responseeditsvg.status_code == 200:
                result = responseeditsvg.json()
                svg_url = result['result']['url']

                dest_folder = os.path.join(parentFolder, 'DownloadedSVG')
                os.makedirs(dest_folder, exist_ok=True)
                dest_file = os.path.join(dest_folder, 'Downloaded.svg')

                gdown.download(svg_url, dest_file, quiet=False)
                return True
            else:
                self.errorVar.append(f"editsvg Error {responseeditsvg.status_code}: {responseeditsvg.text}")
                messageArr = f"Error Response: {responseeditsvg.text}"
                self.createAPILog(parentFolder, messageArr)
                return False

        except Exception as e:
            self.errorVar.append(f"Error occurred: {str(e)}")
            messageArr = f"Error occurred: {str(e)}"
            self.createAPILog(parentFolder, messageArr)
            return False

    def prepareaddingcertificatetemp(self, wbObservation, parentFolder, accessToken, solutionId,baseTemplate_id, programdetails, userRole):
        try:
            tasksLevelEvidance = []
            projectMinNooEvide = None
            projectLevelEvidance = []
            taskMinNooEvide =[]
            programID = programdetails.get('_id')
            projectdetails = wbObservation.get('Project upload')
            projectLevelMinNooEvidence = projectdetails.get("Minimum No. of Evidence")
            projectLevelEvidance = projectdetails.get("Project Level Evidence").lower()
            if projectLevelMinNooEvidence == "":
                projectLevelMinNooEvidence = 1  # Set default value to 1
                projectMinNooEvide = int(projectLevelMinNooEvidence)
            else:
                projectMinNooEvide = int(projectLevelMinNooEvidence)

            # projectTaskdetails = wbObservation.get('Tasks upload')               
            for projectTaskdetails in wbObservation.get('Tasks upload'):         
                taskLevelEvidence = projectTaskdetails.get("Task Level Evidence req. for certificate criteria").lower()
                minNoOfEvidence = projectTaskdetails.get("Minimum No. of Evidence for task level evidence criteria")
                TaskEvidenceOperator =  projectTaskdetails.get("Evidence required for any task for certificate criteria")
                AnyTaskEvidenceNo = projectTaskdetails.get("Minimum No. of Evidence for any task criteria")
                if AnyTaskEvidenceNo == "":
                    AnyTaskEvidenceNo == 1
                if TaskEvidenceOperator.lower() == "yes":
                    tasksLevelEvidance.append(projectTaskdetails.get("TaskTitle")) 
                else:
                    if taskLevelEvidence == "yes":
                        tasksLevelEvidance.append(projectTaskdetails.get("TaskTitle"))
                        if minNoOfEvidence == "":
                            minNoOfEvidence = 1  # Set default value to 1
                            taskMinNooEvide.append(minNoOfEvidence)
                        else:
                            taskMinNooEvide.append(minNoOfEvidence)

            addcetificateFilePath = parentFolder + '/addCertificate/'
            if not os.path.exists(addcetificateFilePath):
                os.mkdir(addcetificateFilePath)

            urladdcertificate = elevateprojecthost + addcertificatetemplate
            headeraddcertificateApi = apiHeader.headers().headeraddcertificateApi(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)

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
            
            projectCertificatedetails = wbObservation.get('Certificate details')
            if projectCertificatedetails.get('Certificate issuer'):
                certificateissuer = projectCertificatedetails.get('Certificate issuer')
            else:
                self.errorVar.append("\"Certificate issuer\" must not be Empty in \"Certificate details\" sheet")
            payload["issuer"]["name"] = certificateissuer

            # Typeofcertificate = dictDetailsEnv['Type of certificate'] if dictDetailsEnv['Type of certificate'] in ["One Logo - One Signature", "One Logo - Two Signature", "Two Logo - One Signature","Two Logo - Two Signature"] else Elevateproject.terminatingMessage("\"Type of certificate\" must not be Empty in \"Certificate details\" sheet")
            if projectCertificatedetails.get('Type of certificate') in ["One Logo - One Signature", "One Logo - Two Signature", "Two Logo - One Signature", "Two Logo - Two Signature"]:
                Typeofcertificate = projectCertificatedetails.get('Type of certificate')
            else:
                self.errorVar.append("\"Type of certificate\" must not be Empty or invalid in \"Certificate details\" sheet")
            payload["baseTemplateId"]=baseTemplate_id
                    


            projectInternalfile = open(parentFolder + '/projectUpload/projectInternal.csv', mode='r',encoding='utf-8')
            projectInternalfile = csv.DictReader(projectInternalfile)
            for projectInternal in projectInternalfile:
                projectExternalId = projectInternal["externalId"]
                project_id = projectInternal["_SYSTEM_ID"]

            taskinternalfile = open(parentFolder + '/taskUpload/taskInternal.csv', mode='r',encoding='utf-8')
            taskinternalfile = csv.DictReader(taskinternalfile)
            projectTemplatefile = open(parentFolder + '/solutionDetails/solutionDetails.csv', mode='r',encoding='utf-8')
            projectTemplatefile = csv.DictReader(projectTemplatefile)
            for Projecttemp in projectTemplatefile:
                projectTemplateId = Projecttemp["duplicateTemplate_id"]
            c = 2
            for task in taskinternalfile:
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
                            print(payload)
                            
                        else:
                            c = c + 1
                            cn = "C" + str(c)
                            taskconditions = {
                                cn: {
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
            responseaddcertificateUploadApi = requests.post(url=urladdcertificate, headers=headeraddcertificateApi, data=json.dumps(payload))
            messageArr = ["Add certificate json is prepared",
                        "File path : " + parentFolder + '/addCertificate/Addcertificate.text']
            messageArr.append("URL : " + str(responseaddcertificateUploadApi))
            messageArr.append("Upload status code : " + str(responseaddcertificateUploadApi.status_code))
            self.createAPILog(parentFolder, messageArr)
            with open(parentFolder + '/addCertificate/Addcertificatejson.json',
                    'w+',encoding='utf-8') as tasksRes:
                tasksRes.write(json.dumps(payload))

            if responseaddcertificateUploadApi.status_code == 200:
                responseaddcetificate = responseaddcertificateUploadApi.json()
                certificatetemplateid = responseaddcetificate['result']['_id']
                print("-->Certificate template id generated <--", certificatetemplateid)
                with open(parentFolder + '/addCertificate/Addcertificate.text',
                        'w+',encoding='utf-8') as tasksRes:
                    tasksRes.write(responseaddcertificateUploadApi.text)

            else:
                if responseaddcertificateUploadApi.status_code in [400, 401, 403, 404, 422]:
                    self.errorVar(f"addcertificateUploadApi-Client Error {responseaddcertificateUploadApi.status_code}: {responseaddcertificateUploadApi.text}")
                elif responseaddcertificateUploadApi.status_code in [500, 502, 503, 504]:
                    self.errorVar(f"addcertificateUploadApi-Server Error {responseaddcertificateUploadApi.status_code}: {responseaddcertificateUploadApi.text}")
                else:
                    self.errorVar(f"addcertificateUploadApi-Unexpected Error {responseaddcertificateUploadApi.status_code}: {responseaddcertificateUploadApi.text}")
                print(self.errorVar)
                print("Add certificate mission failed please check logs")
                messageArr.append("Error Response : " + str(responseaddcertificateUploadApi.text))
                self.createAPILog(parentFolder, messageArr)
                return False

            urluploadcertificatepi = elevateprojecthost + uploadcertificatetosvg + certificatetemplateid
            headeruploadcertificateApi = apiHeader.headers().headeruploadcertificateApi(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)
            task_payload = {}
            task_file = []
            certificateaddtotemplate = ('file', ( 'Dowloaded.svg',open(parentFolder + '/DownloadedSVG/Downloaded.svg', 'rb'), 'image/svg+xml'))
            task_file.append(certificateaddtotemplate)


            responseDownloadsvgApi = requests.request("POST",url=urluploadcertificatepi, headers=headeruploadcertificateApi,
                                                data=task_payload, files=task_file)
            messageArr = ["Upload certificate to svg API called.",
                        "URL : " + str(urluploadcertificatepi),
                        "Response : " + str(responseDownloadsvgApi.text),
                        "Status Code : " + str(responseDownloadsvgApi.status_code)]
            self.createAPILog(parentFolder, messageArr)
            if responseDownloadsvgApi.status_code == 200:
                responseeditsvg = responseDownloadsvgApi.json()
                svgid = responseeditsvg['result']['data']['templateId']

                urlsolutionupdateapi = elevateprojecthost + solutionupdateapi + solutionId
                headersolutionupdateApi = apiHeader.headers().headerSolutionUpdateAPI(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)

                certificate_payload = json.dumps({
                    'certificateTemplateId':certificatetemplateid
                })
                responseupdatecertificateApi = requests.request("POST", url=urlsolutionupdateapi,
                                                        headers=headersolutionupdateApi,
                                                        data=certificate_payload)

                messageArr = ["Update solution API called.",
                            "URL : " + str(urlsolutionupdateapi),
                            "Response : " + str(responseupdatecertificateApi.text),
                            "Status Code : " + str(responseupdatecertificateApi.status_code)]
                if responseupdatecertificateApi.status_code == 200:
                    print("--->certificate added to the solution<---")
                    messageArr.append("--->certificate added to the solution<---")
                    self.createAPILog(parentFolder, messageArr)

                else:
                    if responseupdatecertificateApi.status_code in [400, 401, 403, 404, 422]:
                        self.errorVar.append(f"updatecertificateApi-Client Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}")
                    elif responseupdatecertificateApi.status_code in [500, 502, 503, 504]:
                        self.errorVar.append(f"updatecertificateApi-Server Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}")
                    else:
                        self.errorVar.append(f"updatecertificateApi-Unexpected Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}")
                    print(self.errorVar)
                    print("error in updating solution")
                    messageArr.append(f"Error Response: {self.errorVar[-1]}")
                    self.createAPILog(parentFolder, messageArr)
                    return False

                urlprojecttemplateapi = elevateprojecthost + updateprojecttemplate + projectTemplateId
                headerprojectrtemplateupdateApi = apiHeader.headers().headerProjectTemplateUpdateAPI(programdetails.get('TenantID'), programdetails.get('OrgForAPIs'), accessToken, userRole)

                certificate_payload = json.dumps({
                    'certificateTemplateId': certificatetemplateid
                })
                responseupdatecertificateApi = requests.request("POST", url=urlprojecttemplateapi,
                                                                headers=headerprojectrtemplateupdateApi,
                                                                data=certificate_payload)
                messageArr = ["Update project template API called.",
                            "URL : " + str(urlprojecttemplateapi),
                            "Response : " + str(responseupdatecertificateApi.text),
                            "Status Code : " + str(responseupdatecertificateApi.status_code)]
                if responseupdatecertificateApi.status_code == 200:
                    print("--->Certificate added to project<---")
                    messageArr.append("--->Certificate added to project<---")
                    self.createAPILog(parentFolder, messageArr)
                    return True
                else:
                    if responseupdatecertificateApi.status_code in [400, 401, 403, 404, 422]:
                        self.errorVar.append(f"TasksUploadApi-Client Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}")
                    elif responseupdatecertificateApi.status_code in [500, 502, 503, 504]:
                        self.errorVar.append(f"TasksUploadApi-Server Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}")
                    else:
                        self.errorVar.append(f"TasksUploadApi-Unexpected Error {responseupdatecertificateApi.status_code}: {responseupdatecertificateApi.text}")
                    print("error in updating certificate with project")
                    messageArr.append(f"Error Response: {self.errorVar[-1]}")
                    self.createAPILog(parentFolder, messageArr)
                    return False
            else:
                if responseDownloadsvgApi.status_code in [400, 401, 403, 404, 422]:
                    self.errorVar.append(f"DownloadsvgApi-Client Error {responseDownloadsvgApi.status_code}: {responseDownloadsvgApi.text}")
                elif responseDownloadsvgApi.status_code in [500, 502, 503, 504]:
                    self.errorVar.append(f"DownloadsvgApi-Server Error {responseDownloadsvgApi.status_code}: {responseDownloadsvgApi.text}")
                else:
                    self.errorVar.append(f"DownloadsvgApi-Unexpected Error {responseDownloadsvgApi.status_code}: {responseDownloadsvgApi.text}")
                messageArr.append(f"Error Response: {self.errorVar[-1]}")
                self.createAPILog(parentFolder, messageArr)
                print("error in updating certificate with project")
                return False
        except Exception as e:
            self.errorVar.append(f"Error occurred: {str(e)}")
            messageArr = f"Error occurred: {str(e)}"
            self.createAPILog(parentFolder, messageArr)
            return False

    def CreateProject(self, resource, parentFolder, accessToken, ProgramGlobalDict, programdetails, MainFilePath, programFile, userRole):
        if resource.get('typeofSolution') == 4:
            print("Creating Project Solution...")
            print(resource)
            projectSolutionlink = ""
            wbObservation = resource.get("ResourceCre")
            print(wbObservation)
            projectUpload = wbObservation.get('Project upload')
            if (projectUpload.get('has certificate')).lower() == 'no':
                if not self.prepareProjectAndTasksSheets(wbObservation, parentFolder, accessToken, programdetails, userRole):
                    self.errorVar.append('Unable to prepare Project And Tasks Upload Sheets')
                    return projectSolutionlink, self.errorVar
                print("Prepared project and task upload sheet success...")
                if not self.projectUpload(parentFolder, accessToken, programdetails, userRole):
                    self.errorVar.append('Project Upload Failed...')
                    return projectSolutionlink, self.errorVar
                print("projectUpload")
                if not self.taskUpload(parentFolder, accessToken, programdetails, userRole):
                    self.errorVar.append('Task Upload Failed...')
                    return projectSolutionlink, self.errorVar
                print("Task Upload Success...")
                
                listOfFoundRoles = []
                ProjectSolutionResp = self.solutionCreationAndMapping(parentFolder, wbObservation, listOfFoundRoles, accessToken,programFile, programdetails, ProgramGlobalDict, userRole)
                if not ProjectSolutionResp:
                    self.errorVar.append('Solution creation and mapping Failed...')
                    return projectSolutionlink, self.errorVar
                ProjectSolutionExternalId = ProjectSolutionResp[0]
                ProjectSolutionId = ProjectSolutionResp[1]
                projectSolutionlink = self.prepareProgramSuccessSheet(MainFilePath, parentFolder, programFile, ProjectSolutionExternalId, ProjectSolutionId, accessToken, programdetails, userRole)
                if not projectSolutionlink:
                    self.errorVar.append('Fetching DeepLink Failed...')
                    return projectSolutionlink, self.errorVar
                else:
                    print("solution created successfully!")
                    return projectSolutionlink, self.errorVar
            else:
                print("---->this is certificate with project<---")
                baseTemplate_id=self.fetchCertificateBaseTemplate(wbObservation,accessToken,parentFolder, programdetails, userRole)
                if not baseTemplate_id:
                    self.errorVar.append('Unable to fetch Certificate Base Template...')
                    return projectSolutionlink, self.errorVar
                print("Base template id found -", baseTemplate_id)
                if not self.downloadlogosign(wbObservation,parentFolder):
                    self.errorVar.append('Unable to download logo and sign...')
                    return projectSolutionlink, self.errorVar
                print("Successfully downloaded logos and signs...")
                if not self.editsvg(accessToken,wbObservation,parentFolder,baseTemplate_id, programdetails, userRole):
                    self.errorVar.append('Unable to edit svg...')
                    return projectSolutionlink, self.errorVar
                print("Successfully Edited SVGs...")
                if not self.prepareProjectAndTasksSheets(wbObservation, parentFolder, accessToken, programdetails, userRole):
                    self.errorVar.append('Unable to prepare Project And Tasks Upload Sheets')
                    return projectSolutionlink, self.errorVar
                print("Prepared project and task upload sheet success...")
                if not self.projectUpload(parentFolder, accessToken, programdetails, userRole):
                    self.errorVar.append('Project Upload Failed...')
                    return projectSolutionlink, self.errorVar
                print("projectUpload")
                if not self.taskUpload(parentFolder, accessToken, programdetails, userRole):
                    self.errorVar.append('Task Upload Failed...')
                    return projectSolutionlink, self.errorVar
                print("Task Upload Success...")
                
                listOfFoundRoles = []
                ProjectSolutionResp = self.solutionCreationAndMapping(parentFolder, wbObservation, listOfFoundRoles, accessToken,programFile, programdetails, ProgramGlobalDict, userRole)
                if not ProjectSolutionResp:
                    self.errorVar.append('Solution creation and mapping Failed...')
                    return projectSolutionlink, self.errorVar
                ProjectSolutionExternalId = ProjectSolutionResp[0]
                ProjectSolutionId = ProjectSolutionResp[1]
                certificatetemplateid= self.prepareaddingcertificatetemp(wbObservation,parentFolder, accessToken,ProjectSolutionId,baseTemplate_id,programdetails, userRole)
                if not certificatetemplateid:
                    self.errorVar.append('prepare adding certificate temp failed...')
                    return projectSolutionlink, self.errorVar
                print("certificate added...")
                projectSolutionlink = self.prepareProgramSuccessSheet(MainFilePath, parentFolder, programFile, ProjectSolutionExternalId, ProjectSolutionId, accessToken, programdetails, userRole)
                print(projectSolutionlink,"projectSolutionlink")
                if not projectSolutionlink:
                    self.errorVar.append('Fetching DeepLink Failed...')
                    return projectSolutionlink, self.errorVar
                else:
                    print("solution created successfully!")
                    return projectSolutionlink, self.errorVar
                
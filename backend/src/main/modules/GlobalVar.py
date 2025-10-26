import xlrd, os, wget, sys, time, shutil, requests, csv, jwt, re
from dotenv import load_dotenv
from pathlib import Path
from backend.src.main.modules.common_config import *
from backend.src.main.modules.Programs import *


env_path = Path(__file__).resolve().parents[1] / "apiServices" / "src" / "main" / ".env"

# Load the .env file
load_dotenv(dotenv_path=env_path)

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

# class ProjectValidator:
#     def __init__(self):
#         self.errorVar = []

class GlobalVariables:
    def __init__(self):
        self.errorVar = []

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
    
    def createFileStructre(self, MainFilePath):
        if not os.path.isdir(MainFilePath + '/SolutionFiles'):
            os.mkdir(MainFilePath + '/SolutionFiles')
        
        # Extract the file name regardless of the path format
        # fileNameSplit = os.path.basename(str(addResourceSolution))
        
        # if ".xlsx" in fileNameSplit:
            # ts = str(time.time()).replace(".", "_")
            # folderName = fileNameSplit.replace(".xlsx", "-" + str(ts))
            # os.mkdir(MainFilePath + '/SolutionFiles/' + str(folderName))
        path = MainFilePath + '/SolutionFiles'
        path = os.path.join(path, str('apiHitLogs'))
        os.mkdir(path)
        # else:
        #     self.errorVar = "File Error.offff"        
        returnPathStr = os.path.join(MainFilePath + '/SolutionFiles')

        # if not os.path.isdir(returnPathStr + "/user_input_file"):
        #     os.mkdir(returnPathStr + "/user_input_file")

        # shutil.copy(addResourceSolution, os.path.join(returnPathStr, "user_input_file"))
        return returnPathStr
    
    def typeofresource(self, addResourceSolution):
        wbObservation1 = xlrd.open_workbook(addResourceSolution, on_demand=True)
        sheetNames1 = wbObservation1.sheet_names()
        rubrics_sheet_names = ['Instructions', 'details', 'framework', 'ECMs or Domains', 'questions','Criteria_Rubric-Scoring', 'Domain(theme)_rubric_scoring']
        rubrics_sheet_IMP_names = ['Instructions', 'details', 'framework', 'ECMs or Domains', 'questions','Criteria_Rubric-Scoring', 'Domain(theme)_rubric_scoring', 'Imp mapping']
        observation_sheet_names = ['Instructions', 'details', 'criteria', 'questions']
        survey_sheet_names = ['Instructions', 'details', 'questions']
        project_sheet_names = ['Instructions', 'Project upload', 'Tasks upload','Certificate details']

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
            self.errorVar = ("No Resources Detected in the Resource sheet.")
        return typeofSolution
    
    def fetchUserDetails(self, accessToken, ppd):
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
                self.errorVar = error_message
                return False
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            print(self.errorVar)

    def check_sequence(arr):
        for i in range(1, len(arr)):
            if arr[i] != arr[i - 1] + 1:
                return False
        return True
    
    def row_is_empty(sheet, r):
            """Return True if every cell in row r is blank/whitespace."""
            return all(str(sheet.cell(r, c).value).strip() == "" for c in range(sheet.ncols))
    
    def ObservationFileRead(self, addResourceSolution, impLedObsFlag, typeofSolution,PRName):
        ObservationDict = {
            "details": {},
            "framework": [],
            "ecms or domains": [],
            "questions": []
        }
        questionsequenceArr = []
        ecmIds = []
        criteriaLevels = []
        criteriaExternalIds = []
        pointBasedValue = ""
        wbObservation         = xlrd.open_workbook(addResourceSolution, on_demand=True)
        ObservationSheetNames = wbObservation.sheet_names()

        # Iterate through all sheets
        for sheet_name in ObservationSheetNames:
            sheet = wbObservation.sheet_by_name(sheet_name)
            sheet_lc = sheet_name.strip().lower()

            # Skip Instructions sheet
            if sheet_name.lower().strip() == "instructions":
                continue

            # ---------- DETAILS ----------
            if sheet_lc == "details":
                print("--->Checking details sheet...")

                keys = [sheet.cell(1, c).value for c in range(sheet.ncols)]

                mandatory_cols = [
                    "observation_solution_name",
                    "observation_solution_description",
                    "Username/user id/email id/phone no. of the Content creator",
                    "Name_of_the_creator",
                    "language",
                    "allow_multiple_submissions",
                    "keywords",
                    "scoring_system",
                    "entity_type"
                ]

                if "details" not in ObservationDict:
                    ObservationDict["details"] = {}

                if sheet.nrows < 3:
                    self.errorVar.append(f"No data row found in sheet '{sheet_name}'.")
                elif sheet.nrows > 3:
                    self.errorVar.append(f"Only one data row is allowed in sheet '{sheet_name}', found multiple.")

                r = 2
                if GlobalVariables.row_is_empty(sheet, r):
                    self.errorVar.append(f"Row {r+1} in sheet '{sheet_name}' cannot be empty.")

                for c in range(sheet.ncols):
                    col_name = keys[c]
                    value = sheet.cell(r, c).value

                    if col_name in mandatory_cols and (value is None or str(value).strip() == ""):
                        self.errorVar.append(f"Mandatory column '{col_name}' cannot be empty in row {r+1}.")

                    if col_name == "observation_solution_name" and str(value).strip() != str(PRName).strip():
                        self.errorVar.append("The Resource name in program template does not match the name on Observation Template")

                    ObservationDict["details"][col_name] = value
                pointBasedValue = ObservationDict["details"].get("scoring_system", "")    

            # ---------- FRAMEWORK ----------
            elif sheet_lc == "framework":
                print("--->Checking Framework sheet...")
                keys = [sheet.cell(1, c).value for c in range(sheet.ncols)]

                if "framework" not in ObservationDict:
                    ObservationDict["framework"] = []

                mandatory_cols = ["Domain ID", "Domain Name", "Criteria ID", "Criteria Name"]

                for r in range(2, sheet.nrows):
                    if GlobalVariables.row_is_empty(sheet, r):
                        self.errorVar.append(f"Row {r+1} in sheet '{sheet_name}' cannot be completely empty.")

                    row_data = {
                        keys[c]: sheet.cell(r, c).value
                        for c in range(sheet.ncols)
                    }

                    for col in mandatory_cols:
                        if not str(row_data.get(col)).strip():
                            self.errorVar.append(
                                f"Mandatory column '{col}' cannot be empty "
                                f"in sheet '{sheet_name}', row {r+1}."
                            )
                    ObservationDict["framework"].append(row_data)
                
            # ---------- ECMS OR DOMAINS ----------
            elif sheet_lc.strip().lower() == "ecms or domains":
                print("--->Checking ecms or domains sheet...")
                keys = [sheet.cell(1, c).value for c in range(sheet.ncols)]
                if "ecms or domains" not in ObservationDict:
                    ObservationDict["ecms or domains"] = []

                for r in range(2, sheet.nrows):
                    if GlobalVariables.row_is_empty(sheet, r):
                        self.errorVar.append(
                            f"Row {r + 1} in sheet '{sheet_name}' cannot be empty."
                        )
                    dictDetailsEnv = {
                        keys[c]: sheet.cell(r, c).value for c in range(sheet.ncols)
                    }

                    ECMorDomainDict = {}
                    required_fields = [
                        "ECM Id/Domian ID",
                        "section_id",
                        "section_name",
                        "ECM Name/Domain Name",
                        "Is ECM Mandatory?",
                    ]

                    for field in required_fields:
                        val = dictDetailsEnv.get(field)
                        if val is None or str(val).strip() == "":
                            self.errorVar.append(
                                f"'{field}' cannot be empty in sheet '{sheet_name}' (row {r + 1})."
                            )
                        ECMorDomainDict[field] = val

                    ecm_id_lower = str(ECMorDomainDict["ECM Id/Domian ID"]).lower()
                    if ecm_id_lower not in ecmIds:
                        ecmIds.append(ecm_id_lower)

                    ObservationDict["ecms or domains"].append(ECMorDomainDict)
                
            # ---------- QUESTIONS ----------
            elif sheet_lc == "questions":
                
                MANDATORY_COLUMNS = [
                    "section_id", "criteria_id", "question_id",
                    "question_sequence", "question_primary_language",
                    "question_response_type", "response_required",
                    "question_weightage"
                ]

                CONDITIONAL_COLUMNS = {
                    "parent_question_id":
                        lambda row: bool(row.get("parent_question_value") or row.get("show_when_parent_question_value_is")),
                    "show_when_parent_question_value_is":
                        lambda row: bool(row.get("parent_question_id")),
                    "parent_question_value":
                        lambda row: bool(row.get("parent_question_id")),
                    "instance_parent_question_id":
                        lambda row: bool(row.get("instance_identifier")),
                    "instance_identifier":
                        lambda row: row.get("question_response_type", "").strip().lower() == "matrix",
                    "date_auto_capture":
                        lambda row: row.get("question_response_type", "").strip().lower() == "date",
                    "min_number_value":
                        lambda row: row.get("question_response_type", "").strip().lower() == "slider",
                    "max_number_value":
                        lambda row: row.get("question_response_type", "").strip().lower() == "slider",
                }

                OPTIONAL_COLUMNS = [
                    "section_header", "page", "question_number",
                    "question_secondory_language", "question_tip", "question_hint",
                    "file_upload", "show_remarks"
                ]

                keys = [sheet.cell(1, c).value.strip() for c in range(sheet.ncols)]
                response_col_pattern = re.compile(r"^(response\(R\d+\)|Score for R\d+|response\(R\d+\)_hint)$", re.IGNORECASE)

                for r in range(2, sheet.nrows):
                    if GlobalVariables.row_is_empty(sheet, r):
                        self.errorVar.append(f"Row {r+1} in sheet '{sheet_name}' cannot be empty.")

                    row_data = {keys[c]: sheet.cell(r, c).value for c in range(sheet.ncols)}
                    for col in MANDATORY_COLUMNS:
                        if not str(row_data.get(col)).strip():
                            self.errorVar.append(
                                f"'{col}' is mandatory in sheet '{sheet_name}', row {r+1}."
                            )

                    for col, condition in CONDITIONAL_COLUMNS.items():
                        if condition(row_data) and not str(row_data.get(col)).strip():
                            self.errorVar.append(
                                f"'{col}' is mandatory when its condition is met "
                                f"in sheet '{sheet_name}', row {r+1}."
                            )

                    if row_data.get("question_response_type", "").strip().lower() in ("radio", "multiselect"):
                        responses_present = [
                            k for k in keys
                            if re.match(r"^response\(R\d+\)$", k, re.IGNORECASE)
                            and str(row_data.get(k)).strip()
                        ]
                        if len(responses_present) < 2:
                            self.errorVar.append(
                                f"At least 2 response options are required for "
                                f"radio/multiselect questions (row {r+1})."
                            )
                    ObservationDict["questions"].append(row_data)
            
            # ---------- IMP MAPPING ----------
            elif sheet_lc.strip().lower() == "imp mapping" and typeofSolution == 5:
                print("--->Checking Imp mapping sheet...")
                countImps = 1
                keysEnv = [sheet.cell(1, col_index_env).value
                        for col_index_env in range(sheet.ncols)]
                if "imp mapping" not in ObservationDict:
                    ObservationDict["imp mapping"] = []
                for row_index_env in range(2, sheet.nrows):

                    first_col_value = str(sheet.cell(row_index_env, 0).value).strip()
                    if not first_col_value:
                        self.errorVar.append(
                            f"criteriaId (first column) is mandatory and cannot be empty "
                            f"at row {row_index_env + 1} in sheet '{sheet_lc}'."
                        )
                    dictDetailsEnv = {
                        keysEnv[col_index_env]: sheet.cell(row_index_env, col_index_env).value
                        for col_index_env in range(sheet.ncols)
                    }
                    for eachCols in dictDetailsEnv.keys():
                        if eachCols.strip() == "L" + str(countImps) + "-improvement-projects":
                            countImps += 1

                    ObservationDict["imp mapping"].append(dictDetailsEnv)
                countImps -= 1

            # ---------- CRITERIA-RUBRIC-SCORING ----------
            elif sheet_lc.strip().lower() == "criteria_rubric-scoring" and pointBasedValue.lower() != "null":
                print("--->Checking Criteria Rubric sheet...")
                keysEnv = [sheet.cell(1, c).value for c in range(sheet.ncols)]
                dynamic_L_cols = []
                pattern = re.compile(r"^L\d+\s+SCORE$", re.IGNORECASE)
                for key in keysEnv:
                    if pattern.match(key.strip()):
                        dynamic_L_cols.append(key.strip())
                mandatory_cols = ["criteriaId", "weightage"] + dynamic_L_cols
                optional_cols = ["Ln SCORE"]
                for key in keysEnv:
                    if key not in mandatory_cols + optional_cols:
                        print(f"---> {key} : unexpected column detected (will be ignored).")
                if "criteria_rubric-scoring" not in ObservationDict:
                    ObservationDict["criteria_rubric-scoring"] = []

                cR_extIds = []
                for r in range(2, sheet.nrows):
                    row_dict = {keysEnv[c]: sheet.cell(r, c).value for c in range(sheet.ncols)}
                    for col in dynamic_L_cols:
                        val = row_dict.get(col)
                        if val is None or str(val).strip() == "":
                            row_dict[col] = ""
                        elif isinstance(val, (int, float)):
                            min_val = int(val) - 5
                            max_val = int(val)
                            row_dict[col] = f"{min_val}<SCORE<={max_val}"
                        else:
                            row_dict[col] = str(val)

                    for col in mandatory_cols:
                        if not str(row_dict.get(col)).strip():
                            self.errorVar.append(
                                f"Mandatory column '{col}' cannot be empty "
                                f"in sheet '{sheet_name}', row {r+1}."
                            )
                    cR_extIds.append(str(row_dict["criteriaId"]).strip().lower())

                    for col in optional_cols:
                        if col not in row_dict or row_dict[col] is None:
                            row_dict[col] = ""
                        else:
                            row_dict[col] = str(row_dict[col])
                    ObservationDict["criteria_rubric-scoring"].append(row_dict)
                if len(cR_extIds) != len(set(cR_extIds)):
                    self.errorVar.append("Duplicate criteriaId detected in criteria_rubric-scoring sheet.")
        
            # ---------- CRITERIA-RUBRIC-SCORING ----------            
            elif sheet_lc.strip().lower() == 'domain(theme)_rubric_scoring' and pointBasedValue.lower() != "null":
                print("--->Checking Theme Rubrics sheet")
                keysEnv = [sheet.cell(1, c).value for c in range(sheet.ncols)]
                dynamic_L_cols = []
                pattern = re.compile(r"^L\d+$", re.IGNORECASE)
                for key in keysEnv:
                    if pattern.match(key.strip()):
                        dynamic_L_cols.append(key.strip())

                mandatory_cols = ['domain_Id', 'domain_name', 'weightage'] + dynamic_L_cols
                optional_cols = ['Ln']

                for key in keysEnv:
                    if key not in mandatory_cols + optional_cols:
                        print(f"---> {key} : unexpected column detected (will be ignored).")

                if "theme_rubric_scoring" not in ObservationDict:
                    ObservationDict["theme_rubric_scoring"] = []

                for r in range(2, sheet.nrows):
                    row_dict = {keysEnv[c]: sheet.cell(r, c).value for c in range(sheet.ncols)}

                    for col in dynamic_L_cols:
                        val = row_dict.get(col)
                        if val is None or str(val).strip() == "":
                            row_dict[col] = ""
                        elif isinstance(val, (int, float)):
                            min_val = int(val) - 5
                            max_val = int(val)
                            row_dict[col] = f"{min_val}<SCORE<={max_val}"
                        else:
                            row_dict[col] = str(val)

                    for col in mandatory_cols:
                        if not str(row_dict.get(col)).strip():
                            self.errorVar.append(f"Mandatory column '{col}' cannot be empty in sheet '{sheet_lc}', row {r+1}.")

                    for col in optional_cols:
                        if col not in row_dict or row_dict[col] is None:
                            row_dict[col] = ""
                        else:
                            row_dict[col] = str(row_dict[col])
                    ObservationDict["theme_rubric_scoring"].append(row_dict)
        if self.errorVar:
            print("Validation failed with the following errors:")
            for err in self.errorVar:
                print(" -", err)
            return False
        return ObservationDict

    def ObservationWORFileRead(self,addResourceSolution,PRName):
        ObservationWORDict = {
            "details": {},
            "criteria": [],
            "questions": []
        }
        questionsequenceArr = []
        criteria_id_arr = []
        ecmIds = []
        criteriaLevels = []
        criteriaExternalIds = []
        pointBasedValue = ""
        wbObservation = xlrd.open_workbook(addResourceSolution, on_demand=True)
        ObservationSheetNames = wbObservation.sheet_names()

        # Iterate through all sheets
        for sheet_name in ObservationSheetNames:
            sheet = wbObservation.sheet_by_name(sheet_name)
            sheet_lc = sheet_name.strip().lower()

            # Skip Instructions sheet
            if sheet_name.lower().strip() == "instructions":
                continue
            elif sheet_lc == "details":
                print("--->Checking details sheet...")

                keys = [sheet.cell(1, c).value for c in range(sheet.ncols)]

                mandatory_cols = [
                    "observation_solution_name",
                    "observation_solution_description",
                    "Username/user id/email id/phone no. of the Content creator",
                    "Name_of_the_creator",
                    "language",
                    "keywords",
                    "scoring_system",
                    "entity_type"
                ]

                if "details" not in ObservationWORDict:
                    ObservationWORDict["details"] = {}

                if sheet.nrows < 3:
                    self.errorVar.append(f"No data row found in sheet '{sheet_name}'.")
                elif sheet.nrows > 3:
                    self.errorVar.append(f"Only one data row is allowed in sheet '{sheet_name}', found multiple.")

                r = 2
                if GlobalVariables.row_is_empty(sheet, r):
                    self.errorVar.append(f"Row {r+1} in sheet '{sheet_name}' cannot be empty.")

                for c in range(sheet.ncols):
                    col_name = keys[c]
                    value = sheet.cell(r, c).value

                    if col_name in mandatory_cols and (value is None or str(value).strip() == ""):
                        self.errorVar.append(f"Mandatory column '{col_name}' cannot be empty in row {r+1}.")

                    if col_name == "observation_solution_name" and str(value).strip() != str(PRName).strip():
                        self.errorVar.append("The Resource name in program template does not match the name on Observation Template")
                    
                    ObservationWORDict["details"][col_name] = value
                pointBasedValue = ObservationWORDict["details"].get("scoring_system", "")

            elif sheet_lc == "criteria":
                print("--->Checking criteria sheet...")
                detailsEnvSheet = wbObservation.sheet_by_name(sheet_lc)
                keysEnv = [detailsEnvSheet.cell(1, col_index).value.strip() for col_index in range(detailsEnvSheet.ncols)]
                mandatory_criteria_cols = ["criteria_id", "criteria_name"]
                # Initialize the criteria block in the main dict
                if "criteria" not in ObservationWORDict:
                    ObservationWORDict["criteria"] = []
                criteria_id_arr = []
                # Iterate through each data row
                for row_index in range(2, detailsEnvSheet.nrows):
                    row_data = {
                        keysEnv[col_index]: detailsEnvSheet.cell(row_index, col_index).value
                        for col_index in range(detailsEnvSheet.ncols)
                    }
                    # Trim and validate values
                    criteria_id = str(row_data.get("criteria_id", "")).strip()
                    criteria_name = str(row_data.get("criteria_name", "")).strip()

                    if not criteria_id:
                        self.errorVar.append(f"'criteria_id' must not be empty in row {row_index + 1}")
                    if not criteria_name:
                        self.errorVar.append(f"'criteria_name' must not be empty in row {row_index + 1}")

                    # Only append if both values are present
                    if criteria_id and criteria_name:
                        ObservationWORDict["criteria"].append({
                            "criteria_id": criteria_id,
                            "criteria_name": criteria_name
                        })
                        criteria_id_arr.append(criteria_id)

                # Check for uniqueness
                if len(criteria_id_arr) != len(set(criteria_id_arr)):
                    self.errorVar.append("Duplicate 'criteria_id' values found. They must be unique.")

            elif sheet_lc == "questions":
                print("--->Checking questions sheet...")

                MANDATORY_COLUMNS = [
                    "criteria_id", "question_sequence", "question_id",
                    "page", "question_number", "question_primary_language",
                    "question_response_type", "response_required"
                ]

                CONDITIONAL_COLUMNS = {
                    "parent_question_id":
                        lambda row: bool(row.get("parent_question_value") or row.get("show_when_parent_question_value_is")),
                    "show_when_parent_question_value_is":
                        lambda row: bool(row.get("parent_question_id")),
                    "parent_question_value":
                        lambda row: bool(row.get("parent_question_id")),
                    "instance_parent_question_id":
                        lambda row: bool(row.get("instance_identifier")),
                    "instance_identifier":
                        lambda row: row.get("question_response_type", "").strip().lower() == "matrix",
                    "date_auto_capture":
                        lambda row: row.get("question_response_type", "").strip().lower() == "date",
                    "min_number_value":
                        lambda row: row.get("question_response_type", "").strip().lower() == "slider",
                    "max_number_value":
                        lambda row: row.get("question_response_type", "").strip().lower() == "slider",
                }

                OPTIONAL_COLUMNS = [
                    "question_secondory_language", "question_tip", "question_hint",
                    "file_upload", "show_remarks"
                ]

                keys = [sheet.cell(1, c).value.strip() for c in range(sheet.ncols)]
                response_col_pattern = re.compile(r"^response\(R\d+\)$", re.IGNORECASE)
                response_hint_pattern = re.compile(r"^response\(R\d+\)_hint$", re.IGNORECASE)

                if "questions" not in ObservationWORDict:
                    ObservationWORDict["questions"] = []

                for r in range(2, sheet.nrows):
                    if GlobalVariables.row_is_empty(sheet, r):
                        self.errorVar.append(f"Row {r+1} in sheet '{sheet_name}' cannot be empty.")

                    row_data = {keys[c]: sheet.cell(r, c).value for c in range(sheet.ncols)}

                    # --- Check MANDATORY columns ---
                    for col in MANDATORY_COLUMNS:
                        if not str(row_data.get(col)).strip():
                            self.errorVar.append(f"'{col}' is mandatory in sheet '{sheet_name}', row {r+1}.")

                    # --- Check CONDITIONAL columns ---
                    for col, condition in CONDITIONAL_COLUMNS.items():
                        if condition(row_data) and not str(row_data.get(col)).strip():
                            self.errorVar.append(
                                f"'{col}' is conditionally mandatory in sheet '{sheet_name}', row {r+1}, "
                                f"because its condition was met."
                            )

                    # --- Validate response columns for radio/multiselect ---
                    response_type = row_data.get("question_response_type", "").strip().lower()
                    if response_type in ("radio", "multiselect"):
                        responses_present = [
                            k for k in keys if response_col_pattern.match(k) and str(row_data.get(k)).strip()
                        ]
                        if len(responses_present) < 2:
                            self.errorVar.append(
                                f"At least 2 response options are required for radio/multiselect "
                                f"in sheet '{sheet_name}', row {r+1}."
                            )

                    # --- Append cleaned row to ObservationDict ---
                    ObservationWORDict["questions"].append(row_data)
        if self.errorVar:
            print("Validation failed with the following errors:")
            for err in self.errorVar:
                print(" -", err)
            return False
        return ObservationWORDict

    def SurveyFileRead(self, addResourceSolution,PRName,ResStartDate,ResEndDate):

        SurveyDict = {
            "details": {},
            "questions": []
        }
        wbObservation = xlrd.open_workbook(addResourceSolution, on_demand=True)
        ObservationSheetNames = wbObservation.sheet_names()

        # Iterate through all sheets
        for sheet_name in ObservationSheetNames:
            sheet = wbObservation.sheet_by_name(sheet_name)
            sheet_lc = sheet_name.strip().lower()
            if sheet_name.lower().strip() == "instructions":
                continue
            
            elif sheet_lc == "details":
                print("--->Checking details sheet...")

                keys = [sheet.cell(1, c).value for c in range(sheet.ncols)]

                mandatory_cols = [
                    "survey_solution_name",
                    "survey_solution_description",	
                    "Name_of_the_creator",	
                    "Username/user id/email id/phone no. of the Content creator",
                    "survey_start_date",
                    "survey_end_date"   
                ]

                if "details" not in SurveyDict:
                    SurveyDict["details"] = {}

                if sheet.nrows < 3:
                    self.errorVar.append(f"No data row found in sheet '{sheet_name}'.")
                elif sheet.nrows > 3:
                    self.errorVar.append(f"Only one data row is allowed in sheet '{sheet_name}', found multiple.")

                r = 2
                if GlobalVariables.row_is_empty(sheet, r):
                    self.errorVar.append(f"Row {r+1} in sheet '{sheet_name}' cannot be empty.")

                for c in range(sheet.ncols):
                    col_name = keys[c]
                    value = sheet.cell(r, c).value

                    if col_name in mandatory_cols and (value is None or str(value).strip() == ""):
                        self.errorVar.append(f"Mandatory column '{col_name}' cannot be empty in row {r+1}.")

                    if col_name == "survey_solution_name" and str(value).strip() != str(PRName).strip():
                        self.errorVar.append("The Resource name in program template does not match the name on Survey Template")
                    
                    if col_name == "survey_start_date" and str(value).strip() != str(ResStartDate).strip():
                        self.errorVar.append("The survey_start_date in Survey template does not match the dates on Program Template")
                    
                    if col_name == "survey_end_date" and str(value).strip() != str(ResEndDate).strip():
                        self.errorVar.append("The survey_end_date in Survey template does not match the dates on Program Template")

                    SurveyDict["details"][col_name] = value
                pointBasedValue = SurveyDict["details"].get("scoring_system", "")
                
            elif sheet_lc == "questions":
                print("--->Checking questions sheet...")

                MANDATORY_COLUMNS = [
                    "criteria_id", "question_sequence", "question_id",
                    "page", "question_number", "question_primary_language",
                    "question_response_type", "response_required","question_response_validation"
                ]

                CONDITIONAL_COLUMNS = {
                    "parent_question_id":
                        lambda row: bool(row.get("parent_question_value") or row.get("show_when_parent_question_value_is")),
                    "show_when_parent_question_value_is":
                        lambda row: bool(row.get("parent_question_id")),
                    "parent_question_value":
                        lambda row: bool(row.get("parent_question_id")),
                    "instance_parent_question_id":
                        lambda row: bool(row.get("instance_identifier")),
                    "instance_identifier":
                        lambda row: row.get("question_response_type", "").strip().lower() == "matrix",
                    "date_auto_capture":
                        lambda row: row.get("question_response_type", "").strip().lower() == "date",
                    "min_number_value":
                        lambda row: row.get("question_response_type", "").strip().lower() == "slider",
                    "max_number_value":
                        lambda row: row.get("question_response_type", "").strip().lower() == "slider",
                }

                OPTIONAL_COLUMNS = [
                    "section_header","question_secondory_language", "question_tip", "question_hint",
                    "file_upload", "show_remarks"
                ]

                keys = [sheet.cell(1, c).value.strip() for c in range(sheet.ncols)]
                response_col_pattern = re.compile(r"^response\(R\d+\)$", re.IGNORECASE)
                response_hint_pattern = re.compile(r"^response\(R\d+\)_hint$", re.IGNORECASE)

                if "questions" not in SurveyDict:
                    SurveyDict["questions"] = []

                for r in range(2, sheet.nrows):
                    if GlobalVariables.row_is_empty(sheet, r):
                        self.errorVar.append(f"Row {r+1} in sheet '{sheet_name}' cannot be empty.")

                    row_data = {keys[c]: sheet.cell(r, c).value for c in range(sheet.ncols)}

                    # --- Check MANDATORY columns ---
                    for col in MANDATORY_COLUMNS:
                        if not str(row_data.get(col)).strip():
                            self.errorVar.append(f"'{col}' is mandatory in sheet '{sheet_name}', row {r+1}.")

                    # --- Check CONDITIONAL columns ---
                    for col, condition in CONDITIONAL_COLUMNS.items():
                        if condition(row_data) and not str(row_data.get(col)).strip():
                            self.errorVar.append(
                                f"'{col}' is conditionally mandatory in sheet '{sheet_name}', row {r+1}, "
                                f"because its condition was met."
                            )

                    # --- Validate response columns for radio/multiselect ---
                    response_type = row_data.get("question_response_type", "").strip().lower()
                    if response_type in ("radio", "multiselect"):
                        responses_present = [
                            k for k in keys if response_col_pattern.match(k) and str(row_data.get(k)).strip()
                        ]
                        if len(responses_present) < 2:
                            self.errorVar.append(
                                f"At least 2 response options are required for radio/multiselect "
                                f"in sheet '{sheet_name}', row {r+1}."
                            )

                    # --- Append cleaned row to ObservationDict ---
                    SurveyDict["questions"].append(row_data)
        if self.errorVar:
            print("Validation failed with the following errors:")
            for err in self.errorVar:
                print(" -", err)
            return False
        return SurveyDict

    def validate_identifier(self,identifier, field_name="Field"):
        pattern = r'^[A-Za-z0-9 -]+$'
        if not re.match(pattern, identifier):
            self.errorVar.append(f"Invalid {field_name}: '{identifier}'. Only A-Z, a-z, 0-9, '-', and '_' are allowed.")
            return False
        else:
            print(f"{field_name} '{identifier}' is valid.")
            return True

    def projectValidate(self, addResourceSolution,PRName):
        print("Validating project temp....")
        hasCertificateFlag = False
        ProjectDict = {
            "Project upload": {},
            "Tasks upload": [],
            "Certificate details": {}
        }

        wb = xlrd.open_workbook(addResourceSolution, on_demand=True)
        sheet_names = wb.sheet_names()

        # ---------------------- PROJECT UPLOAD SHEET ----------------------
        for sheet_name in sheet_names:
            sheet_lc = sheet_name.strip().lower()
            if sheet_lc == "project upload":
                print("--->Checking Project Upload sheet...")
                sheet = wb.sheet_by_name(sheet_name)

                # Header keys
                keys = [sheet.cell(1, c).value.strip() for c in range(sheet.ncols)]

                # Expected columns
                projectDetailsCols = [
                    "title", "projectId", "Username/user id/email id/phone no. of content creator",
                    "categories", "objective", "duration", "recommendedFor", "keywords"
                ]

                # Handle dynamic learningResources
                lentasks = (len(keys) - 11) // 2
                for i in range(lentasks):
                    projectDetailsCols.append(f"learningResources{i+1}-name")
                    projectDetailsCols.append(f"learningResources{i+1}-link")

                projectDetailsCols.append("has certificate")

                # Column check
                if set(projectDetailsCols) - set(keys):
                    missing = set(projectDetailsCols) - set(keys)
                    self.errorVar.append(f"Missing columns in 'Project Upload' sheet: {', '.join(missing)}")

                # Must have at least one data row (2nd row is header info)
                if sheet.nrows < 3:
                    self.errorVar.append("No data row found in 'Project Upload' sheet.")

                # Read row data (assuming only one valid project row)
                r = 2
                row_data = {keys[c]: sheet.cell(r, c).value for c in range(sheet.ncols)}

                # Mandatory fields
                mandatory_cols = [
                    "title", "projectId", "Username/user id/email id/phone no. of content creator",
                    "categories", "objective", "duration", "has certificate"
                ]
                for col in mandatory_cols:
                    if not str(row_data.get(col)).strip():
                        self.errorVar.append(f"Mandatory column '{col}' cannot be empty in 'Project Upload' sheet, row {r+1}.")

                # Field-specific validations
                if row_data.get("title") != PRName:
                    self.errorVar.append("Validation failed: Title in resource template does not match title in program template.")
                
                if row_data.get("has certificate").lower() == "yes":
                    hasCertificateFlag = True
                    if row_data.get("Project Level Evidence"):
                        if (row_data.get("Project Level Evidence") or "") == "yes":
                            if not row_data.get("Minimum No. of Evidence"):
                                row_data["Minimum No. of Evidence"] = "1"
                    else:    
                        self.errorVar.append("Project Level Evidence - can not be left empty.")
                    

                projectId = row_data.get("projectId", "")
                if projectId and not self.validate_identifier(projectId):
                    self.errorVar.append("ProjectID should be alphanumeric.")

                ProjectDict["Project upload"] = row_data

            # ---------------------- TASKS UPLOAD SHEET ----------------------
            elif sheet_lc == "tasks upload":
                print("--->Checking Tasks Upload sheet...")
                sheet = wb.sheet_by_name(sheet_name)
                keys = [sheet.cell(1, c).value.strip() for c in range(sheet.ncols)]

                if not hasCertificateFlag:
                    mandatory_cols = [
                        "TaskId", "TaskTitle", "Mandatory task(Yes or No)", "isAnExternalTask"
                    ]
                else:
                    mandatory_cols = [
                        "TaskId", "TaskTitle", "Mandatory task(Yes or No)", "isAnExternalTask",
                        "Evidence required for any task for certificate criteria"
                    ]

                # Track previous non-empty values only for merged columns
                merged_columns = [
                    "Evidence required for any task for certificate criteria",
                    "Minimum No. of Evidence for any task criteria"
                ]
                previous_values = {col: "" for col in merged_columns}

                for r in range(2, sheet.nrows):
                    if GlobalVariables.row_is_empty(sheet, r):
                        self.errorVar.append(f"Row {r+1} in 'Tasks Upload' sheet cannot be empty.")
                        continue

                    row_data = {keys[c]: sheet.cell(r, c).value for c in range(sheet.ncols)}

                    # Fill merged column values only if blank
                    for col in merged_columns:
                        if (row_data.get(col) == "" or row_data.get(col) is None):
                            if previous_values[col]:
                                row_data[col] = previous_values[col]
                        else:
                            previous_values[col] = row_data[col]

                    # --- Mandatory column validation ---
                    for col in mandatory_cols:
                        if not str(row_data.get(col)).strip():
                            self.errorVar.append(
                                f"Mandatory column '{col}' cannot be empty in 'Tasks Upload' sheet, row {r+1}."
                            )

                    # --- External Task validation ---
                    if (row_data.get("isAnExternalTask") or "").lower() == "yes":
                        sol_type = (row_data.get("solutionType") or "").lower()
                        if sol_type in ("observation", "survey"):
                            if row_data.get("Solution Name"):
                                if not row_data.get("Number of submissions for observation"):
                                    row_data["Number of submissions for observation"] = "1"
                            else:
                                self.errorVar.append(f"Solution Name - cannot be empty in row {r+1}")

                    # --- Check conflicts between fields ---
                    lr_links = [row_data.get(f"learningResources{i}-link") for i in range(1, 5)]
                    conflicts = [
                        ("Mitra_Link and Solution as a Task cannot be given to the same task",
                        row_data.get("Solution Name") and row_data.get("Mitra_Link")),
                        ("Mitra_Link and learning Resources cannot be given to the same task",
                        row_data.get("Mitra_Link") and any(lr_links)),
                        ("Solution Name and learning Resources cannot be given to the same task",
                        row_data.get("Solution Name") and any(lr_links))
                    ]
                    self.errorVar += [msg for msg, cond in conflicts if cond]

                    # --- Certificate-related logic ---
                    if hasCertificateFlag:
                        if (row_data.get("Evidence required for any task for certificate criteria") or "").strip().lower() == "yes":
                            row_data["Minimum No. of Evidence for any task criteria"] = (
                                row_data.get("Minimum No. of Evidence for any task criteria") or "1"
                            )
                            row_data["Task Level Evidence req. for certificate criteria"] = ""
                            row_data["Minimum No. of Evidence for task level evidence criteria"] = ""
                        else:
                            task_level_req = (row_data.get("Task Level Evidence req. for certificate criteria") or "").strip().lower()
                            if task_level_req == "yes":
                                row_data["Minimum No. of Evidence for task level evidence criteria"] = (
                                    row_data.get("Minimum No. of Evidence for task level evidence criteria") or "1"
                                )
                            elif not task_level_req:
                                self.errorVar.append(
                                    f"Task Level Evidence req. for certificate criteria - cannot be left empty in row {r+1}."
                                )

                    ProjectDict["Tasks upload"].append(row_data)

            # ---------------------- CERTIFICATE SHEET ----------------------
            elif hasCertificateFlag == True and sheet_lc == "certificate details":
                print("--->Checking Certificate sheet...")
                sheet = wb.sheet_by_name(sheet_name)
                keys = [sheet.cell(1, c).value.strip() for c in range(sheet.ncols)]

                certificateCols = [
                    "Certificate issuer", "Type of certificate", "Logo - 1",
                    "Authorised Signature Image - 1", "Authorised Signatory - 1"
                ]

                if set(certificateCols) - set(keys):
                    missing = set(certificateCols) - set(keys)
                    self.errorVar.append(f"Missing columns in 'Certificate' sheet: {', '.join(missing)}")

                if sheet.nrows < 2:
                    self.errorVar.append("No data row found in 'Certificate' sheet.")
                    continue

                r = 2
                row_data = {keys[c]: sheet.cell(r, c).value for c in range(sheet.ncols)}

                for col in certificateCols:
                    if not str(row_data.get(col)).strip():
                        self.errorVar.append(f"Mandatory column '{col}' cannot be empty in 'Certificate' sheet, row {r+1}.")

                    cert_type = (row_data.get("Type of certificate") or "").strip().lower()
                    if cert_type == "onelogo-onesignature":
                        required_fields = ["Logo - 1", "Authorised Signature Image - 1", "Authorised Signatory - 1"]
                    elif cert_type == "onelogo-twosignature":
                        required_fields = [
                            "Logo - 1",
                            "Authorised Signature Image - 1",
                            "Authorised Signatory - 1",
                            "Authorised Signature Image - 2",
                            "Authorised Signatory - 2",
                        ]
                    elif cert_type == "twologo-onesignature":
                        required_fields = [
                            "Logo - 1",
                            "Logo - 2",
                            "Authorised Signature Image - 1",
                            "Authorised Signatory - 1",
                        ]
                    elif cert_type == "twologo-twosignature":
                        required_fields = [
                            "Logo - 1",
                            "Logo - 2",
                            "Authorised Signature Image - 1",
                            "Authorised Signatory - 1",
                            "Authorised Signature Image - 2",
                            "Authorised Signatory - 2",
                        ]
                    else:
                        required_fields = []
                    if required_fields and not all(row_data.get(field) for field in required_fields):
                        self.errorVar.append(f"One or more required certificate fields are missing: {', '.join(required_fields)}.")

                ProjectDict["Certificate details"] = row_data

        # ---------------------- FINAL VALIDATION ----------------------
        if self.errorVar:
            print("Validation failed with the following errors:")
            for e in self.errorVar:
                print(" -", e)
            return False

        return ProjectDict



    def ReadProgramTemplate(self, programFile):
        MainFilePath = GlobalVariables.createFileStructForProgram(programFile)
        wbPgm = xlrd.open_workbook(programFile, on_demand=True)
        sheetNames = wbPgm.sheet_names()

        pdpmsheet = MainFilePath+ "/pdpmmapping/"
        if not os.path.exists(pdpmsheet):
            os.mkdir(pdpmsheet)
        pdpmcolo1 = ["user","role","entity","entityOperation","keycloak-userId","acl_school","acl_cluster","programOperation",
                    "platform_role","programs","_arrayFields"]
        with open(pdpmsheet + 'mapping.csv', 'w',encoding='utf-8') as file:
            writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC, delimiter=',',lineterminator='\n')
            writer.writerows([pdpmcolo1])

        pgmSheets = ["Instructions", "Program Details", "Resource Details", "Program Manager Details","Role-Subrole Mapping"]
        downloaded_file = {}
        ProgramGlobalDict = {}

        if len(sheetNames) == len(pgmSheets) and sheetNames == pgmSheets:
            print("--->Program Template detected.<---")
            ResStartDate = ""
            ResEndDate = ""
            isProgramnamePresent = True
            for sheetEnv in sheetNames:
                print(sheetEnv,"sheetEnv")
                if sheetEnv.strip().lower() == 'program details':
                    print("--->Checking Program Details sheet...")

                    programDetailsSheet = wbPgm.sheet_by_name(sheetEnv)
                    keysEnv = [programDetailsSheet.cell(1, c).value for c in range(programDetailsSheet.ncols)]

                    mandatory_cols = [
                        "Program ID",
                        "Username/user id/email id/phone no. of Program Designer",
                        "Tenant ID",
                        "Targeted state at program level",
                        "Targeted District at program level",
                        "Title of the Program",
                        "Description of the Program",
                        "Targeted role at program level",
                        "Targeted subrole at program level",
                        "Start date of program",
                        "End date of program"
                    ]

                    optional_cols = [
                        "Targeted Block at program level",
                        "Targeted Cluster at program level",
                        "Targeted School at program level",
                        "Keywords",
                        "Org ID"  # Include as optional — since not all tenants have it
                    ]

                    if "Program Details" not in ProgramGlobalDict:
                        ProgramGlobalDict["Program Details"] = []

                    if programDetailsSheet.nrows < 3:
                        self.errorVar.append("No data row found in 'Program Details' sheet.")
                    else:
                        row_data = {
                            keysEnv[c]: str(programDetailsSheet.cell(2, c).value).strip()
                            for c in range(programDetailsSheet.ncols)
                        }

                        program_dict = {}

                        # Populate mandatory fields
                        for field in mandatory_cols:
                            val = row_data.get(field, "")
                            if not val:
                                self.errorVar.append(f"Mandatory field '{field}' cannot be empty.")
                            program_dict[field.replace(" ", "")] = val

                        # Populate optional fields
                        for field in optional_cols:
                            val = row_data.get(field, "")
                            program_dict[field.replace(" ", "")] = (
                                [v.strip() for v in val.split(",") if v.strip()] if val else []
                            )

                        # --- Tenant-specific OrgID handling ---
                        tenant_id = str(program_dict.get("TenantID", "")).strip().lower()

                        if tenant_id == "shikshalokam":
                            # Org ID must exist and not be empty
                            derived_org = program_dict.get("OrgID", "")
                            if not derived_org:
                                self.errorVar.append("For tenant 'shikshalokam', 'Org ID' is mandatory but missing.")

                            # Ensure OrgID is always a list
                            if isinstance(derived_org, str):
                                program_dict["OrgID"] = [v.strip() for v in derived_org.split(",") if v.strip()]
                            elif isinstance(derived_org, list):
                                program_dict["OrgID"] = [v.strip() for v in derived_org if v.strip()]
                            else:
                                program_dict["OrgID"] = []

                        elif tenant_id == "shikshagraha":
                            # Use Targeted state as Org ID
                            derived_org = program_dict.get("Targetedstateatprogramlevel", "")
                            if isinstance(derived_org, str):
                                program_dict["OrgID"] = [v.strip() for v in derived_org.split(",") if v.strip()]
                            elif isinstance(derived_org, list):
                                program_dict["OrgID"] = [v.strip() for v in derived_org if v.strip()]
                            else:
                                program_dict["OrgID"] = []

                            if not program_dict["OrgID"]:
                                self.errorVar.append(
                                    "For tenant 'shikshagraha', 'Targeted state at program level' is required to derive Org ID."
                                )

                        else:
                            # Optional fallback if unknown tenant
                            derived_org = program_dict.get("Targetedstateatprogramlevel", "")
                            if isinstance(derived_org, str):
                                program_dict["OrgID"] = [v.strip() for v in derived_org.split(",") if v.strip()]
                            elif isinstance(derived_org, list):
                                program_dict["OrgID"] = [v.strip() for v in derived_org if v.strip()]
                            else:
                                program_dict["OrgID"] = []

                            self.errorVar.append(
                                f"Unknown tenant '{tenant_id}'."
                            )

                        # --- Add OrgForAPIs (always a single value) ---
                        if isinstance(program_dict.get("OrgID"), list) and program_dict["OrgID"]:
                            program_dict["OrgForAPIs"] = program_dict["OrgID"][0]
                        else:
                            program_dict["OrgForAPIs"] = ""

                        # Store and print
                        ProgramGlobalDict["Program Details"].append(program_dict)
                        print(ProgramGlobalDict)

                elif sheetEnv.strip().lower() == 'resource details':
                    print("--->Checking Resource Details sheet...")

                    detailsEnvSheet = wbPgm.sheet_by_name(sheetEnv)
                    keysEnv = [detailsEnvSheet.cell(1, c).value for c in range(detailsEnvSheet.ncols)]

                    if "Program Resources" not in ProgramGlobalDict:
                        ProgramGlobalDict["Program Resources"] = []

                    downloaded_file = {}
                    dest_dir = "InputFiles"
                    os.makedirs(dest_dir, exist_ok=True)

                    for row_index_env in range(2, detailsEnvSheet.nrows):
                        dictDetailsEnv = {
                            keysEnv[c]: detailsEnvSheet.cell(row_index_env, c).value
                            for c in range(detailsEnvSheet.ncols)
                        }
                        ResStartDate = dictDetailsEnv.get('Start date of resource')
                        ResEndDate = dictDetailsEnv.get('End date of resource')
                        status = dictDetailsEnv.get('Resource Status')
                        if status is None or str(status).strip().lower() == "existing":
                            continue

                        mandatory_fields = [
                            'Name of resources in program',
                            'Type of resources',
                            'Resource Link',
                            'Resource Status',
                            'Target role at the resource level',
                            'Targeted subrole at resource level',
                            'Start date of resource',
                            'End date of resource'
                        ]

                        resource_dict = {}
                        for field in mandatory_fields:
                            value = dictDetailsEnv.get(field)
                            if value is None or str(value).strip() == "":
                                self.errorVar.append(f"Mandatory field '{field}' cannot be empty in 'Resource Details' sheet, row {row_index_env + 1}.")
                            resource_dict[field.replace(" ", "")] = str(value).strip()
                            resource_dict["typeofSolution"] = 0
                        ProgramGlobalDict["Program Resources"].append(resource_dict)

                        # Handle file download for non-course resources
                        if str(resource_dict['Typeofresources']).lower() != "course":
                            print(f"--->Resource Name: {resource_dict['Nameofresourcesinprogram']}")
                            
                            # Extract file ID safely
                            match = re.search(r"/d/([a-zA-Z0-9-_]+)", resource_dict['ResourceLink'])
                            if not match:
                                self.errorVar.append(f"Invalid Resource Link format for '{resource_dict['Nameofresourcesinprogram']}' at row {row_index_env + 1}.")
                                continue

                            file_id = match.group(1)
                            file_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"

                            download_file = wget.download(file_url, out=dest_dir)
                            downloaded_file[download_file] = resource_dict['Nameofresourcesinprogram']

                elif sheetEnv.strip().lower() == 'program manager details':
                    print("--->Checking Program Manager Details sheet...")

                    detailsEnvSheet = wbPgm.sheet_by_name(sheetEnv)
                    keysEnv = [detailsEnvSheet.cell(1, c).value for c in range(detailsEnvSheet.ncols)]

                    # Ensure master key exists
                    if "ProgramManagerDetails" not in ProgramGlobalDict:
                        ProgramGlobalDict["ProgramManagerDetails"] = []

                    email_ids = []

                    for row_index_env in range(2, detailsEnvSheet.nrows):
                        dictDetailsEnv = {
                            keysEnv[c]: detailsEnvSheet.cell(row_index_env, c).value
                            for c in range(detailsEnvSheet.ncols)
                        }

                        # Skip empty rows
                        if all(str(v).strip() == "" for v in dictDetailsEnv.values()):
                            continue

                        # Validate Email ID
                        email_id = dictDetailsEnv.get('Email ID')
                        if email_id is None or str(email_id).strip() == "":
                            self.errorVar.append(f"Mandatory field 'Email ID' cannot be empty in 'Program Manager Details' sheet, row {row_index_env + 1}.")
                            continue

                        # Add valid emails
                        email_ids.append(str(email_id).encode('utf-8').decode('utf-8').strip())

                    # Store all emails under a single dictionary
                    ProgramGlobalDict["ProgramManagerDetails"] = [{"EmailID": email_ids}]

            print("--->Solution input file successfully downloaded: " + str(downloaded_file))
            parentFolder = self.createFileStructre(MainFilePath)
            for addResourceSolution, PRName in downloaded_file.items():
                print(f"Processing file: {addResourceSolution} for resource: {PRName}")
                typeofSolution = self.typeofresource(addResourceSolution)
                print(typeofSolution)
                for resourceDict in ProgramGlobalDict["Program Resources"]:
                    if resourceDict.get("Nameofresourcesinprogram") == PRName:
                        resourceDict["typeofSolution"] = typeofSolution
                        break
                if typeofSolution == 1 or typeofSolution == 5:
                    if typeofSolution == 5:
                        impLedObsFlag = True
                    else:
                        impLedObsFlag = False
                    ObservationDict = self.ObservationFileRead(addResourceSolution,impLedObsFlag,typeofSolution,PRName)
                    if not ObservationDict:
                        print(self.errorVar)
                    else:
                        solutionName = ObservationDict['details']['observation_solution_name']
                        print(solutionName,"solutionName")
                        for resource in ProgramGlobalDict['Program Resources']:
                            if resource.get('Nameofresourcesinprogram') == solutionName:
                                # resource.pop('ResourceLink', None)
                                resource["ResourceCre"] = ObservationDict

                elif typeofSolution == 2:
                    ObservationWORDict = self.ObservationWORFileRead(addResourceSolution,PRName)
                    if not ObservationWORDict:
                        print(self.errorVar)
                    else:
                        solutionName = ObservationWORDict['details']['observation_solution_name']
                        print(solutionName,"solutionName")
                        for resource in ProgramGlobalDict['Program Resources']:
                            if resource.get('Nameofresourcesinprogram') == solutionName:
                                # resource.pop('ResourceLink', None)
                                resource["ResourceCre"] = ObservationWORDict

                elif typeofSolution == 3:
                    SurveyDict = self.SurveyFileRead(addResourceSolution,PRName,ResStartDate,ResEndDate)
                    if not SurveyDict:
                        print(self.errorVar)
                    else:
                        solutionName = SurveyDict['details']['survey_solution_name']
                        print(solutionName,"solutionName")
                        for resource in ProgramGlobalDict['Program Resources']:
                            if resource.get('Nameofresourcesinprogram') == solutionName:
                                # resource.pop('ResourceLink', None)
                                resource["ResourceCre"] = SurveyDict
                
                elif typeofSolution == 4:
                    ProjectDict = self.projectValidate(addResourceSolution,PRName)
                    if not ProjectDict:
                        print(self.errorVar)
                    else:
                        solutionName = ProjectDict['Project upload']['title']
                        print(solutionName,"solutionName")
                        for resource in ProgramGlobalDict['Program Resources']:
                            if resource.get('Nameofresourcesinprogram') == solutionName:
                                # resource.pop('ResourceLink', None)
                                resource["ResourceCre"] = ProjectDict

            ProgramsInstance = Programs()
            programcreation = ProgramsInstance.programCheckCreate(programFile, MainFilePath, parentFolder, ProgramGlobalDict, PRName)
                
            finalObsRubricSolutionLink = programcreation
            result = {
                "solutionDict": finalObsRubricSolutionLink,
                "programName": program_dict.get('TitleoftheProgram') 
            }
            print(result)
            return result
            # sys.exit()

            #     # CurrentResourceName = PRName
            #     solutionSL = Elevateproject.mainFunc(MainFilePath, programFile, addObservationSolution,resourceName, millisecond, isProgramnamePresent, isCourse,
            #  scopeEntityType=scopeEntityType)
            #     print(solutionSL)
            #     print(solutionSL.items(),"3400")
            #     for resourceName, solutionLink in solutionSL.items():
            #         solutionDict[resourceName] = solutionLink
            #         print()
            # downloaded_file = {}
            # print()
        else:
            print("The provided Template is not a Program Template...")

import os
import csv
import time
import json
import threading
import requests
from backend.src.main.modules.common_config import *
from datetime import datetime
from requests import get, post
from dotenv import load_dotenv
from pathlib import Path

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


class SurveyCreate:
    def __init__(self):
        # keep same interface, just initialize errorVar
        self.errorVar = []

    # ----------------- helpers -----------------

    def _parse_date(self, date_str):
        """Safely parse ISO date; return minimal date on failure."""
        if not date_str:
            return datetime.min
        for fmt in ("%Y-%m-%dT%H:%M:%S.%fZ", "%Y-%m-%dT%H:%M:%SZ"):
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        return datetime.min

    def _build_payload(self, resourceType, include_isReusable=True):
        """Build query payload based on resourceType."""
        query = {"status": "active","isReusable":False,"tenantId":"shikshalokam","orgId":"sot"}
        if resourceType == "observation with rubrics":
            query.update({
                "type": "observation",
                "isRubricDriven": True
            })
            if include_isReusable:
                query["isReusable"] = False
        elif resourceType == "observation without rubrics":
            query.update({
                "type": "observation",
                "isRubricDriven": False
            })
            # isReusable intentionally not forced here (matching your original comment)
        elif resourceType == "survey":
            query.update({
                "type": "survey"
            })
            if include_isReusable:
                query["isReusable"] = False
        elif resourceType == "improvementProject":
            query.update({
                "type": "improvementProject"
            })
            if include_isReusable:
                query["isReusable"] = False
        else:
            query.update({
                "type": "programs"
            })
            if include_isReusable:
                query["isReusable"] = False

        payload = {
            "query": query,
            "mongoIdKeys": [
                "_id",
                "solutionId",
                "metaInformation.solutionId"
            ],
            "limit": 10000
        }

        print(payload,"payload--102")
        return payload
    

    # ----------------- access token -----------------

    def generate_access_token(self):
        header_keyclock_user = {'Content-Type': keyclockapicontent_type}
        try:
            headerKeyClockUser = {
                'Content-Type': "application/x-www-form-urlencoded",
                'origin': "default-qa.tekdinext.com"
            }

            loginBody = {
                'identifier': identifier,
                'password': password
            }
            responseKeyClockUser = requests.post(
                userLoginHost + keyclockapiurl,
                headers=headerKeyClockUser,
                data=loginBody
            )
            messageArr = []
            messageArr.append("URL : " + str(keyclockapiurl))
            messageArr.append("Body : " + str(loginBody))
            messageArr.append("Status Code : " + str(responseKeyClockUser.status_code))

            if responseKeyClockUser.status_code == 200:
                responseKeyClockUser = responseKeyClockUser.json()
                accessTokenUser = responseKeyClockUser['result']['access_token']
                messageArr.append("Acccess Token : " + str(accessTokenUser))
                print("--->Access Token Generated!")
                return accessTokenUser

            else:
                print("Error in generating Access token")
                print("Status code : " + str(responseKeyClockUser.status_code))
                self.errorVar.append(responseKeyClockUser.text)
                return False
        except Exception as e:
            print(f"Error occurred: {str(e)}")
            self.errorVar.append(str(e))
            print(self.errorVar, "---> API-Error")
            return False

    # ----------------- fetch in memory -----------------

    def fetch_solution_id(self, access_token, resourceType):
        if not access_token:
            return None

        solution_update_api = f"{elevateprojecthost}{dbfind_view}solutions"
        solution_survey_api = f"{internal_kong_ip}{dbfind_view}solutions"
        program_details_api = f"{internal_kong_ip}{dbfind_view}programs"

        headers = {
            'Content-Type': 'application/json',
            'Authorization': authorization,
            'X-auth-token': access_token,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token
        }
        print(headers, "headers")

        payload = self._build_payload(resourceType, include_isReusable=True)
        print(resourceType,"resourceType")

        try:
            # choose URL based on resourceType
            if resourceType == "improvementProject":
                url = solution_update_api
            elif resourceType == "programs":
                url = program_details_api
            else:
                url = solution_survey_api

            response = requests.post(
                url=url,
                headers=headers,
                data=json.dumps(payload)
            )

            print(solution_update_api, "solution_update_api")
            print(payload, "data")
            print(response.text, "response")
            response.raise_for_status()

            result = response.json().get('result', [])
            result2 = response.json().get('count', [])
            print(result2)

        except requests.RequestException as e:
            print(f"Error fetching solutions: {e}")
            self.errorVar.append(str(e))
            return None

        # all_solution_ids is not needed anywhere, so dropped for efficiency
        all_parent_solution_ids = {
            item.get('parentSolutionId')
            for item in result
            if 'parentSolutionId' in item
        }

        solutions_data = []
        for item in result:
            solution_id = item.get('_id', 'N/A')
            parent_solution_id = item.get('parentSolutionId', 'N/A')
            if solution_id in all_parent_solution_ids:
                continue

            solution_data = {
                'Link': item.get('link', 'None'),
                'SOLUTION_NAME': item.get('name', 'N/A'),
                'SOLUTION_CREATED_DATE': item.get('createdAt') if item.get('createdAt') != 'None' else None,
                'START_DATE': item.get('startDate') if item.get('startDate') != 'None' else None,
                'END_DATE': item.get('endDate') if item.get('endDate') != 'None' else None,
                'PROGRAM_NAME': item.get('programName', 'None'),
                'ORGID': item.get('orgId', 'None'),
                'TENANTID': item.get('tenantId', 'None'),
            }
            solutions_data.append(solution_data)

        # safer sort (handles None or bad format)
        solutions_data.sort(
            key=lambda x: self._parse_date(x['SOLUTION_CREATED_DATE']),
            reverse=True
        )
        return solutions_data

    # ----------------- fetch to CSV -----------------

    def fetch_solution_id_csv(self, access_token, resurceType, csv_file_path='solutions.csv'):
        print(resurceType, "resurceType")
        if not access_token:
            return None
        print(access_token,"access_token")
        solution_update_api = f"{elevateprojecthost}{dbfind_view}solutions"
        headers = {
            'Content-Type': 'application/json',
            'Authorization': authorization,
            'X-auth-token': access_token,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token
        }
        print(headers,"headers")

        # reuse helper; for CSV we don't force isReusable flag
        if resurceType == "observation with rubrics":
            payload = self._build_payload("observation with rubrics", include_isReusable=False)
        elif resurceType == "observation without rubrics":
            payload = self._build_payload("observation without rubrics", include_isReusable=False)
        elif resurceType == "survey":
            payload = self._build_payload("survey", include_isReusable=False)
        else:
            payload = self._build_payload("improvementProject", include_isReusable=False)

        try:
            response = requests.post(
                url=solution_update_api,
                headers=headers,
                data=json.dumps(payload)
            )
            response.raise_for_status()
            result = response.json().get('result', [])
        except requests.RequestException as e:
            print(f"Error fetching solutions: {e}")
            self.errorVar.append(str(e))
            return None

        # sort using safe parser
        result.sort(
            key=lambda x: self._parse_date(x.get('createdAt')),
            reverse=True
        )

        all_parent_solution_ids = {
            item['parentSolutionId']
            for item in result
            if 'parentSolutionId' in item
        }

        # ensure we always work with the actual path we return
        csv_filepath = os.path.abspath(csv_file_path)

        with open(csv_filepath, mode='w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['SOLUTION_ID', 'SOLUTION_NAME', 'SOLUTION_CREATED_DATE', 'START_DATE', 'END_DATE']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()

            for item in result:
                solution_id = item.get('_id', 'N/A')
                parent_solution_id = item.get('parentSolutionId', 'N/A')
                if solution_id in all_parent_solution_ids:
                    continue

                solution_name = item.get('name', 'N/A')
                solution_createdat = item.get('createdAt', 'N/A')
                startdate = item.get('startDate', 'None')
                endate = item.get('endDate', 'None')

                writer.writerow({
                    'SOLUTION_ID': solution_id,
                    'SOLUTION_NAME': solution_name,
                    'SOLUTION_CREATED_DATE': solution_createdat,
                    'START_DATE': startdate,
                    'END_DATE': endate,
                    'PROGRAM_NAME': item.get('programName', 'None'),
                    'ORGID': item.get('orgId', 'None'),
                    'TENANTID': item.get('tenantId', 'None')
                })

        print("Data written to CSV successfully.")
        local = os.getcwd()
        print(local)
        downloadcsv = csv_filepath
        print(f"CSV file is created at: {csv_filepath}")
        self.schedule_deletion(csv_filepath)
        return downloadcsv

    # ----------------- file deletion -----------------

    def schedule_deletion(self, file_path):
        def delete_file():
            try:
                time.sleep(60)
                if os.path.exists(file_path):
                    os.remove(file_path)
                    print(f"File {file_path} deleted successfully.")
                else:
                    print(f"File {file_path} not found.")
            except Exception as e:
                print(f"Error deleting file: {e}")

        threading.Thread(target=delete_file, daemon=True).start()

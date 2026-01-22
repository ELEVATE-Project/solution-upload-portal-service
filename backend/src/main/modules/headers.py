import backend.src.main.modules.common_config as config
from dotenv import load_dotenv
from pathlib import Path
import os

env_path = Path(__file__).resolve().parents[1] / "apiServices" / "src" / "main" / ".env"
load_dotenv(dotenv_path=env_path)
x_channel_id = os.getenv("x_channel_id")
adminTokenHeaderName = os.getenv("adminTokenHeaderName")
projAdminAccessToken = os.getenv("projAdminAccessToken")
adminAccessToken = os.getenv("adminAccessToken")
internal_access_token = os.getenv("internal_access_token")
jwtTokenSecret = os.getenv("jwtTokenSecret")
authorization = os.getenv("authorization")
authorizationforhost = os.getenv("authorizationforhost")
appname = os.getenv("appname")
host = os.getenv("host")
userLoginHost = os.getenv("userLoginHost")
internal_kong_ip = os.getenv("internal_kong_ip")
elevateprojecthost = os.getenv("elevateprojecthost")
elevateentityhost = os.getenv("elevateentityhost")
identifier = os.getenv("identifier")
password = os.getenv("password")
origin = os.getenv("origin")

class headers:
    #Program Api Headers
    def programSearchHeaders(self, accessToken):
        headersProgramSearch = {
            'Content-Type': config.content_type,
            'X-auth-token': accessToken
        }
        return headersProgramSearch
    
    def validateRoleHeaders(self, tenant_id):
        validateRoleHeaders = {
            'Content-Type': config.content_type,
            'tenantId': tenant_id,
            'X-Channel-id': x_channel_id
        }
        return validateRoleHeaders

    def PheaderFetchEntitytype(self):
        headerFetchEntityListApi = {
            'Content-Type': config.content_type,
            'internal-access-token': internal_access_token
        }
        return headerFetchEntityListApi
    
    def headerFetchEntityDetails(self,tenant_id):
        headerFetchEntityDetailsApi = {
            'tenantId': tenant_id
        }
        return headerFetchEntityDetailsApi
    def headerFetchDetailsEntity(self, TenantID, accessToken):
        headerFetchDetailsEntityApi = {
            'Content-Type': config.content_type,
            'tenantId': TenantID,
            'Authorization': f'Bearer {accessToken}'
            }
        return headerFetchDetailsEntityApi
    def headerProgramCreate(self, accessToken,TenantID,OrgForAPIs, userRole):
        if userRole == 'superadmin':
            headerProgramCreateApi = {
            'internal-access-token': internal_access_token,
            'X-auth-token': accessToken,
            'Content-Type': config.content_type,
            'Authorization': authorization,
            'tenantId': TenantID,
            'orgid': OrgForAPIs,
            adminTokenHeaderName: projAdminAccessToken
        }
        else:
            headerProgramCreateApi = {
                'internal-access-token': internal_access_token,
                'X-auth-token': accessToken,
                'Content-Type': config.content_type,
                'Authorization': authorization,
                'tenantId': TenantID,
                'orgid': OrgForAPIs
            }
        return headerProgramCreateApi
    
    #Project Api Headers
    def headerCheckEntityOfSolution(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            checkEntityOfSolutionHeaders ={
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: projAdminAccessToken
            }
        else:
            checkEntityOfSolutionHeaders ={
                    'Authorization': authorization,
                    'X-auth-token': accessToken,
                    'X-Channel-id': x_channel_id,
                    'internal-access-token': internal_access_token,
                    'Content-Type': config.content_type,
                    'tenantId': tenantid,
                    'orgId' : orgid
            }
        return checkEntityOfSolutionHeaders

    def headersObservationsolutionUpdate(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            observationSolutionUpdateHeaders = {
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: adminAccessToken
            }
        else:
            observationSolutionUpdateHeaders = {
                    'X-auth-token': accessToken,
                    'X-Channel-id': x_channel_id,
                    'internal-access-token': internal_access_token,
                    'Content-Type': config.content_type,
                    'tenantId': tenantid,
                    'orgId' : orgid
            }
        return observationSolutionUpdateHeaders

    def headersprojectUpload(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerProjectUploadApi = {
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: projAdminAccessToken
            }
        else:
            headerProjectUploadApi = {
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headerProjectUploadApi

    def headersFetchSol(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerfetchProjectIdApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: projAdminAccessToken
            }
        else:
            headerfetchProjectIdApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headerfetchProjectIdApi

    def headersTaskUpload(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerTaskUploadApi = {
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: projAdminAccessToken
            }
        else:
            headerTaskUploadApi = {
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headerTaskUploadApi

    def headerCreateSolutionApi(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerCreateSolutionApi = {
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: projAdminAccessToken
            }
        else:
            headerCreateSolutionApi = {
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headerCreateSolutionApi

    def headerMapSolutionProjectAPI(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerMapSolutionProject = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: projAdminAccessToken
            }
        else:
            headerMapSolutionProject = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headerMapSolutionProject

    def headerFetchSolutionDetailFromProgramSheetApi(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerFetchSolutionApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: projAdminAccessToken
            }
        else:
            headerFetchSolutionApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headerFetchSolutionApi

    def header_validate_roles_against_api(self, tenantid):
        headerFetchSolutionApi = {
            'X-Channel-id': x_channel_id,
            'Content-Type': config.content_type,
            'tenantId': tenantid
        }
        return headerFetchSolutionApi

    def headerSolutionUpdateAPI(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerUpdateSolutionApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: projAdminAccessToken
            }
        else:
            headerUpdateSolutionApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headerUpdateSolutionApi

    def headerfetchUserDetailsAPI(self, accessToken):
        headerfetchUserDetails = {
            'internal-access-token': internal_access_token,
            'X-auth-token': accessToken
        }
        return headerfetchUserDetails

    def headerUrlFetchSolutionApi(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerurlFetchSolution = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: projAdminAccessToken
            }
        else:
            headerurlFetchSolution = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headerurlFetchSolution

    def headerURLFetchSolutionLinkApi(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerFetchSolutionLinkApi = {
                'X-auth-token': accessToken,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: projAdminAccessToken
            }
        else:
            headerFetchSolutionLinkApi = {
                'X-auth-token': accessToken,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headerFetchSolutionLinkApi

    def headereditingsvgApi(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headereditingsvgApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: projAdminAccessToken
            }
        else:
            headereditingsvgApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headereditingsvgApi

    def headeraddcertificateApi(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headeraddcertificateApi = {
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: projAdminAccessToken
            }
        else:
            headeraddcertificateApi = {
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headeraddcertificateApi

    def headeruploadcertificateApi(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headeruploadcertificateApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: projAdminAccessToken
            }
        else:
            headeruploadcertificateApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headeruploadcertificateApi

    def headerProjectTemplateUpdateAPI(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerprojectrtemplateupdateApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: projAdminAccessToken
            }
        else:
            headerprojectrtemplateupdateApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headerprojectrtemplateupdateApi
    
    #Survey Api Headers
    def headersCreateSurveySolution(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerCreateSurveySolutionApi = {
                'X-auth-token': accessToken,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: adminAccessToken
            }
        else:
            headerCreateSurveySolutionApi = {
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headerCreateSurveySolutionApi
    
    def headersFetchSolutionDetails(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerFetchSolutionApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: adminAccessToken
            }
        else:
            headerFetchSolutionApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'Content-Type': config.content_type,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headerFetchSolutionApi
    
    def headersFetchSolutionLink(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerFetchSolutionLinkApi = {
                'X-auth-token': accessToken,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: adminAccessToken
            }
        else:
            headerFetchSolutionLinkApi = {
                'X-auth-token': accessToken,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headerFetchSolutionLinkApi
    
    def headersQuestionUpload(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerQuestionUploadApi = {
                "internal-access-token": internal_access_token,
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: adminAccessToken
            }
        else:
            headerQuestionUploadApi = {
                "internal-access-token": internal_access_token,
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headerQuestionUploadApi
    
    def headersImportSurveySolutionTemplate(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerImportSoluTemplateApi = {
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: adminAccessToken
            }
        else:
            headerImportSoluTemplateApi = {
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headerImportSoluTemplateApi
    
    def headersSurveyProgramMapping(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headeSurveyProgramMappingApi = {
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid,
                adminTokenHeaderName: adminAccessToken
            }
        else:
            headeSurveyProgramMappingApi = {
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgId' : orgid
            }
        return headeSurveyProgramMappingApi
    
    #Observation Api Headers
    def headerFetchEntitytype(self, accessToken):
        headerFetchEntityListApi = {
            'Content-Type': config.content_type,
            'X-auth-token': accessToken
        }
        return headerFetchEntityListApi
    
    def headersCriteriaUpload(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerCriteriaUploadApi = {
                'internal-access-token': internal_access_token,
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'tenantId': tenantid,
                'orgid' : orgid
            }
        else:
            headerCriteriaUploadApi = {
                'internal-access-token': internal_access_token,
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'tenantId': tenantid,
                'orgid' : orgid
            }
        return headerCriteriaUploadApi
    
    def headersFrameworkUpload(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerFrameworkUploadApi = {
                'Authorization': authorization,
                "internal-access-token": internal_access_token,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'tenantId': tenantid,
                'orgid' : orgid,
                adminTokenHeaderName: adminAccessToken
            }
        else:
            headerFrameworkUploadApi = {
                'Authorization': authorization,
                "internal-access-token": internal_access_token,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'tenantId': tenantid,
                'orgid' : orgid
            }
        return headerFrameworkUploadApi
    
    def headersThemesUpload(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerThemesUploadApi = {
                'Authorization': authorization,
                "internal-access-token": internal_access_token,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'tenantId': tenantid,
                'orgid' : orgid,
                adminTokenHeaderName: adminAccessToken
            }
        else:
            headerThemesUploadApi = {
                'Authorization': authorization,
                "internal-access-token": internal_access_token,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'tenantId': tenantid,
                'orgid' : orgid
            }
        return headerThemesUploadApi
    
    def headersCreateSolutionFromFramework(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerCreateSolutionApi = {
                'Content-Type': config.content_type,
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgid' : orgid,
                adminTokenHeaderName: adminAccessToken
            }
        else:
            headerCreateSolutionApi = {
                'Content-Type': config.content_type,
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgid' : orgid
            }
        return headerCreateSolutionApi
    
    def headersFetchSolutionCriteria(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headers = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgid' : orgid,
                adminTokenHeaderName: adminAccessToken
            }
        else:
            headers = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgid' : orgid
            }
        return headers
    
    def headersCriteriaRubricUpload(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerCriteriaRubricUploadApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgid' : orgid,
                adminTokenHeaderName: adminAccessToken
            }
        else:
            headerCriteriaRubricUploadApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgid' : orgid
            }
        return headerCriteriaRubricUploadApi
    
    def headersThemeRubricUpload(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headerThemeRubricUploadApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgid' : orgid,
                adminTokenHeaderName: adminAccessToken
            }
        else:
            headerThemeRubricUploadApi = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'X-Channel-id': x_channel_id,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgid' : orgid
            }
        return headerThemeRubricUploadApi
    
    def headersSolutionToProgramMapping(self, tenantid, orgid, accessToken, userRole):
        if userRole == 'superadmin':
            headersSol_prog_mapping = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'Content-Type': config.content_type,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgid' : orgid,
                adminTokenHeaderName: adminAccessToken
            }
        else:
            headersSol_prog_mapping = {
                'Authorization': authorization,
                'X-auth-token': accessToken,
                'Content-Type': config.content_type,
                'internal-access-token': internal_access_token,
                'tenantId': tenantid,
                'orgid' : orgid
            }
        return headersSol_prog_mapping
    
    

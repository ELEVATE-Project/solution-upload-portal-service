#Host for user Services
host = ""
userLoginHost =""
# Host URL for the elevate Samiksha service
internal_kong_ip = ""
elevateprojecthost = ""
# Host URL for the elevate entity service
elevateentityhost = ""

# JSON body for Keycloak API login
keyclockapibody = ''

# Email for login credentials
email = ""

# Password for login credentials
passwordold = ""


identifier = "",
password = ""

# Internal access token used for internal API calls
#internal-access-token = Fqn0m0HQ0gXydRtBCg5l
internal_access_token = ""
authorization = ""
authorizationforhost = ""
# ; appname = diksha
appname = "samiksha"
x_channel_id = ""
# ; internal-access-token = 183c219f984050c0e89d
keyclockapicontent_type = "application/x-www-form-urlencoded"
content_type = "application/json"
# ; keyclockapiurl = /auth/realms/sunbird/protocol/openid-connect/token
keyclockapiurl = "user/v1/account/login"
 
origin= "shikshagrah-qa.tekdinext.com"

# Endpoint for user login API
# Default content type for API requests
# ; content-type = application/json

# Endpoint for fetching user information
userinfoapiurl = "user/v1/user/read"
prouserinfoapiurl = "v1/profile/read"
# Endpoint for searching locations in the entity management service
searchforlocation = "entity-management/v1/entities/find"
criteriauploadapiurl = "v1/criteria/upload"
themeuploadapiurl = "v1/frameworks/uploadThemes/"
solutioncreationapiurl = "v1/observations/importFromFramework"
surveysolutioncreationapiurl = "v1/surveys/createSolutionTemplate"
questionuploadapiurl = "v1/questions/bulkCreate"
criteriarubricuploadapiurl = "v1/solutions/uploadCriteriaRubricExpressions/"
themerubricuploadapiurl = "v1/solutions/uploadThemesRubricExpressions/"
importsurveysolutiontemplateurl = "v1/surveys/importSurveryTemplateToSolution/"
importsurveysolutiontoprogramurl = "v1/surveys/mapSurverySolutionToProgram/"
solutiontoprogrammappingapiurl = "v1/solutions/importFromSolution"
# ; fetchprograminfoapiurl = v1/programs/list?page=1&limit=5&search=
fetchprograminfoapiurl="v1/admin/dbFind/programs"
fetchsolutiondetails = "v1/solutions/list?type="
frameworkcreationapi = "v1/frameworks/create"
solutionupdateapi = "v1/solutions/update/"
listofrolesapi = "v1/userRoles/list"
ferchsolutioncriteria = "v1/solutionDetails/criteria/"
projectuploadapi = "v1/project/templates/bulkCreate"
taskuploadapi = "v1/project/templateTasks/bulkCreate/"
projectsolutioncreationapi = "v1/solutions/create"
mapsolutiontoproject = "v1/project/templates/importProjectTemplate/"
fetchsolutiondoc = "v1/solutions/getDetails/"
programcreationurl = "v1/programs/create"
pdpmurl = "v1/userExtension/bulkUpload"
fetchlink = "v1/solutions/fetchLink/"
courseprogrammapping = "v1/solutions/create"
readcourseurl = "/api/content/v1/read/"
fetchsolutiondump = "v1/solutions/getDetails/"
fetchorgdetails = "api/org/v1/search"
fetchprojectlist = "api/private/mlprojects/v1/library/categories/projects"
dbfindapi_url = "v1/admin/dbFind/"
addingbasetemp = "v1/certificateBaseTemplates/createOrUpdate"
dbfindapi = "v1/admin/dbFind/certificateBaseTemplates"
addcertificatetemplate = "v1/certificateTemplates/createOrUpdate"
editsvgtemp = "v1/certificateTemplates/createSvg?baseTemplateId="
uploadcertificatetosvg = "v1/certificateTemplates/uploadTemplate/"
updatecertificatesolu = "v1/solutions/update/"
updateprojecttemplate = "v1/project/templates/update/"
fetchprofessionalRole = "entity-management/v1/entities/entityListBasedOnEntityType?entityType=professional_role"
fetchDetailsEntity = "entity-management/v1/entities/details/"
tenantFetch = "user/v1/tenant/read/"
dbfindapi_projectTemplate = "v1/admin/dbFind/projectTemplates"
certificatetypeof = {
    "onelogo-onesignature": "onelogo_onesign",
    "onelogo-twosignature": "onelogo_twosign",
    "twologo-onesignature": "twologo_onesign",
    "twologo-twosignature": "twologo_twosign"
}
jwtTokenSecret = ""
adminTokenHeaderName = ""
adminAccessToken = ""

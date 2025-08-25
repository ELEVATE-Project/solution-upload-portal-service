#Host for user Services
host = ""
userLoginHost =""
# Host URL for the elevate Samiksha service
internal_kong_ip = ""
elevateprojecthost = ""

# Host URL for the elevate entity service
elevateentityhost = ""

# JSON body for Keycloak API login
keyclockapibody = '{"email": "prajwal@tunerlabs.com","password": "Password1@"}'

# Email for login credentials
email = "prajwal@tunerlabs.com"

# Password for login credentials
identifier = ""
password = ""

keyclockapicontent_type = "application/x-www-form-urlencoded"
content_type = "application/json"
keyclockapiurl ="user/v1/admin/login"
 
origin= "app.shikshagraha.org"

userinfoapiurl = "user/v1/user/read"

# Endpoint for searching locations in the entity management service
searchforlocation = "entity-management/v1/entities/find"

# Endpoint for fetching program information
fetchprograminfoapiurl = "admin/dbFind/programs"

# Endpoint for fetching solution details
fetchsolutiondetails = "solutions/list?type="

# Endpoint for updating a solution
solutionupdateapi = "solutions/update/"

# Endpoint for uploading project templates in bulk
projectuploadapi = "project/templates/bulkCreate"


# Endpoint for uploading project template tasks in bulk
taskuploadapi = "project/templateTasks/bulkCreate/"


# Endpoint for creating a project solution
projectsolutioncreationapi = "solutions/create"


# Endpoint for mapping a solution to a project
mapsolutiontoproject = "project/templates/importProjectTemplate/"


# Endpoint for fetching solution details
fetchsolutiondoc = "solutions/getDetails/"


# Endpoint for creating a program
programcreationurl = "programs/create"

fetchDetailsEntity = "entity-management/v1/entities/details/"

# Endpoint for performing a database find operation for certificate base templates
dbfindapi = "admin/dbFind/certificateBaseTemplates"


# Endpoint for creating or updating a certificate template
addcertificatetemplate = "certificateTemplates/createOrUpdate"


# Endpoint for uploading a certificate template as SVG
uploadcertificatetosvg = "certificateTemplates/uploadTemplate/"


# Endpoint for editing an SVG template
editsvgtemp = "certificateTemplates/createSvg?baseTemplateId="


# Endpoint for updating a project template
updateprojecttemplate = "project/templates/update/"


# Endpoint for fetching a link related to a solution
fetchlink = "solutions/fetchLink/"


# Endpoint for reading course details
readcourseurl = "api/content/v1/read/"


# Endpoint for fetching organization details
fetchorgdetails = "api/org/v1/search"

tenantFetch = "user/v1/tenant/read/"

# Configuration for different logo-signature combinations (likely used in templates)
certificatetypeof = {
    "onelogo-onesignature": "onelogo_onesign",
    "onelogo-twosignature": "onelogo_twosign",
    "twologo-onesignature": "twologo_onesign",
    "twologo-twosignature": "twologo_twosign"
}

fetchprofessionalRole = "entity-management/v1/entities/entityListBasedOnEntityType?entityType=professional_role"

# Endpoint for searching locations in the entity management service
criteriauploadapiurl = "criteria/upload"
themeuploadapiurl = "frameworks/uploadThemes/"
solutioncreationapiurl = "observations/importFromFramework"
surveysolutioncreationapiurl = "surveys/createSolutionTemplate"
questionuploadapiurl = "questions/bulkCreate"
criteriarubricuploadapiurl = "solutions/uploadCriteriaRubricExpressions/"
themerubricuploadapiurl = "solutions/uploadThemesRubricExpressions/"
importsurveysolutiontemplateurl = "surveys/importSurveyTemplateToSolution/"
importsurveysolutiontoprogramurl = "surveys/mapSurveySolutionToProgram/"
solutiontoprogrammappingapiurl = "solutions/importFromSolution"
frameworkcreationapi = "frameworks/create"
listofrolesapi = "userRoles/list"
ferchsolutioncriteria = "solutionDetails/criteria/"
pdpmurl = "userExtension/bulkUpload"
courseprogrammapping = "solutions/create"
readcourseurl = "api/content/v1/read/"
fetchsolutiondump = "solutions/getDetails/"
fetchprojectlist = "api/private/mlprojects/v1/library/categories/projects"
dbfindapi_url = "admin/dbFind/"
addingbasetemp = "certificateBaseTemplates/createOrUpdate"

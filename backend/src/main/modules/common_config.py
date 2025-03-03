# Host URL for the elevate user service
elevateuserhost = "https://project-qa.elevate-apis.shikshalokam.org/"

# Host URL for the elevate project service
elevateprojecthost = "https://project-qa.elevate-apis.shikshalokam.org/project/v1/"

# Host URL for the elevate entity service
elevateentityhost = "https://project-qa.elevate-apis.shikshalokam.org/"

internal_kong_ip = "https://qa.elevate-apis.shikshalokam.org/survey/"

# JSON body for Keycloak API login
keyclockapibody =  '{"email": "Vishnu@tunerlabs.com","password": "Parayilla#2"}'

# Email for login credentials
email = "Vishnu@tunerlabs.com"

# Password for login credentials
password = "Parayilla#2"

# Internal access token used for internal API calls
internal_access_token = "Fqn0m0HQ0gXydRtBCg5l"

# Endpoint for user login API
userlogin = "user/v1/account/login"

# Authorization token for API requests (to be added)
authorization = "Bearer "  # ADD_YOUR_TOKEN_HERE

# Authorization token specifically for host API requests (to be added)
authorizationforhost = "Bearer " # ADD_YOUR_TOKEN_HERE

# Application name, typically used in headers or logs
appname = "diksha"

# Channel ID for identifying the request source
x_channel_id = "0125747659358699520"

# Content type for Keycloak API requests
keyclockapicontent_type = "application/x-www-form-urlencoded"

# Default content type for API requests
content_type = "application/json"

# Endpoint for fetching user information
userinfoapiurl = "profile/read"

criteriauploadapiurl = "v1/criteria/upload"
frameworkcreationapi = "v1/frameworks/create"
themeuploadapiurl = "v1/frameworks/uploadThemes/"
solutioncreationapiurl = "v1/observations/importFromFramework"
solutionupdateapiObs = "v1/solutions/update/"
questionuploadapiurl = "v1/questions/bulkCreate"
ferchsolutioncriteria = "v1/solutionDetails/criteria/"
criteriarubricuploadapiurl = "v1/solutions/uploadCriteriaRubricExpressions/"
themerubricuploadapiurl = "v1/solutions/uploadThemesRubricExpressions/"
fetchsolutiondocobs = "v1/solutions/getDetails/"
solutiontoprogrammappingapiurl = "v1/solutions/importFromSolution"
fetchlinkobs = "v1/solutions/fetchLink/"
# Endpoint for searching locations in the entity management service
searchforlocation = "entity-management/v1/entities/find"

# Endpoint for fetching program information
fetchprograminfoapiurl = "programs/list?page=1&limit=5&search="
fetchprograminfoapiurlobs= "v1/admin/dbFind/programs"

# Endpoint for fetching solution details
fetchsolutiondetails = "solutions/list?page=1&limit=100&search=&type=observation&subType"

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
programcreationurlobs = "v1/programs/create"

# Endpoint for performing a database find operation for certificate base templates
dbfindapi = "/admin/dbFind/certificateBaseTemplates"

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

# Endpoint for fetching URL
fetchPreSignedUrl = "cloud-services/files/preSignedUrls"

# Configuration for different logo-signature combinations (likely used in templates)
certificatetypeof = {
    "onelogo-onesignature": "onelogo_onesign",
    "onelogo-twosignature": "onelogo_twosign",
    "twologo-onesignature": "twologo_onesign",
    "twologo-twosignature": "twologo_twosign"
}

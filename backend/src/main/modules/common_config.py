keyclockapicontent_type = "application/x-www-form-urlencoded"
content_type = "application/json"

keyclockapiurl ="user/v1/admin/login"
 
origin= "app.shikshagraha.org"

# Endpoint for user login API
keyclockapiurl ="user/v1/admin/login"



# Endpoint for fetching user information
userinfoapiurl = "user/v1/user/read"

# Endpoint for searching locations in the entity management service
fetchDetailsEntity = "entity-management/v1/entities/details/"
searchforlocation = "entity-management/v1/entities/find"
criteriauploadapiurl = "v1/criteria/upload"
themeuploadapiurl = "v1/frameworks/uploadThemes/"
solutioncreationapiurl = "v1/observations/importFromFramework"
surveysolutioncreationapiurl = "v1/surveys/createSolutionTemplate"
questionuploadapiurl = "v1/questions/bulkCreate"
criteriarubricuploadapiurl = "v1/solutions/uploadCriteriaRubricExpressions/"
themerubricuploadapiurl = "v1/solutions/uploadThemesRubricExpressions/"
importsurveysolutiontemplateurl = "v1/surveys/importSurveyTemplateToSolution/"
importsurveysolutiontoprogramurl = "v1/surveys/mapSurveySolutionToProgram/"
solutiontoprogrammappingapiurl = "v1/solutions/importFromSolution"
fetchprograminfoapiurl="v1/admin/dbFind/programs"
dbfindapi_projectTemplate = "v1/admin/dbFind/projectTemplates"
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
fetchprojectlist = "v1/library/categories/projects"
dbfindapi_url = "v1/admin/dbFind/solutions"
addingbasetemp = "v1/certificateBaseTemplates/createOrUpdate"
dbfindapi = "v1/admin/dbFind/certificateBaseTemplates"
addcertificatetemplate = "v1/certificateTemplates/createOrUpdate"
editsvgtemp = "v1/certificateTemplates/createSvg?baseTemplateId="
uploadcertificatetosvg = "v1/certificateTemplates/uploadTemplate/"
tenantFetch = "user/v1/tenant/read/"
updatecertificatesolu = "v1/solutions/update/"
updateprojecttemplate = "v1/project/templates/update/"
fetchprofessionalRole = "entity-management/v1/entities/entityListBasedOnEntityType?entityType=professional_role"
certificatetypeof = {
    "onelogo-onesignature": "onelogo_onesign",
    "onelogo-twosignature": "onelogo_twosign",
    "twologo-onesignature": "twologo_onesign",
    "twologo-twosignature": "twologo_twosign"
}
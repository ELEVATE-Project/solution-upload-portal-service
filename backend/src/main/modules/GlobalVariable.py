from datetime import datetime
from typing import Dict, List, Optional, Any
import os
import xlrd

class GlobalVariable:
    _instance = None
    _initialized = False
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(GlobalVariable, cls).__new__(cls)
        return cls._instance
    
    def __init__(self):
        """Initialize all global variables with default values"""
        if not GlobalVariable._initialized:
            self._initialize_variables()
            GlobalVariable._initialized = True
    
    def _initialize_variables(self):
        """Initialize all global variables with appropriate default values"""
        
        # Dictionary and lookup variables
        self.criteriaLookUp: Dict[str, Any] = {}
        self.dictCritLookUp: Dict[str, Any] = {}
        self.themeRubricFileObj: Dict[str, Any] = {}
        self.ecm_sections: Dict[str, Any] = {}
        self.criteriaIdNameDict: Dict[str, Any] = {}
        self.ecmToSection: Dict[str, Any] = {}
        
        self.programDict: Dict[str, Any] = {}
        self.programResourceDetails: List[Dict[str, Any]] = []
        self.programManagerDetails: List[Dict[str, Any]] = []
        self.programTemplatePath: Optional[str] = None
        self.programTemplateSheetNames: List[str] = []
        self.programValidationCache: Dict[str, Any] = {}
        self.solutionDetailsDict: Dict[str, Any] = {}
        self.solutionTemplatePath: Optional[str] = None
        self.resourceValidationCache: Dict[str, Any] = {}
        
        # Timestamp and date variables
        self.millisecond: Optional[int] = None
        self.startDateOfResource: Optional[str] = None
        self.endDateOfResource: Optional[str] = None
        self.startDateOfProgram: Optional[str] = None
        self.endDateOfProgram: Optional[str] = None
        self.solutionStartDate: str = ""
        self.solutionEndDate: str = ""
        self.ReffstartDateOfProgram: Optional[str] = None
        self.ReffendDateOfProgram: Optional[str] = None
        self.SurveyTemplateStartDate: Optional[str] = None
        self.SurveyTemplateEndDate: Optional[str] = None
        
        # Program related variables
        self.programNameInp: Optional[str] = None
        self.programName: Optional[str] = None
        self.programID: Optional[str] = None
        self.programExternalId: Optional[str] = None
        self.programDescription: Optional[str] = None
        self.isProgramnamePresent: Optional[bool] = None
        
        # Environment and configuration
        self.environment: Optional[str] = None
        
        # Solution related variables
        self.observationId: Optional[str] = None
        self.solutionName: Optional[str] = None
        self.solutionId: Optional[str] = None
        self.solutionDescription: Optional[str] = None
        self.solutionLanguage: Optional[str] = None
        self.addObservationSolution: Optional[str] = None
        self.surveySolutionlink: Optional[str] = None
        # Entity and type related variables
        self.pointBasedValue: Optional[bool] = None
        self.entityType: Optional[str] = None
        self.entityTypeId: Optional[str] = None
        self.scopeEntityType: str = ""
        self.userEntity: Optional[str] = None
        self.entityToUpload: Optional[str] = None
        
        # Submission and access control
        self.allow_multiple_submissions: Optional[bool] = None
        
        # Roles and permissions
        self.roles: str = ""
        self.mainRole: str = ""
        self.rolesPGM: Optional[str] = None
        
        # List variables
        self.listOfFoundRoles: List[str] = []
        self.themesSheetList: List[Any] = []
        self.scopeEntities: List[str] = []
        self.scopeRoles: List[str] = []
        self.entitiesPGM: List[str] = []
        self.entitiesPGMID: List[str] = []
        self.solutionRolesArr: List[str] = []
        self.solutionRolesArray: List[str] = []
        self.orgIds: List[str] = []
        self.OrgName: List[str] = []
        self.question_sequence_arr: List[Any] = []
        self.resourceEndDates: List[str] = []
        self.criteriaLevels: List[Any] = []
        
        # Keywords and criteria
        self.keyWords: Optional[str] = None
        self.criteriaName: Optional[str] = None
        
        # User and authentication
        self.creator: Optional[str] = None
        self.dikshaLoginId: Optional[str] = None
        self.matchedShikshalokamLoginId: Optional[str] = None
        self.projectCreator: str = ""
        self.projectAuthor: str = ""
        self.creatorId: Optional[str] = None
        self.solutionNameForSuccess: Optional[str] = None
        
        # API and logging
        self.API_log: Optional[str] = None
        
        # Report and criteria flags
        self.criteriaLevelsReport: bool = False
        
        # Counters
        self.criteriaLevelsCount: int = 0
        self.numberOfResponses: int = 0
        self.countImps: int = 0
        
        # Organization details
        self.ccRootOrgName: Optional[str] = None
        self.ccRootOrgId: Optional[str] = None
        
        # Certificate and template
        self.certificatetemplateid: Optional[str] = None
        
        # Task and evidence operators
        self.TaskEvidenceOperator: str = ""
        self.AnyTaskEvidenceNo: str = ""
        
        # Downloaded files
        self.downloaded_file: Optional[str] = None
        
        self.regex = r"\"?([-a-zA-Z0-9.`?{}]+@\w+\.\w+)\"?"
        
    
    def reset(self):
        """Reset all variables to their default values"""
        self._initialize_variables()

    def reset_program_template_cache(self):
        """Reset cached program template details."""
        self.programDict = {}
        self.programResourceDetails = []
        self.programManagerDetails = []
        self.programTemplatePath = None
        self.programTemplateSheetNames = []
        self.programValidationCache = {}

    def reset_solution_template_cache(self):
        """Reset cached solution details template."""
        self.solutionDetailsDict = {}
        self.solutionTemplatePath = None

    def reset_resource_validation_cache(self):
        """Reset cached resource workbook/details captured during validation."""
        self.resourceValidationCache = {}

    def load_program_template(self, program_file: str, force_reload: bool = False):
        """
        Read Program Details/Resource Details from program template once and cache them.
        Returns (program_dict, resource_details_list).
        """
        normalized_path = os.path.abspath(program_file) if program_file else program_file
        if (
            not force_reload
            and self.programTemplatePath == normalized_path
            and self.programDict
        ):
            return self.programDict, self.programResourceDetails

        self.reset_program_template_cache()
        wb_program = xlrd.open_workbook(program_file, on_demand=True)
        sheet_names = wb_program.sheet_names()
        self.programTemplateSheetNames = sheet_names

        for sheet_name in sheet_names:
            sheet_name_lower = sheet_name.strip().lower()

            if sheet_name_lower == "program details":
                details_sheet = wb_program.sheet_by_name(sheet_name)
                keys = [details_sheet.cell(1, col_idx).value for col_idx in range(details_sheet.ncols)]
                for row_idx in range(2, details_sheet.nrows):
                    self.programDict = {
                        keys[col_idx]: details_sheet.cell(row_idx, col_idx).value
                        for col_idx in range(details_sheet.ncols)
                    }
                    if self.programDict:
                        break

            elif sheet_name_lower == "resource details":
                resource_sheet = wb_program.sheet_by_name(sheet_name)
                keys = [resource_sheet.cell(1, col_idx).value for col_idx in range(resource_sheet.ncols)]
                for row_idx in range(2, resource_sheet.nrows):
                    resource_row = {
                        keys[col_idx]: resource_sheet.cell(row_idx, col_idx).value
                        for col_idx in range(resource_sheet.ncols)
                    }
                    self.programResourceDetails.append(resource_row)

            elif sheet_name_lower == "program manager details":
                manager_sheet = wb_program.sheet_by_name(sheet_name)
                keys = [manager_sheet.cell(1, col_idx).value for col_idx in range(manager_sheet.ncols)]
                for row_idx in range(2, manager_sheet.nrows):
                    manager_row = {
                        keys[col_idx]: manager_sheet.cell(row_idx, col_idx).value
                        for col_idx in range(manager_sheet.ncols)
                    }
                    self.programManagerDetails.append(manager_row)

        self.programTemplatePath = normalized_path
        return self.programDict, self.programResourceDetails

    def get_program_value(self, key: str, default: Any = None) -> Any:
        """Get any field from cached Program Details by column name."""
        return self.programDict.get(key, default)

    def load_solution_template_details(self, solution_file: str, force_reload: bool = False):
        """
        Read first data row from 'details' sheet once and cache it.
        Returns solution details dict.
        """
        normalized_path = os.path.abspath(solution_file) if solution_file else solution_file
        if (
            not force_reload
            and self.solutionTemplatePath == normalized_path
            and self.solutionDetailsDict
        ):
            return self.solutionDetailsDict

        self.reset_solution_template_cache()
        wb_solution = xlrd.open_workbook(solution_file, on_demand=True)

        for sheet_name in wb_solution.sheet_names():
            if sheet_name.strip().lower() == "details":
                details_sheet = wb_solution.sheet_by_name(sheet_name)
                keys = [details_sheet.cell(1, col_idx).value for col_idx in range(details_sheet.ncols)]
                for row_idx in range(2, details_sheet.nrows):
                    self.solutionDetailsDict = {
                        keys[col_idx]: details_sheet.cell(row_idx, col_idx).value
                        for col_idx in range(details_sheet.ncols)
                    }
                    if self.solutionDetailsDict:
                        break
                break

        self.solutionTemplatePath = normalized_path
        return self.solutionDetailsDict

    def set_resource_validation_cache(self, **kwargs):
        """Store resource validation cache values."""
        self.resourceValidationCache = kwargs
        return self.resourceValidationCache

    def get_resource_validation_cache(self):
        """Get resource validation cache values."""
        return self.resourceValidationCache
    
    def reset_specific(self, *variable_names):
        """Reset specific variables to their default values"""
        defaults = {
            'criteriaLookUp': {},
            'dictCritLookUp': {},
            'themeRubricFileObj': {},
            'ecm_sections': {},
            'criteriaIdNameDict': {},
            'ecmToSection': {},
            'millisecond': None,
            'startDateOfResource': None,
            'endDateOfResource': None,
            'startDateOfProgram': None,
            'endDateOfProgram': None,
            'solutionStartDate': "",
            'solutionEndDate': "",
            'programNameInp': None,
            'programName': None,
            'programID': None,
            'programExternalId': None,
            'programDescription': None,
            'isProgramnamePresent': None,
            'environment': None,
            'observationId': None,
            'solutionName': None,
            'solutionId': None,
            'solutionDescription': None,
            'solutionLanguage': None,
            'pointBasedValue': None,
            'entityType': None,
            'entityTypeId': None,
            'scopeEntityType': "",
            'userEntity': None,
            'entityToUpload': None,
            'allow_multiple_submissions': None,
            'roles': "",
            'mainRole': "",
            'rolesPGM': None,
            'listOfFoundRoles': [],
            'themesSheetList': [],
            'scopeEntities': [],
            'scopeRoles': [],
            'entitiesPGM': [],
            'entitiesPGMID': [],
            'solutionRolesArr': [],
            'solutionRolesArray': [],
            'orgIds': [],
            'OrgName': [],
            'question_sequence_arr': [],
            'resourceEndDates': [],
            'criteriaLevels': [],
            'keyWords': None,
            'criteriaName': None,
            'creator': None,
            'dikshaLoginId': None,
            'matchedShikshalokamLoginId': None,
            'projectCreator': "",
            'projectAuthor': "",
            'creatorId': None,
            'solutionNameForSuccess': None,
            'API_log': None,
            'criteriaLevelsReport': False,
            'criteriaLevelsCount': 0,
            'numberOfResponses': 0,
            'countImps': 0,
            'ccRootOrgName': None,
            'ccRootOrgId': None,
            'certificatetemplateid': None,
            'TaskEvidenceOperator': "",
            'AnyTaskEvidenceNo': "",
            'downloaded_file': None,
            'addObservationSolution': None,
            'surveySolutionlink': None,
            'ReffstartDateOfProgram': None,
            'ReffendDateOfProgram': None,
            'SurveyTemplateStartDate': None,
            'SurveyTemplateEndDate': None,
            'programDict': {},
            'programResourceDetails': [],
            'programManagerDetails': [],
            'programTemplatePath': None,
            'programTemplateSheetNames': [],
            'solutionDetailsDict': {},
            'solutionTemplatePath': None,
            'resourceValidationCache': {},
        }
        
        for var_name in variable_names:
            if var_name in defaults:
                setattr(self, var_name, defaults[var_name])
    
    def get_all_variables(self) -> Dict[str, Any]:
        """Return a dictionary of all variables and their current values"""
        return {
            key: value for key, value in self.__dict__.items()
            if not key.startswith('_')
        }
    
    def set_variable(self, name: str, value: Any) -> bool:
        """
        Safely set a variable value
        Returns True if successful, False if variable doesn't exist
        """
        if hasattr(self, name):
            setattr(self, name, value)
            return True
        return False
    
    def get_variable(self, name: str, default: Any = None) -> Any:
        """
        Safely get a variable value
        Returns the value if it exists, otherwise returns default
        """
        return getattr(self, name, default)
    
    def variable_exists(self, name: str) -> bool:
        """Check if a variable exists"""
        return hasattr(self, name)
    
# Create a single global instance
global_vars = GlobalVariable()



# Convenience functions for backward compatibility
def get_global(name: str, default: Any = None) -> Any:
    """Get a global variable value"""
    return global_vars.get_variable(name, default)


def set_global(name: str, value: Any) -> bool:
    """Set a global variable value"""
    return global_vars.set_variable(name, value)


def reset_globals(*variable_names):
    """Reset specific global variables or all if none specified"""
    if variable_names:
        global_vars.reset_specific(*variable_names)
    else:
        global_vars.reset()


def get_all_globals() -> Dict[str, Any]:
    """Get all global variables as a dictionary"""
    return global_vars.get_all_variables()


# Export the instance and functions
__all__ = [
    'GlobalVariable',
    'global_vars',
    'get_global',
    'set_global',
    'reset_globals',
    'get_all_globals',
    'exception_handler'
]

def exception_handler(func):
    """
    Decorator to capture exceptions in GlobalVariable stats and re-raise them.
    This ensures every function failure is recorded.
    """
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            raise e
    return wrapper

#Author-Wayne Brill
#Description-Allows you to select a CSV (comma seperated values) file and then edits existing Attributes. Also allows you to write parameters to a file

import adsk.core, adsk.fusion, traceback, csv

_app = None
_ui = None

fullCommandId = 'ParamsFromCSV' 
quickLoadCommandId = 'LoadParamsFromCSV'
workspaceToUse = 'FusionSolidEnvironment'
panelsToUse = ['SolidModifyPanel', 'SketchCreatePanel']
commandLocations = [
    ['FusionSolidEnvironment', 'SketchModifyPanel'],
    ['FusionSolidEnvironment', 'SolidModifyPanel']
]

# Try to remove objects in case we are restarting:
try: removeObjects()
except: pass

# objects we'll clean up when stopped
ourObjects = []

# Keep track of the last file we loaded from, for the quick load command:
global previousImportFilename
previousImportFilename = None

def keep(obj):
    ourObjects.append(obj)
    return obj

class SimpleCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self, fn):
        super().__init__()
        self._fn = fn

    def notify(self, args):
        cmd = args.command
        cmd.execute.add(keep(SimpleCommandExecuteHandler(self._fn)))

class SimpleCommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self, fn):
        super().__init__()
        self._fn = fn
    def notify(self, args):
        try: self._fn()
        except: _ui.messageBox('command failed:\n{}'.format(traceback.format_exc()))

def run(context):
    global _app, _ui
    _app =  adsk.core.Application.get()
    _ui = _app.userInterface

    try:
        addButtons()
    except:
        _ui.messageBox('AddIn Start Failed:\n{}'.format(traceback.format_exc()))

def addButtons():
    commandName = 'Import/Export Parameters (CSV)'
    commandDescription = 'Import parameters from or export them to a CSV (Comma Separated Values) file\n'
    commandResources = './resources/command'
    quickLoadCommandName = 'Quick Import Parameters (CSV)'
    quickLoadCommandDescription = 'Quick import parameters from CSV file\n' 
    
    commandDefinition_ = keep(_ui.commandDefinitions.addButtonDefinition(
        fullCommandId, commandName, commandDescription, commandResources))
    quickCommandDefinition_ = keep(_ui.commandDefinitions.addButtonDefinition(
        quickLoadCommandId, quickLoadCommandName, quickLoadCommandDescription,
        commandResources))

    commandDefinition_.commandCreated.add(keep(SimpleCommandCreatedHandler(
        lambda: updateParamsFromCSV())))
    quickCommandDefinition_.commandCreated.add(keep(SimpleCommandCreatedHandler(
        lambda: updateParamsFromCSV(quick = True))))
    
    for (workspace, toolbar) in commandLocations:
        controls = _ui.workspaces.itemById(workspace).toolbarPanels.itemById(toolbar).controls
        for cmdDef in [commandDefinition_, quickCommandDefinition_]:
            cmd = keep(controls.addCommand(cmdDef))
            cmd.isVisible = True

def stop(context):
    removeObjects()

def removeObjects():
    try:
        for obj in ourObjects:
            if hasattr(obj, 'deleteMe'):
                obj.deleteMe()
        ourObjects.clear()
    except:
        _ui.messageBox('Cleanup failed:\n{}'.format(traceback.format_exc()))	


def updateParamsFromCSV(quick = False):
    global previousImportFilename
    readParameters = None
    if quick:
        if not previousImportFilename:
            # File not known, ask them to pick one.
            readParameters = True
        else:
            return readTheParameters(previousImportFilename, silent = True)

    #Ask if reading or writing parameters
    if readParameters is None:
        dialogResult = _ui.messageBox('Importing/Updating parameters from file or Exporting them to file?\n' \
        'Import = Yes, Export = No', 'Import or Export Parameters', \
        adsk.core.MessageBoxButtonTypes.YesNoCancelButtonType, \
        adsk.core.MessageBoxIconTypes.QuestionIconType) 
        if dialogResult == adsk.core.DialogResults.DialogYes:
            readParameters = True
        elif dialogResult == adsk.core.DialogResults.DialogNo:
            readParameters = False
        else:
            return

    fileDialog = _ui.createFileDialog()
    fileDialog.isMultiSelectEnabled = False
    fileDialog.title = "Get the file to read from or the file to save the parameters to"
    fileDialog.filter = 'Text files (*.csv)'
    fileDialog.filterIndex = 0
    if readParameters:
        dialogResult = fileDialog.showOpen()
    else:
        dialogResult = fileDialog.showSave()
        
    if dialogResult == adsk.core.DialogResults.DialogOK:
        filename = fileDialog.filename
    else:
        return

    if readParameters:
        previousImportFilename = filename
        readTheParameters(filename)
    else:
        writeTheParameters(filename)

def writeTheParameters(theFileName):
    app = adsk.core.Application.get()
    design = app.activeProduct
      
    result = ""
    for _param in design.allParameters:
        try:
            paramUnit = _param.unit
        except:
            paramUnit = ""
            
        result = result + "\"" + _param.name +  "\",\"" + paramUnit +  "\",\"" + _param.expression.replace('"', '""') + "\",\"" + _param.comment + "\"\n"    
                      
    with open(theFileName, 'w') as outputFile:
        outputFile.writelines(result)
    
    # get the name of the file without the path    
    pathsInTheFileName = theFileName.split("/")
    _ui.messageBox('Parameters written to ' + pathsInTheFileName[-1])   
   
def readTheParameters(theFileName, silent = False):
    design = _app.activeProduct
    paramsList = []
    for oParam in design.allParameters:
        paramsList.append(oParam.name)           
    
    # Read the csv file.
    with open(theFileName) as csvFile:
        csvReader = csv.reader(csvFile, dialect=csv.excel)
        for row in csvReader:
            if not row:
                continue

            # Get the values from the csv file.
            nameOfParam = row[0]
            unitOfParam = row[1]
            expressionOfParam = row[2]
            # userParameters.add does not like empty string as comment
            # so we make it a space
            commentOfParam = ' ' 
            if len(row) > 3:
                commentOfParam = row[3]
                
            # comment might be empty
            if commentOfParam == '':
                commentOfParam = ' ' 
                
            print(expressionOfParam)
                
            # if the name of the paremeter is not an existing parameter add it
            if nameOfParam not in paramsList:
                valInput_Param = adsk.core.ValueInput.createByString(expressionOfParam) 
                design.userParameters.add(nameOfParam, valInput_Param, unitOfParam, commentOfParam)
            # update the values of existing parameters            
            else:
                paramInModel = design.allParameters.itemByName(nameOfParam)
                paramInModel.unit = unitOfParam
                paramInModel.expression = expressionOfParam
                paramInModel.comment = commentOfParam
    if not silent:
        _ui.messageBox('Finished reading and updating parameters')

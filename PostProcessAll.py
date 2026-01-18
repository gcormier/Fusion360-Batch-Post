#Author-Tim Paterson
#Description-Post process all CAM setups, using the setup name as the output file name.

import adsk.core, adsk.fusion, adsk.cam, traceback, shutil, json, os, os.path, time, re, pathlib, enum, tempfile

# Version number of settings as saved in documents and settings file
# update this whenever settings content changes
version = 12

# Initial default values of settings
defaultSettings = {
    "version" : version,
    "ncProgram": "",
    "output" : "",
    "sequence" : True,
    "twoDigits" : False,
    "delFiles" : False,
    "delFolder" : False,
    "splitSetup" : False,
    "combineTool" : False,
    "combineSetups" : False,
    "fastZ" : False,
    "toolChange" : "M9 G30",
    "numericName" : False,
    "endCodes" : "M5 M9 M30",
    "onlySelected" : False,
    "skipFirstToolchange" : False,
    "appendOriginLocation" : True,
    # Groups are expanded or not
    "groupPersonal" : True,
    "groupPost" : False,
    "groupAdvanced" : False,
    "groupRename" : False,
    # Retry policy
    "initialDelay" : 0.2,
    "postRetries" : 3
}

# Constants
constCmdName = "Post Process All"
constCmdDefId = "PatersonTech_PostProcessAll"
constCAMWorkspaceId = "CAMEnvironment"
constCAMActionsPanelId = "CAMActionPanel"
constPostProcessControlId = "IronPostProcess"
constCAMProductId = "CAMProductType"
constAttrGroup = constCmdDefId
constAttrName = "settings"
constAttrCompressedName = "CompressedName"
constSettingsFileExt = ".settings"
constPostLoopDelay = 0.1
constBodyTmpFile = "gcodeBody"
constOpTmpFile = "8910"   # in case name must be numeric
constRapidZgcode = 'G00 Z{} (Changed from: "{}")\n'
constRapidXYgcode = 'G00 {} (Changed from: "{}")\n'
constFeedZgcode = 'G01 Z{} F{} (Changed from: "{}")\n'
constFeedXYgcode = 'G01 {} F{} (Changed from: "{}")\n'
constFeedXYZgcode = 'G01 {} Z{} F{} (Changed from: "{}")\n'
constAddFeedGcode = " F{} (Feed rate added)\n"
constMotionGcodeSet = {0,1,2,3,33,38,73,76,80,81,82,84,85,86,87,88,89}
constHomeGcodeSet = {28, 30}
constLineNumInc = 5
constNcProgramName = "PostProcessAll NC Program"

# Tool tip text
toolTip = (
    "Post process all setups into G-code for your machine.\n\n"
    "The name of the setup is used for the name of the output "
    "file adding the appropriate extension. A colon (':') in the name indicates "
    "the preceding portion is the name of a subfolder. Multiple "
    "colons can be used to nest subfolders. Spaces around colons "
    "are removed.\n\n"
    "Setups within a folder are optionally preceded by a "
    "sequence number. This identifies the order in which the "
    "setups appear. The sequence numbers for each folder begin "
    "with 1."
    )

# Global list to keep all event handlers in scope.
# This is only needed with Python.
handlers = []

# Global settingsMgr object
settingsMgr = None

def run(context):
    global settingsMgr
    ui = None
    try:
        settingsMgr = SettingsManager()
        app = adsk.core.Application.get()
        ui  = app.userInterface
        InitAddIn()

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def stop(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface

        # Clean up the UI.
        cmdDef = ui.commandDefinitions.itemById(constCmdDefId)
        if cmdDef:
            cmdDef.deleteMe()
            
        addinsPanel = ui.allToolbarPanels.itemById(constCAMActionsPanelId)
        cmdControl = addinsPanel.controls.itemById(constCmdDefId)
        if cmdControl:
            cmdControl.deleteMe()
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))	


class SettingsManager:
    def __init__(self):
        self.default = None
        self.path = None
        self.fMustSave = False
        self.inputs = None

    def GetSettings(self, docAttr):
        docSettings = None
        attr = docAttr.itemByName(constAttrGroup, constAttrName)
        if attr:
            try:
                docSettings = json.loads(attr.value)
                if docSettings["version"] == version:
                    return docSettings
            except Exception:
                pass
            
        # Document does not have valid settings, get defaults
        if not self.default:
            # Haven't read the settings file yet
            file = None
            try:
                file = open(self.GetPath())
                self.default = json.load(file)
                # never allow delFiles or delFolder to default to True
                self.default["delFiles"] = False
                self.default["delFolder"] = False
                if self.default["version"] != version:
                    self.UpdateSettings(defaultSettings, self.default)
            except Exception:
                self.default = dict(defaultSettings)
                self.fMustSave = True
            finally:
                if file:
                    file.close
        
        if not docSettings:
            docSettings = dict(self.default)
        else:
            self.UpdateSettings(self.default, docSettings)
        return docSettings

    def SaveDefault(self, docSettings):
        self.fMustSave = False
        self.default = dict(docSettings)
        # never allow delFiles or delFolder to default to True
        self.default["delFiles"] = False
        self.default["delFolder"] = False
        try:
            strSettings = json.dumps(docSettings)
            file = open(self.GetPath(), "w")
            file.write(strSettings)
            file.close
        except Exception:
            pass

    def SaveSettings(self, docAttr, docSettings):
        if self.fMustSave:
            self.SaveDefault(docSettings)
        docAttr.add(constAttrGroup, constAttrName, json.dumps(docSettings))
            
    def UpdateSettings(self, src, dst):
        if "homeEndsOp" in dst:
            if dst["homeEndsOp"] and not ("endCodes" in dst):
                dst["endCodes"] = "M5 M9 M30 G28 G30"
            del dst["homeEndsOp"]
        for item in src:
            if not (item in dst):
                dst[item] = src[item]
        dst["version"] = src["version"]

    def GetPath(self):
        if not self.path:
            pos = __file__.rfind(".")
            if pos == -1:
                pos = len(__file__)
            self.path = __file__[0:pos] + constSettingsFileExt
        return self.path


def InitAddIn():
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface

        # Create a button command definition.
        cmdDefs = ui.commandDefinitions
        cmdDef = cmdDefs.addButtonDefinition(constCmdDefId, constCmdName, toolTip, "resources/Command")
        
        # Connect to the commandCreated event.
        commandEventHandler = CommandEventHandler()
        cmdDef.commandCreated.add(commandEventHandler)
        handlers.append(commandEventHandler)
        
        # Get the Actions panel in the Manufacture workspace.
        workSpace = ui.workspaces.itemById(constCAMWorkspaceId)
        addInsPanel = workSpace.toolbarPanels.itemById(constCAMActionsPanelId)
        
        # Add the button right after the Post Process command.
        cmdControl = addInsPanel.controls.addCommand(cmdDef, constPostProcessControlId, False)
        cmdControl.isPromotedByDefault = True
        cmdControl.isPromoted = True

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def CountOutputFolderFiles(folder, limit, fileExt):
    cntFiles = 0
    cntNcFiles = 0
    for path, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith(fileExt):
                cntNcFiles += 1
            else:
                cntFiles += 1
        if cntFiles > limit:
            return "many files that are not G-code"
        if cntNcFiles > limit * 1.5:
            return "many more G-code files than are produced by this design"
    return None


def ExpandFileName(file):
    return os.path.expanduser(file).replace("\\", "/")


def CompressFileName(file):
    # normalize whacks 
    base = os.path.expanduser("~").replace("\\", "/")
    newFile = file.replace("\\", "/").removeprefix(base)
    if len(file) != len(newFile) and newFile[0] == "/":
        file = "~" + newFile
    return file


def AssignOutputFolder(parameters, folder):
    parameters.itemByName("nc_program_output_folder").value.value = folder
    result = parameters.itemByName("nc_program_output_folder").value.value
    if result != folder and folder[0:2] == "\\\\":
        parameters.itemByName("nc_program_output_folder").value.value = "\\\\" + folder    # double up leading "\"
    return None


def GetSetups(cam, settings, setups):
    if len(setups) == 0 or not settings["onlySelected"]:
        setups = []
        # move all setups into a list
        for setup in cam.setups:
            setups.append(setup)
    return setups


def GetNcProgram(cam, settings):
    for program in cam.ncPrograms:
        if program.name == settings["ncProgram"]:
            return program
    return cam.ncPrograms.item(0)


def RenameSetups(settings, setups, find, replace, isRegex):
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        doc = app.activeDocument
        cam = adsk.cam.CAM.cast(doc.products.itemByProductType(constCAMProductId))
        setups = GetSetups(cam, settings, setups)
        
        for setup in setups:
            if isRegex:
                newName = re.sub(find, replace, setup.name)
            else:
                if find == "":
                    # special case, prepend
                    newName = replace + setup.name
                else:
                    newName = setup.name.replace(find, replace)

            if setup.name != newName:
                setup.name = newName

        # Save settings in document attributes
        settingsMgr.SaveSettings(doc.attributes, settings)

    except:
        pass


# Event handler for the commandCreated event.
class CommandEventHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        try:
            eventArgs = adsk.core.CommandCreatedEventArgs.cast(args)
            cmd = eventArgs.command

            # Get document attributes that will set initial values
            app = adsk.core.Application.get()
            cam = adsk.cam.CAM.cast(app.activeDocument.products.itemByProductType(constCAMProductId))
            docSettings  = settingsMgr.GetSettings(app.activeDocument.attributes)

            # See if we're doing only selected setups
            selectedSetups = []
            for setup in cam.setups:
                if setup.isSelected:
                    selectedSetups.append(setup)

            # Get the NCProgram
            programs = cam.ncPrograms
            if programs.count == 0:
                ncInput = programs.createInput()
                ncInput.displayName = constNcProgramName
                program = programs.add(ncInput)
                program.postConfiguration = program.postConfiguration
                outputFolder = docSettings["output"]
                program.attributes.add(constAttrGroup, constAttrCompressedName, outputFolder)
                AssignOutputFolder(program.parameters, ExpandFileName(outputFolder))
                program.parameters.itemByName("nc_program_createInBrowser").value.value = True
            elif programs.count == 1:
                program = programs.item(0)
            else:
                haveProgram = False
                for program in programs:
                    if program.name == docSettings["ncProgram"]:
                        haveProgram = True
                        break
                if not haveProgram:
                    program = programs.item(0)              
            docSettings["ncProgram"] = program.name

            # Connect to the execute event.
            onExecute = CommandExecuteHandler(docSettings, selectedSetups)
            cmd.execute.add(onExecute)
            handlers.append(onExecute)

            # Add inputs that will appear in a dialog
            inputs = cmd.commandInputs

            # text box as a label for NC Program
            input = inputs.addTextBoxCommandInput("ncProgramLabel", 
                                                   "", 
                                                   "NC Program:",
                                                   1,
                                                   True)
            input.isFullWidth = True
            label = input

            input = inputs.addDropDownCommandInput("ncProgram", 
                                                   "NC Program",
                                                   adsk.core.DropDownStyles.TextListDropDownStyle)
            for listItem in programs:
                input.listItems.add(listItem.name, listItem.name == program.name)
            #input.isFullWidth = True
            input.tooltip = "NC Program to Use"
            input.tooltipDescription = (
                "Post processing will use the settings from the selected NC Program."
            )
            label.tooltip = input.tooltip
            label.tooltipDescription = input.tooltipDescription

            # check box to use only selected setups
            input = inputs.addBoolValueInput("onlySelected", 
                                             "Only selected setups", 
                                             True, 
                                             "", 
                                             docSettings["onlySelected"])
            input.tooltip = "Only Process Selected Setups"
            input.tooltipDescription = (
                "Only setups selected in the browser will be processed. Note "
                "that a selected setup will be highlighted, not simply activated. "
                "Selecting individual operations within a setup has no effect."
            )
            input.isEnabled = len(selectedSetups) != 0

            # check box to delete existing files
            input = inputs.addBoolValueInput("delFiles", 
                                             "Delete existing files", 
                                             True, 
                                             "", 
                                             docSettings["delFiles"])
            input.tooltip = "Delete Existing Files in Each Folder"
            input.tooltipDescription = (
                "Delete all files in each output folder before post processing. "
                "This will help prevent accumulation of G-code files which are "
                "no longer used."
                "<p>For example, you could decide to add sequence numbers after "
                "already post processing without them. If this option is not "
                "checked, you will have two of each file, a newer one with a "
                "sequence number and older one without. With this option checked, "
                "all previous files will be deleted so only current results will "
                "be present.</p>"
                "<p>This option will only delete the files in folders in which new "
                "G-code files are being written. If you change the name of a "
                "folder, for example, it will not be deleted.</p>")

            # check box to delete entire output folder
            input = inputs.addBoolValueInput("delFolder", 
                                             "Delete output folder", 
                                             True, 
                                             "", 
                                             docSettings["delFolder"] and docSettings["delFiles"])
            input.isEnabled = docSettings["delFiles"] # enable only if delete existing files
            input.tooltip = "Delete Entire Output Folder First"
            input.tooltipDescription = (
                "Delete the entire output folder before post processing. This "
                "deletes all files and subfolders regardless of whether or not "
                "new G-code files are written to a particular folder."
                "<p><b>WARNING!</b> Be absolutely sure the output folder is set "
                "correctly before selecting this option. Run the command once "
                "before setting this option and verify the results are in the "
                "correct folder. An incorrect setting of the output folder with "
                "this option selected could result in unintentionally wiping out "
                "a vast number of files.</p>")

            # check box to prepend sequence numbers
            input = inputs.addBoolValueInput("sequence", 
                                             "Prepend sequence number", 
                                             True, 
                                             "", 
                                             docSettings["sequence"])
            input.tooltip = "Add Sequence Numbers to Name"
            input.tooltipDescription = (
                "Begin each file name with a sequence number. The numbering "
                "represents the order that the setups appear in the browser tree. "
                "Each folder has its own sequence numbers starting with 1.")

            # check box to select 2-digit sequence numbers
            input = inputs.addBoolValueInput("twoDigits", 
                                             "Use 2-digit numbers", 
                                             True, 
                                             "", 
                                             docSettings["twoDigits"])
            input.isEnabled = docSettings["sequence"] # enable only if using sequence numbers
            input.tooltip = "Use 2-Digit Sequence Numbers"
            input.tooltipDescription = (
                "Sequence numbers 0 - 9 will have a leading zero added, becoming"
                '"01" to "09". This could be useful for formatting or sorting.')

            # "Personal Use" version
            # check box to split up setup into individual operations
            inputGroup = inputs.addGroupCommandInput("groupPersonal", "Personal Use")
            input = inputGroup.children.addBoolValueInput("splitSetup",
                                                          "Use individual operations",
                                                          True,
                                                          "",
                                                          docSettings["splitSetup"])
            input.tooltip = "Split Setup Into Individual Operations"
            input.tooltipDescription = (
                "Generate output for each operation individually. This is usually "
                "REQUIRED when using Fusion for Personal Use, because tool "
                "changes are not allowed. The individual operations will be "
                "grouped back together into the same file, eliminating this "
                "limitation. You will get an error if there is a tool change "
                "in a setup and this options is not selected.")

            # check box to combine operation that use the same tool
            input = inputGroup.children.addBoolValueInput("combineTool",
                                                          "Combine operations using same tool",
                                                          True,
                                                          "",
                                                          docSettings.get("combineTool", False))
            input.isEnabled = docSettings["splitSetup"] # enable only if using individual operations
            input.tooltip = "Combine Consecutive Operations That Use the Same Tool"
            input.tooltipDescription = (
                "If consecutive operations use the same tool, have Fusion generate "
                "their output together. This can optimize G-code for some routers. "
                "However, it will cause the logic that restores rapid moves to also "
                "treat it as one operation, which can have negative effects if the "
                "feed heights for the operations are different.")

            # check box to combine setups into one file with tool optimization
            input = inputGroup.children.addBoolValueInput("combineSetups",
                                                          "Combine setups (minimize tool changes)",
                                                          True,
                                                          "",
                                                          docSettings.get("combineSetups", False))
            input.isEnabled = docSettings["splitSetup"] # enable only if using individual operations
            input.tooltip = "Combine Multiple Setups Into One File"
            input.tooltipDescription = (
                "Combine all selected setups into a single output file, reordering "
                "operations by tool number to minimize tool changes. Operations using "
                "the same tool across different setups will be grouped together. "
                "Your post processor will output the correct WCS (G54, G55, etc.) "
                "for each operation based on the setup's WCS setting."
                "<p>For example, if Setup 1 uses T1, T2, T5 and Setup 2 uses T1, T2, T4, "
                "the output will run all T1 operations (from both setups), then T2, etc.</p>"
                "<p>Requires 'Use individual operations' to be enabled.</p>")

            # text box as a label for tool change command
            input = inputGroup.children.addTextBoxCommandInput("toolLabel", 
                                                               "", 
                                                               "G-code for tool change:",
                                                               1,
                                                               True)
            input.isFullWidth = True
            label = input

            # enter G-code for tool change
            input = inputGroup.children.addStringValueInput("toolChange", "", docSettings["toolChange"])
            input.isEnabled = docSettings["splitSetup"] # enable only if using individual operations
            input.isFullWidth = True
            input.tooltip = "G-code to Precede Tool Change"
            input.tooltipDescription = (
                "Allows inserting a line of code before tool changes. For example, "
                "you might want M5 (spindle stop), M9 (coolant stop), and/or G28 or G30 "
                "(return to home). The code will be placed on the line before the "
                "tool change. You can get mulitple lines by separating them with "
                "a colon (:)."
                "<p>If you want a line number, just put a dummy line number in front. "
                "If you use the colon to get multiple lines, only put the dummy line "
                "number on the first line. For example, <b>N10 M9:G30</b> will give "
                "you two lines, both with properly sequenced line numbers.</p>"
            )
            label.tooltip = input.tooltip
            label.tooltipDescription = input.tooltipDescription
           
            # text box as a label for operation end commands
            input = inputGroup.children.addTextBoxCommandInput("endLabel", 
                                                               "", 
                                                               "G-codes that mark ending sequence:",
                                                               1,
                                                               True)
            input.isFullWidth = True
            label = input

            # enter G-codes for end of operation
            input = inputGroup.children.addStringValueInput("endCodes", "", docSettings["endCodes"])
            input.isEnabled = docSettings["splitSetup"] # enable only if using individual operations
            input.isFullWidth = True
            input.tooltip = "G-codes That Mark the Ending Sequence"
            input.tooltipDescription = (
                "To combine operations generated individually, the ending sequence "
                "(which should only appear once) must be found. This entry is the "
                "list of G-codes that start this ending sequence. For example, M30 "
                "(end program) would normally be here, but it may not be the first "
                "G-code of the ending sequence. M5 (spindle stop), M9 (coolant "
                "stop) and G28/G30 (move home) are also candidates, but you should "
                "look at the code from your post processor to determine what "
                "will work in your case. Any one of the G-codes you enter here "
                "will mark the start of ending sequence."
            )
            label.tooltip = input.tooltip
            label.tooltipDescription = input.tooltipDescription
           
            # check box to enable restoring rapid moves
            input = inputGroup.children.addBoolValueInput("fastZ",
                                                          "Restore rapid moves",
                                                          True,
                                                          "",
                                                          docSettings["fastZ"])
            input.isEnabled = docSettings["splitSetup"] # enable only if using individual operations
            input.tooltip = "Restore Rapid Moves (Experimental)"
            input.tooltipDescription = (
                "Replace appropriate moves at feed rate with rapid (G0) moves. "
                "In Fusion for Personal Use, moves that could be rapid are "
                "now limited to the current feed rate. When this option is selected, "
                "the G-code will be analyzed to find moves at or above the feed "
                "height and replace them with rapid moves."
                "<p><b>WARNING!<b> This option should be used with caution. "
                "Review the G-code to verify it is correct. Comments have been "
                "added to indicate the changes.")

            # check box to skip first toolchange
            input = inputGroup.children.addBoolValueInput("skipFirstToolchange",
                                                          "Skip first toolchange",
                                                          True,
                                                          "",
                                                          docSettings["skipFirstToolchange"])
            input.isEnabled = docSettings["splitSetup"] # enable only if using individual operations
            input.tooltip = "Skip First Tool Change"
            input.tooltipDescription = (
                "Skip the first tool change when post processing operations. "
                "This is useful when the first tool is already loaded in the "
                "spindle and you don't want to generate unnecessary tool change "
                "commands at the beginning of the program.")

            # check box to append origin location to filename
            input = inputGroup.children.addBoolValueInput("appendOriginLocation",
                                                          "Append origin location to filename",
                                                          True,
                                                          "",
                                                          docSettings.get("appendOriginLocation", True))
            input.tooltip = "Add Origin Location to Filename"
            input.tooltipDescription = (
                "Append the WCS origin location to the filename based on the stock point setting. "
                "For example, if the origin is at the back-right corner on top of the stock, "
                "'-BR-TOP' will be added to the filename. Corner positions are: FL (Front-Left), "
                "FR (Front-Right), BL (Back-Left), BR (Back-Right). If the origin is at the center "
                "in XY, only the Z position will be added (e.g., '-TOP' or '-BOT'). If the origin "
                "is fully custom, nothing will be appended.")
           
            inputGroup.isExpanded = docSettings["groupPersonal"]

            # Rename
            inputGroup = inputs.addGroupCommandInput("groupRename", "Rename Setups")

            # check box to use regular expressions
            input = inputGroup.children.addBoolValueInput("regex",
                                                          "Use Python regular expressions",
                                                          True,
                                                          "",
                                                          False)
            input.tooltip = "Search With Regular Expressions"
            input.tooltipDescription = (
                "Treat the search string as a Python regular expression (regex). "
                "This is extremely flexible but also very technical. Refer to "
                "Python documentation for details."
                "<p>One example is to put $ in the search box. This special "
                "symbol searches for the end of the setup name. Then the replacement "
                "string will be appended to the existing name."
            )

            # text box as a label for search field
            input = inputGroup.children.addTextBoxCommandInput("searchLabel", 
                                                               "", 
                                                               "Search for this string:",
                                                               1,
                                                               True)
            input.isFullWidth = True
            label = input

            # Find
            input = inputGroup.children.addStringValueInput("findString", "")
            input.isFullWidth = True
            input.tooltip = "String to find in setup name"
            input.tooltipDescription = (
                "Replace all occurences of this string with the replacement string. "
                "If this is left blank, the replacement string will be prepended to "
                "each setup name."
            )
            label.tooltip = input.tooltip
            label.tooltipDescription = input.tooltipDescription

            # text box as a label for replace field
            input = inputGroup.children.addTextBoxCommandInput("replaceLabel", 
                                                               "", 
                                                               "Replace with this string:",
                                                               1,
                                                               True)
            input.isFullWidth = True
            label = input

            # Replace
            input = inputGroup.children.addStringValueInput("replaceString", "")
            input.isFullWidth = True
            input.tooltip = "String to use as replacement"
            input.tooltipDescription = (
                "Replace all occurences of the Find string with this string."
            )
            label.tooltip = input.tooltip
            label.tooltipDescription = input.tooltipDescription

            # button to execute search & replace
            input = inputGroup.children.addBoolValueInput("replace", "Search and replace", False)
            input.resourceFolder = "resources/Rename"
            input.tooltip = "Execute search and replace"
            input.tooltipDescription = (
                "Search for all strings matching the Find box and replace them "
                "with the string in the Replace box.")
            inputGroup.isExpanded = docSettings["groupRename"]

            # Advanced -- retry settings
            inputGroup = inputs.addGroupCommandInput("groupAdvanced", "Advanced")
            # Time delay
            input = inputGroup.children.addFloatSpinnerCommandInput("initialDelay", 
                "Initial time allowance", "s", 0.1, 1.0, 0.1, docSettings["initialDelay"])
            input.tooltip = "Initial Time to Post Process an Operation"
            input.tooltipDescription = (
                "Initial delay to wait for post processor. Doubled for each retry.")
            # Retry count
            input = inputGroup.children.addIntegerSpinnerCommandInput("postRetries", 
                "Number of retries", 1, 9, 1, docSettings["postRetries"])
            input.tooltip = "Number of Retries"
            input.tooltipDescription = (
                "Retries if post processing failed. Time delay is doubled each retry.")
            inputGroup.isExpanded = docSettings["groupAdvanced"]
            
            # post processor
            inputGroup = inputs.addGroupCommandInput("groupPost", "Post Processor")
            inputGroup.isExpanded = docSettings["groupPost"]

            # Numeric name required?
            input = inputGroup.children.addBoolValueInput("numericName",
                                                          "Name must be numeric",
                                                          True,
                                                          "",
                                                          docSettings["numericName"])
            input.tooltip = "Output File Name Must Be Numeric"
            input.tooltipDescription = (
                "The name of the setup will not be used in the file name, "
                "only sequence numbers. The option to prepend sequence numbers "
                "will have no effect.")

            # button to save default settings
            input = inputs.addBoolValueInput("save", "Save as default", False)
            input.resourceFolder = "resources/Save"
            input.tooltip = "Save These Settings as System Default"
            input.tooltipDescription = (
                "Save these settings to use as the default for each new design.")

            # text box for error messages
            input = inputs.addTextBoxCommandInput("error", "", "", 3, True)
            input.isFullWidth = True
            input.isVisible = False

            # Connect to the inputChanged event.
            onInputChanged = CommandInputChangedHandler(docSettings, selectedSetups)
            cmd.inputChanged.add(onInputChanged)
            handlers.append(onInputChanged)

            # Connect to the validateInputs event.
            onValidateInputs = CommandValidateInputsHandler()
            cmd.validateInputs.add(onValidateInputs)
            handlers.append(onValidateInputs)
        except:
            ui = app.userInterface
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

# Event handler for the inputChanged event.
class CommandInputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self, docSettings, selectedSetups):
        self.docSettings = docSettings
        self.selectedSetups = selectedSetups
        super().__init__()

    def notify(self, args):
        app = adsk.core.Application.get()
        ui  = app.userInterface
        try:
            eventArgs = adsk.core.InputChangedEventArgs.cast(args)
            cmd = eventArgs.input.parentCommand
            inputs = eventArgs.inputs

            doc = app.activeDocument
            cam = adsk.cam.CAM.cast(doc.products.itemByProductType(constCAMProductId))

            # See if button clicked
            input = eventArgs.input
            if input.id == "save":
                settingsMgr.SaveDefault(self.docSettings)

            elif input.id == "replace":
                cmd.doExecute(False)    # do it in execute handler for Undo
                return

            elif input.id in self.docSettings:
                if input.objectType == adsk.core.GroupCommandInput.classType():
                    self.docSettings[input.id] = input.isExpanded
                elif input.objectType == adsk.core.DropDownCommandInput.classType():
                    self.docSettings[input.id] = input.selectedItem.name
                else:
                    self.docSettings[input.id] = input.value

            # Enable twoDigits only if sequence is true
            if input.id == "sequence":
                inputs.itemById("twoDigits").isEnabled = input.value

            # Enable delFolder only if delFiles is true
            if input.id == "delFiles":
                item = inputs.itemById("delFolder")
                item.value = input.value and item.value
                item.isEnabled = input.value

            # Options for splitSetup
            if input.id == "splitSetup":
                inputs.itemById("combineTool").isEnabled = input.value
                inputs.itemById("toolChange").isEnabled = input.value
                inputs.itemById("toolLabel").isEnabled = input.value
                inputs.itemById("endCodes").isEnabled = input.value
                inputs.itemById("endLabel").isEnabled = input.value
                inputs.itemById("fastZ").isEnabled = input.value
                inputs.itemById("skipFirstToolchange").isEnabled = input.value
                inputs.itemById("combineSetups").isEnabled = input.value
                # If splitSetup is disabled, also disable combineSetups
                if not input.value:
                    inputs.itemById("combineSetups").value = False
                    self.docSettings["combineSetups"] = False

            # combineSetups requires splitSetup
            if input.id == "combineSetups":
                if input.value and not inputs.itemById("splitSetup").value:
                    # Auto-enable splitSetup when combineSetups is checked
                    inputs.itemById("splitSetup").value = True
                    self.docSettings["splitSetup"] = True
                    # Enable the splitSetup dependent controls
                    inputs.itemById("toolChange").isEnabled = True
                    inputs.itemById("toolLabel").isEnabled = True
                    inputs.itemById("endCodes").isEnabled = True
                    inputs.itemById("endLabel").isEnabled = True
                    inputs.itemById("fastZ").isEnabled = True
                    inputs.itemById("skipFirstToolchange").isEnabled = True
                    inputs.itemById("combineSetups").isEnabled = True

        except:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


# Event handler for the validateInputs event.
class CommandValidateInputsHandler(adsk.core.ValidateInputsEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        app = adsk.core.Application.get()
        ui  = app.userInterface

        # No validation currently performed. Skeleton code retained.
        try:
            eventArgs = adsk.core.ValidateInputsEventArgs.cast(args)
            inputs = eventArgs.firingEvent.sender.commandInputs

        except:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


# Event handler for the execute event.
class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self, docSettings, selectedSetups):
        self.docSettings = docSettings
        self.selectedSetups = selectedSetups
        super().__init__()

    def notify(self, args):
        eventArgs = adsk.core.CommandEventArgs.cast(args)
        cmd = eventArgs.command
        inputs = cmd.commandInputs

        # Code to react to the event.
        button = inputs.itemById("replace")
        if button.value:
            RenameSetups(self.docSettings, 
                        self.selectedSetups, 
                        inputs.itemById("findString").value, 
                        inputs.itemById("replaceString").value,
                        inputs.itemById("regex").value)
            button.value = False
        else:
            PerformPostProcess(self.docSettings, self.selectedSetups)


def GetOriginLocationSuffix(setup, debugComments=None):
    """
    Analyze the setup's WCS origin relative to stock and return a filename suffix.
    Returns a string like "-FL-TOP" (Front-Left, Top), "-BR-BOT" (Back-Right, Bottom),
    "-TOP", "-BOT", or empty string if custom location.
    debugComments is a list that will be populated with debug info if provided.
    """
    try:
        if debugComments is not None:
            debugComments.append("(DEBUG: Analyzing origin location for setup: {})\n".format(setup.name))
            # List all attributes of setup for debugging
            attrs = [attr for attr in dir(setup) if not attr.startswith('_')]
            debugComments.append("(DEBUG: Setup attributes: {})\n".format(", ".join(attrs[:20])))  # first 20
            if len(attrs) > 20:
                debugComments.append("(DEBUG: Setup attributes (cont): {})\n".format(", ".join(attrs[20:40])))
            if len(attrs) > 40:
                debugComments.append("(DEBUG: Setup attributes (cont): {})\n".format(", ".join(attrs[40:])))
        
        # Try to access the WCS (Work Coordinate System)
        # The stock offset point is typically in setup.parameters
        parameters = setup.parameters
        if debugComments is not None:
            debugComments.append("(DEBUG: setup.parameters exists = {})\n".format(parameters is not None))
        
        if not parameters:
            return ""
            
        # Look for stock-related parameters
        stockPoint = None
        
        # Try common parameter names for stock point settings
        paramNames = ['wcs_origin_boxPoint', 'job_stockPointX', 'job_stockPointY', 'job_stockPointZ',
                      'wcs_stock_point', 'stockPoint']
        
        if debugComments is not None:
            debugComments.append("(DEBUG: Looking for stock point parameters...)\n")
            
        for paramName in paramNames:
            try:
                param = parameters.itemByName(paramName)
                if param:
                    if debugComments is not None:
                        debugComments.append("(DEBUG: Found parameter: {} = {})\n".format(paramName, param.value.value))
            except:
                pass
        
        # The actual parameter is likely 'wcs_origin_boxPoint' which is an enum
        try:
            boxPointParam = parameters.itemByName('wcs_origin_boxPoint')
            if boxPointParam:
                stockPoint = boxPointParam.value.value
                if debugComments is not None:
                    debugComments.append("(DEBUG: wcs_origin_boxPoint value = {})\n".format(stockPoint))
                    debugComments.append("(DEBUG: wcs_origin_boxPoint type = {})\n".format(type(stockPoint).__name__))
        except Exception as e:
            if debugComments is not None:
                debugComments.append("(DEBUG: Error accessing wcs_origin_boxPoint: {})\n".format(str(e)))
        
        if not stockPoint:
            if debugComments is not None:
                debugComments.append("(DEBUG: stockPoint is None, trying alternative approach)\n")
                # Try to list all parameters
                try:
                    allParams = []
                    for i in range(min(parameters.count, 30)):  # first 30 params
                        param = parameters.item(i)
                        allParams.append(param.name)
                    debugComments.append("(DEBUG: Available parameters: {})\n".format(", ".join(allParams)))
                except:
                    pass
            return ""
        
        # The stockPoint value is a string like "top 4" or "bottom 6"
        # Format appears to be: "[top|bottom] [0-9]"
        # Where the number represents the XY corner position
        
        if debugComments is not None:
            debugComments.append("(DEBUG: stockPoint string value = '{}', type = {})\n".format(stockPoint, type(stockPoint).__name__))
        
        # Parse the string to extract Z position and corner number
        stockPointStr = str(stockPoint).lower().strip()
        parts = stockPointStr.split()
        
        if debugComments is not None:
            debugComments.append("(DEBUG: Parsed parts = {})\n".format(parts))
        
        if len(parts) < 2:
            if debugComments is not None:
                debugComments.append("(DEBUG: Could not parse stockPoint string)\n")
            return ""
        
        zPos = parts[0]  # "top" or "bottom"
        try:
            cornerNum = int(parts[1])  # 0-9
        except:
            if debugComments is not None:
                debugComments.append("(DEBUG: Could not parse corner number)\n")
            return ""
        
        if debugComments is not None:
            debugComments.append("(DEBUG: Z position = '{}', corner number = {})\n".format(zPos, cornerNum))
        
        # Map corner numbers to XY positions
        # 0 = Center, 1 = Front-Left, 2 = Front-Right, 3 = Back-Left, 4 = Back-Right
        cornerMap = {
            0: "",      # Center (no XY suffix)
            1: "FL",    # Front-Left
            2: "FR",    # Front-Right
            3: "BL",    # Back-Left
            4: "BR"     # Back-Right
        }
        
        xyPos = cornerMap.get(cornerNum, None)
        
        if xyPos is None:
            if debugComments is not None:
                debugComments.append("(DEBUG: Unknown corner number: {})\n".format(cornerNum))
            return ""
        
        # Build the suffix
        suffix = ""
        if zPos == "top":
            if xyPos:
                suffix = "-" + xyPos + "-TOP"
            else:
                suffix = "-TOP"
        elif zPos == "bottom":
            if xyPos:
                suffix = "-" + xyPos + "-BOT"
            else:
                suffix = "-BOT"
        
        if debugComments is not None:
            debugComments.append("(DEBUG: Final suffix = '{}')\n".format(suffix))
        
        return suffix
        
    except Exception as e:
        # If any error occurs, log it and return empty string
        if debugComments is not None:
            debugComments.append("(DEBUG: ERROR in GetOriginLocationSuffix: {})\n".format(str(e)))
        return ""


def PerformPostProcess(docSettings, setups):
    ui = None
    progress = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        doc = app.activeDocument
        cam = adsk.cam.CAM.cast(doc.products.itemByProductType(constCAMProductId))

        cntFiles = 0
        cntSkipped = 0
        lstSkipped = ""

        program = GetNcProgram(cam, docSettings);
        parameters = program.parameters
        setups = GetSetups(cam, docSettings, setups)

        # normalize output folder for this user
        # "\" is converted to "/"
        outputFolder = parameters.itemByName("nc_program_output_folder").value.value.replace("\\", "/")
        # keep leading "\\" for file share
        if outputFolder[0:2] == "//":
            outputFolder = "\\\\" + outputFolder[2:]
        try:
            pathlib.Path(outputFolder).mkdir(exist_ok=True)
        except Exception as exc:
            # see if we can map it to folder with compressed user
            compressedName = program.attributes.itemByName(constAttrGroup, constAttrCompressedName).value
            if compressedName[0] == "~" and compressedName[1:] == outputFolder[-(len(compressedName) - 1):]:
                # yes, it matches
                outputFolder = ExpandFileName(compressedName)

        compressedName = CompressFileName(outputFolder)
        program.attributes.add(constAttrGroup, constAttrCompressedName, compressedName)
        docSettings["output"] = compressedName

        # Save settings in document attributes
        settingsMgr.SaveSettings(doc.attributes, docSettings)

        if len(setups) != 0 and cam.allOperations.count != 0:
            # make sure we're not going to delete too much
            if not docSettings["delFiles"]:
                docSettings["delFolder"] = False

            if docSettings["delFolder"]:
                fileExt = parameters.itemByName("nc_program_nc_extension").value.value
                strMsg = CountOutputFolderFiles(outputFolder, len(setups), fileExt)
                if strMsg:
                    docSettings["delFolder"] = False
                    strMsg = (
                        "The output folder contains {}. "
                        "It will not be deleted. You may wish to make sure you selected "
                        "the correct folder. If you want the folder deleted, you must "
                        "do it manually."
                        ).format(strMsg)
                    res = ui.messageBox(strMsg, 
                                        constCmdName,
                                        adsk.core.MessageBoxButtonTypes.OKCancelButtonType,
                                        adsk.core.MessageBoxIconTypes.WarningIconType)
                    if res == adsk.core.DialogResults.DialogCancel:
                        return  # abort!

            if docSettings["delFolder"]:
                try:
                    shutil.rmtree(outputFolder, True)
                except:
                    pass #ignore errors

            progress = ui.createProgressDialog()
            progress.isCancelButtonShown = True
            progressMsg = "{} files written to " + outputFolder
            progress.show("Post Processing...", "", 0, len(setups))
            progress.progressValue = 1 # try to get it to display
            progress.progressValue = 0

            # Check if we should combine setups into one file
            if docSettings.get("combineSetups", False) and len(setups) > 1:
                # Use combined processing mode
                progress.message = "Combining setups..."
                status = PostProcessCombinedSetups(setups, outputFolder, docSettings, program, progress)
                if status == None:
                    cntFiles = 1
                else:
                    cntSkipped = len(setups)
                    lstSkipped = status
                progress.hide()
                # restore program output folder
                AssignOutputFolder(parameters, outputFolder)
            else:
                # Normal per-setup processing
                cntSetups = 0
                seqDict = dict()

                # We pass through all setups even if only some are selected
                # so numbering scheme doesn't change.
                for setup in cam.setups:
                    if progress.wasCancelled:
                        break
                    if not setup.isSuppressed and setup.allOperations.count != 0:
                        nameList = setup.name.split(':')    # folder separator
                        setupFolder = outputFolder
                        cnt = len(nameList) - 1
                        i = 0
                        while i < cnt:
                            setupFolder += "/" + nameList[i].strip()
                            i += 1
                    
                        # keep a separate sequence number for each folder
                        if setupFolder in seqDict:
                            seqDict[setupFolder] += 1
                            # skip if we're not actually including this setup
                            if setup not in setups:
                                continue
                        else:
                            # first file for this folder
                            seqDict[setupFolder] = 1
                            # skip if we're not actually including this setup
                            if setup not in setups:
                                continue

                            if (docSettings["delFiles"]):
                                # delete all the files in the folder
                                try:
                                    for entry in os.scandir(setupFolder):
                                        if entry.is_file():
                                            try:
                                                os.remove(entry.path)
                                            except:
                                                pass #ignore errors
                                except:
                                    pass #ignore errors

                        # prepend sequence number if enabled
                        fname = nameList[i].strip()
                        if docSettings["sequence"] or docSettings["numericName"]:
                            seq = seqDict[setupFolder]
                            seqStr = str(seq)
                            if docSettings["twoDigits"] and seq < 10:
                                seqStr = "0" + seqStr
                            if docSettings["numericName"]:
                                fname = seqStr
                            else:
                                fname = seqStr + ' ' + fname

                        # append origin location suffix based on WCS relative to stock
                        if docSettings.get("appendOriginLocation", True):
                            # debugComments = []  # Set to [] to enable debug output in G-code files
                            originSuffix = GetOriginLocationSuffix(setup, None)
                            if originSuffix:
                                fname = fname + originSuffix

                        # append NOFIRSTTOOL if skipFirstToolchange is enabled
                        if docSettings["skipFirstToolchange"] and docSettings["splitSetup"]:
                            fname = fname + "-NOFIRSTTOOL"

                        # post the file
                        status = PostProcessSetup(fname, setup, setupFolder, docSettings, program, None)
                        if status == None:
                            cntFiles += 1
                        else:
                            cntSkipped += 1
                            lstSkipped += "\nFailed on setup " + setup.name + ": " + status
                        
                    cntSetups += 1
                    progress.message = progressMsg.format(cntFiles)
                    progress.progressValue = cntSetups

                progress.hide()
                # restore program output folder
                AssignOutputFolder(parameters, outputFolder)

        # done with setups, report results
        if cntSkipped != 0:
            ui.messageBox("{} files were written. {} Setups were skipped due to error:{}".format(cntFiles, cntSkipped, lstSkipped), 
                constCmdName, 
                adsk.core.MessageBoxButtonTypes.OKButtonType,
                adsk.core.MessageBoxIconTypes.WarningIconType)
            
        elif cntFiles == 0:
            ui.messageBox('No CAM operations posted', 
                constCmdName, 
                adsk.core.MessageBoxButtonTypes.OKButtonType,
                adsk.core.MessageBoxIconTypes.WarningIconType)
            

    except:
        if progress:
            progress.hide()
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def PostProcessCombinedSetups(setups, outputFolder, docSettings, program, progress):
    """
    Combine multiple setups into a single output file, reordering operations
    by tool number to minimize tool changes. Operations using the same tool
    across different setups will be grouped together.
    
    Returns None on success, or an error message string on failure.
    """
    ui = None
    fileHead = None
    fileBody = None
    fileOp = None
    retVal = "Fusion reported an exception"

    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        doc = app.activeDocument
        cam = adsk.cam.CAM.cast(doc.products.itemByProductType(constCAMProductId))
        parameters = program.parameters

        # Generate filename from first setup's name
        firstSetup = setups[0]
        nameList = firstSetup.name.split(':')
        fname = nameList[-1].strip() + "-COMBINED"
        
        # Verify file name is valid by creating it now
        fileExt = parameters.itemByName("nc_program_nc_extension").value.value
        path = outputFolder + "/" + fname + fileExt
        try:
            pathlib.Path(outputFolder).mkdir(parents=True, exist_ok=True)
            fileHead = open(path, "w")
        except Exception as exc:
            return "Unable to create output file '" + path + "'. Make sure the setup name is valid as a file name."
        
        # Make sure toolpaths are valid for all setups
        for setup in setups:
            if not cam.checkToolpath(setup):
                genStat = cam.generateToolpath(setup)
                while not genStat.isGenerationCompleted:
                    time.sleep(.1)

        # Collect all operations from all setups and group by tool number
        # Each entry: (tool_number, setup, operation)
        allOps = []
        for setup in setups:
            for i in range(setup.allOperations.count):
                op = setup.allOperations.item(i)
                if op.isSuppressed:
                    continue
                # Get tool number (Manual NC operations don't have tools)
                if op.hasToolpath:
                    toolNum = int(op.tool.parameters.itemByName("tool_number").value.value)
                else:
                    # Manual NC - assign tool number 0 or find adjacent operation's tool
                    toolNum = 0
                allOps.append((toolNum, setup, op))
        
        if len(allOps) == 0:
            fileHead.close()
            os.remove(path)
            return "No operations found in selected setups"

        # Sort operations by tool number, maintaining order within same tool
        # This groups all operations using the same tool together
        allOps.sort(key=lambda x: x[0])

        # Build list of operations grouped by tool for processing
        # We need to process in groups where each group can be sent to postProcess
        opGroups = []  # List of (toolNum, [(setup, op), ...])
        currentTool = None
        currentGroup = []
        
        for toolNum, setup, op in allOps:
            if currentTool != toolNum:
                if currentGroup:
                    opGroups.append((currentTool, currentGroup))
                currentTool = toolNum
                currentGroup = [(setup, op)]
            else:
                currentGroup.append((setup, op))
        
        if currentGroup:
            opGroups.append((currentTool, currentGroup))

        # Set up temporary output location
        opName = constOpTmpFile
        opFolder = tempfile.gettempdir().replace("\\", "/")
        opPath = opFolder + "/" + opName + fileExt
        
        parameters.itemByName("nc_program_openInEditor").value.value = False
        AssignOutputFolder(parameters, opFolder)
        parameters.itemByName("nc_program_filename").value.value = opName
        parameters.itemByName("nc_program_name").value.value = fname

        # Set up for file processing
        fileBody = open(opFolder + "/" + constBodyTmpFile + fileExt, "w")
        fFirst = True
        fBlankOk = False
        lineNum = 10
        regToolComment = re.compile(r"\(T[0-9]+\s")
        fFastZenabled = docSettings["fastZ"]
        regBody = re.compile(r""
            r"(?P<N>N[0-9]+ *)?" # line number
            r"(?P<line>"         # line w/o number
            r"(M(?P<M>[0-9]+) *)?" # M-code
            r"(G(?P<G>[0-9]+) *)?" # G-code
            r"(T(?P<T>[0-9]+))?" # Tool
            r".+)",              # to end of line
            re.IGNORECASE | re.DOTALL)
        toolChange = docSettings["toolChange"]
        fToolChangeNum = False
        if len(toolChange) != 0:
            toolChange = toolChange.replace(":", "\n")
            toolChange += "\n"
            match = regBody.match(toolChange).groupdict()
            if match["N"] != None:
                fToolChangeNum = True
                toolChange = match["line"]
                toolChange = toolChange.splitlines(True)
        
        # Parse end code list
        endCodes = docSettings["endCodes"]
        endGcodes = re.findall("G([0-9]+)", endCodes)
        endGcodeSet = set()
        for code in endGcodes:
            endGcodeSet.add(int(code))
        endMcodes = re.findall("M([0-9]+)", endCodes)
        endMcodeSet = set()
        for code in endMcodes:
            endMcodeSet.add(int(code))

        if fFastZenabled:
            regParseLine = re.compile(r""
                r"(G(?P<G>[0-9]+(\.[0-9]*)?)[^XYZF]*)?"
                r"(?P<XY>((X-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?"
                r"((Y-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?)"
                r"(Z(?P<Z>-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?"
                r"(F(?P<F>-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?",
                re.IGNORECASE)
            regGcodes = re.compile(r"G([0-9]+(?:\.[0-9]*)?)")

        tailGcode = ""
        pendingStopCmds = []
        totalOps = sum(len(group[1]) for group in opGroups)
        processedOps = 0

        # Track current machine state to suppress redundant commands
        currentToolNum = None
        currentSpindleSpeed = None
        currentSetup = None  # Track current setup to detect WCS changes
        
        # Regex patterns for detecting commands to potentially suppress
        regSpindleSpeed = re.compile(r'\bS(\d+)', re.IGNORECASE)
        regSpindleStart = re.compile(r'\bM\s*3\b', re.IGNORECASE)
        regCoolantOff = re.compile(r'\bM\s*9\b', re.IGNORECASE)
        regCoolantOn = re.compile(r'\bM\s*[78]\b', re.IGNORECASE)
        regDwell = re.compile(r'\bG\s*4\b', re.IGNORECASE)
        regReturnHome = re.compile(r'\bG\s*(28|53)\b', re.IGNORECASE)  # G28 or G53 (machine coords)
        # Fusion Personal Use warning message to suppress
        regPersonalUseWarning = re.compile(r'\(When using Fusion for Personal Use|\(moves is reduced to match|\(which can increase machining time|\(are available with a Fusion Subscription', re.IGNORECASE)

        # Process each tool group
        for toolNum, opsInGroup in opGroups:
            # Determine if this is a real tool change (first op in group with different tool)
            fRealToolChange = (currentToolNum is None or currentToolNum != toolNum)
            # Update current tool at start of group so subsequent ops know they're same tool
            currentToolNum = toolNum
            
            # Process each operation in this tool group
            for idx, (setup, op) in enumerate(opsInGroup):
                # Only the first operation in the group gets the actual tool change
                fRealToolChangeThisOp = fRealToolChange and (idx == 0)
                
                # Detect if WCS is changing (different setup = different WCS)
                fWcsChanging = (currentSetup is not None and currentSetup != setup)
                
                if progress and progress.wasCancelled:
                    fileBody.close()
                    fileHead.close()
                    os.remove(fileBody.name)
                    os.remove(path)
                    return "Cancelled by user"

                # Post process this operation
                opList = [op]
                
                retries = docSettings["postRetries"]
                delay = docSettings["initialDelay"]
                while True:
                    try:
                        program.operations = opList
                        if not program.postProcess(adsk.cam.NCProgramPostProcessOptions.create()):
                            retVal = "Fusion reported an error processing operation: " + op.name
                            return retVal
                    except Exception as exc:
                        retVal += " in operation " + op.name + ": " + str(exc)
                        return retVal

                    time.sleep(delay)
                    try:
                        fileOp = open(opPath, encoding="utf8", errors='replace')
                        break
                    except:
                        delay *= 2
                        retries -= 1
                        if retries > 0:
                            continue
                        return "Unable to open " + opPath
                
                # Parse and combine the gcode (similar to PostProcessSetup)
                if not fFirst and fBlankOk:
                    fileBody.write("\n")

                for stopCmd in pendingStopCmds:
                    fileBody.write("M9 (Coolant off for program stop)\n")
                    fileBody.write(stopCmd)
                pendingStopCmds = []

                # % at start only
                line = fileOp.readline()
                if len(line) > 0 and line[0] == "%":
                    if fFirst:
                        fileHead.write(line)
                    line = fileOp.readline()

                # Header comments and tool info
                while len(line) > 0 and (line[0] == "(" or line[0] == "O" or line[0] == "\n"):
                    if line[0] == "\n":
                        fBlankOk = True
                    # Skip Fusion Personal Use warning comments
                    if regPersonalUseWarning.search(line):
                        line = fileOp.readline()
                        continue
                    if regToolComment.match(line) != None:
                        fileHead.write(line)
                        line = fileOp.readline()
                        break
                    if fFirst:
                        pos = line.upper().find(opName.upper())
                        if pos != -1:
                            pos += len(opName)
                            line = line[0] + fname + line[pos:]
                        fileHead.write(line)
                    line = fileOp.readline()

                if len(line) == 0:
                    fileOp.close()
                    os.remove(fileOp.name)
                    fileOp = None
                    processedOps += 1
                    if progress:
                        progress.progressValue = int((processedOps / totalOps) * len(setups))
                    continue

                # Find tool change line and process preamble
                # Track whether this operation uses the same tool as previous
                # Within a tool group, only the first operation needs full setup
                fSuppressToolSetup = (idx > 0)  # Not the first operation in this tool group
                
                toolChangePattern = re.compile(r'(?=.*\bM\s*6\b)(?=.*\bT\s*\d{1,3}\b)', re.IGNORECASE)
                
                # For same-tool operations, we need to skip the preamble (coolant off, return home, 
                # tool change, spindle start, dwell, coolant on) and go straight to motion
                fInPreamble = True
                fFoundToolChange = False
                fPreambleComplete = False
                
                while True:
                    if len(line) == 0:
                        break
                    match = regBody.match(line)
                    if match is None:
                        line = fileOp.readline()
                        continue
                    match = match.groupdict()
                    lineContent = match["line"]
                    fNum = match["N"] != None

                    # Check if this is the tool change line
                    if toolChangePattern.search(lineContent):
                        fFoundToolChange = True
                        # Extract spindle speed from upcoming lines if this is a real tool change
                        if fRealToolChangeThisOp:
                            # Add tool change G-codes (full sequence including M9)
                            if not fFirst and len(toolChange) != 0:
                                if fToolChangeNum:
                                    for code in toolChange:
                                        fileBody.write("N" + str(lineNum) + " " + code)
                                        lineNum += constLineNumInc
                                else:
                                    fileBody.write(toolChange)
                            elif fFirst and not docSettings["skipFirstToolchange"] and len(toolChange) != 0:
                                if fToolChangeNum:
                                    for code in toolChange:
                                        fileBody.write("N" + str(lineNum) + " " + code)
                                        lineNum += constLineNumInc
                                else:
                                    fileBody.write(toolChange)
                        elif fWcsChanging and len(toolChange) != 0:
                            # Same tool but WCS changing - output G28/G30 for safety, but NOT M9
                            # Parse toolChange and filter out coolant commands
                            toolChangeLines = toolChange if isinstance(toolChange, list) else toolChange.strip().split('\n')
                            for code in toolChangeLines:
                                codeStr = code.strip() if isinstance(code, str) else code
                                # Skip M9 (coolant off) - we want coolant to stay on
                                if regCoolantOff.search(codeStr):
                                    continue
                                # Output G28, G30, or other return-home codes
                                if codeStr:
                                    if fToolChangeNum:
                                        fileBody.write("N" + str(lineNum) + " " + codeStr + "\n")
                                        lineNum += constLineNumInc
                                    else:
                                        fileBody.write(codeStr + "\n")
                        break

                    # Suppress preamble lines when not changing tools
                    if fSuppressToolSetup:
                        # Skip coolant off (M9)
                        if regCoolantOff.search(lineContent):
                            line = fileOp.readline()
                            continue
                        # Skip return home (G28) UNLESS WCS is changing (safety measure)
                        if regReturnHome.search(lineContent) and not fWcsChanging:
                            line = fileOp.readline()
                            continue
                    
                    # Skip Fusion Personal Use warning comments
                    if regPersonalUseWarning.search(lineContent):
                        line = fileOp.readline()
                        continue
                    
                    # Keep comments and program stop commands
                    if fFirst or lineContent[0] == "(" or re.search(r'\bM\s*[01]\b', lineContent, re.IGNORECASE):
                        if fNum:
                            fileBody.write("N" + str(lineNum) + " ")
                            lineNum += constLineNumInc
                        fileBody.write(lineContent)
                    line = fileOp.readline()
                    if len(line) == 0:
                        break
                    if line[0] == "\n":
                        fBlankOk = True

                # Process body - now we're past the tool change line
                fFastZ = fFastZenabled
                Gcode = None
                Zcur = None
                Zlast = None
                Zfeed = None
                fZfeedNotSet = True
                feedCur = 0
                fNeedFeed = False
                fLockSpeed = False
                
                # Track if we're still in the post-tool-change setup phase
                # (spindle start, dwell, coolant on, WCS selection)
                fInToolSetup = True
                linesSinceToolChange = 0

                while len(line) > 0:
                    match = regBody.match(line)
                    if match is None:
                        lineFull = fileOp.readline()
                        if len(lineFull) == 0:
                            break
                        line = lineFull
                        continue
                    
                    match = match.groupdict()
                    line = match["line"]
                    fNum = match["N"] != None
                    
                    linesSinceToolChange += 1
                    
                    # After ~15 lines from tool change, we're past the setup phase
                    if linesSinceToolChange > 15:
                        fInToolSetup = False

                    # Check for end markers
                    endMark = match["M"]
                    if endMark != None:
                        endMark = int(endMark)
                        if endMark in endMcodeSet:
                            break
                        if endMark == 49:
                            fLockSpeed = True
                        elif endMark == 48:
                            fLockSpeed = False
                    endMark = match["G"]
                    if endMark != None:
                        endMark = int(endMark)
                        if endMark in endGcodeSet:
                            break

                    # Determine if this line should be skipped
                    skipCurrentLine = False
                    
                    # Skip first tool change if option enabled
                    if fFirst and docSettings["skipFirstToolchange"]:
                        if toolChangePattern.search(line):
                            skipCurrentLine = True
                    
                    # When tool isn't actually changing, suppress setup commands
                    if fSuppressToolSetup and fInToolSetup:
                        # Skip tool change line (T# M6)
                        if toolChangePattern.search(line):
                            skipCurrentLine = True
                        # Skip spindle speed + start (S#### M3) if same speed
                        elif regSpindleStart.search(line):
                            speedMatch = regSpindleSpeed.search(line)
                            if speedMatch:
                                newSpeed = int(speedMatch.group(1))
                                if newSpeed == currentSpindleSpeed:
                                    skipCurrentLine = True
                                else:
                                    currentSpindleSpeed = newSpeed
                            else:
                                # M3 without S - skip if spindle already running
                                if currentSpindleSpeed is not None:
                                    skipCurrentLine = True
                        # Skip dwell (G4) - only used for spindle spinup
                        elif regDwell.search(line):
                            skipCurrentLine = True
                        # Skip coolant commands (M7, M8) - coolant still on from previous op
                        elif regCoolantOn.search(line):
                            skipCurrentLine = True
                    else:
                        # Track spindle speed for future comparisons
                        if regSpindleStart.search(line):
                            speedMatch = regSpindleSpeed.search(line)
                            if speedMatch:
                                currentSpindleSpeed = int(speedMatch.group(1))
                    
                    # Once we see actual motion (G0, G1) we're past setup
                    if re.search(r'\bG\s*[01]\b', line, re.IGNORECASE) and re.search(r'[XYZ]', line, re.IGNORECASE):
                        fInToolSetup = False
                    
                    # Analyze code for chances to make rapid moves (fastZ feature)
                    if fFastZ and not skipCurrentLine:
                        matchFast = regParseLine.match(line)
                        if matchFast and matchFast.end() != 0:
                            try:
                                matchFast = matchFast.groupdict()
                                Gcodes = regGcodes.findall(line)
                                fNoMotionGcode = True
                                fHomeGcode = False
                                for GcodeTmp in Gcodes:
                                    GcodeTmp = int(float(GcodeTmp))
                                    if GcodeTmp in constHomeGcodeSet:
                                        fHomeGcode = True
                                        break

                                    if GcodeTmp in constMotionGcodeSet:
                                        fNoMotionGcode = False
                                        Gcode = GcodeTmp
                                        if Gcode == 0:
                                            fNeedFeed = False
                                        break

                                if not fHomeGcode:
                                    Ztmp = matchFast["Z"]
                                    if Ztmp != None:
                                        Zlast = Zcur
                                        Zcur = float(Ztmp)

                                    feedTmp = matchFast["F"]
                                    if feedTmp != None:
                                        feedCur = float(feedTmp)

                                    XYcur = matchFast["XY"].rstrip("\n ")

                                    if (Zfeed == None or fZfeedNotSet) and (Gcode == 0 or Gcode == 1) and Ztmp != None and len(XYcur) == 0:
                                        # Figure out Z feed
                                        if (Zfeed != None):
                                            fZfeedNotSet = False
                                        Zfeed = Zcur
                                        if Gcode != 0:
                                            # Replace line with rapid move
                                            line = constRapidZgcode.format(Zcur, line[:-1])
                                            fNeedFeed = True
                                            Gcode = 0

                                    if Gcode == 1 and not fLockSpeed:
                                        if Ztmp != None:
                                            if len(XYcur) == 0 and (Zcur >= Zlast or Zcur >= Zfeed or feedCur == 0):
                                                # Upward move, above feed height, or anomalous feed rate.
                                                # Replace with rapid move
                                                line = constRapidZgcode.format(Zcur, line[:-1])
                                                fNeedFeed = True
                                                Gcode = 0

                                        elif Zcur >= Zfeed:
                                            # No Z move, at/above feed height
                                            line = constRapidXYgcode.format(XYcur, line[:-1])
                                            fNeedFeed = True
                                            Gcode = 0

                                    elif fNeedFeed and fNoMotionGcode:
                                        # No G-code present, changing to G1
                                        if Ztmp != None:
                                            if len(XYcur) != 0:
                                                # Not Z move only - back to G1
                                                line = constFeedXYZgcode.format(XYcur, Zcur, feedCur, line[:-1])
                                                fNeedFeed = False
                                                Gcode = 1
                                            elif Zcur < Zfeed and Zcur <= Zlast:
                                                # Not up nor above feed height - back to G1
                                                line = constFeedZgcode.format(Zcur, feedCur, line[:-1])
                                                fNeedFeed = False
                                                Gcode = 1
                                                
                                        elif len(XYcur) != 0 and Zcur < Zfeed:
                                            # No Z move, below feed height - back to G1
                                            line = constFeedXYgcode.format(XYcur, feedCur, line[:-1])
                                            fNeedFeed = False
                                            Gcode = 1

                                    if (Gcode != 0 and fNeedFeed):
                                        if (feedTmp == None):
                                            # Feed rate not present, add it
                                            line = line[:-1] + constAddFeedGcode.format(feedCur)
                                        fNeedFeed = False

                                    if Zcur != None and Zfeed != None and Zcur >= Zfeed and Gcode != None and \
                                        Gcode != 0 and len(XYcur) != 0 and (Ztmp != None or Gcode != 1):
                                        # We're at or above the feed height, but made a cutting move.
                                        # Feed height is wrong, bring it up
                                        Zfeed = Zcur + 0.001
                            except:
                                fFastZ = False # Just skip changes
                    
                    if not skipCurrentLine:
                        if fNum:
                            fileBody.write("N" + str(lineNum) + " ")
                            lineNum += constLineNumInc
                        fileBody.write(line)
                    
                    lineFull = fileOp.readline()
                    if len(lineFull) == 0:
                        break
                    match = regBody.match(lineFull)
                    if match is None:
                        line = lineFull
                        continue
                    match = match.groupdict()
                    line = match["line"]
                    fNum = match["N"] != None

                # Save tail from first operation
                tailContent = fileOp.read()
                tailLines = tailContent.splitlines(True)
                remainingTail = []
                for tailLine in tailLines:
                    if re.search(r'\bM\s*[01]\b', tailLine, re.IGNORECASE):
                        match = regBody.match(tailLine)
                        if match:
                            match = match.groupdict()
                            if match["N"] != None:
                                stopCmd = "N" + str(lineNum) + " " + match["line"]
                                lineNum += constLineNumInc
                            else:
                                stopCmd = tailLine
                            pendingStopCmds.append(stopCmd)
                    else:
                        remainingTail.append(tailLine)
                
                if fFirst:
                    tailGcode = "".join(remainingTail)
                fFirst = False
                currentSetup = setup  # Update current setup for WCS change detection
                fileOp.close()
                os.remove(fileOp.name)
                fileOp = None
                
                processedOps += 1
                if progress:
                    progress.progressValue = int((processedOps / totalOps) * len(setups))

        # Write remaining pending commands
        for stopCmd in pendingStopCmds:
            fileBody.write("M9 (Coolant off for program stop)\n")
            fileBody.write(stopCmd)

        # Add tail
        if len(tailGcode) != 0:
            tailGcode = tailGcode.splitlines(True)
            for code in tailGcode:
                match = regBody.match(code)
                if match:
                    match = match.groupdict()
                    if match["N"] != None:
                        fileBody.write("N" + str(lineNum) + " " + match["line"])
                        lineNum += constLineNumInc
                    else:
                        fileBody.write(code)
                else:
                    fileBody.write(code)

        # Copy body to head
        fileBody.close()
        fileBody = open(fileBody.name)
        while True:
            block = fileBody.read(10240)
            if len(block) == 0:
                break
            fileHead.write(block)
        fileBody.close()
        os.remove(fileBody.name)
        fileBody = None
        fileHead.close()
        fileHead = None

        return None

    except:
        if fileHead:
            try:
                fileHead.close()
                os.remove(fileHead.name)
            except:
                pass
        if fileBody:
            try:
                fileBody.close()
                os.remove(fileBody.name)
            except:
                pass
        if fileOp:
            try:
                fileOp.close()
                os.remove(fileOp.name)
            except:
                pass
        if ui:
            retVal += " " + traceback.format_exc()
        return retVal


def PostProcessSetup(fname, setup, setupFolder, docSettings, program, debugComments=None):
    ui = None
    fileHead = None
    fileBody = None
    fileOp = None
    retVal = "Fusion reported an exception"

    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        doc = app.activeDocument
        cam = adsk.cam.CAM.cast(doc.products.itemByProductType(constCAMProductId))
        parameters = program.parameters

        # Verify file name is valid by creating it now
        fileExt = parameters.itemByName("nc_program_nc_extension").value.value
        path = setupFolder + "/" + fname + fileExt
        try:
            pathlib.Path(setupFolder).mkdir(parents=True, exist_ok=True)
            fileHead = open(path, "w")
            
            # Write debug comments at the top of the file if available
            if debugComments and len(debugComments) > 0:
                fileHead.write("(========== DEBUG INFORMATION ==========)\n")
                for comment in debugComments:
                    fileHead.write(comment)
                fileHead.write("(========================================)\n")
                fileHead.write("\n")
                
        except Exception as exc:
            return "Unable to create output file '" + path + "'. Make sure the setup name is valid as a file name."
        
        # Make sure toolpaths are valid
        if not cam.checkToolpath(setup):
            genStat = cam.generateToolpath(setup)
            while not genStat.isGenerationCompleted:
                time.sleep(.1)

        # set up NCProgram parameters
        opName = fname
        opFolder = setupFolder
        if docSettings["splitSetup"]:
            opName = constOpTmpFile
            opFolder = tempfile.gettempdir()    # e.g., C:\Users\Tim\AppData\Local\Temp
            opFolder = opFolder.replace("\\", "/")

        parameters.itemByName("nc_program_openInEditor").value.value = False
        AssignOutputFolder(parameters, opFolder)
        parameters.itemByName("nc_program_filename").value.value = opName
        parameters.itemByName("nc_program_name").value.value = fname

        # Do it all at once?
        if not docSettings["splitSetup"]:
            fileHead.close()
            try:
                program.operations = [setup]
                if not program.postProcess(adsk.cam.NCProgramPostProcessOptions.create()):
                    return "Fusion reported an error."
                time.sleep(constPostLoopDelay) # files missing sometimes unless we slow down (??)
                return None
            except Exception as exc:
                retVal += ": " + str(exc)
                return retVal

        # Split setup into individual operations
        opPath = opFolder + "/" + opName + fileExt
        fileBody = open(opFolder + "/" + constBodyTmpFile + fileExt, "w")
        fFirst = True
        fBlankOk = False
        lineNum = 10
        regToolComment = re.compile(r"\(T[0-9]+\s")
        fFastZenabled = docSettings["fastZ"]
        regBody = re.compile(r""
            r"(?P<N>N[0-9]+ *)?" # line number
            r"(?P<line>"         # line w/o number
            r"(M(?P<M>[0-9]+) *)?" # M-code
            r"(G(?P<G>[0-9]+) *)?" # G-code
            r"(T(?P<T>[0-9]+))?" # Tool
            r".+)",              # to end of line
            re.IGNORECASE | re.DOTALL)
        toolChange = docSettings["toolChange"]
        fToolChangeNum = False
        if len(toolChange) != 0:
            toolChange = toolChange.replace(":", "\n")
            toolChange += "\n"
            match = regBody.match(toolChange).groupdict()
            if match["N"] != None:
                fToolChangeNum = True
                toolChange = match["line"]
                # split into individual lines to add line numbers
                toolChange = toolChange.splitlines(True)
        # Parse end code list, splitting into G-codes and M-codes
        endCodes = docSettings["endCodes"]
        endGcodes = re.findall("G([0-9]+)", endCodes)
        endGcodeSet = set()
        for code in endGcodes:
            endGcodeSet.add(int(code))
        endMcodes = re.findall("M([0-9]+)", endCodes)
        endMcodeSet = set()
        for code in endMcodes:
            endMcodeSet.add(int(code))

        if fFastZenabled:
            regParseLine = re.compile(r""
                r"(G(?P<G>[0-9]+(\.[0-9]*)?)[^XYZF]*)?"
                r"(?P<XY>((X-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?"
                r"((Y-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?)"
                r"(Z(?P<Z>-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?"
                r"(F(?P<F>-?[0-9]+(\.[0-9]*)?)[^XYZF]*)?",
                re.IGNORECASE)
            regGcodes = re.compile(r"G([0-9]+(?:\.[0-9]*)?)")

        # Pending M0/M1 commands to write at start of next operation
        pendingStopCmds = []

        i = 0
        ops = setup.allOperations
        while i < ops.count:
            op = ops[i]
            i += 1
            if op.isSuppressed:
                continue

            # Look ahead for operations without a toolpath. This can happen
            # with a manual operation. Group it with current operation.
            # Or if first, group it with subsequent ones.
            # Also optionally group together operations with the same tool number
            # BUT: Don't group Manual NC (no toolpath) - process them separately
            opHasTool = None
            curTool = -1
            hasTool = op.hasToolpath
            
            # If this is a Manual NC operation (no toolpath), process it alone
            if not hasTool:
                opList = [op]
                # Find the next operation with a toolpath to get tool info for grouping
                while i < ops.count:
                    nextOp = ops[i]
                    if not nextOp.isSuppressed and nextOp.hasToolpath:
                        opHasTool = nextOp
                        opList.append(nextOp)
                        curTool = nextOp.tool.parameters.itemByName("tool_number").value.value
                        i += 1
                        break
                    i += 1
            else:
                opHasTool = op
                curTool = op.tool.parameters.itemByName("tool_number").value.value
                opList = [op]
                while i < ops.count:
                    op = ops[i]
                    if not op.isSuppressed:
                        # Stop grouping if we hit a Manual NC (no toolpath) or different tool
                        if not op.hasToolpath:
                            break  # Don't include Manual NC in this group
                        if not docSettings.get("combineTool", False) or op.tool.parameters.itemByName("tool_number").value.value != curTool:
                            break
                        opList.append(op)
                    i += 1

            retries = docSettings["postRetries"]
            delay = docSettings["initialDelay"]
            while True:
                try:
                    program.operations = opList
                    if not program.postProcess(adsk.cam.NCProgramPostProcessOptions.create()):
                        retVal = "Fusion reported an error processing operation"
                        if (opHasTool != None):
                            retVal += ": " +  opHasTool.name
                        return retVal
                except Exception as exc:
                    if (opHasTool != None):
                        retVal += " in operation " +  opHasTool.name
                    retVal += ": " + str(exc)
                    return retVal

                time.sleep(delay) # wait for it to finish (??)
                try:
                    fileOp = open(opPath, encoding="utf8", errors='replace')
                    break
                except:
                    delay *= 2
                    retries -= 1
                    if retries > 0:
                        continue
                    # Maybe the file name extension is wrong
                    for file in os.listdir(opFolder):
                        if file.startswith(opName):
                            ext = file[len(opName):]
                            if ext != fileExt:
                                return ("Unable to open output file. "
                                    "Found the file with extension '{}' instead "
                                    "of '{}'. Make sure you have the correct file "
                                    "extension set in the Post Process All "
                                    "dialog.".format(ext, fileExt))
                            break
                    return "Unable to open " + opPath
            
            # Parse the gcode. We expect a header like this:
            #
            # % <optional>
            # (<comments>) <0 or more lines>
            # (<Txx tool comment>) <optional>
            # <comments or G-code initialization, up to Txx>
            #
            # This header is stripped from all files after the first,
            # except the tool comment is put in a list at the top.
            # The header ends when we find the body, which starts with:
            #
            # Txx ...   (optionally preceded by line number Nxx)
            #
            # We copy all the body, looking for the tail. The start
            # of the tail is denoted by any of a list of G-codes
            # entered by the user. The defaults are:
            # M30 - end program
            # M5 - stop spindle
            # M9 - stop coolant
            # The tail is stripped until the last operation is done.

            # Space between operations
            if not fFirst and fBlankOk:
                fileBody.write("\n")

            # Write any pending M0/M1 commands from previous operation's Manual NC
            # Add M9 (coolant off) before each stop command for safety
            for stopCmd in pendingStopCmds:
                fileBody.write("M9 (Coolant off for program stop)\n")
                fileBody.write(stopCmd)
            pendingStopCmds = []

            # % at start only
            line = fileOp.readline()
            if line[0] == "%":
                if fFirst:
                    fileHead.write(line)
                line = fileOp.readline()

            # check for initial comments and tool
            # send it to header
            while line[0] == "(" or line[0] == "O" or line[0] == "\n":
                if line[0] == "\n":
                    fBlankOk = True
                if regToolComment.match(line) != None:
                    fileHead.write(line)
                    line = fileOp.readline()
                    break

                if fFirst:
                    pos = line.upper().find(opName.upper())
                    if pos != -1:
                        pos += len(opName)
                        if docSettings["numericName"]:
                            fill = "0" * (pos - len(fname) - 1)
                        else:
                            fill = ""
                        line = line[0] + fill + fname + line[pos:]    # correct file name
                    fileHead.write(line)
                line = fileOp.readline()

            # Body starts at tool code, T - look for tool change line (M6 + T###)
            toolChangePattern = re.compile(r'(?=.*\bM\s*6\b)(?=.*\bT\s*\d{1,3}\b)', re.IGNORECASE)
            while True:
                match = regBody.match(line).groupdict()
                line = match["line"]        # filter off line number if present
                fNum = match["N"] != None

                # Check for tool change line using more specific pattern
                if toolChangePattern.search(line):
                    # Add tool change G-codes if not first operation
                    if not fFirst and len(toolChange) != 0:
                        # have tool change G-codes to add
                        if fToolChangeNum:
                            # Add line number to tool change
                            for code in toolChange:
                                fileBody.write("N" + str(lineNum) + " " + code)
                                lineNum += constLineNumInc
                        else:
                            fileBody.write(toolChange)
                    elif fFirst and not docSettings["skipFirstToolchange"] and len(toolChange) != 0:
                        # First operation and skipFirstToolchange is disabled - add tool change codes
                        if fToolChangeNum:
                            # Add line number to tool change
                            for code in toolChange:
                                fileBody.write("N" + str(lineNum) + " " + code)
                                lineNum += constLineNumInc
                        else:
                            fileBody.write(toolChange)
                    break

                # Preserve line if: first operation, comment, or M0/M1 (program stop) command
                if fFirst or line[0] == "(" or re.search(r'\bM\s*[01]\b', line, re.IGNORECASE):
                    if (fNum):
                        fileBody.write("N" + str(lineNum) + " ")
                        lineNum += constLineNumInc
                    fileBody.write(line)
                line = fileOp.readline()
                if len(line) == 0:
                    return "Tool change G-code (Txx) not found; this post processor is not compatible with Post Process All."
                if line[0] == "\n":
                    fBlankOk = True

            # We're done with the head, move on to the body
            # Initialize rapid move optimizations
            fFastZ = fFastZenabled
            Gcode = None
            Zcur = None
            Zfeed = None
            fZfeedNotSet = True
            feedCur = 0
            fNeedFeed = False
            fLockSpeed = False

            # Note that match, line, and fNum are already set
            while True:
                # End of program marker?
                endMark = match["M"]
                if endMark != None:
                    endMark = int(endMark)
                    if endMark in endMcodeSet:
                        break
                    # When M49/M48 is used to turn off speed changes, disable fast moves as well
                    if endMark == 49:
                        fLockSpeed = True
                    elif endMark == 48:
                        fLockSpeed = False
                endMark = match["G"]
                if endMark != None:
                    endMark = int(endMark)
                    if endMark in endGcodeSet:
                        break

                if fFastZ:
                    # Analyze code for chances to make rapid moves
                    match = regParseLine.match(line)
                    if match.end() != 0:
                        try:
                            match = match.groupdict()
                            Gcodes = regGcodes.findall(line)
                            fNoMotionGcode = True
                            fHomeGcode = False
                            for GcodeTmp in Gcodes:
                                GcodeTmp = int(float(GcodeTmp))
                                if GcodeTmp in constHomeGcodeSet:
                                    fHomeGcode = True
                                    break

                                if GcodeTmp in constMotionGcodeSet:
                                    fNoMotionGcode = False
                                    Gcode = GcodeTmp
                                    if Gcode == 0:
                                        fNeedFeed = False
                                    break

                            if not fHomeGcode:
                                Ztmp = match["Z"]
                                if Ztmp != None:
                                    Zlast = Zcur
                                    Zcur = float(Ztmp)

                                feedTmp = match["F"]
                                if feedTmp != None:
                                    feedCur = float(feedTmp)

                                XYcur = match["XY"].rstrip("\n ")

                                if (Zfeed == None or fZfeedNotSet) and (Gcode == 0 or Gcode == 1) and Ztmp != None and len(XYcur) == 0:
                                    # Figure out Z feed
                                    if (Zfeed != None):
                                        fZfeedNotSet = False
                                    Zfeed = Zcur
                                    if Gcode != 0:
                                        # Replace line with rapid move
                                        line = constRapidZgcode.format(Zcur, line[:-1])
                                        fNeedFeed = True
                                        Gcode = 0

                                if Gcode == 1 and not fLockSpeed:
                                    if Ztmp != None:
                                        if len(XYcur) == 0 and (Zcur >= Zlast or Zcur >= Zfeed or feedCur == 0):
                                            # Upward move, above feed height, or anomalous feed rate.
                                            # Replace with rapid move
                                            line = constRapidZgcode.format(Zcur, line[:-1])
                                            fNeedFeed = True
                                            Gcode = 0

                                    elif Zcur >= Zfeed:
                                        # No Z move, at/above feed height
                                        line = constRapidXYgcode.format(XYcur, line[:-1])
                                        fNeedFeed = True
                                        Gcode = 0

                                elif fNeedFeed and fNoMotionGcode:
                                    # No G-code present, changing to G1
                                    if Ztmp != None:
                                        if len(XYcur) != 0:
                                            # Not Z move only - back to G1
                                            line = constFeedXYZgcode.format(XYcur, Zcur, feedCur, line[:-1])
                                            fNeedFeed = False
                                            Gcode = 1
                                        elif Zcur < Zfeed and Zcur <= Zlast:
                                            # Not up nor above feed height - back to G1
                                            line = constFeedZgcode.format(Zcur, feedCur, line[:-1])
                                            fNeedFeed = False
                                            Gcode = 1
                                            
                                    elif len(XYcur) != 0 and Zcur < Zfeed:
                                        # No Z move, below feed height - back to G1
                                        line = constFeedXYgcode.format(XYcur, feedCur, line[:-1])
                                        fNeedFeed = False
                                        Gcode = 1

                                if (Gcode != 0 and fNeedFeed):
                                    if (feedTmp == None):
                                        # Feed rate not present, add it
                                        line = line[:-1] + constAddFeedGcode.format(feedCur)
                                    fNeedFeed = False

                                if Zcur != None and Zfeed != None and Zcur >= Zfeed and Gcode != None and \
                                    Gcode != 0 and len(XYcur) != 0 and (Ztmp != None or Gcode != 1):
                                    # We're at or above the feed height, but made a cutting move.
                                    # Feed height is wrong, bring it up
                                    Zfeed = Zcur + 0.001
                        except:
                            fFastZ = False # Just skip changes

                # copy line to output
                # Skip T code line if this is first operation and skipFirstToolchange is enabled
                skipCurrentLine = False
                if fFirst and docSettings["skipFirstToolchange"]:
                    # Check for tool change line: must contain both M6 and T### (1-3 digits)
                    if toolChangePattern.search(line):
                        skipCurrentLine = True
                
                if not skipCurrentLine:
                    if (fNum):
                        fileBody.write("N" + str(lineNum) + " ")
                        lineNum += constLineNumInc
                    fileBody.write(line)
                lineFull = fileOp.readline()
                if len(lineFull) == 0:
                    break
                match = regBody.match(lineFull).groupdict()
                line = match["line"]        # filter off line number if present
                fNum = match["N"] != None

            # Found tail of program - scan for M0/M1 commands to preserve in sequence
            tailContent = lineFull + fileOp.read()
            tailLines = tailContent.splitlines(True)
            remainingTail = []
            for tailLine in tailLines:
                if re.search(r'\bM\s*[01]\b', tailLine, re.IGNORECASE):
                    # Preserve M0/M1 (program stop) commands from Manual NC operations
                    # Save them to be written at the START of next operation
                    match = regBody.match(tailLine).groupdict()
                    if match["N"] != None:
                        stopCmd = "N" + str(lineNum) + " " + match["line"]
                        lineNum += constLineNumInc
                    else:
                        stopCmd = tailLine
                    pendingStopCmds.append(stopCmd)
                else:
                    remainingTail.append(tailLine)
            
            if fFirst:
                # Save remaining tail (without M0/M1) for the very end
                tailGcode = "".join(remainingTail)
            fFirst = False
            fileOp.close()
            os.remove(fileOp.name)
            fileOp = None

        # Write any remaining pending M0/M1 commands before final tail
        for stopCmd in pendingStopCmds:
            fileBody.write("M9 (Coolant off for program stop)\n")
            fileBody.write(stopCmd)

        # Completed all operations, add tail
        # Update line numbers if present
        if len(tailGcode) != 0:
            tailGcode = tailGcode.splitlines(True)
            for code in tailGcode:
                match = regBody.match(code).groupdict()
                if match["N"] != None:
                    fileBody.write("N" + str(lineNum) + " " + match["line"])
                    lineNum += constLineNumInc
                else:
                    fileBody.write(code)

        # Copy body to head
        fileBody.close()
        fileBody = open(fileBody.name)  # open for reading
        # copy in chunks
        while True:
            block = fileBody.read(10240)
            if len(block) == 0:
                break
            fileHead.write(block)
            block = None    # free memory
        fileBody.close()
        os.remove(fileBody.name)
        fileBody = None
        fileHead.close()
        fileHead = None

        return None

    except:
        if fileHead:
            try:
                fileHead.close()
                os.remove(fileHead.name)
            except:
                pass

        if fileBody:
            try:
                fileBody.close()
                os.remove(fileBody.name)
            except:
                pass

        if fileOp:
            try:
                fileOp.close()
                os.remove(fileOp.name)
            except:
                pass

        if ui:
            retVal += " " + traceback.format_exc()

        return retVal

import adsk.core, adsk.cam
import os
import sys
import re
from ...lib import fusion360utils as futil
from ... import config
app = adsk.core.Application.get()
ui = app.userInterface


# TODO *** Specify the command identity information. ***
CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_cmdDialog'
CMD_NAME = 'Rename Operations'
CMD_Description = '''Rename all operations in the activated setup (including operations in any folder within that setup) to a specified format.\n
Example output (with "Add Operation Strategy" checked):\nOP 1 - drill\nOP 2 - chamfer2d\n\nOr (with "Add Operation Strategy" unchecked):\nOP 1 \nOP 2'''
# Specify that the command will be promoted to the panel.
IS_PROMOTED = True

# prints the report
PRINT_REPORT = False

# TODO *** Define the location where the command button will be created. ***
# This is done by specifying the workspace, the tab, and the panel, and the 
# command it will be inserted beside. Not providing the command to position it
# will insert it at the end.
WORKSPACE_ID = 'CAMEnvironment'
PANEL_ID = 'CAMActionPanel'
#COMMAND_BESIDE_ID = 'ScriptsManagerCommand'

# Resource location for command icons, here we assume a sub folder in this directory named "resources".
ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')
# imagePath = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', 'help.png')

# Local list of event handlers used to maintain a reference so
# they are not released and garbage collected.
local_handlers = []


# Executed when add-in is run.
def start():
    # Create a command Definition.
    cmd_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_Description, ICON_FOLDER)
    # cmd_def.toolClipFilename = imagePath

    # Define an event handler for the command created event. It will be called when the button is clicked.
    futil.add_handler(cmd_def.commandCreated, command_created)

    # ******** Add a button into the UI so the user can run the command. ********
    # Get the target workspace the button will be created in.
    workspace = ui.workspaces.itemById(WORKSPACE_ID)

    # Get the panel the button will be created in.
    panel = workspace.toolbarPanels.itemById(PANEL_ID)

    # Create the button command control in the UI after the specified existing command.
    #control = panel.controls.addCommand(cmd_def, COMMAND_BESIDE_ID, False)
    control = panel.controls.addCommand(cmd_def)

    # Specify if the command is promoted to the main toolbar. 
    control.isPromoted = IS_PROMOTED


# Executed when add-in is stopped.
def stop():
    # Get the various UI elements for this command
    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    panel = workspace.toolbarPanels.itemById(PANEL_ID)
    command_control = panel.controls.itemById(CMD_ID)
    command_definition = ui.commandDefinitions.itemById(CMD_ID)

    # Delete the button command control
    if command_control:
        command_control.deleteMe()

    # Delete the command definition
    if command_definition:
        command_definition.deleteMe()


# Function that is called when a user clicks the corresponding button in the UI.
# This defines the contents of the command dialog and connects to the command related events.
def command_created(args: adsk.core.CommandCreatedEventArgs):
    # General logging for debug.
    #futil.log(f'{CMD_NAME} Command Created Event')

    cmd = args.command
    # avoid the OK event to be raised when the form is closed by an external input
    cmd.isExecutedWhenPreEmpted = False

    # https://help.autodesk.com/view/fusion360/ENU/?contextId=CommandInputs
    inputs = cmd.commandInputs
    
    # Create inputs
    inputs.addTextBoxCommandInput('input_prefix', 'Prefix', 'OP', 1, False)
    inputs.addIntegerSpinnerCommandInput('input_start', 'Start Number', 0, 1000, 1, 1)
    inputs.addIntegerSpinnerCommandInput('input_increment', 'Increment', 1, 1000, 1, 1)
    inputs.addBoolValueInput('input_stratname', 'Add Operation Strategy', True, '', True)

    # TODO Connect to the events that are needed by this command.
    futil.add_handler(cmd.execute, command_execute, local_handlers=local_handlers)
    futil.add_handler(cmd.inputChanged, command_input_changed, local_handlers=local_handlers)
    futil.add_handler(cmd.executePreview, command_preview, local_handlers=local_handlers)
    futil.add_handler(cmd.validateInputs, command_validate_input, local_handlers=local_handlers)
    futil.add_handler(cmd.destroy, command_destroy, local_handlers=local_handlers)


# This event handler is called when the user clicks the OK button in the command dialog or 
# is immediately called after the created event not command inputs were created for the dialog.
def command_execute(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    #futil.log(f'{CMD_NAME} Command Execute Event')

    # Get a reference to command's inputs.
    cmd = args.command
    inputs = cmd.commandInputs
    input_prefix: adsk.core.TextBoxCommandInput = inputs.itemById('input_prefix')
    input_start: adsk.core.IntegerSpinnerCommandInput = inputs.itemById('input_start')
    input_increment: adsk.core.IntegerSpinnerCommandInput = inputs.itemById('input_increment')
    input_stratname: adsk.core.BoolValueCommandInput = inputs.itemById('input_stratname')

    doc = app.activeDocument
    products = doc.products
    # Get the CAM product
    cam = adsk.cam.CAM.cast(products.itemByProductType("CAMProductType"))

    # get setups
    setups = cam.setups

    # search active setup
    setup = None
    for s in setups:
        if s.isActive:
            setup = s
            break

    # check if setup found
    if not setup:
        #app.log('WARNING: No operation renamed!')  
        ui.messageBox('Ensure the setup to rename is active and try again...', 'Fusion 360',
                        adsk.core.MessageBoxButtonTypes.OKButtonType, adsk.core.MessageBoxIconTypes.WarningIconType)
        return

    # check setup has operation
    if setup.allOperations.count < 1:
        #app.log('WARNING: No operation renamed!')  
        ui.messageBox('Ensure there are operations in the active setup and try again...', 'Fusion 360',
                       adsk.core.MessageBoxButtonTypes.OKButtonType, adsk.core.MessageBoxIconTypes.WarningIconType)
        return

    # process setup 
    numOp = setup.allOperations.count

    # Show dialog
    progressDialog = ui.createProgressDialog()
    progressDialog.show('Renaming in Progress ', 'Renaming operation %v/%m (%p%)', 0, numOp, 0) 

# rename loop
    counter = input_start.value
    for i, op in enumerate(setup.allOperations):
        op = adsk.cam.Operation.cast(op)
        prefix = input_prefix.text
        tempInitName = op.name

 
        parts = tempInitName.split()
        
        # Convert strategy name
        strategy = convert_strategy(op.strategy)
        strategy_parts = strategy.split()
        modified_parts = []

        for part in parts:
            if prefix.lower() in part.lower(): # Check if there is already a prefix
                continue
            elif part in strategy_parts or op.strategy.lower() in part.lower(): # Check if there is already an internal strategy name
                continue
            elif strategy.lower() in part.lower(): # Check if the human readable strategy is already in the name
                continue
            elif part == op.name: # Check if the part is the original name
                modified_parts.append(op.name) # Add the original name
            else:
                modified_parts.append(part) # Add the part

 


        merged_str = ' '.join(modified_parts) # Merge the parts
        strategy_name = f"{strategy} " if input_stratname.value else "" # Add the strategy name if the user wants it
        op.name = f"{prefix}{counter} {strategy_name}{format_comment(merged_str)}" # Set the new name
        
        counter = counter + input_increment.value # Increment the Prefix counter
        progressDialog.progressValue = i
    progressDialog.hide()
    
    # report
    if PRINT_REPORT:
        indentString = '  '
        app.log('\n')
        app.log('Renaming Statistics...')
        app.log(indentString + 'Active setup = ' + str(setup.name))
        app.log(indentString + '  Number of operations renamed = ' + str(numOp))
        firstOut = adsk.cam.Operation.cast(setup.allOperations.item(0)).name
        lastOut = adsk.cam.Operation.cast(setup.allOperations.item(setup.allOperations.count - 1)).name
        app.log(indentString + '  First operation              = ' + str(firstOut))
        app.log(indentString + '  Last operation               = ' + str(lastOut))
        app.log('Renaming complete!')


# This event handler is called when the command needs to compute a new preview in the graphics window.
def command_preview(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    #futil.log(f'{CMD_NAME} Command Preview Event')
    inputs = args.command.commandInputs


# This event handler is called when the user changes anything in the command dialog
# allowing you to modify values of other inputs based on that change.
def command_input_changed(args: adsk.core.InputChangedEventArgs):
    changed_input = args.input
    inputs = args.inputs

    # General logging for debug.
    #futil.log(f'{CMD_NAME} Input Changed Event fired from a change to {changed_input.id}')


# This event handler is called when the user interacts with any of the inputs in the dialog
# which allows you to verify that all of the inputs are valid and enables the OK button.
def command_validate_input(args: adsk.core.ValidateInputsEventArgs):
    # General logging for debug.
    #futil.log(f'{CMD_NAME} Validate Input Event')
    pass
        

# This event handler is called when the command terminates.
def command_destroy(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    #futil.log(f'{CMD_NAME} Command Destroy Event')

    global local_handlers
    local_handlers = []
    
    
# Operation names with special formatting
strategy_dict = {
    "adaptive2d" : "2D Adaptive",
    "pocket2d" : "2D Pocket",
    "contour2d" : "2D Contour",
    "path3d" : "Trace",
    "chamfer2d" : "2D Chamfer",
    "profile2d" : "2D Profile",
    "adaptive" : "3D Adaptive",
    "three_plus_two" : "3+2",
    "contour3d": "3D Contour",
}

# Return the human readable strategy if it exists, otherwise format the strategy name
def convert_strategy(strategy):
    return strategy_dict.get(strategy, format_comment(strategy)) 

# Remove numbers from the end of the string and capitalize the first letter
def format_comment(text):   
    char = r'\b[a-zA-Z]+\d+\b|\(\d+\)'
    filtered_string = re.sub(char, '', text)
    # Remove any extra spaces left by the removal
    filtered_string = re.sub(r'\s+', ' ', filtered_string).strip()
    return filtered_string.capitalize()
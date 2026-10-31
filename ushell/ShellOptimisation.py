#Author: Owen Peckham
#Description: An add-in that optimises the thickness of an outside shell of a body to retain the body's mass with the aim of lightweighting.
#Version: 1.0

"""This is a Fusion 360 add-in designed to be placed (within its Microchannels parent folder) in the Fusion 360 Add-ins folder (typically: 'C:/Users/$user$/AppData/Roaming/Autodesk/Autodesk Fusion 360/API/AddIns')."""

import adsk.core, adsk.fusion, adsk.cam, traceback
import os, datetime, timeit

# Global list to keep all event handlers in scope.

# Command inputs.
_initialThickness = adsk.core.ValueCommandInput.cast(None)
_tolerance = adsk.core.ValueCommandInput.cast(None)
_maxIterations = adsk.core.IntegerSpinnerCommandInput.cast(None)
_errMessage = adsk.core.TextBoxCommandInput.cast(None)
_bodySelection = adsk.core.SelectionCommandInput.cast(None)
_wasSurface = False
_debug = True

# This is only needed for Python.
handlers = []


def debugToConsole(message):
    global _debug

    if _debug:
        app = adsk.core.Application.get()
        ui = app.userInterface

        textPalette = ui.palettes.itemById('TextCommands')
        if not textPalette.isVisible:
            textPalette.isVisible = True
        textPalette.writeText(message + '\n')


# Function to get the overall mass of the whole active component.
def weighComponent():
    app = adsk.core.Application.get()
    design = app.activeProduct

    # Check if we have a valid design
    if not design or not isinstance(design, adsk.fusion.Design):
        debugToConsole('No active Fusion 360 design found.')
        return None

    activeComponent = design.activeComponent

    # Get the mass of the whole component
    totalMass = 0
    for bRepBody in activeComponent.bRepBodies:
        totalMass += bRepBody.physicalProperties.mass
    
    return totalMass


# Function to undo the last shell feature created on a body.
def undoShellFeatures():
    try:
        # Get the root component of the active design
        app = adsk.core.Application.get()
        ui = app.userInterface
        design = app.activeProduct

        global _debug, _wasSurface
        
        # Check if we have a valid design
        if not design or not isinstance(design, adsk.fusion.Design):
            debugToConsole('No active Fusion 360 design found.')
            return None
        
        # Get the active component
        activeComponent = design.activeComponent

        debugToConsole(f"Attempting to undo shell features on '{activeComponent.name}'.")
        
        message = ""
        if _wasSurface:

            # Get the timeline object
            timeline = design.timeline
            features_to_delete = []
            shell_feature_found = False

            # Iterate over the timeline in chronological order
            for timeline_obj in timeline:
                # Skip groups in the timeline
                if timeline_obj.isGroup:
                    continue
                
                # Get the actual feature from the timeline object
                feature = timeline_obj.entity

                # If the shell feature for 'Body 1' is found, we start deleting subsequent features
                if isinstance(feature, adsk.fusion.ShellFeature) and feature.bodies[0].name == 'Selected_Body':
                    shell_feature_found = True
                    features_to_delete.append(feature)
                
                # Once shell feature is found, add all following features to the delete list
                elif shell_feature_found:
                    if isinstance(feature, (adsk.fusion.StitchFeature, adsk.fusion.CombineFeature)):
                        features_to_delete.append(feature)

            # Delete the features in reverse order (starting with the latest feature)
            for feature in reversed(features_to_delete):
                feature.deleteMe()  # Safely delete the feature entity

            message += 'All features related to Combined_Body have been removed, starting from the shell feature.'

            _wasSurface = False # Reset the flag
        
        else:
            for bRepBody in activeComponent.bRepBodies:
                if bRepBody.name == 'Selected_Body':
                    body = bRepBody
                    break
            # Get the last shell feature created on the body
            shellFeatures = activeComponent.features.shellFeatures
            shellFeature = None
            for feature in shellFeatures:
                if body in feature.bodies:
                    shellFeature = feature
                    break
            
            # Delete the shell feature
            if shellFeature:
                try:
                    shellFeature.deleteMe()
                    message += f"Successfully deleted shell feature from '{body.name}'. "
                except:
                    message += f"Failed to delete shell feature from '{body.name}'. "
        
        debugToConsole(message)

        return True
    
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
        stop(None)


# Function to patch a surface body to create a solid body.
def patchSurface():

    def is_boundary_closed(boundaryEdges):
        # tolerance = 0.0001  # Tolerance for considering edges as connected
        points = set()

        # Collect all start and end points of the edges
        for edge in boundaryEdges:
            points.add((edge.startVertex.geometry.asArray(), edge.endVertex.geometry.asArray()))
        
        # Check if each point is connected to another point
        for point1, point2 in points:
            is_connected = False
            for other_point1, other_point2 in points:
                if (point1 != other_point1) and (point1 == other_point2 or point2 == other_point1):
                    is_connected = True
                    break
            if not is_connected:
                return False
        
        return True

    try:
        # Get the active Fusion 360 application and user interface
        app = adsk.core.Application.get()
        ui  = app.userInterface
        design = app.activeProduct

        # Check if we have a valid design
        if not design or not isinstance(design, adsk.fusion.Design):
            debugToConsole('No active Fusion 360 design found.')
            return None

        activeComponent = design.activeComponent

        # Collect any non-solid bodies (there should only be one in this case)
        surfaceBody = None
        for bRepBody in activeComponent.bRepBodies:
            if not bRepBody.isSolid:
                surfaceBody = bRepBody
                break

        # Ensure that a non-solid body was found
        if not surfaceBody:
            debugToConsole("No non-solid bodies found to patch.")
            return None

        # Create an ObjectCollection to hold the boundary edges
        boundaryEdgesCollection = adsk.core.ObjectCollection.create()

        # Extract the boundary edges of the surface body
        for edge in surfaceBody.edges:
            boundaryEdgesCollection.add(edge)

        if boundaryEdgesCollection.count == 0:
            debugToConsole("No boundary edges found to patch.")
            return None

        # Check if the boundary is closed
        if not is_boundary_closed(boundaryEdgesCollection):
            debugToConsole("The boundary edges are not closed. A valid closed loop is required for patching.")
            return None

        # Create the patch feature input using the boundary edges
        patches = activeComponent.features.patchFeatures
        patchInput = patches.createInput(boundaryEdgesCollection, adsk.fusion.FeatureOperations.NewBodyFeatureOperation)

        # Try to create the patch feature
        patchFeature = patches.add(patchInput)

        # Check if the patch was successful and returned a solid body
        if patchFeature.bodies.count > 0 and patchFeature.bodies.item(0).isSolid:
            # Cache the patch body name
            patchBodyName = patchFeature.bodies.item(0).name
            debugToConsole(f"Successfully patched a solid body: {patchBodyName}.")
        else:
            # Patch failed, clean up the patch feature and log the error
            if patchFeature.bodies.count > 0:
                debugToConsole(f"Patch failed. Deleting failed patch.")
                patchFeature.deleteMe()
            
            return None
    
    except Exception as e:
        ui.messageBox(f"Failed to patch a solid body: {traceback.format_exc()}")
        return None


# Function to turn any surface results into a solid body.
def surfaceToSolid(selectedBody=None):

    app = adsk.core.Application.get()
    ui  = app.userInterface
    design = app.activeProduct

    global _debug

    # Check if we have a valid design
    if not design or not isinstance(design, adsk.fusion.Design):
        debugToConsole('No active Fusion 360 design found.')
        return None

    activeComponent = design.activeComponent

    # Create an ObjectCollection to hold the surface bodies
    surfacesCollection = adsk.core.ObjectCollection.create()

    # Collect any non-solid bodies (there should only be one in this case)
    for bRepBody in activeComponent.bRepBodies:
        if not bRepBody.isSolid:
            surfaceBody = bRepBody
            surfacesCollection.add(bRepBody)
            break

    # Ensure that the surfacesCollection is not empty
    if surfacesCollection.count == 0:
        debugToConsole("No non-solid bodies found to stitch.")
        return None
    else:
        # Create the stitch feature input
        stitches = activeComponent.features.stitchFeatures
        tols = ['0.1 mm', '1 mm', '10 mm']
        ii, patchTried = 0, False
        while ii < len(tols):
            tol = tols[ii]
            stitchInput = stitches.createInput(surfacesCollection, adsk.core.ValueInput.createByString(tol), adsk.fusion.FeatureOperations.NewBodyFeatureOperation)
            tolFailMessage = f"Failed to stitch a solid body from {surfaceBody.name} with tolerance: {tol}."

            # Create the stitch feature
            try:
                stitchFeature = stitches.add(stitchInput)
                if stitchFeature.bodies.count > 0:
                    if stitchFeature.bodies.item(0).isSolid:
                        break
                    else:
                        # Delete the failed stitch feature
                        stitchFeature.deleteMe()
                        if not patchTried:
                            patchTried = True
                            debugToConsole(tolFailMessage + f" Attempting to patch the surface.")
                            patchSurface()
                            ii = 0
                        else:
                            debugToConsole(tolFailMessage)
                            ii += 1
                else:
                    debugToConsole(tolFailMessage)
                    ii += 1
            except:
                if not patchTried:
                    patchTried = True
                    debugToConsole(tolFailMessage + f" Attempting to patch the surface.")
                    patchSurface()
                    ii = 0
                else:
                    debugToConsole(tolFailMessage)
                    ii += 1
    
    if stitchFeature.bodies.count > 0:
        stitchedBody = stitchFeature.bodies.item(0)
        stitchedBody.name = 'Stitched_Body'
        # Cache the stitched body name
        stitchedBodyName = stitchedBody.name
    else:
        debugToConsole(f'Failed to create a solid body from {surfaceBody.name}.')
        return None
    
    # Cache selected body name
    selectedBodyName = selectedBody.name

    toolBodies = adsk.core.ObjectCollection.create()
    toolBodies.add(selectedBody)
    # Cut the selected body from the solid body
    combineFeatures = activeComponent.features.combineFeatures
    combineInput = combineFeatures.createInput(stitchedBody, toolBodies)
    combineInput.operation = adsk.fusion.FeatureOperations.CutFeatureOperation
    combineInput.isKeepToolBodies = True
    combineInput.isNewComponent = False
    combineFeature = combineFeatures.add(combineInput)

    if combineFeature:
        if combineFeature.bodies.count > 0:
            combinedBody = combineFeature.bodies.item(0)
            combinedBody.name = 'Combined_Body'

            # Return the combined body
            return combinedBody
        else:
            debugToConsole(f'Failed to cut {selectedBodyName} from {stitchedBodyName}.')
            return None
    else:
        debugToConsole('Combine feature failed to execute.')
        return None


# Function to create a shell feature for a body.
def createShellFeature(body, thickness, preUndo=True, iteration=None):

    if preUndo:
        if not undoShellFeatures():  # Undo the last features applied to the body in the last iteration
            return None

    # Get the root component of the active design
    app = adsk.core.Application.get()
    ui  = app.userInterface
    design = app.activeProduct

    global _wasSurface
    
    # Check if we have a valid design
    if not design or not isinstance(design, adsk.fusion.Design):
        debugToConsole('No active Fusion 360 design found.')
        return None
    
    # Get the active component
    activeComponent = design.activeComponent
    
    # Check if there are any selected bodies
    if not body:
        debugToConsole('No bodies are selected.')
        return None
    
    if not body.isSolid:
        debugToConsole(f'Body {body.name} is not a solid body and cannot be shelled.')
        return None
    
    # Create a collection of input entities for the shell feature
    inputEntities = adsk.core.ObjectCollection.create()
    inputEntities.add(body)  # Add the selected body to the collection

    # Create a shell feature input
    shellFeatureInput = activeComponent.features.shellFeatures.createInput(inputEntities)

    # Set the outside thickness of the shell
    shellFeatureInput.outsideThickness = adsk.core.ValueInput.createByString(f'{thickness} mm')
    shellFeatureInput.shellType = adsk.fusion.ShellTypes.RoundedOffsetShellType
    shellFeatureInput.isTangentChain = True

    # Create the shell feature
    shellFeature = activeComponent.features.shellFeatures.add(shellFeatureInput)

    # Check if the shell feature was created successfully
    if shellFeature:
        # Check if any of the bodies in the design are not solid
        for bRepBody in activeComponent.bRepBodies:
            if not bRepBody.isSolid:
                bRepBody.name = 'Surface_Shell'
                body = surfaceToSolid(selectedBody=body)  # Convert the surface body to a solid body
                if body:
                    _wasSurface = True
                    break
                else:
                    # debugToConsole(f'Failed to convert {bRepBody.name} to a solid body.')
                    return None
        debugToConsole(f'Successfully created a shell feature for {body.name}.')
        # Return the mass of the shelled body
        return weighComponent() # body.physicalProperties.mass
    else:
        if ui:
            if iteration:
                debugToConsole(f'Failed at iteration {iteration}:\n{traceback.format_exc()}')
            else:
                debugToConsole(f'Failed:\n{traceback.format_exc()}')
        return None


# Apply objective function
def objectiveFunction(solidMass, body, thickness, preUndo=True, iteration=None):
    app = adsk.core.Application.get()

    shellMass = createShellFeature(body, thickness, preUndo=preUndo, iteration=iteration)
    if shellMass:
        return (shellMass - solidMass)**2
    else:
        debugToConsole(f"Failed to apply outside shell feature for {body.name} with thickness {thickness} mm.\nReturning a large objective function value.")
        return 1e6 # None


# Function to optimise the shell thickness of a body to retain the body's mass.
def optimiseThickness(eventArgs, bodies=None):
    app = adsk.core.Application.get()
    ui  = app.userInterface

    try:
        global _initialThickness, _tolerance, _maxIterations, _errMessage, _undoTest

        _initialThickness.value *= 10 # Convert from mm to cm (bodge)
        
        if not bodies:
            debugToConsole('No bodies are selected.')
            return None

        # Optimise the shell thickness of each selected body
        for body in bodies:

            cachedName = body.name
            body.name = 'Selected_Body'
            
            # Check if 'log' directory exists
            scriptDir = os.path.dirname(os.path.realpath(__file__))
            logDir = os.path.join(scriptDir, 'log')
            if not os.path.exists(logDir):
                os.makedirs(logDir)

            # Get the mass of the body
            solidMass = weighComponent() # body.physicalProperties.mass # + 11e-3 # Add 11 g to account for the mass lost in the combine feature

            # Create a logging file
            now = datetime.datetime.now()
            logName = f"{cachedName}_{now.strftime('%d-%m-%Y_%H-%M-%S')}.txt"
            logPath = os.path.join(logDir, logName)
            startMessage = f"Optimisation of {cachedName} shell thickness to maintain mass of {round(1e3*solidMass, 6)} g.\n{now.strftime('%d-%m-%Y %H:%M:%S')}\n"
            with open(logPath, 'w') as logFile:
                logFile.write(startMessage)

            textPalette = ui.palettes.itemById('TextCommands')
            if not textPalette.isVisible:
                textPalette.isVisible = True  # Open the Text Command window if it's not already open
            textPalette.writeText(startMessage)

            # Setup the Nelder-Mead optimisation algorithm
            t0 = timeit.default_timer()
            alpha, gamma, rho, sigma = 1.0, 2.0, 0.5, 0.5 # Reflection, expansion, contraction, shrinkage
            simplex = [_initialThickness.value, _initialThickness.value*1.1, _initialThickness.value*1.2]
            # simplex = [1.5153125, 1.5010937499999994, 1.5158203124999994]
            iteration, iterations = 0, []

            # get initial simplex values
            try:
                simplex_values = []
                for thickness in simplex:
                    value = objectiveFunction(solidMass, body, thickness, iteration=iteration)
                    simplex_values.append(value)
            except Exception as inner_e:
                raise Exception(f"Failed to evaluate objective function for thickness {thickness} mm:\n{inner_e}")

            while iteration <= _maxIterations.value:
                # Sort the simplex values
                sorted_indices = sorted(range(len(simplex_values)), key=lambda i: simplex_values[i]) # This is where the error is
                simplex = [simplex[i] for i in sorted_indices]
                simplex_values = [simplex_values[i] for i in sorted_indices]
                
                centroid = sum(simplex[:-1]) / len(simplex[:-1])
                
                # Reflection
                reflected_thickness = centroid + alpha * (centroid - simplex[-1])
                reflected_value = objectiveFunction(solidMass, body, reflected_thickness, iteration=iteration)
                
                if simplex_values[0] <= reflected_value < simplex_values[-2]:
                    simplex[-1] = reflected_thickness
                    simplex_values[-1] = reflected_value
                # Expansion
                elif reflected_value < simplex_values[0]:
                    expanded_thickness = centroid + gamma * (reflected_thickness - centroid)
                    expanded_value = objectiveFunction(solidMass, body, reflected_thickness, iteration=iteration)
                    if expanded_value < reflected_value:
                        simplex[-1] = expanded_thickness
                        simplex_values[-1] = expanded_value
                    else:
                        simplex[-1] = reflected_thickness
                        simplex_values[-1] = reflected_value
                # Outside contraction
                elif simplex_values[-2] <= reflected_value < simplex_values[-1]:
                    # Contraction
                    contracted_thickness = centroid + rho * (simplex[-1] - centroid)
                    contracted_value = objectiveFunction(solidMass, body, reflected_thickness, iteration=iteration)
                    if contracted_value < simplex_values[-1]:
                        simplex[-1] = contracted_thickness
                        simplex_values[-1] = contracted_value
                    # Shrink
                    else:
                        for i in range(1, len(simplex)):
                            simplex[i] = simplex[0] + sigma * (simplex[i] - simplex[0])
                            simplex_values[i] = objectiveFunction(solidMass, body, reflected_thickness, iteration=iteration)
                # Inside contraction
                elif reflected_value >= simplex_values[-1]:
                    contracted_thickness = centroid - rho * (simplex[-1] - centroid)
                    contracted_value = objectiveFunction(solidMass, body, reflected_thickness, iteration=iteration)
                    if contracted_value < simplex_values[-1]:
                        simplex[-1] = contracted_thickness
                        simplex_values[-1] = contracted_value
                    # Shrink
                    else:
                        for i in range(1, len(simplex)):
                            simplex[i] = simplex[0] + sigma * (simplex[i] - simplex[0])
                            simplex_values[i] = objectiveFunction(solidMass, body, reflected_thickness, iteration=iteration)
                
                debugToConsole(f"Simplex: {simplex}")
                
                t1 = timeit.default_timer()
                iterations.append(simplex[0])
                iteration += 1

                shellMass = createShellFeature(body, simplex[0], preUndo=True, iteration=iteration)
                message = f"Iteration: {iteration}\tThickness: {round(simplex[0], 6)} mm\t Mass: {round(1e3 * shellMass, 6)} g\n"
                # Write the iteration to the log file
                with open(logPath, 'a') as logFile:
                    logFile.write(message)
                
                debugToConsole(message)

                # Check convergence
                if abs(shellMass - solidMass) < _tolerance.value:
                    break

            body.name = cachedName

            debugToConsole(f"Nelder-Mead shell thickness optimisation completed in {iteration} iterations and {round(t1-t0, 3)} seconds.\nOptimal shell thickness for {body.name} is {round(simplex[0], 4)} mm.\nInitial mass: {round(1e3*solidMass, 6)} g\nFinal mass: {round(1e3*shellMass, 6)} g.")

            # Finish the log file
            with open(logPath, 'a') as logFile:
                logFile.write(f"\nOptimal shell thickness for {body.name} is {round(simplex[0], 6)} mm.\nInitial mass: {round(1e3*solidMass, 6)} g\tFinal mass: {round(1e3*shellMass, 6)} g\nOptimisation completed in {iteration} iterations and {round(t1-t0, 3)} seconds.")
    
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
        # write error to log file if it exists
        if logPath:
            with open(logPath, 'a') as logFile:
                logFile.write(f"\nFailed:\n{traceback.format_exc()}")
        stop(None)


# Function to start the add-in.
def run(context):
    app = adsk.core.Application.get()
    ui  = app.userInterface
    
    try:
        # Get the CommandDefinitions collection.
        cmdDefs = ui.commandDefinitions
        
        # Create a button command definition.
        button = cmdDefs.addButtonDefinition('ShellOptimiserButtonid', 
                                                'Shell Optimiser', 
                                                   'Creates and optimises outside shell thickness.',
                                                   './resources/logos/')
        
        # Connect to the command created event.
        CommandCreated = EventHandler()
        button.commandCreated.add(CommandCreated)
        handlers.append(CommandCreated)
        
        # Get the ADD-INS panel in the model workspace. 
        addInsPanel = ui.allToolbarPanels.itemById('SolidScriptsAddinsPanel')
        
        # Add the button to the bottom of the panel.
        buttonControl = addInsPanel.controls.addCommand(button)

        # Get the solid create panel in the model workspace. 
        addInsPanel = ui.allToolbarPanels.itemById('SolidCreatePanel')
        
        # Add the button to the bottom of the panel.
        buttonControl = addInsPanel.controls.addCommand(button)
        buttonControl.isPromotedByDefault = True

        #Adds a toolbar for the MicroChannels
        workSpace = ui.workspaces.itemById('FusionSolidEnvironment')
        tbPanels = workSpace.toolbarPanels

        tbPanel = tbPanels.itemById('MicroPanel')
        if tbPanel:
            tbPanel.deleteMe()
        tbPanel = tbPanels.add('MicroPanel', 'Shell Optimiser', 'SelectPanel', False)

        # Add the button to the bottom of the panel.
        Microtool = tbPanel.controls.addCommand(button)
        Microtool.isPromotedByDefault = True
        
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
        stop(None)


# Function to stop the add-in.
def stop(context):
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface

        # Delete the command definition.
        cmdDef = ui.commandDefinitions.itemById('ShellOptimiserButtonid')
        if cmdDef:
            cmdDef.deleteMe()
        
        # Remove the controls from the panels.
        addInsPanel = ui.allToolbarPanels.itemById('SolidScriptsAddinsPanel')
        if addInsPanel:
            buttonControl = addInsPanel.controls.itemById('ShellOptimiserButtonid')
            if buttonControl:
                buttonControl.deleteMe()
        
        solidCreatePanel = ui.allToolbarPanels.itemById('SolidCreatePanel')
        if solidCreatePanel:
            createPanelControl = solidCreatePanel.controls.itemById('ShellOptimiserButtonid')
            if createPanelControl:
                createPanelControl.deleteMe()
        
        workspace = ui.workspaces.itemById('FusionSolidEnvironment')
        if workspace:
            tbPanels = workspace.toolbarPanels
            tbPanel = tbPanels.itemById('MicroPanel')
            if tbPanel:
                tbPanel.deleteMe()

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


# Event handler.
class EventHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        app = adsk.core.Application.get()
        ui  = app.userInterface

        try:
            event_args = adsk.core.CommandCreatedEventArgs.cast(args)

            # Call global variables that are being editted within the class
            global _initialThickness, _tolerance, _maxIterations, _errMessage, _bodySelection

            cmd = args.command
            cmd.isExecutedWhenPreEmpted = False
            inputs = cmd.commandInputs

            # Show shell image
            _imgshell = inputs.addImageCommandInput('ImageShell', '', './resources/logos/shell.png')
            _imgshell.isFullWidth = True
            _imgshell.isVisible = True

            # Create a body selection input
            _bodySelection = inputs.addSelectionInput('bodySelection', 'Select Body', 'Select the body to apply the shell to.')
            
            # Create a value input for the thickness
            _initialThickness = inputs.addValueInput('initialThickness', 'Initial Thickness', 'mm', adsk.core.ValueInput.createByString('1.0 mm'))

            # Create a value input for the tolerance
            _tolerance = inputs.addValueInput('tolerance', 'Tolerance', 'g', adsk.core.ValueInput.createByString('3.0 g'))

            # Create an integer spinner for the max iterations
            _maxIterations = inputs.addIntegerSpinnerCommandInput('maxIterations', 'Max Iterations', 1, 500, 10, 50)

            # Error message input
            _errMessage = inputs.addTextBoxCommandInput('errMessage', '', '', 2, True)
            _errMessage.isFullWidth = True
            
            # Connect to the execute event.
            onExecute = ExecuteHandler()
            cmd.execute.add(onExecute)
            handlers.append(onExecute)

            # Connect to the validateInputs event.
            onValidateInputs = ValidateInputsHandler()
            cmd.validateInputs.add(onValidateInputs)
            handlers.append(onValidateInputs)
            
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
            stop(None)
            raise

# Event handler for the execute event.
class ExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            app = adsk.core.Application.get()
            ui  = app.userInterface

            eventArgs = adsk.core.CommandEventArgs.cast(args)

            global _errMessage, _bodySelection

            # Create a collection to store the selected bodies
            bodies = adsk.core.ObjectCollection.create()

            # Loop through the selected entities and add them to the collection
            for ii in range(_bodySelection.selectionCount):
                bodies.add(_bodySelection.selection(ii).entity)

            # Call the function to create the shell feature
            optimiseThickness(eventArgs, bodies=bodies)

        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
            stop(None)
            raise

# Event handler for the validateInputs event.
class ValidateInputsHandler(adsk.core.ValidateInputsEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:

            app = adsk.core.Application.get()
            ui  = app.userInterface

            eventArgs = adsk.core.ValidateInputsEventArgs.cast(args)
            
            _errMessage.text = ''

            # Checks that the entered values are greater than 0          
            if _initialThickness.value <= 0:
                _errMessage.text = 'The channel width must be a positive value.'
                eventArgs.areInputsValid = False
                return
            
            if _tolerance.value < 1e-6:
                _errMessage.text = 'The tolerance is too small.'
                eventArgs.areInputsValid = False
                return

            if _maxIterations.value < 1:
                _errMessage.text = 'The max iterations must be at least 1.'
                eventArgs.areInputsValid = False
                return
            
            if _bodySelection.selectionCount < 1:
                _errMessage.text = 'No bodies are selected.'
                eventArgs.areInputsValid = False
                return
            
            if _bodySelection.selectionCount > 1:
                _errMessage.text = 'Only one body can be selected.'
                eventArgs.areInputsValid = False
                return
            
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
            stop(None)
            raise

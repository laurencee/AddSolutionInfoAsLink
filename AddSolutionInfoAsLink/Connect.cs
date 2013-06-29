using System;
using System.Collections.Generic;
using System.Linq;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;

namespace AddSolutionInfoAsLink
{
    /// <summary>The object for implementing an Add-in.</summary>
    /// <seealso class='IDTExtensibility2' />
    public class Connect : IDTExtensibility2, IDTCommandTarget
    {
        private DTE2 _applicationObject;
        private AddIn _addInInstance;

        /// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
        /// <param term='application'>Root object of the host application.</param>
        /// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
        /// <param term='addInInst'>Object representing this Add-in.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _applicationObject = (DTE2)application;
            _addInInstance = (AddIn)addInInst;

            if(connectMode == ext_ConnectMode.ext_cm_UISetup)
            {
                Commands2 commands = (Commands2)_applicationObject.Commands;

                //Place the command on the tools menu.
                //Find the MenuBar command bar, which is the top-level command bar holding all the main menu items:
                CommandBar menuBarCommandBar = ((CommandBars)_applicationObject.CommandBars)["MenuBar"];

                //Find the Tools command bar on the MenuBar command bar:
                CommandBarControl toolsControl = menuBarCommandBar.Controls["Tools"];
                CommandBarPopup toolsPopup = (CommandBarPopup)toolsControl;

                try
                {
                    //Add a command to the Commands collection:
                    Command command = commands.AddNamedCommand2(
                        AddInInstance: _addInInstance, 
                        Name: "AddSolutionInfoAsLink", 
                        ButtonText: "Add SolutionInfo as link",
                        Tooltip: "Add SolutionInfo/GlobalAssemblyInfo file to all non-test projects as a link", 
                        MSOButton: true,
                        Bitmap: 2308);

                    //Add a control for the command to the tools menu:
                    if((command != null) && (toolsPopup != null))
                    {
                        command.AddControl(toolsPopup.CommandBar);
                    }
                }
                catch(ArgumentException)
                {
                    //If we are here, then the exception is probably because a command with that name
                    //  already exists. If so there is no need to recreate the command and we can 
                    //  safely ignore the exception.
                }
            }
        }

        /// <summary>Implements the QueryStatus method of the IDTCommandTarget interface. This is called when the command's availability is updated</summary>
        /// <param term='commandName'>The name of the command to determine state for.</param>
        /// <param term='neededText'>Text that is needed for the command.</param>
        /// <param term='status'>The state of the command in the user interface.</param>
        /// <param term='commandText'>Text requested by the neededText parameter.</param>
        /// <seealso class='Exec' />
        public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
        {
            if (neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
            {
                // Only show the button when a solution that has projects is opened
                if (commandName == "AddSolutionInfoAsLink.Connect.AddSolutionInfoAsLink" && _applicationObject.Solution.IsOpen && _applicationObject.Solution.Projects.Count > 0)
                {
                    status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                    return;
                }
            }
        }

        /// <summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
        /// <param term='commandName'>The name of the command to execute.</param>
        /// <param term='executeOption'>Describes how the command should be run.</param>
        /// <param term='varIn'>Parameters passed from the caller to the command handler.</param>
        /// <param term='varOut'>Parameters passed from the command handler to the caller.</param>
        /// <param term='handled'>Informs the caller if the command was handled or not.</param>
        /// <seealso class='Exec' />
        public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
        {
            handled = false;
            if (executeOption != vsCommandExecOption.vsCommandExecOptionDoDefault || commandName != "AddSolutionInfoAsLink.Connect.AddSolutionInfoAsLink")
                return;

            if (_applicationObject.Solution.IsOpen && _applicationObject.Solution.Projects.Count > 0)
            {
                // Set the solution info file from common name conventions
                ProjectItem solutionInfo = _applicationObject.Solution.FindProjectItem("SolutionInfo.cs") ?? _applicationObject.Solution.FindProjectItem("GlobalAssemblyInfo.cs");

                if (solutionInfo != null)
                {
                    // Filter out test projects and unloaded projects
                    IEnumerable<Project> validProjects = _applicationObject.Solution.Projects.Cast<Project>()
                        .Where(x => x.ConfigurationManager != null &&
                                    x.Name.IndexOf("test", StringComparison.OrdinalIgnoreCase) < 0);


                    foreach (var validProject in validProjects)
                    {
                        // Don't try to add to projects that already have the solution info file added
                        if (validProject.ProjectItems.Cast<ProjectItem>().All(x => x.Name != solutionInfo.Name))
                        {
                            // Add solution info as a link to the project
                            validProject.ProjectItems.AddFromFile(solutionInfo.FileNames[1]);
                        }
                    }
                }
            }

            handled = true;
            return;
        }

        #region unimplemented interface methods

        /// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
        /// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
        }

        /// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />		
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref Array custom)
        {
        }

        /// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref Array custom)
        {
        }

        #endregion
    }
}
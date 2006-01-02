// Comment Reflower Main plugin Connect class
// Copyright (C) 2004  Ian Nowland
// 
// This program is free software; you can redistribute it and/or
// modify it under the terms of the GNU General Public License
// as published by the Free Software Foundation; either version 2
// of the License, or (at your option) any later version.
// 
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program; if not, write to the Free Software
// Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.


using System;
using Microsoft.Office.Core;
using Extensibility;
using System.Runtime.InteropServices;
using EnvDTE;
using System.Windows.Forms;
using CommentReflowerLib;


namespace CommentReflower
{
    /// <summary>
    ///   The object for implementing the Add-in.
    /// </summary>
    /// <seealso class='IDTExtensibility2' />
    [GuidAttribute("D1432269-9ED2-4506-89FA-87E6E005134A"), ProgId("CommentReflower.Connect")]
    public class Connect : Object, Extensibility.IDTExtensibility2, IDTCommandTarget
    {
        /// <summary>
        ///     Implements the constructor for the Add-in object.
        ///     Place your initialization code within this method.
        /// </summary>
        public Connect()
        {
            mCanonicalFileName = Application.LocalUserAppDataPath + @"\..\CommentReflowerSetup.xml";
            try
            {
                mParams = new ParameterSet(mCanonicalFileName);
            }
            catch (Exception)
            {
                mParams = new ParameterSet();
            }
        }

        /// <summary>
        ///      Implements the OnConnection method of the IDTExtensibility2 interface.
        ///      Receives notification that the Add-in is being loaded.
        /// </summary>
        /// <param term='application'>
        ///      Root object of the host application.
        /// </param>
        /// <param term='connectMode'>
        ///      Describes how the Add-in is being loaded.
        /// </param>
        /// <param term='addInInst'>
        ///      Object representing this Add-in.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(
            object application, 
            Extensibility.ext_ConnectMode connectMode, 
            object addInInst, 
            ref System.Array custom)
        {
            mApplicationObject = (_DTE)application;
            mAddInInstance = (AddIn)addInInst;
            if(connectMode == Extensibility.ext_ConnectMode.ext_cm_UISetup)
            {
                object []contextGUIDS = new object[] { };
                Commands commands = mApplicationObject.Commands;
                _CommandBars commandBars = mApplicationObject.CommandBars;

                // When run, the Add-in wizard prepared the registry for the Add-in.
                // At a later time, the Add-in or its commands may become unavailable for reasons such as:
                //   1) You moved this project to a computer other than which is was originally created on.
                //   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
                //   3) You add new commands or modify commands already defined.
                // You will need to re-register the Add-in by building the CommentReflowerSetup project,
                // right-clicking the project in the Solution Explorer, and then choosing install.
                // Alternatively, you could execute the ReCreateCommands.reg file the Add-in Wizard generated in
                // the project directory, or run 'devenv /setup' from a command prompt.
                try
                {
                    //create command objects
                    Command reflowPointCommand = null;
                    Command reflowSelectionCommand = null;
                    Command reflowSettingsCommand = null;

                    for (int i=0; i < 2; i++)
                    {
                        try
                        {
                            reflowPointCommand  = commands.AddNamedCommand(mAddInInstance, 
                                                                           "PointCommentReflower", 
                                                                           "Reflow Comment Containing Cursor", 
                                                                           "Reflows the comment containing the cursor", 
                                                                           false, 
                                                                           104, 
                                                                           ref contextGUIDS, 
                                                                           (int)vsCommandStatus.vsCommandStatusSupported+(int)vsCommandStatus.vsCommandStatusEnabled);
                            break;
                        }
                        catch (Exception)
                        {
                            if (i == 0)
                            {
                                commands.Item("CommentReflower.Connect.PointCommentReflower",-1).Delete();
                            }
                            else
                            {
                                throw;
                            }
                        }
                    }

                    for (int i=0; i < 2; i++)
                    {
                        try
                        {
                            reflowSelectionCommand  = commands.AddNamedCommand(mAddInInstance, 
                                                                               "SelectionCommentReflower", 
                                                                               "Reflow All Comments in Selected", 
                                                                               "Reflows comments in the selected text", 
                                                                               false, 
                                                                               103, 
                                                                               ref contextGUIDS, 
                                                                               (int)vsCommandStatus.vsCommandStatusSupported+(int)vsCommandStatus.vsCommandStatusEnabled);
                            break;
                        }
                        catch (Exception)
                        {
                            if (i==0)
                            {
                                commands.Item("CommentReflower.Connect.SelectionCommentReflower",-1).Delete();
                            }
                            else
                            {
                                throw;
                            }
                        }
                    }
                    for (int i=0; i < 2; i++)
                    {
                        try
                        {
                            reflowSettingsCommand  = commands.AddNamedCommand(mAddInInstance, 
                                                                              "CommentReflowerSettings", 
                                                                              "Comment Reflower Settings", 
                                                                              "Settings dialog for Comment Reflower", 
                                                                              false, 
                                                                              102, 
                                                                              ref contextGUIDS, 
                                                                              (int)vsCommandStatus.vsCommandStatusSupported+(int)vsCommandStatus.vsCommandStatusEnabled);
                            break;
                        }
                        catch (Exception)
                        {
                            if (i==0)
                            {
                                commands.Item("CommentReflower.Connect.CommentReflowerSettings",-1).Delete();
                            }
                            else
                            {
                                throw;
                            }
                        }
                    }

                    reflowSettingsCommand.AddControl(commandBars["Tools"],1);
                    reflowSelectionCommand.AddControl(commandBars["Tools"],1);
                    reflowPointCommand.AddControl(commandBars["Tools"],1).BeginGroup = true;

                    reflowSelectionCommand.AddControl(commandBars["Code Window"],1);
                    reflowPointCommand.AddControl(commandBars["Code Window"],1);
                }
                catch(System.Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
            }
        }

        /// <summary>
        ///     Implements the OnDisconnection method of the IDTExtensibility2 interface.
        ///     Receives notification that the Add-in is being unloaded.
        /// </summary>
        /// <param term='disconnectMode'>
        ///      Describes how the Add-in is being unloaded.
        /// </param>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
        {
        }

        /// <summary>
        ///      Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
        ///      Receives notification that the collection of Add-ins has changed.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnAddInsUpdate(ref System.Array custom)
        {
        }

        /// <summary>
        ///      Implements the OnStartupComplete method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application has completed loading.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref System.Array custom)
        {
        }

        /// <summary>
        ///      Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application is being unloaded.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref System.Array custom)
        {
        }

        /// <summary>
        ///      Implements the QueryStatus method of the IDTCommandTarget interface.
        ///      This is called when the command's availability is updated
        /// </summary>
        /// <param term='commandName'>
        ///     The name of the command to determine state for.
        /// </param>
        /// <param term='neededText'>
        ///     Text that is needed for the command.
        /// </param>
        /// <param term='status'>
        ///     The state of the command in the user interface.
        /// </param>
        /// <param term='commandText'>
        ///     Text requested by the neededText parameter.
        /// </param>
        /// <seealso class='Exec' />
        public void QueryStatus(
            string commandName, 
            EnvDTE.vsCommandStatusTextWanted neededText, 
            ref EnvDTE.vsCommandStatus status, 
            ref object commandText)
        {
            if(neededText == EnvDTE.vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
            {
                if(commandName == "CommentReflower.Connect.PointCommentReflower")
                {
                    if ((mApplicationObject.ActiveDocument == null) ||
                        (mParams.getBlocksForFileName(mApplicationObject.ActiveDocument.Name).Count == 0))
                    {
                        status = vsCommandStatus.vsCommandStatusUnsupported;
                    }
                    else
                    {
                        status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | 
                                                  vsCommandStatus.vsCommandStatusEnabled;
                    }
                }
                else if(commandName == "CommentReflower.Connect.SelectionCommentReflower")
                {
                    if ((mApplicationObject.ActiveDocument == null) ||
                        (mParams.getBlocksForFileName(mApplicationObject.ActiveDocument.Name).Count == 0))
                    {
                        status = vsCommandStatus.vsCommandStatusUnsupported;
                    }
                    else
                    {
                        TextSelection sel = (TextSelection)mApplicationObject.ActiveDocument.Selection;
                        if (sel.IsEmpty)
                        {
                            status = vsCommandStatus.vsCommandStatusSupported;
                        }
                        else
                        {
                            status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                        }
                    }
                }
                else if(commandName == "CommentReflower.Connect.CommentReflowerSettings")
                {
                    status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                }
                else if(commandName == "CommentReflower.Connect.ParameterAligner")
                {
                    if ((mApplicationObject.ActiveDocument == null) ||
                        (mParams.getBlocksForFileName(mApplicationObject.ActiveDocument.Name).Count == 0))
                    {
                        status = vsCommandStatus.vsCommandStatusUnsupported;
                    }
                    else
                    {
                        status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                    }
                }
            }
        }


        /// <summary>
        ///      Implements the Exec method of the IDTCommandTarget interface.
        ///      This is called when the command is invoked.
        /// </summary>
        /// <param term='commandName'>
        ///     The name of the command to execute.
        /// </param>
        /// <param term='executeOption'>
        ///     Describes how the command should be run.
        /// </param>
        /// <param term='varIn'>
        ///     Parameters passed from the caller to the command handler.
        /// </param>
        /// <param term='varOut'>
        ///     Parameters passed from the command handler to the caller.
        /// </param>
        /// <param term='handled'>
        ///     Informs the caller if the command was handled or not.
        /// </param>
        /// <seealso class='Exec' />
        public void Exec(
            string commandName, 
            EnvDTE.vsCommandExecOption executeOption, 
            ref object varIn, 
            ref object varOut, 
            ref bool handled)
        {
            handled = false;
            if(executeOption == EnvDTE.vsCommandExecOption.vsCommandExecOptionDoDefault)
            {
                if(commandName == "CommentReflower.Connect.PointCommentReflower")
                {
                    handled = true;
                    TextSelection sel = (TextSelection)mApplicationObject.ActiveDocument.Selection;

                    sel.DTE.UndoContext.Open("Reflowing comments in selection",false);
                    try
                    {
                        if (!CommentReflowerObj.WrapBlockContainingPoint(mParams,
                                                                         mApplicationObject.ActiveDocument.Name,
                                                                         sel.ActivePoint))
                        {
                            MessageBox.Show("No comment found");
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Exception processing command: " + e.ToString());
                    }
                    finally
                    {
                        sel.DTE.UndoContext.Close();
                    }
                    return;
                }
                else if(commandName == "CommentReflower.Connect.SelectionCommentReflower")
                {
                    handled = true;
                    TextSelection sel = (TextSelection)mApplicationObject.ActiveDocument.Selection;
                    sel.DTE.UndoContext.Open("Reflowing comments in selection",false);
                    try
                    {
                        if (!CommentReflowerObj.WrapAllBlocksInSelection(mParams,
                                                                         mApplicationObject.ActiveDocument.Name,
                                                                         sel.TopPoint,
                                                                         sel.BottomPoint))
                        {
                            MessageBox.Show("No comment found");
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Exception processing command: " + e.ToString());
                    }
                    finally
                    {
                        sel.DTE.UndoContext.Close();
                    }
                    return;
                }
                else if(commandName == "CommentReflower.Connect.CommentReflowerSettings")
                {
                    handled = true;
                    CommentReflowerSetup setup = new CommentReflowerSetup(mParams,
                                                                          mApplicationObject,
                                                                          mAddInInstance);
                    if (setup.ShowDialog() == DialogResult.OK)
                    {
                        mParams = setup.mpset;
                        mParams.writeToXmlFile(mCanonicalFileName);
                    }
                    setup.Dispose();
                    return;
                }
                else if(commandName == "CommentReflower.Connect.ParameterAligner")
                {
                    TextSelection sel = (TextSelection)mApplicationObject.ActiveDocument.Selection;
                    sel.DTE.UndoContext.Open("Reflowing comments in selection",false);
                    try
                    {
                        EnvDTE.EditPoint outPt;
                        if (!ParameterAlignerObj.go(sel.ActivePoint, out outPt))
                        {
                            MessageBox.Show("Function call not found.");
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Exception processing command: " + e.ToString());
                    }
                    finally
                    {
                        sel.DTE.UndoContext.Close();
                    }
                }
            }
        }
        private _DTE mApplicationObject;
        private AddIn mAddInInstance;
        private ParameterSet mParams;
        private string mCanonicalFileName;
    }
}

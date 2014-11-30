using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using System.Diagnostics.CodeAnalysis;

using Microsoft.VSPowerToys.ResourceRefactor;
using System.Windows.Forms;
using EnvDTE80;
using EnvDTE;
using Microsoft.VSPowerToys.ResourceRefactor.Common;

namespace ResourceRefactoring.ResourceRefactor2013
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidResourceRefactor2013PkgString)]
    public sealed class ResourceRefactor2013Package : Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public ResourceRefactor2013Package()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }



        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Debug.WriteLine (string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(GuidList.guidResourceRefactor2013CmdSet, (int)PkgCmdIDList.cmdidExtractResource);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID );
                mcs.AddCommand( menuItem );
            }
        }
        #endregion

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            DTE2 applicationObject = Package.GetGlobalService(typeof(DTE)) as DTE2;

            TextSelection selection = (TextSelection)(applicationObject.ActiveDocument.Selection);
            if (applicationObject.ActiveDocument.ProjectItem.Object != null)
            {

                BaseHardCodedString stringInstance = BaseHardCodedString.GetHardCodedString(applicationObject.ActiveDocument);

                if (stringInstance == null)
                {
                    MessageBox.Show(
                        Strings.UnsupportedFile + " (" + applicationObject.ActiveDocument.Language + ")",
                        Strings.WarningTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                MatchResult scanResult = stringInstance.CheckForHardCodedString(
                   selection.Parent,
                   selection.AnchorPoint.AbsoluteCharOffset - 1,
                   selection.BottomPoint.AbsoluteCharOffset - 1);

                if (!scanResult.Result && selection.AnchorPoint.AbsoluteCharOffset < selection.BottomPoint.AbsoluteCharOffset)
                {
                    scanResult.StartIndex = selection.AnchorPoint.AbsoluteCharOffset - 1;
                    scanResult.EndIndex = selection.BottomPoint.AbsoluteCharOffset - 1;
                    scanResult.Result = true;
                }
                if (scanResult.Result)
                {
                    stringInstance = stringInstance.CreateInstance(applicationObject.ActiveDocument.ProjectItem, scanResult.StartIndex, scanResult.EndIndex);
                    if (stringInstance != null && stringInstance.Parent != null)
                    {
                        PerformAction(stringInstance);
                    }
                }
                else
                {
                    MessageBox.Show(Strings.NotStringLiteral, Strings.WarningTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>Shows "Extract to Resource" dialog to user and extracts the string to resource files</summary>
        /// <param name="stringInstance">String to refactor</param>
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions")]
        private void PerformAction(BaseHardCodedString stringInstance)
        {
            ExtractToResourceActionSite site = new ExtractToResourceActionSite(stringInstance);
            if (site.ActionObject != null)
            {
                RefactorStringDialog dialog = new RefactorStringDialog();
                dialog.ShowDialog(site);
            }
            else
            {
                MessageBox.Show(Strings.UnsupportedFile, Strings.WarningTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}

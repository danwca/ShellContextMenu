/********************************** Module Header **********************************\
Module Name:  FileContextMenuExt.cs
Project:      CSShellExtContextMenuHandler
Copyright (c) Microsoft Corporation.

The FileContextMenuExt.cs file defines a context menu handler by implementing the 
IShellExtInit and IContextMenu interfaces.

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***********************************************************************************/

#region Using directives
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using CSAppShellExt.Properties;
using System.Drawing;

using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Linq;
using System.Diagnostics;

#endregion


namespace ShellExtContextMenuHandler
{
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("B1F1405D-94A1-4692-B72F-FC8CAF8B8700"), ComVisible(true)]
    public class MyShellMenu : IShellExtInit, IContextMenu
    {

        // NOTE parts of this are purposely "wordy" to make the 
        //   code more recognizable to VB or others not fluent in C#


        // Everything you need to change should be between the star bars
        //
        //         *****  *****  *****
   

        // {PL} list of file exts to associate with this handler
        // ...leave it as an array and use lowercase only
        private static string[] MyExts = {".txt", ".foo", ".log", ".lst", ".bar"};

        // The base name of the app to run - do not add or hard code a path!
        // The code will figure out where it is.
        private string AppName = "MyMainShApp.exe";

        private string menuText = "Open in MainApp";
        private string verb = "openmyapp";
        private string verbCanonicalName = "OpenInMainApp";
        private string verbHelpText = "Open Files in MainApp";
        
        // If you want to send the first file ONLY, change to TRUE
        private bool FeedFirstFileOnly = false;


        // simple step logger for debugging...you'll find it in MyDocuments
        private string LogFile = "ShellExtDebug.log";
        // Be sure to set to false for release version
        private bool LogActive  = true;
        // T/F whether you want to Append or Start a new log file each time (single entry)
        private bool LogAppend  = false;

        // T/F, to add a seperator after your entry
        private bool AddSeperator  = false;

        // name of the image resource to use
        private string bmpName = "curlyBraces";

        // oldBackColor == the back color of the Resource image,
        //   typically Transparent or Fuchsia
        // newBackColor == Color to change all oldBackColor pixels to
        //   null (default) == use default ContextMenuStrip.BackColor
        private Color oldBackColor = Color.Fuchsia;
        private Color? newBackColor = null;

        // End of variable block and stuff to change
        // *****************************************


        // The list of the selected files
        private List<string> selectedFiles;

        private string targetApp  = "";
        private IntPtr menuBmp = IntPtr.Zero;
        private uint IDM_DISPLAY = 0;

        private string AsmHandlerName = "";

        // ToDo: change ctor to match class name
        // 'class name' + '()'
        public MyShellMenu()
        {
            LogFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    LogFile);

            Type t = this.GetType();
            AsmHandlerName = String.Format("{0}.{1}.{2} Class", t.Assembly.GetName().Name, t.Namespace, t.Name);

            OpenLog();
            WriteToLog("@ctor");

            selectedFiles = new List<string>();

           // almost always reports System/32
            WriteToLog(String.Format("Current dir: {0}",
                                     Environment.CurrentDirectory));
                        
            //WriteToLog(String.Format("    AsmName: {0}", AsmHandlerName));
            
            string mypath = Path.GetDirectoryName( Assembly.GetAssembly( this.GetType()).Location);
            targetApp = Path.Combine(mypath, Path.GetFileName(AppName));

            if (File.Exists(targetApp))
                WriteToLog("  Target found: " + targetApp);
            else
                WriteToLog(String.Format("  Target app NOT found: [{0}]",
                                     targetApp));

            // Load the bitmap for the menu item.
            //Bitmap bmp;

            if (!string.IsNullOrEmpty(bmpName))
            { 
                Bitmap bmp = GetPseudoTransparentBitmap();
                menuBmp = bmp.GetHbitmap();
                //try
                //{
                //    bmp = (Bitmap)Resources.ResourceManager.GetObject(bmpName);
                //    if (bmp != null)
                //    {
                //        bmp.MakeTransparent(oldBackColor);
                //        this.menuBmp = bmp.GetHbitmap();
                //    }
                //    else WriteToLog("bmp is null");
                //}
                //catch (Exception ex)
                //{
                //    WriteToLog(ex.Message);
                //}
            }
        }

        // ToDo: change dtor to match class name
        // '~' + 'class name' + '()'
        ~MyShellMenu()
        {
            WriteToLog("@Finalize/dtor");
            if (this.menuBmp != IntPtr.Zero)
            {
                NativeMethods.DeleteObject(this.menuBmp);
                this.menuBmp = IntPtr.Zero;
            }
        }

        //  the bitmaps do not show up transparent; even Adobe has problems with this.
        //  Neither PNG nor ICONs work.
        //  Use ContextMenuStrip backcolor (default) and replace the specified BG pixels
        private Bitmap GetPseudoTransparentBitmap()
        {
            Bitmap bmp = (Bitmap)Resources.ResourceManager.GetObject(bmpName);
            Int32 px;

            if (newBackColor.HasValue == false)
            {
                using (System.Windows.Forms.ContextMenuStrip cms = new System.Windows.Forms.ContextMenuStrip())
                { 
                    newBackColor = cms.BackColor;
                }
            }
            Int32 pxT = oldBackColor.ToArgb();

            for (Int32 w = 0; w < bmp.Width; w++)
            {
                for (Int32 h = 0; h < bmp.Height; h++)
                {
                    px = bmp.GetPixel(w, h).ToArgb();
                    //bmp.SetPixel(w, h, (px == pxT) ? newBackColor.Value : oldBackColor);
                    if (pxT == px) bmp.SetPixel(w, h, newBackColor.Value);
                }
            }
            return bmp;
        }


        void OnExecuteMenuCommand(IntPtr hWnd)
        {
            string args = "";
            string Quote = "\"";

            WriteToLog("@ExecuteMenuCommand");

            if (String.IsNullOrEmpty(targetApp))
            {
                WriteToLog("cannot find targetApp: " + targetApp);
            }
                
            WriteToLog("  FeedFirstFileOnly == " + FeedFirstFileOnly.ToString());
            WriteToLog("  File Count == " + selectedFiles.Count.ToString());

            if (selectedFiles.Count > 0)
            {
                if (FeedFirstFileOnly)
                    args = String.Format(" {0}{1}{0} ", Quote, selectedFiles[0]);
                else
                {
                    for (int n = 0; n < selectedFiles.Count; n++)
                    {
                        args += String.Format(" {0}{1}{0} ", Quote, selectedFiles[0]);
                    }
                }
                
                WriteToLog(String.Format("  Cmd to execute: {0}", targetApp));
                WriteToLog(String.Format("  Args: {0}", args));

                ProcessStartInfo pStartInfo = new ProcessStartInfo(targetApp);
                Process proc;

                try
                {
                    pStartInfo.Arguments = args;
                    pStartInfo.WorkingDirectory = Path.GetDirectoryName(targetApp);
                    proc = Process.Start(pStartInfo);
                    if (proc == null)
                        WriteToLog("Process is null");
                    else
                        WriteToLog("Process started; ID = " + proc.Id.ToString());
                }
                catch (Exception ex)
                {
                    WriteToLog(ex.Message);
                }

                //// alternatively, invoke a Class or Form to 
                //// do Wonderful Things here in the DLL
                //ShellForm myFrm = new ShellForm();
                //myFrm.lbfiles.Items.AddRange(selectedFiles.ToArray);
                //myFrm.Show();  
                // (ill-advised)

            }
            else
            {
                // might mean something is broke in Initialize
                WriteToLog("Selected Files Count = 0 ");
            }
            WriteToLog("Workingset: " + Environment.WorkingSet.ToString());
            WriteToLog("***  Normal Termination  ***");

        }


        #region Shell Extension Registration

        // used by RegAsm for file associations
        [ComRegisterFunction()]
        public static void Register(Type T)
        {

            string AsmHandlerName = String.Format("{0}.{1}.{2} Class", 
                        T.Assembly.GetName().Name, T.Namespace, T.Name);
            try
            {
                foreach (string s in MyExts)
                { 
                    ShellExtReg.RegisterShellExtContextMenuHandler(T.GUID, s,
                                        AsmHandlerName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message); // Log the error
                throw;  // Re-throw the exception
            }
        }

        // used by RegAsm to remove file associations
        [ComUnregisterFunction()]
        public static void Unregister(Type t)
        {
            try
            {
                foreach (string s in MyExts)
                { 
                    ShellExtReg.UnregisterShellExtContextMenuHandler(t.GUID, s);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message); // Log the error
                throw;  // Re-throw the exception
            }
        }

        // get the master list of extensions
        public static string[] GetExtensions()
        {
            return MyExts;
        }

        public static void RegisterExtensions(Type T, string[] exts)
        { 
            string thisExt = "";

            // detach existing ones
            Unregister(T);
            
            string AsmHandlerName = String.Format("{0}.{1}.{2} Class", 
                        T.Assembly.GetName().Name, T.Namespace, T.Name);

            try
            {
                foreach (string s in exts)
                {
                    thisExt = s.ToLowerInvariant();
                    if (MyExts.Contains(thisExt))
                    {
                        ShellExtReg.RegisterShellExtContextMenuHandler(T.GUID, 
                            thisExt, AsmHandlerName);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message); // Log the error
                throw;
            }
        
        }



        #endregion

        #region IShellExtInit Members

        /// <summary>
        /// Initialize the context menu handler.
        /// </summary>
        /// <param name="pidlFolder">
        /// A pointer to an ITEMIDLIST structure that uniquely identifies a folder.
        /// </param>
        /// <param name="pDataObj">
        /// A pointer to an IDataObject interface object that can be used to retrieve 
        /// the objects being acted upon.
        /// </param>
        /// <param name="hKeyProgID">
        /// The registry key for the file object or folder type.
        /// </param>
        public void Initialize(IntPtr pidlFolder, IntPtr pDataObj, IntPtr hKeyProgID)
        {
            if (pDataObj == IntPtr.Zero)
            {
                throw new ArgumentException();
            }

            FORMATETC fe = new FORMATETC();
            fe.cfFormat = (short)CLIPFORMAT.CF_HDROP;
            fe.ptd = IntPtr.Zero;
            fe.dwAspect = DVASPECT.DVASPECT_CONTENT;
            fe.lindex = -1;
            fe.tymed = TYMED.TYMED_HGLOBAL;
            STGMEDIUM stm = new STGMEDIUM();

            // The pDataObj pointer contains the objects being acted upon. In this 
            // example, we get an HDROP handle for enumerating the selected files 
            // and folders.
            IDataObject dataObject = (IDataObject)Marshal.GetObjectForIUnknown(pDataObj);
            dataObject.GetData(ref fe, out stm);

            try
            {
                // Get an HDROP handle.
                IntPtr hDrop = stm.unionmember;
                if (hDrop == IntPtr.Zero)
                {
                    throw new ArgumentException();
                }

                // Determine how many files are involved in this operation.
                uint nFiles = NativeMethods.DragQueryFile(hDrop, UInt32.MaxValue, null, 0);

                // This code sample displays the custom context menu item when only 
                // one file is selected. 
                if (nFiles > 0)
                {
                    // Get the path of the file.
                    StringBuilder sbfileName;
                    string f = "";

                    for (UInt32 n = 0; n < nFiles; n++)
                    {
                        sbfileName = new StringBuilder(260);

                        if (0 == NativeMethods.DragQueryFile(hDrop, 0, sbfileName,
                            sbfileName.Capacity))
                        {
                            Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                        }

                        f = sbfileName.ToString();
                        selectedFiles.Add(f);
                        WriteToLog("    added: " + f);
                    }

                }
                else
                {
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                }        
             
            }
            finally
            {
                NativeMethods.ReleaseStgMedium(ref stm);
            }
        }

        #endregion

        #region IContextMenu Members

        /// <summary>
        /// Add commands to a shortcut menu.
        /// </summary>
        /// <param name="hMenu">A handle to the shortcut menu.</param>
        /// <param name="iMenu">
        /// The zero-based position at which to insert the first new menu item.
        /// </param>
        /// <param name="idCmdFirst">
        /// The minimum value that the handler can specify for a menu item ID.
        /// </param>
        /// <param name="idCmdLast">
        /// The maximum value that the handler can specify for a menu item ID.
        /// </param>
        /// <param name="uFlags">
        /// Optional flags that specify how the shortcut menu can be changed.
        /// </param>
        /// <returns>
        /// If successful, returns an HRESULT value that has its severity value set 
        /// to SEVERITY_SUCCESS and its code value set to the offset of the largest 
        /// command identifier that was assigned, plus one.
        /// </returns>
        public int QueryContextMenu(IntPtr hMenu,
                                uint iMenu, 
                                uint idCmdFirst,
                                uint idCmdLast,
                                uint uFlags) 
        {
            // If uFlags include CMF_DEFAULTONLY then we should not do anything.
            if (((uint)CMF.CMF_DEFAULTONLY & uFlags) != 0)
            {
                return WinError.MAKE_HRESULT(WinError.SEVERITY_SUCCESS, 0, 0);
            }

            // Use either InsertMenu or InsertMenuItem to add menu items.
            MENUITEMINFO mii = new MENUITEMINFO();
            mii.cbSize = (uint)Marshal.SizeOf(mii);
            mii.fMask = MIIM.MIIM_BITMAP | MIIM.MIIM_STRING | MIIM.MIIM_FTYPE | 
                MIIM.MIIM_ID | MIIM.MIIM_STATE;
            mii.wID = idCmdFirst + IDM_DISPLAY;
            mii.fType = MFT.MFT_STRING;
            mii.dwTypeData = this.menuText;
            mii.fState = MFS.MFS_ENABLED;
            mii.hbmpItem = this.menuBmp;
            if (!NativeMethods.InsertMenuItem(hMenu, iMenu, true, ref mii))
            {
                return Marshal.GetHRForLastWin32Error();
            }

            // Add a separator.
            if (AddSeperator)
            { 
                MENUITEMINFO sep = new MENUITEMINFO();
                sep.cbSize = (uint)Marshal.SizeOf(sep);
                sep.fMask = MIIM.MIIM_TYPE;
                sep.fType = MFT.MFT_SEPARATOR;
                if (!NativeMethods.InsertMenuItem(hMenu, iMenu + 1, true, ref sep))
                {
                    return Marshal.GetHRForLastWin32Error();
                }
            }


            // Return an HRESULT value with the severity set to SEVERITY_SUCCESS. 
            // Set the code value to the offset of the largest command identifier 
            // that was assigned, plus one (1).
            return WinError.MAKE_HRESULT(WinError.SEVERITY_SUCCESS, 0,
                IDM_DISPLAY + 1);
        }

        /// <summary>
        /// Carry out the command associated with a shortcut menu item.
        /// </summary>
        /// <param name="pici">
        /// A pointer to a CMINVOKECOMMANDINFO or CMINVOKECOMMANDINFOEX structure 
        /// containing information about the command. 
        /// </param>
        public void InvokeCommand(IntPtr pici)
        {
            bool isUnicode = false;

            // Determine which structure is being passed in, CMINVOKECOMMANDINFO or 
            // CMINVOKECOMMANDINFOEX based on the cbSize member of lpcmi. Although 
            // the lpcmi parameter is declared in Shlobj.h as a CMINVOKECOMMANDINFO 
            // structure, in practice it often points to a CMINVOKECOMMANDINFOEX 
            // structure. This struct is an extended version of CMINVOKECOMMANDINFO 
            // and has additional members that allow Unicode strings to be passed.
            CMINVOKECOMMANDINFO ici = (CMINVOKECOMMANDINFO)Marshal.PtrToStructure(
                pici, typeof(CMINVOKECOMMANDINFO));
            CMINVOKECOMMANDINFOEX iciex = new CMINVOKECOMMANDINFOEX();
            if (ici.cbSize == Marshal.SizeOf(typeof(CMINVOKECOMMANDINFOEX)))
            {
                if ((ici.fMask & CMIC.CMIC_MASK_UNICODE) != 0)
                {
                    isUnicode = true;
                    iciex = (CMINVOKECOMMANDINFOEX)Marshal.PtrToStructure(pici,
                        typeof(CMINVOKECOMMANDINFOEX));
                }
            }

            // Determines whether the command is identified by its offset or verb.
            // There are two ways to identify commands:
            // 
            //   1) The command's verb string 
            //   2) The command's identifier offset
            // 
            // If the high-order word of lpcmi->lpVerb (for the ANSI case) or 
            // lpcmi->lpVerbW (for the Unicode case) is nonzero, lpVerb or lpVerbW 
            // holds a verb string. If the high-order word is zero, the command 
            // offset is in the low-order word of lpcmi->lpVerb.


            WriteToLog("@InvokeCommand");

            // For the ANSI case, if the high-order word is not zero, the command's 
            // verb string is in lpcmi->lpVerb. 
            if (!isUnicode && NativeMethods.HighWord(ici.verb.ToInt32()) != 0)
            {
                WriteToLog("   Is Ansi");
                // Is the verb supported by this context menu extension?
                if (Marshal.PtrToStringAnsi(ici.verb) == this.verb)
                {
                    OnExecuteMenuCommand(ici.hwnd);
                }
                else
                {
                    // If the verb is not recognized by the context menu handler, it 
                    // must return E_FAIL to allow it to be passed on to the other 
                    // context menu handlers that might implement that verb.
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                }
            }

            // For the Unicode case, if the high-order word is not zero, the 
            // command's verb string is in lpcmi->lpVerbW. 
            else if (isUnicode && NativeMethods.HighWord(iciex.verbW.ToInt32()) != 0)
            {
                WriteToLog("   Is Unicode");
                // Is the verb supported by this context menu extension?
                if (Marshal.PtrToStringUni(iciex.verbW) == this.verb)
                {
                    OnExecuteMenuCommand(ici.hwnd);
                }
                else
                {
                    // If the verb is not recognized by the context menu handler, it 
                    // must return E_FAIL to allow it to be passed on to the other 
                    // context menu handlers that might implement that verb.
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                }
            }

            // If the command cannot be identified through the verb string, then 
            // check the identifier offset.
            else
            {
                WriteToLog("   use IDM_DISPLAY");
                // Is the command identifier offset supported by this context menu 
                // extension?
                if (NativeMethods.LowWord(ici.verb.ToInt32()) == IDM_DISPLAY)
                {
                    OnExecuteMenuCommand(ici.hwnd);
                }
                else
                {
                    // If the verb is not recognized by the context menu handler, it 
                    // must return E_FAIL to allow it to be passed on to the other 
                    // context menu handlers that might implement that verb.
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                }
            }
        }

        /// <summary>
        /// Get information about a shortcut menu command, including the help string 
        /// and the language-independent, or canonical, name for the command.
        /// </summary>
        /// <param name="idCmd">Menu command identifier offset.</param>
        /// <param name="uFlags">
        /// Flags specifying the information to return. This parameter can have one 
        /// of the following values: GCS_HELPTEXTA, GCS_HELPTEXTW, GCS_VALIDATEA, 
        /// GCS_VALIDATEW, GCS_VERBA, GCS_VERBW.
        /// </param>
        /// <param name="pReserved">Reserved. Must be IntPtr.Zero</param>
        /// <param name="pszName">
        /// The address of the buffer to receive the null-terminated string being 
        /// retrieved.
        /// </param>
        /// <param name="cchMax">
        /// Size of the buffer, in characters, to receive the null-terminated string.
        /// </param>
        public void GetCommandString(UIntPtr idCmd, uint uFlags,
                                    IntPtr pReserved,
                                    StringBuilder pszName,
                                    uint cchMax)
        {

                      
            if (idCmd.ToUInt32() == IDM_DISPLAY)
            {
                switch ((GCS)uFlags)
                {
                    case GCS.GCS_VERBW:
                        if (this.verbCanonicalName.Length > cchMax - 1)
                        {
                            Marshal.ThrowExceptionForHR(WinError.STRSAFE_E_INSUFFICIENT_BUFFER);
                        }
                        else
                        {
                            pszName.Clear();
                            pszName.Append(this.verbCanonicalName);
                        }
                        break;

                    case GCS.GCS_HELPTEXTW:
                        if (this.verbHelpText.Length > cchMax - 1)
                        {
                            Marshal.ThrowExceptionForHR(WinError.STRSAFE_E_INSUFFICIENT_BUFFER);
                        }
                        else
                        {
                            pszName.Clear();
                            pszName.Append(this.verbHelpText);
                        }
                        break;
                }
            }
        }

        #endregion
    
        #region Log methods

        private void OpenLog()
        {
            if (this.LogActive)
            { 
                using (StreamWriter sw = new StreamWriter(this.LogFile, this.LogAppend))
                {
                    sw.WriteLine("************");
                    sw.WriteLine(AsmHandlerName);
                    sw.WriteLine(DateTime.Now.ToString("G"));
                }
            }

        }

        private void WriteToLog(string txt)
        {
            if (this.LogActive)
            { 
                using (StreamWriter sw = new StreamWriter(this.LogFile, true))
                {
                    sw.WriteLine(txt);
                }
            }

        }

    #endregion
        
    }

}
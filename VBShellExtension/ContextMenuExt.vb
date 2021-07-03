'****************************** Module Header ******************************'
' Module Name:  FileContextMenuExt.vb
' Project:      VBShellExtContextMenuHandler
' Copyright (c) Microsoft Corporation.
' 
' The FileContextMenuExt.vb file defines a context menu handler by 
' implementing the IShellExtInit and IContextMenu interfaces.
' 
' This source is subject to the Microsoft Public License.
' See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
' All other rights reserved.
' 
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
' EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
' WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
'***************************************************************************'


' see CodeProject: http://www.codeproject.com/Articles/174369/How-to-Write-Windows-Shell-Extension-with-NET-Lang
' also:  stackoverflow:  http://stackoverflow.com/a/27230246/1070452
' 
'

#Region "Imports directives"

Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Drawing

'{PL} Added:
Imports System.Diagnostics
Imports System.Collections.Generic
Imports System.IO
Imports System.Reflection

#End Region


Namespace ShellExtContextMenuHandler


    <Guid("679944C8-FEC8-446A-8089-64E0DE515898"),
    ClassInterface(ClassInterfaceType.None),
    ComVisible(True)>
    Public Class MyShellMenu
        '...
        Implements IShellExtInit, IContextMenu


        ' ignore - the code works this out for itself now

        ' The Assembly (DLL), this NameSpace and this Class name
        ' Get assembly name used in Project->Properties->Application
        ' Namespace and Class are just above
        ' Regasm and ShellReg use these

        '' The class figures these out for itself...
        'Private Shared AsmName As String = "MyShellExt"
        'Private Shared NameSpcName As String = "ShellExtContextMenuHandler"
        'Private Shared ClassName As String = "MyShellMenu"


        ' Everything you need to change should be between the star bars
        '*****  *****  *****

        ' {PL} list of file exts to associate with this handler
        ' ...leave it as an array and use lowercase only
        Private Shared MyExts As String() = {".txt", ".foo", ".zig", ".zag", ".bar"}

        ' The base name of the app to run - do not add or hard code a path!
        ' The code will figure out where it is.
        Private AppName As String = "MyMainShApp.exe"

        ' menu text and such
        Private menuText As String = "Open in MainApp"
        Private verb As String = "openmyapp"
        Private verbCanonicalName As String = "OpenInMainApp"
        Private verbHelpText As String = "Open Files in MainApp"

        ' If you want to send the first file ONLY, change to TRUE
        '   use a context menu if you only need one file though
        Private FeedFirstFileOnly As Boolean = False

        ' simple step logger for debugging...you'll find it in MyDocuments
        Private LogFile As String = "ShellExtDebug.log"
        ' Be sure to set to false for release version
        Private LogActive As Boolean = True
        ' T/F whether you want to Append or Start a new log file each time (single entry)
        Private LogAppend As Boolean = False

        ' T/F, to add a seperator after your entry
        Private AddSeperator As Boolean = False

        ' name of the bmp in Resources you wish to use
        Private bmpName As String = "starBlue"

        ' oldBackColor == the back color of the Resource image,
        '   typically  Fuchsia or Transparent 
        ' newBackColor == Color to change all oldBackColor pixels to
        '   Nothing (default) == use default ContextMenuStrip.BackColor
        Private oldBackColor As Color = Color.Fuchsia
        Private newBackColor As Color? = Nothing          ' == Nullable(Of Color)

        ' end of stuff to change region for simply sending files
        '*****  *****  *****

        ' {PL} The names of the selected file
        Private selectedFiles As List(Of String)

        ' full path name
        Private targetApp As String = ""

        Private menuBmp As IntPtr = IntPtr.Zero
        Private IDM_DISPLAY As UInteger = 0

        ' mainly for reporting...***NOT*** for use in REG/UnREG
        Private AsmHandlerName As String = ""


        Public Sub New()
            ' create logfile full path name
            LogFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                              LogFile)
            Dim t As Type = Me.GetType
            Dim AsmHandlerName = String.Format("{0}.{1}.{2} Class",
                                               t.Assembly.GetName.Name, t.Namespace, t.Name)

            ' write dated header
            OpenLog()
            WriteToLog("@ctor")

            ' {PL} - no NREs
            selectedFiles = New List(Of String)

            ' almost always reports System/32
            WriteToLog(String.Format("Current dir: {0}",
                                     Environment.CurrentDirectory))

            ' {PL}
            ' The ShellExt is instanced by Explorer, so things like Environment.CurrentDirectory
            ' may report 'System32` no matter where the FILE actually is.  We know the name of
            ' of the file to run (AppName), so get the Assembly Location for use as the 
            ' folder name to append to AppName.  The ShellExt must be in the same folder as the app
            ' for this to work:
            Dim myPath As String = Path.GetDirectoryName(Assembly.GetAssembly(Me.GetType).Location)
            targetApp = Path.Combine(myPath, Path.GetFileName(AppName))

            If File.Exists(targetApp) Then
                WriteToLog("  Target found: " & targetApp)
            Else
                WriteToLog(String.Format("  Target app NOT found: [{0}]",
                                         targetApp))
            End If

            ' Load the bitmap for the menu item.
            If String.IsNullOrEmpty(bmpName) = False Then
                Dim bmp As Bitmap = GetPseudoTransparentBitmap()
                menuBmp = bmp.GetHbitmap()

                '{PL} MS orginal code does not make BG transparent
                '  hence the hack
                'Dim bmp As Bitmap
                'Try
                '    bmp = CType(My.Resources.ResourceManager.GetObject(bmpName), Bitmap)

                '    If bmp IsNot Nothing Then
                '        bmp.MakeTransparent(oldBackColor)
                '        menuBmp = bmp.GetHbitmap()
                '    Else
                '        WriteToLog(String.Format("   no BMP res named: {0}", bmpName))
                '    End If

                '    menuBmp = bmp.GetHbitmap()
                'Catch ex As Exception
                '    WriteToLog(ex.Message)
                'End Try
            End If


        End Sub

        ' the bitmaps do not show up transparent; even Adobe has probs with this.
        ' Neither PNG nor ICONs work.
        '
        ' I cannot find where Windows stores the Context Menu back color
        '  its not the menu color in the registry.  It probably varies depending
        '  on whether there is a Theme or not.
        '
        ' Until more info is available, ASSUME the menu BG is
        ' the same as a new CMS.  Scan the BMP, replace TRANSPARENT values
        ' with that value.
        Private Function GetPseudoTransparentBitmap() As Bitmap
            Dim bmp As Bitmap = CType(My.Resources.ResourceManager.GetObject(bmpName), Bitmap)
            Dim px As Integer

            If newBackColor.HasValue = False Then
                Dim cms = New Windows.Forms.ContextMenuStrip
                newBackColor = cms.BackColor
                cms.Dispose()
            End If
            Dim pxT = oldBackColor.ToArgb

            For w As Int32 = 0 To bmp.Width - 1
                For h As Int32 = 0 To bmp.Height - 1
                    px = bmp.GetPixel(w, h).ToArgb
                    If px = pxT Then
                        bmp.SetPixel(w, h, newBackColor.Value)
                    End If
                Next
            Next

            Return bmp

        End Function

        Protected Overrides Sub Finalize()
            WriteToLog("@Finalize")
            If (menuBmp <> IntPtr.Zero) Then
                NativeMethods.DeleteObject(menuBmp)
                menuBmp = IntPtr.Zero
            End If
        End Sub


        Private Sub OnExecuteMenuCommand(ByVal hWnd As IntPtr)

            ' {PL} the command line args and a literal for formatting
            Dim args As String = ""
            Dim quote As String = """"

            WriteToLog("@ExecuteMenuCommand")

            If String.IsNullOrEmpty(targetApp) Then
                Throw New FileNotFoundException("cannot find targetApp", targetApp)
            End If

            WriteToLog("  FeedFirstFileOnly == " & FeedFirstFileOnly.ToString)
            WriteToLog("  File Count == " & selectedFiles.Count.ToString)

            If selectedFiles.Count > 0 Then
                ' alternatively, Sub Initialize could be changed to only
                ' collect one file if true
                If FeedFirstFileOnly Then
                    args = String.Format(" {0}{1}{0} ", quote, selectedFiles(0))
                Else
                    ' {PL} concat args, wrap in quotes to preserve spaces
                    '  -- See Note below
                    For n As Integer = 0 To selectedFiles.Count - 1
                        args &= String.Format(" {0}{1}{0} ", quote, selectedFiles(n))
                    Next
                End If

                ' MS-PL test stand-in for feeding a full command line to Some Other App
                ' if you really, really need to.
                'MessageBox.Show("Cmd to execute: " & Environment.NewLine & " args: [" & args & "]", "ShellExtContextMenuHandler")

                ' Note: 
                ' Our menu pops up when the item under the mouse matches one of the registered extensions
                '   but other file types might ALSO be selected.  So, you could filter non supported types out of
                '   the list, but that seems better left to the app.  

                ' report app and command line as is
                WriteToLog(String.Format("  Cmd to execute: {0}", targetApp))
                WriteToLog(String.Format("  Args: {0}", args))

                ' 
                Dim pStartInfo As New ProcessStartInfo(targetApp)
                Dim proc As Process
                Try
                    pStartInfo.Arguments = args
                    pStartInfo.WorkingDirectory = Path.GetDirectoryName(targetApp)
                    proc = Process.Start(pStartInfo)
                    If proc Is Nothing Then
                        WriteToLog("Process is null")
                    Else
                        WriteToLog("Process started; ID = " & proc.Id.ToString)
                    End If
                Catch ex As Exception
                    WriteToLog(ex.Message)
                End Try

                ' if you want to Do Something in this DLL, rather than 
                ' run some other process, you can.  For something simple:
                ' (Not Recommended)

                ''Dim frm As New ShellForm
                ''frm.lbFiles.Items.AddRange(selectedFiles.ToArray)
                ''frm.Show()

                ' NOTE: These are very tedious to debug and test.  It is actually
                ' simpler to start some other App and iteratively debug that.

            Else
                ' might mean something is broke in Initialize
                WriteToLog("Selected Files Count = 0 ")
            End If
            WriteToLog("Working set: " & Environment.WorkingSet.ToString())
            WriteToLog("***  Normal Termination  ***")

        End Sub

        Private Sub OpenLog()
            If LogActive = False Then Exit Sub
            Using sw As New StreamWriter(LogFile, LogAppend)
                sw.WriteLine("************")
                sw.WriteLine(AsmHandlerName)
                sw.WriteLine(DateTime.Now.ToString("G"))
            End Using

        End Sub

        Private Sub WriteToLog(s As String)
            If LogActive = False Then Exit Sub
            Using sw As New StreamWriter(LogFile, True)
                sw.WriteLine(s)
            End Using
        End Sub


#Region "Shell Extension Registration"

        ' this provides the means for the main app to get the 
        ' "master list" of all supported extensions
        Public Shared Function GetExtensions() As String()
            Return MyExts
        End Function

        ' standard register operation used by RegAsm 
        ' DO NOT try to overload it
        <ComRegisterFunction()>
        Public Shared Sub Register(ByVal T As Type)
            Try
                Dim AsmHandlerName = String.Format("{0}.{1}.{2} Class",
                                                   T.Assembly.GetName().Name, T.Namespace, T.Name)
                For Each s As String In MyExts
                    ShellExtReg.RegisterShellExtContextMenuHandler(T.GUID, s, AsmHandlerName)
                Next

            Catch ex As Exception
                Console.WriteLine(ex.Message) ' Log the error
                Throw ' Re-throw the exception
            End Try
        End Sub


        Public Shared Sub RegisterExtensions(T As Type, exts As String())
            Dim thisExt As String = ""

            ' detach all extensions in the static list
            Unregister(T)

            Try
                Dim AsmHandlerName = String.Format("{0}.{1}.{2} Class",
                                                   T.Assembly.GetName().Name, T.Namespace, T.Name)
                For Each s As String In exts
                    thisExt = s.ToLower
                    ' only mess with passed ext if it is in our list
                    ' or we will have trouble removing it
                    If MyExts.Contains(thisExt) Then
                        ShellExtReg.RegisterShellExtContextMenuHandler(T.GUID, thisExt, AsmHandlerName)
                    End If
                Next

            Catch ex As Exception
                Console.WriteLine(ex.Message) ' Log the error
                Throw ' Re-throw the exception
            End Try
        End Sub

        ' standard UNregister operation used by RegAsm 
        <ComUnregisterFunction()>
        Public Shared Sub Unregister(ByVal t As Type)
            Try
                For Each s As String In MyExts
                    ShellExtReg.UnregisterShellExtContextMenuHandler(t.GUID, s)
                Next

            Catch ex As Exception
                Console.WriteLine(ex.Message) ' Log the error
                Throw ' Re-throw the exception
            End Try
        End Sub

#End Region


#Region "IShellExtInit Members"

        ''' <summary>
        ''' Initialize the context menu handler.
        ''' </summary>
        ''' <param name="pidlFolder">
        ''' A pointer to an ITEMIDLIST structure that uniquely identifies a folder.
        ''' </param>
        ''' <param name="pDataObj">
        ''' A pointer to an IDataObject interface object that can be used to 
        ''' retrieve the objects being acted upon.
        ''' </param>
        ''' <param name="hKeyProgID">
        ''' The registry key for the file object or folder type.
        ''' </param>
        Public Sub Initialize(ByVal pidlFolder As IntPtr,
                            ByVal pDataObj As IntPtr,
                            ByVal hKeyProgID As IntPtr) Implements IShellExtInit.Initialize

            If (pDataObj = IntPtr.Zero) Then
                Throw New ArgumentException
            End If

            ' if targetapp....

            '{PL} I think a new instace of this class gets created
            ' each time so the ctor will already have done this. But
            ' in case it doesnt, create a new instance assure the list
            ' is empty
            selectedFiles = New List(Of String)


            Dim fe As New FORMATETC
            With fe
                .cfFormat = Convert.ToInt16(CLIPFORMAT.CF_HDROP)
                .ptd = IntPtr.Zero
                .dwAspect = DVASPECT.DVASPECT_CONTENT
                .lindex = -1
                .tymed = TYMED.TYMED_HGLOBAL
            End With

            Dim stm As New STGMEDIUM

            ' The pDataObj pointer contains the objects being acted upon. In this 
            ' example, we get an HDROP handle for enumerating the selected files 
            ' and folders.
            Dim dataObject As IDataObject = CType(Marshal.GetObjectForIUnknown(pDataObj), IDataObject)
            dataObject.GetData(fe, stm)

            Try
                ' Get an HDROP handle.
                Dim hDrop As IntPtr = stm.unionmember
                If (hDrop = IntPtr.Zero) Then
                    Throw New ArgumentException
                End If

                ' Determine how many files are involved in this operation.
                Dim nFiles As UInteger = NativeMethods.DragQueryFile(hDrop,
                                     UInt32.MaxValue, Nothing, 0)

                WriteToLog("@Initialize")
                WriteToLog(String.Format("DragQueryFile count = {0}", nFiles.ToString))

                ' This code sample displays the custom context menu item when only 
                ' one file is selected. 
                If (nFiles > 0) Then
                    ' Get the path of the file.
                    Dim sbFileName As StringBuilder
                    Dim f As String

                    For n As UInteger = 0 To nFiles - 1UI
                        sbFileName = New StringBuilder(260)

                        If (0 = NativeMethods.DragQueryFile(hDrop, n, sbFileName,
                                           sbFileName.Capacity)) Then
                            Marshal.ThrowExceptionForHR(WinError.E_FAIL)
                        End If

                        f = sbFileName.ToString
                        selectedFiles.Add(f)
                        WriteToLog("    added: " & f)

                    Next

                Else
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL)
                End If

            Finally
                NativeMethods.ReleaseStgMedium((stm))
            End Try
        End Sub

#End Region


#Region "IContextMenu Members"

        ''' <summary>
        ''' Add commands to a shortcut menu.
        ''' </summary>
        ''' <param name="hMenu">A handle to the shortcut menu.</param>
        ''' <param name="iMenu">
        ''' The zero-based position at which to insert the first new menu item.
        ''' </param>
        ''' <param name="idCmdFirst">
        ''' The minimum value that the handler can specify for a menu item ID.
        ''' </param>
        ''' <param name="idCmdLast">
        ''' The maximum value that the handler can specify for a menu item ID.
        ''' </param>
        ''' <param name="uFlags">
        ''' Optional flags that specify how the shortcut menu can be changed.
        ''' </param>
        ''' <returns>
        ''' If successful, returns an HRESULT value that has its severity value 
        ''' set to SEVERITY_SUCCESS and its code value set to the offset of the 
        ''' largest command identifier that was assigned, plus one.
        ''' </returns>
        Public Function QueryContextMenu(ByVal hMenu As IntPtr,
                            ByVal iMenu As UInt32, ByVal idCmdFirst As UInt32,
                            ByVal idCmdLast As UInt32,
                            ByVal uFlags As UInt32) As Integer Implements IContextMenu.QueryContextMenu

            ' If uFlags include CMF_DEFAULTONLY then we should not do anything.
            If ((CMF.CMF_DEFAULTONLY And uFlags) <> 0) Then
                Return WinError.MAKE_HRESULT(WinError.SEVERITY_SUCCESS, 0, 0)
            End If

            ' Use either InsertMenu or InsertMenuItem to add menu items.

            Dim mii As New MENUITEMINFO
            With mii
                .cbSize = Convert.ToUInt32(Marshal.SizeOf(mii))
                .fMask = MIIM.MIIM_BITMAP Or MIIM.MIIM_STRING Or MIIM.MIIM_FTYPE Or
                                MIIM.MIIM_STATE Or MIIM.MIIM_ID
                .wID = idCmdFirst + Me.IDM_DISPLAY
                .fType = MFT.MFT_STRING
                .dwTypeData = menuText
                .fState = MFS.MFS_ENABLED
                .hbmpItem = menuBmp
            End With
            If Not NativeMethods.InsertMenuItem(hMenu, iMenu, True, mii) Then
                Return Marshal.GetHRForLastWin32Error
            End If

            ' Add a separator.
            If AddSeperator Then
                Dim sep As New MENUITEMINFO
                With sep
                    .cbSize = Convert.ToUInt32(Marshal.SizeOf(sep))
                    .fMask = MIIM.MIIM_TYPE
                    .fType = MFT.MFT_SEPARATOR
                End With
                If Not NativeMethods.InsertMenuItem(hMenu, Convert.ToUInt32(iMenu + 1), True, sep) Then
                    Return Marshal.GetHRForLastWin32Error
                End If
            End If


            ' Return an HRESULT value with the severity set to SEVERITY_SUCCESS. 
            ' Set the code value to the offset of the largest command identifier 
            ' that was assigned, plus one (1).
            Return WinError.MAKE_HRESULT(0, 0, Convert.ToUInt32(IDM_DISPLAY + 1))
        End Function


        ''' <summary>
        ''' Carry out the command associated with a shortcut menu item.
        ''' </summary>
        ''' <param name="pici">
        ''' A pointer to a CMINVOKECOMMANDINFO or CMINVOKECOMMANDINFOEX structure 
        ''' containing information about the command. 
        ''' </param>
        Public Sub InvokeCommand(ByVal pici As IntPtr) Implements IContextMenu.InvokeCommand

            Dim isUnicode As Boolean = False

            ' Determine which structure is being passed in, CMINVOKECOMMANDINFO or 
            ' CMINVOKECOMMANDINFOEX based on the cbSize member of lpcmi. Although 
            ' the lpcmi parameter is declared in Shlobj.h as a CMINVOKECOMMANDINFO 
            ' structure, in practice it often points to a CMINVOKECOMMANDINFOEX 
            ' structure. This struct is an extended version of CMINVOKECOMMANDINFO 
            ' and has additional members that allow Unicode strings to be passed.
            Dim ici As CMINVOKECOMMANDINFO = CType(Marshal.PtrToStructure(pici,
                                                            GetType(CMINVOKECOMMANDINFO)), CMINVOKECOMMANDINFO)

            Dim iciex As New CMINVOKECOMMANDINFOEX
            If (ici.cbSize = Marshal.SizeOf(GetType(CMINVOKECOMMANDINFOEX))) Then
                If (ici.fMask And CMIC.CMIC_MASK_UNICODE) <> 0 Then
                    isUnicode = True
                    iciex = CType(Marshal.PtrToStructure(pici,
                                                         GetType(CMINVOKECOMMANDINFOEX)), CMINVOKECOMMANDINFOEX)
                End If
            End If

            ' Determines whether the command is identified by its offset or verb.
            ' There are two ways to identify commands:
            ' 
            '   1) The command's verb string 
            '   2) The command's identifier offset
            ' 
            ' If the high-order word of lpcmi->lpVerb (for the ANSI case) or 
            ' lpcmi->lpVerbW (for the Unicode case) is nonzero, lpVerb or lpVerbW 
            ' holds a verb string. If the high-order word is zero, the command 
            ' offset is in the low-order word of lpcmi->lpVerb.

            ' For the ANSI case, if the high-order word is not zero, the command's 
            ' verb string is in lpcmi->lpVerb. 
            If (Not isUnicode AndAlso (NativeMethods.HighWord(ici.verb.ToInt32) <> 0)) Then
                ' Is the verb supported by this context menu extension?
                If (Marshal.PtrToStringAnsi(ici.verb) = Me.verb) Then
                    OnExecuteMenuCommand(ici.hwnd)
                Else
                    ' If the verb is not recognized by the context menu handler, 
                    ' it must return E_FAIL to allow it to be passed on to the 
                    ' other context menu handlers that might implement that verb.
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL)
                End If

            ElseIf (isUnicode AndAlso (NativeMethods.HighWord(iciex.verbW.ToInt32) <> 0)) Then
                ' For the Unicode case, if the high-order word is not zero, the 
                ' command's verb string is in lpcmi->lpVerbW. 

                ' Is the verb supported by this context menu extension?
                If (Marshal.PtrToStringUni(iciex.verbW) = Me.verb) Then
                    OnExecuteMenuCommand(ici.hwnd)
                Else
                    ' If the verb is not recognized by the context menu handler, 
                    ' it must return E_FAIL to allow it to be passed on to the 
                    ' other context menu handlers that might implement that verb.
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL)
                End If

            Else
                ' If the command cannot be identified through the verb string, 
                ' then check the identifier offset.

                ' Is the command identifier offset supported by this context menu 
                ' extension?
                If (NativeMethods.LowWord(ici.verb.ToInt32) = Me.IDM_DISPLAY) Then
                    OnExecuteMenuCommand(ici.hwnd)
                Else
                    ' If the verb is not recognized by the context menu handler, 
                    ' it must return E_FAIL to allow it to be passed on to the 
                    ' other context menu handlers that might implement that verb.
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL)
                End If
            End If
        End Sub


        ''' <summary>
        ''' Get information about a shortcut menu command, including the help 
        ''' string and the language-independent, or canonical, name for the 
        ''' command.
        ''' </summary>
        ''' <param name="idCmd">Menu command identifier offset.</param>
        ''' <param name="uFlags">
        ''' Flags specifying the information to return. This parameter can have 
        ''' one of the following values: GCS_HELPTEXTA, GCS_HELPTEXTW, 
        ''' GCS_VALIDATEA, GCS_VALIDATEW, GCS_VERBA, GCS_VERBW.
        ''' </param>
        ''' <param name="pReserved">Reserved. Must be IntPtr.Zero</param>
        ''' <param name="pszName">
        ''' The address of the buffer to receive the null-terminated string being 
        ''' retrieved.
        ''' </param>
        ''' <param name="cchMax">
        ''' Size of the buffer, in characters, to receive the null-terminated 
        ''' string.
        ''' </param>
        Public Sub GetCommandString(ByVal idCmd As UIntPtr,
                        ByVal uFlags As UInt32,
                        ByVal pReserved As IntPtr,
                        ByVal pszName As StringBuilder,
                        ByVal cchMax As UInt32) Implements IContextMenu.GetCommandString

            WriteToLog("@GetCommandString")
            WriteToLog(String.Format("idCmd ID: {0}, IDMDisplay: {1}", idCmd.ToUInt32.ToString, IDM_DISPLAY.ToString))

            If (idCmd.ToUInt32 = Me.IDM_DISPLAY) Then


                Select Case DirectCast(uFlags, GCS)
                    Case GCS.GCS_VERBW
                        WriteToLog("verbCanonicalName")
                        If (Me.verbCanonicalName.Length > (cchMax - 1)) Then
                            Marshal.ThrowExceptionForHR(WinError.STRSAFE_E_INSUFFICIENT_BUFFER)
                        Else
                            pszName.Clear()
                            pszName.Append(Me.verbCanonicalName)
                        End If

                    Case GCS.GCS_HELPTEXTW
                        WriteToLog("verbHelpText")
                        If (Me.verbHelpText.Length > (cchMax - 1)) Then
                            Marshal.ThrowExceptionForHR(WinError.STRSAFE_E_INSUFFICIENT_BUFFER)
                        Else
                            pszName.Clear()
                            pszName.Append(Me.verbHelpText)
                        End If

                    Case Else
                        WriteToLog("case else")

                End Select
            End If
        End Sub

#End Region

    End Class

End Namespace
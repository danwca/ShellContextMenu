Imports System.IO
Imports System.Reflection
Imports System.Security.Principal

Public Class ShellReg

    Private AsmName As String
    '' note: the namespace is an integral part of the class name.
    '' e.g. ShellExtContextMenuHandler.MyShellMenu
    Private ClassName As String

    Private RegAsm As String

    ' these are the same for 4.0, 4.5 and 4.51
    'Private regasm32 As String = "Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
    'Private regasm64 As String = "Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

    Private Quote As String = """"

    ''' <summary>
    ''' Creates a new instance of the Shell Registry Helper
    ''' </summary>
    ''' <param name="aName">Name of your Shell Extension DLL</param>
    ''' <param name="cName">Name of the main Shell Extension Class </param>
    ''' <remarks></remarks>
    Public Sub New(aName As String, cName As String)
        AsmName = aName
        ClassName = cname


        RegAsm = Path.Combine(System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory(),
                              "regasm.exe")

        ' alternative:
        'Dim winPath = Environment.GetFolderPath(Environment.SpecialFolder.Windows)

        ' assumes the file is there, admin did not remove etc
        ' whereas the above relies on NET framework which must be there

        'If Environment.Is64BitOperatingSystem And Environment.Is64BitProcess Then
        '    RegAsm = Path.Combine(winPath, regasm64)

        'ElseIf (Environment.Is64BitOperatingSystem = False) Then 'And (Environment.Is64BitProcess = False) Then
        '    RegAsm = Path.Combine(winPath, regasm32)
        'End If


    End Sub

    ' Register the context menu helper with the
    ' extensions defined in the shellName Class
    Public Function RegisterShellExt() As Boolean

        Return InvokeRegAsm("/codebase")

    End Function

    ' UnRegister the indicated shell helper
    Public Sub UnRegisterShellExt()
        InvokeRegAsm("/u")
    End Sub

    Public Shared Function IsAdminProcess() As Boolean

        Return New WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator)

    End Function

    Public Shared Function OSBitsMatch() As Boolean

        If Environment.Is64BitOperatingSystem Then
            Return (Environment.Is64BitProcess)
        Else
            Return False
        End If

    End Function


    Private Function InvokeRegAsm(arg As String) As Boolean

        If String.IsNullOrEmpty(RegAsm) Then
            Return False
        End If

        If File.Exists(RegAsm) = False Then
            Return False
        End If

        ' do not execute if app is 32bit but OS is 64bit.
        ' the demo prevents this by disabling the buttons
        ' but be sure to prevent it outside the demo:
        If (Environment.Is64BitOperatingSystem) And (Environment.Is64BitProcess = False) Then
            Return False
        End If

        Dim myPath As String = Path.GetDirectoryName(GetExecutableFileName())

        'Dim myDLLTarget As String = Path.Combine(Path.GetDirectoryName(myPath), AsmName)
        Dim myArgs As String = String.Format("{0}{1}{0} {2}", Quote, AsmName, arg)


        Dim proc = New Process
        proc.StartInfo.FileName = RegAsm
        proc.StartInfo.WorkingDirectory = myPath
        proc.StartInfo.Verb = "runas"
        proc.StartInfo.Arguments = myArgs
        proc.StartInfo.UseShellExecute = True
        proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden

        Try
            proc.Start()

        Catch ex As Exception
            Return False
        End Try

        Return True

    End Function


    ' Utility function to get the "master list" of extensions defined in
    ' your shell menu class.  Prevents having to define the list in 2 projects
    ' and possibly get out of synch
    Public Function GetShellExtList() As String()
        Dim myExt As String()


        If File.Exists(AsmName) Then
            Dim fi As New FileInfo(AsmName)
            AsmName = fi.FullName
        Else
            Return New String() {}
        End If

        Dim asm As Assembly = Assembly.LoadFile(AsmName)

        Dim tShell As Type = asm.GetType(ClassName)

        If tShell IsNot Nothing Then
            Try
                Dim g As Guid = tShell.GUID
                ' get the MI for the method we need
                Dim mi As MethodInfo = tShell.GetMethod("GetExtensions",
                                                BindingFlags.Public Or BindingFlags.Static,
                                                Nothing,
                                                New Type() {},
                                                Nothing)

                If mi IsNot Nothing Then
                    myExt = CType(mi.Invoke(tShell, New Object() {}), String())
                    Return myExt
                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If

        Return New String() {}

    End Function

    ' Register the shell helper, but ignore the extensions defined there and use
    ' the ones passed.
    Public Sub RegisterExtensions(newExts As String())
        ' register with a specific set of Extensions.  
        ' The DLL/helper may already be
        ' registered with a set, so the method called will first
        ' unregister the "master" set of extensions, then
        ' register the new set passed


        If File.Exists(AsmName) Then
            Dim fi As New FileInfo(AsmName)
            AsmName = fi.FullName
        Else
            Exit Sub
        End If

        Dim asm As Assembly = Assembly.LoadFile(AsmName)
        Dim tShell As Type = asm.GetType(ClassName)

        ' must run as Admin so that the Register method which
        ' accesses the registry is allowed.  otherwise an exception is thrown.
        ' This is handled by this app running elevated just to do this:
        If tShell IsNot Nothing Then
            Try
                'Dim g As Guid = tShell.GUID
                ' get the MI for the method we need
                Dim mi As MethodInfo = tShell.GetMethod("RegisterExtensions",
                                                BindingFlags.Public Or BindingFlags.Static,
                                                Nothing,
                                                New Type() {GetType(Type), GetType(String())},
                                                Nothing)

                If mi IsNot Nothing Then
                    mi.Invoke(tShell, New Object() {tShell, newExts})
                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If

    End Sub

    Public Function RunAppElevated(args As String()) As Boolean

        Dim fname As String = GetExecutableFileName()
        'fname = Assembly.GetExecutingAssembly().Location

        Dim proc As Process
        Dim psi As New ProcessStartInfo(fname)

        Dim cmdline = "/regext"
        For Each s As String In args
            cmdline &= " " & s
        Next

        psi.Arguments = cmdline
        psi.Verb = "runas"
        psi.WorkingDirectory = String.Format("{0}{1}{0}", Quote, Path.GetDirectoryName(fname))
        psi.UseShellExecute = True
        psi.WindowStyle = ProcessWindowStyle.Hidden

        Try
            proc = Process.Start(psi)

        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    Private Function GetExecutableFileName() As String
        ' if ShellReg is moved to a DLL, the following should
        ' probably change to GetCallingAssembly
        Dim myEXE As String = Assembly.GetExecutingAssembly().Location

        ' "fixing" it to look like a runtime ASM
        ' REQUIRES the correct test directry be entered in Project Properties
        If Debugger.IsAttached Then
            myEXE = Path.Combine(Environment.CurrentDirectory, Path.GetFileName(myEXE))
        End If

        Return myEXE


    End Function


End Class

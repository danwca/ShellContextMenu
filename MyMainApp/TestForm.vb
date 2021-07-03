Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices



Public Class TestForm


    ' the Assembly name of the Shell Extension for this app 
    Public Shared AsmName As String = "MyShellExt.dll"         ' name of your Shell Ext DLL
    ' Note that the Namespace is integral to the class name:
    Public Shared ClassName As String = "ShellExtContextMenuHandler.MyShellMenu"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim args = Environment.GetCommandLineArgs()

        ' getting them from Environment, the EXE is arg(0)
        For n As Int32 = 1 To args.Length - 1
            lbFiles.Items.Add(args(n))
        Next

        ' the shell DLL is 64bit/AnyCPU so
        ' you can only reg/unreg from the app
        ' if this app is 64bit
        If Environment.Is64BitOperatingSystem Then
            btnUnReg.Enabled = Environment.Is64BitProcess
            btnReg.Enabled = Environment.Is64BitProcess
            btnExReg.Enabled = Environment.Is64BitProcess
        End If

    End Sub

    ' these are intendend to start a new instance of this app with /reg or /unreg
    ' on the commandline.  When Sub Main sees these, it invokes the related
    ' function in the DLL from Program.RegisterShellExt or Program.UnRegisterShellExt

    Private Sub btnReg_Click(sender As Object, e As EventArgs) Handles btnReg.Click

        Dim shreg As New ShellReg(AsmName, ClassName)
        shreg.RegisterShellExt()

    End Sub

    Private Sub btnUnReg_Click(sender As Object, e As EventArgs) Handles btnUnReg.Click

        Dim shreg As New ShellReg(AsmName, ClassName)
        shreg.UnRegisterShellExt()
    End Sub


    Private Sub btnExReg_Click(sender As Object, e As EventArgs) Handles btnExReg.Click
        ' register the "bar" file ext.

        ' I hate dialogs, but the button text may not be clear what is being demoed
        If MessageBox.Show("This will remove the '.bar' extension from the DLL's handled list.", "Question", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) <> Windows.Forms.DialogResult.OK Then
            Exit Sub
        End If

        Dim shReg As New ShellReg(AsmName, ClassName)

        Dim temp As List(Of String) = (shReg.GetShellExtList()).ToList
        'If temp.Count = 0 Then Exit Sub

        ' if the method fails, it returns empty array
        If temp.Count > 0 Then
            temp.RemoveAt(temp.IndexOf(".bar"))

            ' no real reason ShellReg must run it
            ' but we already have the method
            shReg.RunAppElevated(temp.ToArray)
        End If



    End Sub


    
End Class

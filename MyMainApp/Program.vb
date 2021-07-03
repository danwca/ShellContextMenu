Imports System.IO
Imports System.Reflection
Imports System.Security.Principal

Module Program

    Private Const SWREG = "/regext"

    ' Note
    ' Starting from Sub Main means that this app cannot use the 
    ' VB App Framework which includes the 'Make Single Instance` option.
    ' 
    ' Alternatives
    ' a) move the Sub Main code to process the commandline 
    '    to a small console app and run that as Admin instead.
    '    change ShellReg.RunAppElevated to 
    ' b) Add your own Single Instance code so that your app
    '    cant do both respond to /regext when present or forward
    '    any other commandline to the first instance


    Public Sub Main(args As String())
        ' arg(0) = switch
        ' arg(1...N) exts

        ' first arg must be /regext and
        ' there has to be 1 or more ext passed
        If (args.Count > 1) AndAlso
                (args(0).ToLower = SWREG.ToLower) Then


            ' dont even try if not Admin process and OS match
            If ShellReg.IsAdminProcess AndAlso ShellReg.OSBitsMatch Then
                ' get args1...N 
                Dim cmdLine(args.Count - 2) As String
                Array.Copy(args, 1, cmdLine, 0, args.Count - 1)

                Dim shReg As New ShellReg(TestForm.AsmName, TestForm.ClassName)
                shReg.RegisterExtensions(cmdLine)

            End If

            ' doesnt really do anything - there is no Application object yet
            Application.Exit()

        Else

            ' normal winforms startup
            Application.EnableVisualStyles()
            Application.Run(New TestForm)

        End If

    End Sub

   

End Module

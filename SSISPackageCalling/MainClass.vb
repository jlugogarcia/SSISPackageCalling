
'****INSTRUCTIONS****'''

'1-Install Microsoft SQL Server Data Tools from Microsoft Visual Studio 2015 installation wizard. 
'  This ensures to add Microsoft.SqlServer.ManagedDTS to the GAC of the computer/server
'2- Add the DLL Microsoft.SqlServer.ManagedDTS.dll to the references
'of the project.
'This DLL can be often found at:
'C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.SqlServer.ManagedDTS\v4.0_13.0.0.0__89845dcd8080cc91\Microsoft.SqlServer.ManagedDTS.dll
Imports System.IO
Imports Microsoft.SqlServer.Dts.Runtime

Module MainClass

    Sub Main()
        Dim pkSSIS = Nothing
        Dim error_Msg As String = Nothing

        Dim test = Directory.GetCurrentDirectory

        pkSSIS = "Path_to_Your_SSIS_Package\Package.dtsx"
        Console.WriteLine("The package is executing...")
        Dim pkg As Package = Nothing
        Dim app As Microsoft.SqlServer.Dts.Runtime.Application
        Dim result As DTSExecResult
        Try
            app = New Microsoft.SqlServer.Dts.Runtime.Application()
            pkg = app.LoadPackage(pkSSIS, Nothing)
            result = pkg.Execute()
            If result = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure Then
                For Each dt_error As Microsoft.SqlServer.Dts.Runtime.DtsError In pkg.Errors
                    error_Msg += dt_error.Description.ToString()
                Next
                Console.WriteLine(error_Msg)
            End If
            If result = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success Then
                Console.WriteLine("Success!")
                Console.ReadLine()
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

End Module

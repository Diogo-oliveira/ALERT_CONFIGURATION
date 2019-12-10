Imports System
Imports System.IO
Imports System.Text
Imports Oracle.DataAccess.Client
Public Class SCRIPT_GENERATOR
    Public Shared CURRENT_PATH As String = System.AppDomain.CurrentDomain.BaseDirectory()
    Public Shared SCRIPT_FOLDER As String = "ALERT_CONFIGURATION_SCRIPTS"
    Public Shared SCRIPT_HISTORY As String = "\HISTORY"
    Public Shared SCRIPT_FILE As String = "\SCRIPT_" & System.DateTime.Now.Day & "-" & System.DateTime.Now.Month & "-" & System.DateTime.Now.Year & "_" & System.DateTime.Now.Hour & "h" & System.DateTime.Now.Minute & "m" & System.DateTime.Now.Second & "s.txt"

    Function CLEAN_SCRIPT()

        Dim script_files = My.Computer.FileSystem.GetFiles(CURRENT_PATH & SCRIPT_FOLDER)

        Dim file_name As String = ""

        Dim path_length As Int64 = CURRENT_PATH.Count() + SCRIPT_FOLDER.Count()

        For i As Integer = 0 To script_files.Count() - 1

            file_name = CStr(script_files(i)).Substring(path_length + 1, CStr(script_files(i).Count() - (path_length + 1)))

            My.Computer.FileSystem.MoveFile(CURRENT_PATH & SCRIPT_FOLDER & "\" & file_name, CURRENT_PATH & SCRIPT_FOLDER & SCRIPT_HISTORY & "\" & file_name)

        Next

    End Function

    Function CREATE_SCRIPT_FOLDER()

        If (Not System.IO.Directory.Exists(CURRENT_PATH & SCRIPT_FOLDER)) Then
            System.IO.Directory.CreateDirectory(CURRENT_PATH & SCRIPT_FOLDER)
        End If

        If (Not System.IO.Directory.Exists(CURRENT_PATH & SCRIPT_FOLDER & SCRIPT_HISTORY)) Then
            System.IO.Directory.CreateDirectory(CURRENT_PATH & SCRIPT_FOLDER & SCRIPT_HISTORY)
        End If

    End Function

    Public Shared Function CREATE_SCRIPT_FILE()

        System.IO.File.Create(CURRENT_PATH & SCRIPT_FOLDER & SCRIPT_FILE).Dispose()

    End Function

    Public Shared Function INIT_SCRIPT()

        Dim l_debug_file As String = CURRENT_PATH & SCRIPT_FOLDER & SCRIPT_FILE
        Dim sw As StreamWriter

        sw = File.AppendText(l_debug_file)
        sw.WriteLine("###########################################################################################")
        sw.WriteLine("                            ALERT_ENVIRONMENTS_CONFIGURATION")
        sw.WriteLine(" ")
        sw.WriteLine(DateTime.Now & " - " & "STARTING SCRIPT")
        sw.Dispose()
        sw.Close()

    End Function
End Class

Imports System
Imports System.IO
Imports System.Text

Public Class DEBUGGER

    Dim CURRENT_PATH As String = System.AppDomain.CurrentDomain.BaseDirectory()
    Dim DEBUG_FOLDER As String = "ALERT_CONFIGURATION_DEBUGGER"
    Dim DEBUG_HISTORY As String = "\HISTORY"
    Dim DEBUG_FILE As String = "\DEBUG_" & System.DateTime.Now.Day & "-" & System.DateTime.Now.Month & "-" & System.DateTime.Now.Year & "_" & System.DateTime.Now.Hour & "h" & System.DateTime.Now.Minute & "m" & System.DateTime.Now.Second & "s.txt"

    Function CLEAN_DEBUG()

        Dim debug_files = My.Computer.FileSystem.GetFiles(CURRENT_PATH & DEBUG_FOLDER)

        Dim file_name As String = ""

        Dim path_length As Int64 = CURRENT_PATH.Count() + DEBUG_FOLDER.Count()

        For i As Integer = 0 To debug_files.Count() - 1

            file_name = CStr(debug_files(i)).Substring(path_length + 1, CStr(debug_files(i).Count() - (path_length + 1)))

            My.Computer.FileSystem.MoveFile(CURRENT_PATH & DEBUG_FOLDER & "\" & file_name, CURRENT_PATH & DEBUG_FOLDER & DEBUG_HISTORY & "\" & file_name)

        Next

    End Function

    Function CREATE_DEBUG_FOLDER()

        If (Not System.IO.Directory.Exists(CURRENT_PATH & DEBUG_FOLDER)) Then
            System.IO.Directory.CreateDirectory(CURRENT_PATH & DEBUG_FOLDER)
        End If

        If (Not System.IO.Directory.Exists(CURRENT_PATH & DEBUG_FOLDER & DEBUG_HISTORY)) Then
            System.IO.Directory.CreateDirectory(CURRENT_PATH & DEBUG_FOLDER & DEBUG_HISTORY)
        End If

    End Function

    Function CREATE_DEBUG_FILE()

        System.IO.File.Create(CURRENT_PATH & DEBUG_FOLDER & DEBUG_FILE)

    End Function

End Class

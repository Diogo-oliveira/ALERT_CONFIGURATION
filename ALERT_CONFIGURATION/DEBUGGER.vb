Imports System
Imports System.IO
Imports System.Text

Public Class DEBUGGER

    Public Shared CURRENT_PATH As String = System.AppDomain.CurrentDomain.BaseDirectory()
    Public Shared DEBUG_FOLDER As String = "ALERT_CONFIGURATION_DEBUGGER"
    Public Shared DEBUG_HISTORY As String = "\HISTORY"
    Public Shared DEBUG_FILE As String = "\DEBUG_" & System.DateTime.Now.Day & "-" & System.DateTime.Now.Month & "-" & System.DateTime.Now.Year & "_" & System.DateTime.Now.Hour & "h" & System.DateTime.Now.Minute & "m" & System.DateTime.Now.Second & "s.txt"

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

    Public Shared Function CREATE_DEBUG_FILE()

        System.IO.File.Create(CURRENT_PATH & DEBUG_FOLDER & DEBUG_FILE).Dispose()

    End Function

    Public Shared Function INIT_DEBUG()

        Dim l_debug_file As String = CURRENT_PATH & DEBUG_FOLDER & DEBUG_FILE
        Dim sw As StreamWriter

        sw = File.AppendText(l_debug_file)
        sw.WriteLine("###########################################################################################")
        sw.WriteLine("                            ALERT_ENVIRONMENTS_CONFIGURATION")
        sw.WriteLine(" ")
        sw.WriteLine(DateTime.Now & " - " & "STARTING DEBUGGER")
        sw.Close()

    End Function

    Public Shared Function SET_DEBUG(ByVal i_debug As String)

        Dim l_debug_file As String = CURRENT_PATH & DEBUG_FOLDER & DEBUG_FILE
        Dim sw As StreamWriter

        sw = File.AppendText(l_debug_file)
        sw.WriteLine(DateTime.Now & " - " & i_debug)
        sw.Close()

    End Function

    Public Shared Function SET_DEBUG_ERROR_INIT(ByVal i_screen As String)

        Dim l_debug_file As String = CURRENT_PATH & DEBUG_FOLDER & DEBUG_FILE
        Dim sw As StreamWriter

        sw = File.AppendText(l_debug_file)
        sw.WriteLine("##############         ERROR FOUND IN " & i_screen & "      ##############")
        sw.Close()

    End Function

    Public Shared Function SET_DEBUG_ERROR_CLOSE()

        Dim l_debug_file As String = CURRENT_PATH & DEBUG_FOLDER & DEBUG_FILE
        Dim sw As StreamWriter

        sw = File.AppendText(l_debug_file)
        sw.WriteLine("###############################################################")
        sw.Close()

    End Function

    Public Shared Function SET_DEBUG_NEW_FORM()

        Dim l_debug_file As String = CURRENT_PATH & DEBUG_FOLDER & DEBUG_FILE
        Dim sw As StreamWriter

        sw = File.AppendText(l_debug_file)
        sw.WriteLine(" ")
        sw.Close()

    End Function

End Class

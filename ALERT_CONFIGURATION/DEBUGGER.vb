Imports System
Imports System.IO
Imports System.Text
Imports Oracle.DataAccess.Client

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
        sw.Dispose()
        sw.Close()

    End Function

    Public Shared Function SET_DEBUG(ByVal i_debug As String)

        Dim l_debug_file As String = CURRENT_PATH & DEBUG_FOLDER & DEBUG_FILE
        Dim sw As StreamWriter

        sw = File.AppendText(l_debug_file)
        sw.WriteLine(DateTime.Now & " - " & i_debug & " - open cursors: " & GET_OPEN_CURSORS())
        sw.Dispose()
        sw.Close()

    End Function

    Public Shared Function SET_DEBUG_ERROR_INIT(ByVal i_screen As String)

        Dim l_debug_file As String = CURRENT_PATH & DEBUG_FOLDER & DEBUG_FILE
        Dim sw As StreamWriter

        sw = File.AppendText(l_debug_file)
        sw.WriteLine("##############         ERROR FOUND IN " & i_screen & "      ##############")
        sw.Dispose()
        sw.Close()

    End Function

    Public Shared Function SET_DEBUG_ERROR_CLOSE()

        Dim l_debug_file As String = CURRENT_PATH & DEBUG_FOLDER & DEBUG_FILE
        Dim sw As StreamWriter

        sw = File.AppendText(l_debug_file)
        sw.WriteLine("###########################################################################################")
        sw.Dispose()
        sw.Close()

    End Function

    Public Shared Function SET_DEBUG_NEW_FORM()

        Dim l_debug_file As String = CURRENT_PATH & DEBUG_FOLDER & DEBUG_FILE
        Dim sw As StreamWriter

        sw = File.AppendText(l_debug_file)
        sw.WriteLine(" ")
        sw.Dispose()
        sw.Close()

    End Function

    Public Shared Function GET_ARRAY_STRING(ByVal i_values As String()) As String

        Dim l_array As String = ""

        For i As Integer = 0 To i_values.Count - 1
            If i = 0 Then
                l_array = l_array & "["
            End If

            l_array = l_array & i_values(i)

            If i = i_values.Count - 1 Then
                l_array = l_array & "]"
            Else
                l_array = l_array & ", "
            End If
        Next

        Return l_array

    End Function

    Public Shared Function GET_ARRAY_NUMBER(ByVal i_values As Int64()) As String

        Dim l_array As String = ""

        For i As Integer = 0 To i_values.Count - 1
            If i = 0 Then
                l_array = l_array & "["
            End If

            l_array = l_array & i_values(i)

            If i = i_values.Count - 1 Then
                l_array = l_array & "]"
            Else
                l_array = l_array & ", "
            End If
        Next

        Return l_array

    End Function

    Public Shared Function SET_RESPONSE(ByVal i_function_name As String, i_parameters As String(), i_result As OracleDataReader)

        If 1 = 1 Then  'to_do: criar configuração para controlar se se apresenta info completa da resposta
            Dim l_debug_file As String = CURRENT_PATH & DEBUG_FOLDER & DEBUG_FILE
            Dim sw As StreamWriter

            Dim l_has_results As Boolean = False

            sw = File.AppendText(l_debug_file)

            Try
                ' sw.WriteLine(DateTime.Now & " - " & i_function_name & " RESPONSE:")
                sw.WriteLine("Parameters:")
                sw.WriteLine(" ")

                For i As Integer = 0 To i_parameters.Count - 1

                    sw.WriteLine("[" & i & "]: " & i_parameters(i).ToUpper)

                Next

                sw.WriteLine(" ")
                sw.WriteLine("Items:")
                sw.WriteLine(" ")

                Dim ii As Integer = 0
                While i_result.Read()
                    sw.WriteLine("[" & ii & "]")
                    For j As Integer = 0 To i_parameters.Count - 1
                        sw.WriteLine("   " & i_parameters(j).ToUpper & ": " & i_result.Item(j))
                    Next
                    sw.WriteLine(" ")
                    ii = ii + 1
                    l_has_results = True
                End While

                If l_has_results = False Then
                    sw.WriteLine("   NULL")
                    sw.WriteLine(" ")
                End If

            Catch ex As Exception
                sw.Close()
            End Try

            sw.Dispose()
            sw.Close()
        End If
    End Function

    Public Shared Function GET_OPEN_CURSORS() As Int64

        Dim l_count As Int64

        Dim sql As String = "   select max(a.value)
                                  from v$sesstat a, v$statname b, v$session s
                                 where a.statistic# = b.statistic#  and s.sid=a.sid
                                   and b.name = 'opened cursors current'
                                   and s.USERNAME='ALERT_CONFIG'"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try
            dr = cmd.ExecuteReader()

            While dr.Read()
                l_count = dr.Item(0)
            End While
            dr.Dispose()
            dr.Close()
            cmd.Dispose()
        Catch ex As Exception
            cmd.Dispose()
        End Try

        Return l_count

    End Function

End Class

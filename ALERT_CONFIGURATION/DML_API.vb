Imports Oracle.DataAccess.Client

Public Class DML_API
    Dim db_access_general As New General

    Function GET_INSERT_PROCESSED(ByVal i_insert As String, ByVal i_values As String) As String

        DEBUGGER.SET_DEBUG("DML_API :: GET_INSERT_PROCESSED(" & i_insert & ", " & i_values & ")")

        Dim l_return As String = "BEGIN" + vbNewLine

        l_return = l_return + "   " + i_insert + vbNewLine + "   " + i_values + vbNewLine

        l_return = l_return + "EXCEPTION
   WHEN DUP_VAL_ON_INDEX THEN
       dbms_output.put_line('WARNING: Operation has been previously executed.');
END;
/"

        Return l_return

    End Function

    Public Function GetNthIndex(s As String, t As Char, n As Integer) As Integer
        Dim count As Integer = 0
        For i As Integer = 0 To s.Length - 1
            If s(i) = t Then
                count += 1
                If count = n Then
                    Return i
                End If
            End If
        Next
        Return -1
    End Function

    Function REMOVE_AUDIT(ByRef io_insert As String, ByRef io_values As String, ByVal i_audit_tags() As String) As Boolean

        DEBUGGER.SET_DEBUG("DML_API :: REMOVE_AUDIT(" & io_insert & ", " & io_values & ")")

        Dim l_pos_init As Integer = -1
        Dim length As Integer = -1
        Dim l_pos As Integer = -1
        Dim l_pos_NthIndexOf As Integer = -1
        Dim l_pos_NthIndexOf_aux As Integer = -1
        Dim l_aux As String
        Dim l_count_sep As Integer
        Dim l_last_element As Boolean = False

        'Dim l_insert As String = io_insert
        ' Dim l_values As String = io_values

        For i As Integer = 0 To i_audit_tags.Count - 1
            If io_insert.Contains(i_audit_tags(i)) Then

                'Determinar posição onde começa a tag
                l_pos_init = io_insert.IndexOf(i_audit_tags(i))

                l_aux = io_insert.Substring(l_pos_init)
                length = l_aux.IndexOf(",")

                If length = -1 Then
                    length = l_aux.IndexOf(")")
                    l_last_element = True
                End If

                l_count_sep = io_insert.Split(",").Length - 1

                For j As Integer = 0 To l_count_sep - 1
                    l_pos_NthIndexOf = GetNthIndex(io_insert, ",", j + 1)
                    If l_pos_NthIndexOf > l_pos_init Then
                        l_pos = j + 1
                        Exit For
                    End If
                Next

                ''Se for o último elemento 
                If l_pos = -1 Then
                    l_pos = l_count_sep + 1
                End If

                l_pos_NthIndexOf = -1

                If l_last_element = True Then
                    l_pos_NthIndexOf = GetNthIndex(io_insert, ",", l_pos - 1)
                    io_insert = io_insert.Remove(l_pos_NthIndexOf, length + (l_pos_init - l_pos_NthIndexOf))
                Else
                    'Remover o espaço antes da tag a ser removida
                    If io_insert.Substring(l_pos_init - 1, 1) = " " Then
                        io_insert = io_insert.Remove((l_pos_init - 1), length + 2)
                    Else
                    io_insert = io_insert.Remove(l_pos_init, length + 1)
                End If
            End If


                    l_pos_NthIndexOf = GetNthIndex(io_values, ",", l_pos - 1)
                l_pos_NthIndexOf_aux = GetNthIndex(io_values, ",", l_pos)
                If l_pos_NthIndexOf_aux = -1 Then
                    l_pos_NthIndexOf_aux = GetNthIndex(io_values, ")", 1)
                End If

                Try
                    io_values = io_values.Remove(l_pos_NthIndexOf, (l_pos_NthIndexOf_aux - l_pos_NthIndexOf))
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If

            l_pos_init = -1
            length = -1
            l_pos = -1
            l_pos_NthIndexOf = -1
            l_pos_NthIndexOf_aux = -1
            l_last_element = False

        Next

        Return True

    End Function

    Function GET_PK_PROCESSED(ByVal i_pk_insert As String) As String

        DEBUGGER.SET_DEBUG("DML_API :: GET_PK_PROCESSED(" & i_pk_insert & ")")

        Dim l_return As String = "BEGIN" + vbNewLine

        l_return = l_return + i_pk_insert + vbNewLine

        l_return = l_return + "END;
/"
        Return l_return

    End Function

    Function GET_DML_IDENTIFIER(ByVal i_insert As String) As String

        DEBUGGER.SET_DEBUG("DML_API :: GET_DML_IDENTIFIER(" & i_insert & " )")

        Dim l_return As String
        Dim l_pos As Int32


        l_return = i_insert.ToLower
        l_return = Replace(l_return, " ", "")

        l_return = Replace(l_return, "insert", "")
        l_return = Replace(l_return, "into", "")

        l_pos = l_return.IndexOf("(")
        l_return = l_return.Substring(0, l_pos)
        Return l_return

    End Function

    Function GET_DML_IDENTIFIER_PK(ByVal i_insert As String) As String

        DEBUGGER.SET_DEBUG("DML_API :: GET_DML_IDENTIFIER_PK(" & i_insert & " )")

        Dim l_aux As String = ""
        Dim l_return As String

        Dim l_pos_init As Int32
        Dim l_pos_final As Int32

        l_aux = i_insert.ToLower
        l_aux = Replace(l_aux, " ", "")

        If l_aux.Contains("insert_into_translation") = True Then
            l_return = "translation_"
        ElseIf l_aux.Contains("insert_into_sys_message") Then
            l_return = "sys_message_"
        ElseIf l_aux.Contains("insert_into_sys_domain") = True Then
            l_return = "sys_domain_"
        ElseIf l_aux.Contains("insert_into_functionality_help") = True Then
            l_return = "functionality_help_"
        End If


        If l_aux.Contains("i_lang") = True Then
            l_pos_init = l_aux.IndexOf(">") + 1
            l_pos_final = l_aux.IndexOf(",")

            l_return = l_return + l_aux.Substring(l_pos_init, (l_pos_final - l_pos_init)) + "|alert|dml"
        Else
            l_pos_init = l_aux.IndexOf("(") + 1
            l_pos_final = l_aux.IndexOf(",")

            l_return = l_return + l_aux.Substring(l_pos_init, (l_pos_final - l_pos_init)) + "|alert|dml"
        End If

        Return l_return

    End Function

    Function GET_DML_INDEX(ByVal i_dml_identifier As String, ByVal i_dml_struct() As DML_TOOLS.dml_struct) As Int32

        DEBUGGER.SET_DEBUG("DML_API :: GET_DML_INDEX(" & i_dml_identifier & ", i_dml_struct )")

        Dim l_ret As Int32 = -1

        Try
            For i As Integer = 0 To i_dml_struct.Count() - 1
                If i_dml_struct(i).dml_identifier = i_dml_identifier Then
                    l_ret = i
                    Exit For
                End If
            Next
        Catch ex As Exception
            Return -1
        End Try

        Return l_ret

    End Function

    Function GET_TAG(ByVal i_dml_identifier As String, ByRef i_Dr As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("DML_API :: GET_TAG(" & i_dml_identifier & ")")

        Dim sql As String = " SELECT lower(fo.owner) AS owner,
                                CASE
                                     WHEN fo.obj_type = 'TABLE' THEN
                                      'dml'
                                     ELSE
                                      NULL
                                 END AS TYPE
                           FROM frmw_objects fo
                          WHERE fo.obj_name = upper('" & i_dml_identifier & "')"

        Dim cmd As New OracleCommand(sql, Connection.conn)

        Try

            cmd.CommandType = CommandType.Text
            i_Dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True

        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("DML_API :: GET_TAG")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            cmd.Dispose()
            Return False

        End Try

    End Function
End Class

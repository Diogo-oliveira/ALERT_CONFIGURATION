Imports Oracle.DataAccess.Client

'GET_INSTITUTION_ID(ByRef i_id_selected_item As Int64, ByVal i_oradb As String) As Int64
'GET_INSTITUTION(ByVal i_ID_INST As Int16, ByVal i_oradb As String) As String
'GET_ALL_INSTITUTIONS(ByVal i_oradb As String) As OracleDataReader
'GET_SOFT_INST(ByVal i_ID_INST As Int16, ByVal i_oradb As String) As OracleDataReader
'GET_CLIN_SERV(ByVal i_ID_INST As Int16, ByVal i_ID_SOFT As Int16, ByVal i_oradb As String) As OracleDataReader
'GET_SELECTED_SOFT(ByVal i_index As Int16, ByVal i_inst As Int16, ByVal i_oradb As String) As Int16
'GET_DEFAULT_TRANSLATION(ByVal i_lang As Int16, ByVal i_code_translation As String, ByVal i_oradb As String) As String
'GET_ID_LANG(ByVal i_id_institution As Int64, ByVal i_oradb As String) As Int16
'SET_TRANSLATION(ByVal i_id_lang As Integer, ByVal i_code_translation As String, ByVal i_desc As String, i_oradb As String) As Boolean

Public Class General

    Public g_notranslation As String = "no_translation"

    Public Function GET_INSTITUTION_ID(ByRef i_id_selected_item As Int64, ByVal i_conn As OracleConnection) As Int64

        Dim sql As String = "select decode(i.id_market,
                      1,
                      T.desc_lang_1,
                      2,
                      T.desc_lang_2,
                      3,
                      T.desc_lang_11,
                      4,
                      T.desc_lang_5,
                      5,
                      T.desc_lang_4,
                      6,
                      T.desc_lang_3,
                      7,
                      T.desc_lang_10,
                      8,
                      T.desc_lang_7,
                      9,
                      T.desc_lang_6,
                      10,
                      T.desc_lang_9,
                      12,
                      T.desc_lang_16,
                      16,
                      T.desc_lang_17,
                      17,
                      T.desc_lang_18,
                      19,
                      T.desc_lang_19),
                      i.id_institution
          from institution i
          join translation t
            on t.code_translation = i.code_institution
         where i.flg_available = 'Y'
           and i.flg_type = 'H'
           and (decode(i.id_market,
                       1,
                       T.desc_lang_1,
                       2,
                       T.desc_lang_2,
                       3,
                       T.desc_lang_11,
                       4,
                       T.desc_lang_5,
                       5,
                       T.desc_lang_4,
                       6,
                       T.desc_lang_3,
                       7,
                       T.desc_lang_10,
                       8,
                       T.desc_lang_7,
                       9,
                       T.desc_lang_6,
                       10,
                       T.desc_lang_9,
                       12,
                       T.desc_lang_16,
                       16,
                       T.desc_lang_17,
                       17,
                       T.desc_lang_18,
                       19,
                       T.desc_lang_19)) is not null
         order by 1 asc"

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Dim l_id_inst As Int64 = 0

        Dim i As Int64 = 0

        While dr.Read()

            If i = i_id_selected_item Then

                l_id_inst = dr.Item(1)

            End If

            i = i + 1

            If i > i_id_selected_item Then

                Exit While

            End If

        End While

        cmd.Dispose()
        dr.Dispose()
        dr.Close()

        Return l_id_inst

    End Function

    Public Function GET_INSTITUTION(ByVal i_ID_INST As Int64, ByVal i_conn As OracleConnection) As String

        Dim l_inst As String = ""

        Dim sql As String = "select decode(i.id_market,
              1,
              T.desc_lang_1,
              2,
              T.desc_lang_2,
              3,
              T.desc_lang_11,
              4,
              T.desc_lang_5,
              5,
              T.desc_lang_4,
              6,
              T.desc_lang_3,
              7,
              T.desc_lang_10,
              8,
              T.desc_lang_7,
              9,
              T.desc_lang_6,
              10,
              T.desc_lang_9,
              12,
              T.desc_lang_16,
              16,
              T.desc_lang_17,
              17,
              T.desc_lang_18,
              19,
              T.desc_lang_19),
              i.id_institution
  from institution i
  join translation t
    on t.code_translation = i.code_institution
 where i.flg_available = 'Y'
   and i.flg_type = 'H'
   and (decode(i.id_market,
              1,
              T.desc_lang_1,
              2,
              T.desc_lang_2,
              3,
              T.desc_lang_11,
              4,
              T.desc_lang_5,
              5,
              T.desc_lang_4,
              6,
              T.desc_lang_3,
              7,
              T.desc_lang_10,
              8,
              T.desc_lang_7,
              9,
              T.desc_lang_6,
              10,
              T.desc_lang_9,
              12,
              T.desc_lang_16,
              16,
              T.desc_lang_17,
              17,
              T.desc_lang_18,
              19,
              T.desc_lang_19)) is not null
and i.id_institution = " & i_ID_INST & "order by 1 asc"

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_inst = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_inst

    End Function


    Public Function GET_ALL_INSTITUTIONS(ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "select decode(i.id_market,
                                                          1,
                                                          T.desc_lang_1,
                                                          2,
                                                          T.desc_lang_2,
                                                          3,
                                                          T.desc_lang_11,
                                                          4,
                                                          T.desc_lang_5,
                                                          5,
                                                          T.desc_lang_4,
                                                          6,
                                                          T.desc_lang_3,
                                                          7,
                                                          T.desc_lang_10,
                                                          8,
                                                          T.desc_lang_7,
                                                          9,
                                                          T.desc_lang_6,
                                                          10,
                                                          T.desc_lang_9,
                                                          12,
                                                          T.desc_lang_16,
                                                          16,
                                                          T.desc_lang_17,
                                                          17,
                                                          T.desc_lang_18,
                                                          19,
                                                          T.desc_lang_19)
                                              from institution i
                                              join translation t
                                                on t.code_translation = i.code_institution
                                             where i.flg_available = 'Y'
                                               and i.flg_type = 'H'
                                               and (decode(i.id_market,
                                                          1,
                                                          T.desc_lang_1,
                                                          2,
                                                          T.desc_lang_2,
                                                          3,
                                                          T.desc_lang_11,
                                                          4,
                                                          T.desc_lang_5,
                                                          5,
                                                          T.desc_lang_4,
                                                          6,
                                                          T.desc_lang_3,
                                                          7,
                                                          T.desc_lang_10,
                                                          8,
                                                          T.desc_lang_7,
                                                          9,
                                                          T.desc_lang_6,
                                                          10,
                                                          T.desc_lang_9,
                                                          12,
                                                          T.desc_lang_16,
                                                          16,
                                                          T.desc_lang_17,
                                                          17,
                                                          T.desc_lang_18,
                                                          19,
                                                          T.desc_lang_19)) is not null
                                             order by 1 asc"
        Try

            Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Public Function GET_SOFT_INST(ByVal i_ID_INST As Int64, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "Select s.id_software, s.id_software || ' - ' ||s.name from alert_core_data.ab_software_institution i
                            join software s on s.id_software=i.id_ab_software
                            where i.id_ab_institution=" & i_ID_INST & "
                            and s.id_software > 0
                            order by 1 asc"

        Try

            Dim cmd As New OracleCommand(sql, i_conn)
            cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()
            cmd.Dispose()

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Public Function GET_CLIN_SERV(ByVal i_ID_INST As Int64, ByVal i_ID_SOFT As Int16, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = GET_ID_LANG(i_ID_INST, i_conn)

        Dim sql As String = "   Select pk_translation.get_translation(" & l_id_language & ",dep.code_department) || ' - ' ||  pk_translation.get_translation(" & l_id_language & ",c.code_clinical_service),
                                      d.id_dep_clin_serv from alert.dep_clin_serv d
                                 join alert.clinical_service c
                                 on c.id_clinical_service=d.id_clinical_service
                                 join alert.department dep on dep.id_department=d.id_department                                 
                                 join software s on s.id_software=dep.id_software
                                 JOIN INSTITUTION I ON I.id_institution=DEP.ID_INSTITUTION                                 
                                 where dep.id_institution= " & i_ID_INST & "
                                 and dep.id_software= " & i_ID_SOFT & "
                                 and dep.flg_available='Y'
                                 and c.flg_available='Y'
                                 and d.flg_available='Y'
                                 and pk_translation.get_translation(" & l_id_language & ",c.code_clinical_service) is not null
                                 order by 1 asc"

        Try

            Dim cmd As New OracleCommand(sql, i_conn)
            cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Function GET_SELECTED_SOFT(ByVal i_index As Int16, ByVal i_inst As Int64, ByVal i_conn As OracleConnection) As Int16

        Dim sql As String = "Select s.id_software, s.id_software || ' - ' ||s.name from alert_core_data.ab_software_institution i
                            join software s on s.id_software=i.id_ab_software
                            where i.id_ab_institution=" & i_inst & "
                            and s.id_software > 0
                            order by 1 asc"

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Dim i As Int16 = 0
        Dim l_soft As Int16 = -1

        While dr.Read()

            If i = i_index Then

                l_soft = dr.Item(0)
                Exit While

            End If

            i = i + 1

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_soft

    End Function

    Function GET_DEFAULT_TRANSLATION(ByVal i_lang As Int16, ByVal i_code_translation As String, ByVal i_conn As OracleConnection) As String

        'IMPORTANTE: Quando se chama esta função é necessário comparar SEMPRE o resultado com a varável g_notranslation - SET TRANSLATION faz isso

        Dim sql As String = "select alert_default.pk_translation_default.get_translation_default(" & i_lang & ",'" & i_code_translation & "') from dual"

        Dim translation As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Try

            While dr.Read()

                translation = dr.Item(0)

            End While

            cmd.Dispose()

        Catch ex As Exception

            cmd.Dispose()
            dr.Dispose()
            dr.Close()
            Return g_notranslation

        End Try

        dr.Dispose()
        dr.Close()
        Return translation

    End Function

    Function GET_ALERT_TRANSLATION(ByVal i_lang As Int16, ByVal i_code_translation As String, ByVal i_conn As OracleConnection) As String

        'IMPORTANTE: Quando se chama esta função é necessário comparar SEMPRE o resultado com a varável g_notranslation - SET TRANSLATION faz isso

        Dim sql As String = "select pk_translation.get_translation(" & i_lang & ",'" & i_code_translation & "') from dual"

        Dim translation As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Try

            While dr.Read()

                translation = dr.Item(0)

            End While

            cmd.Dispose()

        Catch ex As Exception

            cmd.Dispose()
            dr.Dispose()
            dr.Close()
            Return g_notranslation

        End Try

        dr.Dispose()
        dr.Close()
        Return translation

    End Function

    Function GET_ID_LANG(ByVal i_id_institution As Int64, ByVal i_conn As OracleConnection) As Int16

        Dim l_id_market As Int16 = 0

        Dim sql As String = "Select i.id_market from institution i
                             where i.id_institution= " & i_id_institution

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()


        While dr.Read()

            l_id_market = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        If l_id_market = 1 Then

            Return 1

        ElseIf l_id_market = 2 Then

            Return 2

        ElseIf l_id_market = 3 Then

            Return 11

        ElseIf l_id_market = 4 Then

            Return 5

        ElseIf l_id_market = 5 Then

            Return 4

        ElseIf l_id_market = 6 Then

            Return 3

        ElseIf l_id_market = 7 Then

            Return 10

        ElseIf l_id_market = 8 Then

            Return 7

        ElseIf l_id_market = 9 Then

            Return 6

        ElseIf l_id_market = 10 Then

            Return 9

        ElseIf l_id_market = 12 Then

            Return 16

        ElseIf l_id_market = 16 Then

            Return 17

        ElseIf l_id_market = 17 Then

            Return 18

        ElseIf l_id_market = 19 Then

            Return 19

        End If

        Return 0

    End Function

    Function SET_TRANSLATION(ByVal i_id_lang As Integer, ByVal i_code_translation As String, ByVal i_desc As String, ByVal i_conn As OracleConnection) As Boolean

        If i_desc = g_notranslation Then

            Return False

        Else

            Dim Sql = "begin pk_translation.insert_into_translation( " & i_id_lang & " , '" & i_code_translation & "' , '" & i_desc & "' ); end;"

            Dim cmd_insert_trans As New OracleCommand(Sql, i_conn)
            cmd_insert_trans.CommandType = CommandType.Text

            Try

                cmd_insert_trans.ExecuteNonQuery()
                cmd_insert_trans.Dispose()

            Catch ex As Exception

                'Se a inserção falhar, certamente é por causa da existência do caractér '

                Dim l_desc_aux As String = ""

                For i As Integer = 0 To i_desc.Count() - 1

                    If i_desc(i) = "'" Then

                        l_desc_aux = l_desc_aux & "''"

                    Else

                        l_desc_aux = l_desc_aux & i_desc(i)

                    End If

                Next

                Sql = "begin pk_translation.insert_into_translation( " & i_id_lang & " , '" & i_code_translation & "' , '" & l_desc_aux & "' ); end;"

                cmd_insert_trans.Dispose()

                Dim cmd_insert_trans_new As New OracleCommand(Sql, i_conn)
                cmd_insert_trans_new.CommandType = CommandType.Text
                cmd_insert_trans_new.ExecuteNonQuery()
                cmd_insert_trans_new.Dispose()

                Return True

            End Try

            Return True

        End If
    End Function

    Function CHECK_TRANSLATIONS(ByVal i_id_lang As Integer, ByVal i_code_translation_default As String, ByVal i_code_translation_alert As String, ByVal i_conn As OracleConnection) As Boolean

        Dim l_desc_default As String = GET_DEFAULT_TRANSLATION(i_id_lang, i_code_translation_default, i_conn)
        Dim l_desc_alert As String = GET_ALERT_TRANSLATION(i_id_lang, i_code_translation_alert, i_conn)

        If l_desc_default = l_desc_alert And l_desc_default <> g_notranslation Then

            Return True

        Else

            Return False

        End If

    End Function

    Function GET_LAB_ROOMS(ByVal i_institution As Int64, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT pk_translation.get_translation(" & GET_ID_LANG(i_institution, i_conn) & ", r.code_room), r.id_room
                                FROM alert.department d
                                JOIN alert.room r ON r.id_department = d.id_department
                                JOIN translation t ON t.code_translation = r.code_room
                                WHERE d.id_institution = " & i_institution & "
                                AND r.flg_lab = 'Y'
                                AND r.flg_available = 'Y'
                                ORDER BY 1"

        Try

            Dim cmd As New OracleCommand(sql, i_conn)
            cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

End Class

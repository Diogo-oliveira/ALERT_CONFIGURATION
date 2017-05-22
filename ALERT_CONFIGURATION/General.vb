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

    Public Function GET_INSTITUTION_ID(ByRef i_id_selected_item As Int64) As Int64

        DEBUGGER.SET_DEBUG("GENERAL :: GET_INSTITUTION_ID (" & i_id_selected_item & ")")

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

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try
            dr = cmd.ExecuteReader()
        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: GET_INSTITUTION_ID")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG("CONFLICT OF ALERT(R) VERSIONS. TRYING NEW QUERY.")
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            sql = "select decode(i.id_market,
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
                      T.desc_lang_1),
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
                       T.desc_lang_1)) is not null
         order by 1 asc"

            Try
                Dim cmd_Old_version As New OracleCommand(sql, Connection.conn)
                cmd_Old_version.CommandType = CommandType.Text
                dr = cmd_Old_version.ExecuteReader()
                cmd_Old_version.Dispose()

            Catch ex2 As Exception
                DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: GET_INSTITUTION_ID")
                DEBUGGER.SET_DEBUG(sql)
                DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            End Try

        End Try

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

    Public Function GET_INSTITUTION(ByVal i_ID_INST As Int64) As String

        DEBUGGER.SET_DEBUG("GENERAL :: GET_INSTITUTION(" & i_ID_INST & ")")

        Dim l_id_language As Int16 = GET_ID_LANG(i_ID_INST)

        Dim l_inst As String = ""

        Dim sql As String = "SELECT pk_translation.get_translation(" & l_id_language & ", i.code_institution)
                            FROM institution i
                            JOIN translation t ON t.code_translation = i.code_institution
                            WHERE i.id_institution = " & i_ID_INST & "
                            AND i.flg_available = 'Y'
                            and i.flg_type='H'"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try

            dr = cmd.ExecuteReader()

            While dr.Read()

                l_inst = dr.Item(0)

            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: GET_INSTITUTION")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        End Try

        Return l_inst

    End Function

    Public Function GET_ALL_INSTITUTIONS(ByRef i_dr As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("GENERAL :: GET_ALL_INSTITUTIONS")

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

            Dim cmd As New OracleCommand(sql, Connection.conn)
            cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: GET_ALL_INSTITUTIONS")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG("CONFLICT OF ALERT(R) VERSIONS. TRYING NEW QUERY.")
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            Try

                sql = "select decode(i.id_market,
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
                                                          T.desc_lang_1)
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
                                                          T.desc_lang_1)) is not null
                                             order by 1 asc"

                Dim cmd_Old_version As New OracleCommand(sql, Connection.conn)
                cmd_Old_version.CommandType = CommandType.Text

                i_dr = cmd_Old_version.ExecuteReader()

                cmd_Old_version.Dispose()

                Return True

            Catch ex_2 As Exception

                DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: GET_ALL_INSTITUTIONS")
                DEBUGGER.SET_DEBUG(ex.Message)
                DEBUGGER.SET_DEBUG(sql)
                DEBUGGER.SET_DEBUG_ERROR_CLOSE()

                Return False

            End Try

            Return False

        End Try

    End Function

    Public Function GET_SOFT_INST(ByVal i_ID_INST As Int64, ByRef i_dr As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("GENERAL :: GET_SOFT_INST(" & i_ID_INST & ")")

        Dim sql As String = ""

        If i_ID_INST = 0 Then

            sql = "SELECT s.id_software, s.id_software || ' - ' || decode(s.name,'todos','ALL', s.name)
                    From software s
                    Where s.id_software >= 0
                    Order By 1 ASC"

        Else

            sql = "Select s.id_software, s.id_software || ' - ' ||s.name from alert_core_data.ab_software_institution i
                            join software s on s.id_software=i.id_ab_software
                            where i.id_ab_institution=" & i_ID_INST & "
                            and s.id_software > 0
                            order by 1 asc"

        End If

        Try

            Dim cmd As New OracleCommand(sql, Connection.conn)
            cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()
            cmd.Dispose()

            Return True

        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: GET_SOFT_INST")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            Return False

        End Try

    End Function

    Public Function GET_CLIN_SERV(ByVal i_ID_INST As Int64, ByVal i_ID_SOFT As Int16, ByRef i_dr As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("GENERAL :: GET_CLIN_SERV(" & i_ID_INST & ", " & i_ID_SOFT & ")")

        Dim l_id_language As Int16 = GET_ID_LANG(i_ID_INST)

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

            Dim cmd As New OracleCommand(sql, Connection.conn)
            cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: GET_CLIN_SERV")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            Return False

        End Try

    End Function

    Function GET_SELECTED_SOFT(ByVal i_index As Int16, ByVal i_inst As Int64) As Int16

        DEBUGGER.SET_DEBUG("GENERAL :: GET_SELECTED_SOFT(" & i_index & ", " & i_inst & ")")

        Dim sql As String = "Select s.id_software, s.id_software || ' - ' ||s.name from alert_core_data.ab_software_institution i
                            join software s on s.id_software=i.id_ab_software
                            where i.id_ab_institution=" & i_inst & "
                            and s.id_software > 0
                            order by 1 asc"

        Dim cmd As New OracleCommand(sql, Connection.conn)
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

    Function GET_DEFAULT_TRANSLATION(ByVal i_lang As Int16, ByVal i_code_translation As String) As String

        'IMPORTANTE: Quando se chama esta função é necessário comparar SEMPRE o resultado com a varável g_notranslation - SET TRANSLATION faz isso

        DEBUGGER.SET_DEBUG("GENERAL :: GET_DEFAULT_TRANSLATION(" & i_lang & ", " & i_code_translation & ")")

        Dim sql As String = "select alert_default.pk_translation_default.get_translation_default(" & i_lang & ",'" & i_code_translation & "') from dual"

        Dim translation As String = ""

        Dim cmd As New OracleCommand(sql, Connection.conn)
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

    Function GET_ALERT_TRANSLATION(ByVal i_lang As Int16, ByVal i_code_translation As String) As String

        'IMPORTANTE: Quando se chama esta função é necessário comparar SEMPRE o resultado com a varável g_notranslation - SET TRANSLATION faz isso
        DEBUGGER.SET_DEBUG("GENERAL :: GET_ALERT_TRANSLATION(" & i_lang & ", " & i_code_translation & ")")

        Dim sql As String = "select pk_translation.get_translation(" & i_lang & ",'" & i_code_translation & "') from dual"

        Dim translation As String = ""

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Try

            While dr.Read()

                translation = dr.Item(0)

            End While

            cmd.Dispose()

        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: GET_ALERT_TRANSLATION")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            cmd.Dispose()
            dr.Dispose()
            dr.Close()
            Return g_notranslation

        End Try

        dr.Dispose()
        dr.Close()
        Return translation

    End Function

    Function GET_ID_LANG(ByVal i_id_institution As Int64) As Int16

        DEBUGGER.SET_DEBUG("GENERAL :: GET_ID_LANG(" & i_id_institution & ")")

        Dim l_id_market As Int16 = 0

        Dim sql As String = "Select i.id_market from institution i
                             where i.id_institution= " & i_id_institution

        Dim cmd As New OracleCommand(sql, Connection.conn)
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

        ElseIf l_id_market = 18 Then   'Na Prática, o mercado KW usa a lingua 7

            Return 7

        ElseIf l_id_market = 19 Then

            Return 19

        End If

        Return 0

    End Function

    Function SET_TRANSLATION(ByVal i_id_lang As Integer, ByVal i_code_translation As String, ByVal i_desc As String) As Boolean

        DEBUGGER.SET_DEBUG("GENERAL :: SET_TRANSLATION(" & i_id_lang & ", " & i_code_translation & ", " & i_desc & ")")

        If i_desc = g_notranslation Then

            Return False

        Else

            Dim Sql = "begin pk_translation.insert_into_translation( " & i_id_lang & " , '" & i_code_translation & "' , '" & i_desc & "' ); end;"

            Dim cmd_insert_trans As New OracleCommand(Sql, Connection.conn)
            cmd_insert_trans.CommandType = CommandType.Text

            Try

                cmd_insert_trans.ExecuteNonQuery()
                cmd_insert_trans.Dispose()

            Catch ex As Exception

                DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: SET_TRANSLATION")
                DEBUGGER.SET_DEBUG(ex.Message)
                DEBUGGER.SET_DEBUG(Sql)
                DEBUGGER.SET_DEBUG_ERROR_CLOSE()

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

                Dim cmd_insert_trans_new As New OracleCommand(Sql, Connection.conn)
                cmd_insert_trans_new.CommandType = CommandType.Text
                cmd_insert_trans_new.ExecuteNonQuery()
                cmd_insert_trans_new.Dispose()

                Return True

            End Try

            Return True

        End If
    End Function

    Function CHECK_TRANSLATIONS(ByVal i_id_lang As Integer, ByVal i_code_translation_default As String, ByVal i_code_translation_alert As String) As Boolean

        DEBUGGER.SET_DEBUG("GENERAL :: CHECK_TRANSLATIONS(" & i_id_lang & ", " & i_code_translation_default & ", " & i_code_translation_alert & ")")

        Dim l_desc_default As String = GET_DEFAULT_TRANSLATION(i_id_lang, i_code_translation_default)
        Dim l_desc_alert As String = GET_ALERT_TRANSLATION(i_id_lang, i_code_translation_alert)

        If l_desc_default = l_desc_alert And l_desc_default <> g_notranslation Then

            Return True

        Else

            Return False

        End If

    End Function

    Function GET_LAB_ROOMS(ByVal i_institution As Int64, ByRef i_dr As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("GENERAL :: GET_LAB_ROOMS(" & i_institution & ")")

        Dim sql As String = "SELECT pk_translation.get_translation(" & GET_ID_LANG(i_institution) & ", r.code_room), r.id_room
                                FROM alert.department d
                                JOIN alert.room r ON r.id_department = d.id_department
                                JOIN translation t ON t.code_translation = r.code_room
                                WHERE d.id_institution = " & i_institution & "
                                AND r.flg_lab = 'Y'
                                AND r.flg_available = 'Y'
                                ORDER BY 1"

        Try

            Dim cmd As New OracleCommand(sql, Connection.conn)
            cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: GET_LAB_ROOMS")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            Return False

        End Try

    End Function

    Function GET_SYSCONFIG(ByVal i_institution As Int64, ByVal i_id_software As Integer, ByVal i_sysconfig As String, ByRef i_dr As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("GENERAL :: GET_SYSCONFIG(" & i_institution & ", " & i_id_software & ", " & i_sysconfig & ")")

        Dim sql As String = "SELECT alert.pk_sysconfig.get_config('" & i_sysconfig & "', profissional(0, " & i_institution & ", " & i_id_software & "))
                                         FROM dual"

        Try

            Dim cmd As New OracleCommand(sql, Connection.conn)
            cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: GET_SYSCONFIG")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            Return False

        End Try

    End Function

    Function GET_MARKETS(ByRef i_dr As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("GENERAL :: GET_MARKETS")

        Dim sql As String = "SELECT m.id_market, (m.id_market || ' - ' || decode(t.desc_lang_2, 'None', 'ALL', t.desc_lang_2))
                            FROM market m
                            JOIN translation t ON t.code_translation = m.code_market
                            ORDER BY 1 ASC"

        Try

            Dim cmd As New OracleCommand(sql, Connection.conn)
            cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: GET_MARKETS")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            Return False

        End Try

    End Function

    Function SET_SYSCONFIG(ByVal i_id_sysconfig As String, ByVal i_value As String, ByVal i_institution As Int64, ByVal i_sofware As Integer, ByVal i_market As Integer) As Boolean

        DEBUGGER.SET_DEBUG("GENERAL :: SET_SYSCONFIG(" & i_id_sysconfig & ", " & i_value & ", " & i_institution & ", " & i_market & ")")

        Dim Sql As String = "DECLARE

                                l_desc            sys_config.id_sys_config%TYPE;
                                l_fill_type       sys_config.fill_type%TYPE;
                                l_client_config   sys_config.client_configuration%TYPE;
                                l_internal_config sys_config.internal_configuration%TYPE;
                                l_global_config   sys_config.global_configuration%TYPE;
                                l_flg_schema      sys_config.flg_schema%TYPE;
                                l_mvalue          sys_config.mvalue%TYPE;

                            BEGIN

                                SELECT DISTINCT s.desc_sys_config, s.fill_type, s.client_configuration, s.internal_configuration, s.global_configuration, s.flg_schema, s.mvalue
                                INTO l_desc, l_fill_type, l_client_config, l_internal_config, l_global_config, l_flg_schema, l_mvalue
                                FROM sys_config s
                                WHERE s.id_sys_config = '" & i_id_sysconfig & "'
                                and rownum=1;

                                INSERT INTO sys_config
                                    (id_sys_config,
                                     VALUE,
                                     desc_sys_config,
                                     id_institution,
                                     id_software,
                                     fill_type,
                                     client_configuration,
                                     internal_configuration,
                                     global_configuration,
                                     flg_schema,
                                     id_market,
                                     mvalue)
                                VALUES
                                    ('" & i_id_sysconfig & "',
                                     '" & i_value & "',
                                     l_desc,
                                     " & i_institution & ",
                                     " & i_sofware & ",
                                     l_fill_type,
                                     l_client_config,
                                     l_internal_config,
                                     l_global_config,
                                     l_flg_schema,
                                     " & i_market & ",
                                     l_mvalue);

                            EXCEPTION
                                WHEN dup_val_on_index THEN
                                    UPDATE sys_config s
                                    SET s.value = '" & i_value & "'
                                    WHERE s.id_sys_config ='" & i_id_sysconfig & "'
                                    AND s.id_software = " & i_sofware & "
                                    AND s.id_institution = " & i_institution & "
                                    AND s.id_market =" & i_market & ";

                            END;"

        Dim cmd_insert_SC As New OracleCommand(Sql, Connection.conn)
        cmd_insert_SC.CommandType = CommandType.Text

        Try

            cmd_insert_SC.ExecuteNonQuery()


        Catch ex As Exception 'Dá exceção nas versões antigas. Não existe m_value

            DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: SET_SYSCONFIG")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(Sql)
            DEBUGGER.SET_DEBUG("CONFLICT OF ALERT(R) VERSIONS. TRYING NEW QUERY.")
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            Dim Sql_versao_antiga As String = "DECLARE

                                l_desc            sys_config.id_sys_config%TYPE;
                                l_fill_type       sys_config.fill_type%TYPE;
                                l_client_config   sys_config.client_configuration%TYPE;
                                l_internal_config sys_config.internal_configuration%TYPE;
                                l_global_config   sys_config.global_configuration%TYPE;
                                l_flg_schema      sys_config.flg_schema%TYPE;                                

                            BEGIN

                                SELECT DISTINCT s.desc_sys_config, s.fill_type, s.client_configuration, s.internal_configuration, s.global_configuration, s.flg_schema
                                INTO l_desc, l_fill_type, l_client_config, l_internal_config, l_global_config, l_flg_schema
                                FROM sys_config s
                                WHERE s.id_sys_config = '" & i_id_sysconfig & "'
                                and rownum=1;

                                INSERT INTO sys_config
                                    (id_sys_config,
                                     VALUE,
                                     desc_sys_config,
                                     id_institution,
                                     id_software,
                                     fill_type,
                                     client_configuration,
                                     internal_configuration,
                                     global_configuration,
                                     flg_schema,
                                     id_market )
                                VALUES
                                    ('" & i_id_sysconfig & "',
                                     '" & i_value & "',
                                     l_desc,
                                     " & i_institution & ",
                                     " & i_sofware & ",
                                     l_fill_type,
                                     l_client_config,
                                     l_internal_config,
                                     l_global_config,
                                     l_flg_schema,
                                     " & i_market & " );

                            EXCEPTION
                                WHEN dup_val_on_index THEN
                                    UPDATE sys_config s
                                    SET s.value = '" & i_value & "'
                                    WHERE s.id_sys_config ='" & i_id_sysconfig & "'
                                    AND s.id_software = " & i_sofware & "
                                    AND s.id_institution = " & i_institution & "
                                    AND s.id_market =" & i_market & ";
    
                            END;"

            Dim cmd_insert_SC_v_antiga As New OracleCommand(Sql_versao_antiga, Connection.conn)
            cmd_insert_SC_v_antiga.CommandType = CommandType.Text

            Try
                cmd_insert_SC_v_antiga.ExecuteNonQuery()
                cmd_insert_SC_v_antiga.Dispose()

            Catch ex_2 As Exception

                DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: SET_SYSCONFIG")
                DEBUGGER.SET_DEBUG(ex.Message)
                DEBUGGER.SET_DEBUG(Sql)
                DEBUGGER.SET_DEBUG_ERROR_CLOSE()

                cmd_insert_SC_v_antiga.Dispose()
                Return False

            End Try

        End Try

        cmd_insert_SC.Dispose()

        Return True

    End Function

    Function GET_PROFILES(ByVal i_software As Integer, ByVal i_type As String, ByRef o_dr As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("GENERAL :: GET_PROFILES(" & i_software & ", " & i_type & ")")

        Dim sql As String = "SELECT pt.id_profile_template, pt.intern_name_templ, pt.flg_type
                                FROM alert.profile_template pt
                                WHERE pt.flg_available = 'Y'
                                and pt.id_software=" & i_software & "
                                and pt.flg_type is not null "

        If i_type = "-1" Then

            sql = sql & "order by 1 asc"

        ElseIf i_type = "-2" Then

            sql = sql & "   and pt.flg_type not in ('D','N','A')
                            order by 1 asc"

        Else

            sql = sql & "   and pt.flg_type ='" & i_type & "'
                            order by 1 asc"

        End If

        Try

            Dim cmd As New OracleCommand(sql, Connection.conn)
            cmd.CommandType = CommandType.Text

            o_dr = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: GET_PROFILES")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            Return False

        End Try

    End Function

    Function GET_PROFILE_TYPES(ByRef o_dr_pt As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("GENERAL :: GET_PROFILE_TYPES")

        Dim sql As String = "SELECT DISTINCT pt.flg_type
                                FROM alert.profile_template pt
                                WHERE pt.flg_type IS NOT NULL
                                ORDER BY 1 ASC "

        Dim dr As OracleDataReader

        Try

            Dim cmd As New OracleCommand(sql, Connection.conn)
            cmd.CommandType = CommandType.Text

            o_dr_pt = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: GET_PROFILE_TYPES")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            Return False

        End Try

    End Function

    Function GET_PROFILE_TYPE(ByVal i_id_profile_template As Int64) As String

        DEBUGGER.SET_DEBUG("GENERAL :: GET_PROFILE_TYPE")

        Dim sql As String = "SELECT pt.flg_type
                                FROM alert.profile_template pt
                                WHERE pt.flg_type IS NOT NULL
                                AND PT.ID_PROFILE_TEMPLATE=" & i_id_profile_template

        Dim dr As OracleDataReader
        Dim l_pt_type As String

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        dr = cmd.ExecuteReader()

        cmd.Dispose()

        While dr.Read()

            l_pt_type = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()

        Return l_pt_type

    End Function

    Function GET_EPIS_TYPES(ByRef o_dr As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("GENERAL :: GET_EPIS_TYPES")

        Dim sql As String = "SELECT et.id_epis_type, t.desc_lang_2  --É para devolver em inglês
                                FROM alert.epis_type et
                                JOIN translation t ON t.code_translation = et.code_epis_type
                                               AND et.flg_available = 'Y'
                                ORDER BY 1 ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)

        Try
            cmd.CommandType = CommandType.Text
            o_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: GET_EPIS_TYPES")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            Return False
        End Try

    End Function

    Function GET_INSTITUTION_MARKET(ByVal i_institution As Int64) As Integer

        DEBUGGER.SET_DEBUG("GENERAL :: GET_INSTITUTION_MARKET (" & i_institution & ")")

        Dim sql As String = "SELECT i.id_ab_market
                                FROM alert_core_data.ab_institution i
                                WHERE i.id_ab_institution = " & i_institution

        Dim dr As OracleDataReader
        Dim l_id_market As Integer = -1

        Try

            Dim cmd As New OracleCommand(sql, Connection.conn)
            cmd.CommandType = CommandType.Text

            dr = cmd.ExecuteReader()

            cmd.Dispose()

            While dr.Read()

                l_id_market = dr.Item(0)

            End While

            Return l_id_market

        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("GENERAL :: GET_INSTITUTION_MARKET")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

        End Try

    End Function

End Class

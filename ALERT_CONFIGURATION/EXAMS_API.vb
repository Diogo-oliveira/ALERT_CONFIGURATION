Imports Oracle.DataAccess.Client
Public Class EXAMS_API

    Dim db_access_general As New General

    ''Estrutura dos exames carregados do default
    Public Structure exams_default
        Public id_content_category As String
        Public desc_category As String
        Public id_content_exam As String
        Public desc_exam As String
        Public flg_first_result As String
        Public flg_execute As String
        Public flg_timeout As String
        Public flg_result_notes As String
        Public flg_first_execute As String
        Public age_min As Integer
        Public age_max As Integer
        Public gender As String
    End Structure

    Public Structure exams_alert

        Public id_exam As String
        Public desc_exam As String

    End Structure

    Public Structure exams_alert_flg

        Public id_content_exam_cat As String
        Public id_content_exam As String
        Public desc_exam As String
        Public flg_new As String

    End Structure

    Public Function GET_INSTITUTION_ID(ByRef i_id_selected_item As Int64, ByVal i_oradb As String) As Int64

        Dim oradb As String = i_oradb

        Dim conn_new As New OracleConnection(oradb)

        conn_new.Open()

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

        Dim cmd As New OracleCommand(sql, conn_new)
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

        conn_new.Close()

        conn_new.Dispose()

        Return l_id_inst

    End Function

    Public Function GET_INSTITUTION(ByVal i_ID_INST As Int64, ByVal i_oradb As String) As String

        Dim l_inst As String = ""

        Dim oradb As String = i_oradb

        Dim conn_new As New OracleConnection(oradb)

        conn_new.Open()

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

        Dim cmd As New OracleCommand(sql, conn_new)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()


        While dr.Read()

            l_inst = dr.Item(0)

        End While

        conn_new.Close()

        conn_new.Dispose()

        Return l_inst

    End Function


    Public Function GET_ALL_INSTITUTIONS(ByVal i_oradb As String) As OracleDataReader

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

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

        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Return dr

    End Function

    Public Function GET_SOFT_INST(ByVal i_ID_INST As Int64, ByVal i_oradb As String) As OracleDataReader

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

        Dim sql As String = "Select s.id_software, s.id_software || ' - ' ||s.name from alert_core_data.ab_software_institution i
                            join software s on s.id_software=i.id_ab_software
                            where i.id_ab_institution=" & i_ID_INST & "
                            and s.id_software > 0
                            order by 1 asc"

        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Return dr

        conn.Close()

    End Function

    Public Function GET_CLIN_SERV(ByVal i_ID_INST As Int64, ByVal i_ID_SOFT As Int16, ByVal i_oradb As String) As OracleDataReader

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

        Dim sql As String = "   Select (decode(i.id_market,
              1,
              tdep.desc_lang_1,
              2,
              tdep.desc_lang_2,
              3,
              tdep.desc_lang_11,
              4,
              tdep.desc_lang_5,
              5,
              tdep.desc_lang_4,
              6,
              tdep.desc_lang_3,
              7,
              tdep.desc_lang_10,
              8,
              tdep.desc_lang_7,
              9,
              tdep.desc_lang_6,
              10,
              tdep.desc_lang_9,
              12,
              tdep.desc_lang_16,
              16,
              tdep.desc_lang_17,
              17,
              tdep.desc_lang_18,
              19,
              tdep.desc_lang_19) || ' - ' || decode(i.id_market,
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
              T.desc_lang_19)), d.id_dep_clin_serv from alert.dep_clin_serv d
 join alert.clinical_service c
 on c.id_clinical_service=d.id_clinical_service
 join alert.department dep on dep.id_department=d.id_department
 join translation t on t.code_translation=c.code_clinical_service
 join software s on s.id_software=dep.id_software
 JOIN INSTITUTION I ON I.id_institution=DEP.ID_INSTITUTION
 join translation tdep on tdep.code_translation=dep.code_department
 where dep.id_institution= " & i_ID_INST & "
 and dep.id_software= " & i_ID_SOFT & "
 and dep.flg_available='Y'
 and c.flg_available='Y'
 and d.flg_available='Y'
 order by 1 asc"

        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Return dr

    End Function

    Function GET_SELECTED_SOFT(ByVal i_index As Int16, ByVal i_inst As Int64, ByVal i_oradb As String) As Int16

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

        Dim sql As String = "Select s.id_software, s.id_software || ' - ' ||s.name from alert_core_data.ab_software_institution i
                            join software s on s.id_software=i.id_ab_software
                            where i.id_ab_institution=" & i_inst & "
                            and s.id_software > 0
                            order by 1 asc"

        Dim cmd As New OracleCommand(sql, conn)
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

        conn.Close()
        conn.Dispose()

        Return l_soft

    End Function

    Function GET_FREQ_EXAM(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_flg_type As Integer, ByVal i_id_dep_clin_serv As Int64, ByVal i_id_exam_type As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Dim sql As String = ""

        sql = "SELECT DISTINCT ec.id_content, e.id_content, t.desc_lang_" & l_id_language & "

                    FROM alert.exam_dep_clin_serv s

                    JOIN alert.exam e ON e.id_exam = s.id_exam
                    JOIN translation t ON t.code_translation = e.code_exam
                                   AND t.desc_lang_" & l_id_language & " IS NOT NULL
                    JOIN alert.exam_cat ec ON ec.id_exam_cat = e.id_exam_cat
                                       AND ec.flg_available = 'Y'

                    WHERE s.id_software IN (0, " & i_software & ")
                    AND (s.id_institution IN (0, " & i_institution & ") or s.id_institution is null)
                    AND e.flg_available = 'Y'
                    AND e.flg_type = '" & i_id_exam_type & "'
                    and s.id_dep_clin_serv=" & i_id_dep_clin_serv & " "

        If i_flg_type = 0 Then

            sql = sql & " And s.flg_type In ('M','A')
                    ORDER BY 3 ASC;"

        ElseIf i_flg_type = 1 Then

            sql = sql & " And s.flg_type In ('M')
                          ORDER BY 3 ASC;"

        Else

            sql = sql & " And s.flg_type In ('A')
                          ORDER BY 3 ASC;"

        End If

        Dim cmd As New OracleCommand(sql, i_conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception
            cmd.Dispose()
            Return False
        End Try

    End Function


    Function GET_SELECTED_DEP_CLIN_SERV(ByVal i_ID_INST As Int16, ByVal i_ID_SOFT As Int16, ByVal i_index As Int16, ByVal i_oradb As String) As Int64

        Dim oradb As String = i_oradb

        Dim l_dep_clin_serv As Int64 = -1

        Dim dr As OracleDataReader = GET_CLIN_SERV(i_ID_INST, i_ID_SOFT, i_oradb)

        Dim i As Int16 = 0

        While dr.Read()

            If i = i_index Then

                l_dep_clin_serv = dr.Item(1)
                Exit While

            End If

            i = i + 1

        End While

        Return l_dep_clin_serv

    End Function

    Function DELETE_EXAMS_DEP_CLIN_SERV(ByVal i_exam As Int64(), ByVal i_dep_clin_serv As Int64, ByVal i_oradb As String) As Boolean

        Try

            Dim oradb As String = i_oradb

            Dim conn As New OracleConnection(oradb)

            conn.Open()

            For i As Integer = 0 To i_exam.Count() - 1

                Dim sql As String = "delete from ALERT.EXAM_DEP_CLIN_SERV S
                             WHERE S.ID_EXAM = " & i_exam(i) & "                           
                             And S.FLG_TYPE='M'
                             And S.ID_DEP_CLIN_SERV= " & i_dep_clin_serv

                Dim cmd As New OracleCommand(sql, conn)
                cmd.CommandType = CommandType.Text

                cmd.ExecuteNonQuery()

            Next

            conn.Close()
            conn.Dispose()

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Function GET_EXAMS_CAT(ByVal i_id_inst As Int64, ByVal i_id_soft As Integer, ByVal i_exam_type As String, ByVal i_flg_type As Integer, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_id_inst, i_conn)

        Dim sql As String = "Select distinct (tec.desc_lang_" & l_id_language & "),ec.id_exam_cat              
                             from alert.exam_dep_clin_serv d
                             join alert.exam e on e.id_exam=d.id_exam
                             join alert.exam_cat ec on ec.id_exam_cat=e.id_exam_cat
                             join translation tec on tec.code_translation=ec.code_exam_cat                             
                             where d.id_institution in(0, " & i_id_inst & ")
                             and e.flg_type='" & i_exam_type & "'
                             and e.flg_available='Y' and ec.flg_available='Y'
                             and d.id_software= " & i_id_soft & ""

        If i_flg_type = 0 Then

            sql = sql & "And d.flg_type IN ('P', 'M', 'A', 'B') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And d.flg_type IN ('P', 'M') "

        Else

            sql = sql & "And d.flg_type IN ('A', 'B') "

        End If

        sql = sql & "order by 1 asc"

        Dim cmd As New OracleCommand(sql, i_conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception
            cmd.Dispose()
            Return False
        End Try

    End Function

    Function GET_EXAMS(ByVal i_id_inst As Int64, ByVal i_id_soft As Int64, ByVal i_id_content_exam_cat As String, ByVal i_exam_type As String, ByVal i_flg_type As Integer, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_id_inst, i_conn)

        Dim sql As String = ""

        sql = " Select distinct pk_translation.get_translation(" & l_id_language & ",e.code_exam),ec.id_content,e.id_content          
                     from alert.exam_dep_clin_serv d
                     join alert.exam e on e.id_exam=d.id_exam
                     join alert.exam_cat ec on ec.id_exam_cat=e.id_exam_cat and ec.flg_available='Y'
                     join translation t on t.code_translation=e.code_exam and t.desc_lang_" & l_id_language & " is not null                    
                     where d.id_institution = " & i_id_inst & "
                     and e.flg_type='" & i_exam_type & "'
                     and e.flg_available='Y' and ec.flg_available='Y'
                     and d.id_software= " & i_id_soft & " "


        If i_flg_type = 0 Then

            sql = sql & "And d.flg_type IN ('P','M','A','B') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And d.flg_type IN ('P','M') "

        Else

            sql = sql & "And d.flg_type IN ('A','B') "

        End If


        If i_id_content_exam_cat = 0 Then

            sql = sql & " order by 1 asc"

        Else

            sql = sql & " and ec.id_content = " & i_id_content_exam_cat & " 
                          order by 1 asc"

        End If

        Dim cmd As New OracleCommand(sql, i_conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception
            cmd.Dispose()
            Return False
        End Try

    End Function

    Function DELETE_EXAMS(ByVal i_institution As Int64, ByVal i_software As Int64, ByVal i_exam As exams_default, ByVal i_most_freq As Boolean, ByVal i_flg_type As Integer, ByVal i_conn As OracleConnection) As Boolean


        Dim sql As String = "   DELETE from alert.exam_dep_clin_serv dps
                                        where dps.id_exam = (select e.id_exam from alert.exam e where e.id_content='" & i_exam.id_content_exam & "' and e.flg_available='Y')
                                        and (
                                        (dps.id_institution= " & i_institution & " and dps.id_software= " & i_software & " ) 
                                        or 
                                        ((dps.id_institution is null or dps.id_institution=" & i_institution & ") and dps.id_software= " & i_software & " )
                                        ) "

        If i_most_freq = True Then

            If i_flg_type = 0 Then

                sql = sql & "And dps.flg_type IN ('M', 'A') "

            ElseIf i_flg_type = 1 Then

                sql = sql & "And dps.flg_type IN ('M') "

            Else

                sql = sql & "And dps.flg_type IN ('A') "

            End If


        Else

            If i_flg_type = 0 Then

                sql = sql & "And dps.flg_type IN ('P','M','B','A') "

            ElseIf i_flg_type = 1 Then

                sql = sql & "And dps.flg_type IN ('P','M') "

            Else

                sql = sql & "And dps.flg_type IN ('B','A') "

            End If

        End If

        Dim cmd_delete_dep_clin_serv As New OracleCommand(sql, i_conn)

        Try
            cmd_delete_dep_clin_serv.CommandType = CommandType.Text
            cmd_delete_dep_clin_serv.ExecuteNonQuery()
        Catch ex As Exception
            cmd_delete_dep_clin_serv.Dispose()
            Return False
        End Try

        cmd_delete_dep_clin_serv.Dispose()
        Return True

    End Function

    Function GET_EXAMS_CAT_DEFAULT(ByVal i_version As String, ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_exam_type As String, ByVal i_flg_type As Integer, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Dim sql As String = "Select distinct ec.id_content, 
                                    tc.desc_lang_" & l_id_language & "         
                              From alert_default.exam e
                              Join alert_default.exam_mrk_vrs v
                                On v.id_exam = e.id_exam
                                Join alert_default.exam_cat ec on ec.id_exam_cat=e.id_exam_cat
                                Join alert_default.translation tc on tc.code_translation=ec.code_exam_cat
                                Join alert_default.exam_clin_serv ecs on ecs.id_exam=e.id_exam 
                                join institution i on i.id_market=v.id_market
                             where i.id_institution= " & i_institution & "
                               And v.version = '" & i_version & "'
                               And e.flg_type='" & i_exam_type & "'
                               And e.flg_available='Y'
                               And ecs.id_software in  (0," & i_software & ") "

        If i_flg_type = 0 Then

            sql = sql & "And ecs.flg_type IN ('P','M','A','B') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And ecs.flg_type IN ('P','M') "

        Else

            sql = sql & "And ecs.flg_type IN ('A','B') "

        End If


        sql = sql & "order by 2 asc"

        Dim cmd As New OracleCommand(sql, i_conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception
            cmd.Dispose()
            Return False
        End Try

    End Function

    Function GET_DEFAULT_VERSIONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_exam_type As String, ByVal i_flg_type As Integer, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "Select distinct v.version
                              From alert_default.exam e
                              join alert_default.exam_mrk_vrs v
                                on v.id_exam = e.id_exam
                              join alert_default.exam_clin_serv ecs
                                on ecs.id_exam = e.id_exam and ecs.id_software in (0, " & i_software & ")
                               and ecs.flg_type = 'P'
                              join institution i
                                on i.id_market = v.id_market
                             where i.id_institution = " & i_institution & "
                                 and e.flg_type = '" & i_exam_type & "' "

        If i_flg_type = 0 Then

            sql = sql & "And ecs.flg_type IN ('P', 'M', 'B','A') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And ecs.flg_type IN ('P', 'M') "

        Else

            sql = sql & "And ecs.flg_type IN ('B','A') "

        End If

        sql = sql & "  ORDER BY 1 ASC"

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text
        Try
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception
            cmd.Dispose()
            Return False
        End Try

    End Function


    Function GET_EXAMS_DEFAULT_BY_CAT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_version As String, ByVal i_id_cat As String, ByVal i_exam_type As String, ByVal i_flg_type As Integer, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Dim sql As String = "Select  distinct ec.id_content,        
                                                  tc.desc_lang_" & l_id_language & ", 
                                                  e.id_content,       
                                                  te.desc_lang_" & l_id_language & ",              
                                                  ecs.flg_first_result, 
                                                  ecs.flg_execute, 
                                                  ecs.flg_timeout, 
                                                  ecs.flg_result_notes, 
                                                  ecs.flg_first_execute,       
                                                  e.age_min, 
                                                  e.age_max, 
                                                  e.gender
                              from alert_default.exam e
                              join alert_default.exam_mrk_vrs v
                                on v.id_exam = e.id_exam
                                join alert_default.exam_cat ec on ec.id_exam_cat=e.id_exam_cat
                                join alert_default.translation te on te.code_translation=e.code_exam
                                join alert_default.translation tc on tc.code_translation=ec.code_exam_cat
                                join alert_default.exam_clin_serv ecs on ecs.id_exam=e.id_exam
                                join institution i on i.id_market=v.id_market
                             where i.id_institution=  " & i_institution & "
                               and v.version = '" & i_version & "'
                               and e.flg_type='" & i_exam_type & "'
                               and e.flg_available='Y'
                               and ecs.id_software= " & i_software & " "

        If i_flg_type = 0 Then

            sql = sql & "And ecs.flg_type IN ('P','M','A','B') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And ecs.flg_type IN ('P','M') "

        Else

            sql = sql & "And ecs.flg_type IN ('A','B') "

        End If

        If i_id_cat = "0" Then

            sql = sql & " order by 4 asc"

        Else

            sql = sql & " And ec.id_content = '" & i_id_cat & "'
                         order by 2 asc, 4 asc"
        End If

        Dim cmd As New OracleCommand(sql, i_conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception
            cmd.Dispose()
            Return False
        End Try

    End Function

    Function CHECK_CATEGORY_EXISTANCE(ByVal i_id_content_cat, ByVal i_id_institution, ByVal i_oradb) As Boolean

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

        Dim sql As String = "Select count(*)
  from alert.exam_cat ec
 where ec.id_content = '" & i_id_content_cat & "'
   and ec.flg_available = 'Y'"

        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Try

            Dim dr As OracleDataReader = cmd.ExecuteReader()

            Dim l_cat_exist As Integer = 0

            While dr.Read()

                l_cat_exist = dr.Item(0)

            End While

            If l_cat_exist > 0 Then


                conn.Close()

                conn.Dispose()

                Return True

            Else


                conn.Close()

                conn.Dispose()

                Return False

            End If

        Catch ex As Exception


            conn.Close()

            conn.Dispose()

            Return False

        End Try

    End Function

    Function CHECK_EXAM_TRANSLATION_EXISTANCE(ByVal i_id_content_exam, ByVal i_id_institution, ByVal i_oradb) As Boolean

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

        Dim sql As String = "Select count(*)
  from alert.exam e
  left join translation t
    on t.code_translation = e.code_exam
  join institution i
    on i.id_institution = " & i_id_institution & "
 where e.id_content = '" & i_id_content_exam & "'
   and e.flg_available = 'Y'
   and (decode(i.id_market,
              1,
              t.desc_lang_1,
              2,
              t.desc_lang_2,
              3,
              t.desc_lang_11,
              4,
              t.desc_lang_5,
              5,
              t.desc_lang_4,
              6,
              t.desc_lang_3,
              7,
              t.desc_lang_10,
              8,
              t.desc_lang_7,
              9,
              t.desc_lang_6,
              10,
              t.desc_lang_9,
              12,
              t.desc_lang_16,
              16,
              t.desc_lang_17,
              17,
              t.desc_lang_18,
              19,
              t.desc_lang_19)) is not null"

        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Try

            Dim dr As OracleDataReader = cmd.ExecuteReader()

            Dim l_exam_exist As Integer = 0

            While dr.Read()

                l_exam_exist = dr.Item(0)

            End While

            If l_exam_exist > 0 Then


                conn.Close()

                conn.Dispose()

                Return True

            Else


                conn.Close()

                conn.Dispose()

                Return False

            End If

        Catch ex As Exception

            conn.Close()

            conn.Dispose()

            Return False

        End Try

    End Function

    Function GET_LANGUAGE_ID(ByVal i_institution As Int64, ByVal i_oradb As String) As Integer

        Dim l_id_lang As Integer = 0

        Dim Sql = "Select decode(i.id_market,
              1,
              1,
              2,
              2,
              3,
              11,
              4,
              5,
              5,
              4,
              6,
              3,
              7,
              10,
              8,
              7,
              9,
              6,
              10,
              9,
              12,
              16,
              16,
              17,
              17,
              18,
              19,
              19)
  from institution i
    where i.id_institution = " & i_institution

        Dim conn As New OracleConnection(i_oradb)

        conn.Open()

        Dim cmd_get_id_lang As New OracleCommand(Sql, conn)
        cmd_get_id_lang.CommandType = CommandType.Text
        Dim dr As OracleDataReader = cmd_get_id_lang.ExecuteReader()

        While dr.Read()

            l_id_lang = dr.Item(0)

        End While

        conn.Close()

        conn.Dispose()

        Return l_id_lang

    End Function

    Function SET_TRANSLATION(ByVal i_id_lang As Integer, ByVal i_code_translation As String, ByVal i_desc As String, i_oradb As String) As Boolean

        Dim conn As New OracleConnection(i_oradb)

        conn.Open()

        Try

            Dim Sql = "begin pk_translation.insert_into_translation( " & i_id_lang & " , '" & i_code_translation & "' , '" & i_desc & "' ); end;"

            Dim cmd_insert_trans As New OracleCommand(Sql, conn)
            cmd_insert_trans.CommandType = CommandType.Text

            cmd_insert_trans.ExecuteNonQuery()


            conn.Close()

            conn.Dispose()

            Return True

        Catch ex As Exception

            conn.Close()

            conn.Dispose()

            Return False

        End Try


    End Function

    Function CHECK_EXAM_EXISTANCE(ByVal i_id_institution As Int64, ByVal i_id_content_exam As String, ByVal i_conn As OracleConnection) As Boolean

        Dim sql As String = "Select count(*)
                             from alert.exam e
                             where e.id_content = '" & i_id_content_exam & "'
                             and e.flg_available = 'Y'"

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Try

            Dim dr As OracleDataReader = cmd.ExecuteReader()

            Dim l_exam_exists As Integer = 0

            While dr.Read()

                l_exam_exists = dr.Item(0)

            End While

            If l_exam_exists > 0 Then

                Return True

            Else

                Return False

            End If

        Catch ex As Exception

            Return False

        End Try

    End Function

    Function UPDATE_EXAM_CAT(ByVal i_id_content_exam As String, ByVal i_id_content_cat As String, ByVal i_conn As OracleConnection) As Boolean

        Dim Sql = "declare

                        l_id_exam_cat alert.exam_cat.id_exam_cat%type;

                        begin
  
                        select ec.id_exam_cat
                        into l_id_exam_cat
                        from alert.exam_cat ec
                        where ec.id_content='" & i_id_content_cat & "'
                        and ec.flg_available='Y';

                        update alert.exam e
                        set e.id_exam_cat=l_id_exam_cat
                        where e.id_content='" & i_id_content_exam & "'
                        and e.flg_available='Y';

                        end;"

        Dim cmd_insert_trans As New OracleCommand(Sql, i_conn)

        Try

            cmd_insert_trans.CommandType = CommandType.Text
            cmd_insert_trans.ExecuteNonQuery()
            cmd_insert_trans.Dispose()

            Return True

        Catch ex As Exception

            cmd_insert_trans.Dispose()
            Return False

        End Try

    End Function

    Function SET_EXAM_DEP_CLIN_SERV(ByVal i_id_exam As String, ByVal i_id_dep_clin_serv As String, ByVal i_flg_type As String, ByVal i_id_institution As Int64,
                                   ByVal i_id_soft As Int64, ByVal i_flg_first_result As String, ByVal flg_execute As String, ByVal flg_timeout As String,
                                   ByVal flg_result_notes As String, ByVal flg_first_execute As String, ByVal i_oradb As String) As Boolean

        Dim conn As New OracleConnection(i_oradb)

        conn.Open()

        Try

            Dim Sql As String = ""

            If i_id_dep_clin_serv < 0 Then

                Sql = "declare

                        l_id_exam alert.exam.id_exam%type;

                        begin
  
                        select e.id_exam
                        into l_id_exam
                        from alert.exam e
                        where e.id_content='" & i_id_exam & "'
                        and e.flg_available='Y';
  
                        insert into alert.exam_dep_clin_serv (ID_EXAM_DEP_CLIN_SERV, ID_EXAM, ID_DEP_CLIN_SERV, FLG_TYPE, RANK,   ID_INSTITUTION,  ID_SOFTWARE, FLG_FIRST_RESULT, FLG_MOV_PAT, FLG_EXECUTE, FLG_TIMEOUT, FLG_RESULT_NOTES, FLG_FIRST_EXECUTE)
                        values (alert.seq_exam_dep_clin_serv.nextval, l_id_exam, null , '" & i_flg_type & "', 0,  " & i_id_institution & ",  " & i_id_soft & ", '" & i_flg_first_result & "', 'N', '" & flg_execute & "', '" & flg_timeout & "', '" & flg_result_notes & "', '" & flg_first_execute & "');

                        exception
                            when dup_val_on_index then
                            l_id_exam:=-1;
                        end;"
            Else

                Sql = "declare

                        l_id_exam alert.exam.id_exam%type;

                        begin
  
                        select e.id_exam
                        into l_id_exam
                        from alert.exam e
                        where e.id_content='" & i_id_exam & "'
                        and e.flg_available='Y';
  
                        insert into alert.exam_dep_clin_serv (ID_EXAM_DEP_CLIN_SERV, ID_EXAM, ID_DEP_CLIN_SERV, FLG_TYPE, RANK,   ID_INSTITUTION,  ID_SOFTWARE, FLG_FIRST_RESULT, FLG_MOV_PAT, FLG_EXECUTE, FLG_TIMEOUT, FLG_RESULT_NOTES, FLG_FIRST_EXECUTE)
                        values (alert.seq_exam_dep_clin_serv.nextval, l_id_exam, " & i_id_dep_clin_serv & " , '" & i_flg_type & "', 0,  " & i_id_institution & ",  " & i_id_soft & ", '" & i_flg_first_result & "', 'N', '" & flg_execute & "', '" & flg_timeout & "', '" & flg_result_notes & "', '" & flg_first_execute & "');

                        exception
                            when dup_val_on_index then
                            l_id_exam:=-1;
                        end;"

            End If

            Dim cmd_insert_trans As New OracleCommand(Sql, conn)
            cmd_insert_trans.CommandType = CommandType.Text

            cmd_insert_trans.ExecuteNonQuery()

            conn.Close()

            conn.Dispose()

            Return True

        Catch ex As Exception

            conn.Close()

            conn.Dispose()

            Return False

        End Try

    End Function

    Function SET_DEFAULT_EXAM_DEP_CLIN_SERV(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_a_exams() As exams_default, ByVal i_exam_type As String, ByVal i_flg_type As Integer, ByVal i_conn As OracleConnection) As Boolean

        Dim sql As String = "DECLARE

                                l_a_exams table_varchar := table_varchar("

        For i As Integer = 0 To i_a_exams.Count() - 1

            If (i < i_a_exams.Count() - 1) Then

                sql = sql & "'" & i_a_exams(i).id_content_exam & "', "

            Else

                sql = sql & "'" & i_a_exams(i).id_content_exam & "');"

            End If

        Next

        sql = sql & "                          l_id_exam alert.exam.id_exam%TYPE;                      
                        l_a_dep_clin_serv  table_number := table_number();
                        
                        l_a_flg_first_result table_varchar := table_varchar();
                        l_flg_first_result   alert.exam_dep_clin_serv.flg_first_result%type;

                        l_a_flg_mov_pat table_varchar := table_varchar();
                        l_flg_mov_pat   alert.exam_dep_clin_serv.flg_mov_pat%type;                           
                                                
                        l_a_flg_execute      table_varchar := table_varchar();
                        l_flg_execute        alert.exam_dep_clin_serv.flg_execute%type;
                        
                        l_a_flg_timeout      table_varchar := table_varchar();
                        l_flg_timeout        alert.exam_dep_clin_serv.flg_timeout%type;
                        
                        l_a_flg_result_notes table_varchar := table_varchar();
                        l_flg_result_notes   alert.exam_dep_clin_serv.flg_result_notes%type;
                        
                        l_a_flg_first_execute table_varchar := table_varchar();
                        l_flg_first_execute   alert.exam_dep_clin_serv.flg_first_execute%type;

                        l_flg_chargeable alert.interv_dep_clin_serv.flg_chargeable%TYPE;
                        l_a_flg_chargeable table_varchar := table_varchar();
                        
                        l_a_flg_type table_varchar := table_varchar();
                        l_flg_type_indicator  integer := " & i_flg_type & " ;

                        BEGIN
  
                            IF l_flg_type_indicator = 0 THEN

                                FOR i IN 1 .. l_a_exams.count()
                                LOOP
                                    BEGIN
            
                                        SELECT e.id_exam
                                        INTO l_id_exam
                                        FROM alert.exam e
                                        WHERE e.id_content = l_a_exams(i)
                                        AND e.flg_available = 'Y';
            
                                        SELECT DISTINCT ecs.flg_type BULK COLLECT
                                        INTO l_a_flg_type                                        
                                        FROM alert_default.exam_clin_serv ecs
                                        JOIN alert_default.exam de ON de.id_exam = ecs.id_exam
                                        WHERE de.id_content = l_a_exams(i);
            
                                        FOR j IN 1 .. l_a_flg_type.count()
                                        LOOP
                
                                            IF (l_a_flg_type(j) <> 'A' AND l_a_flg_type(j) <> 'M')
                                            THEN
                    
                                                BEGIN
                                                    SELECT ecs.flg_first_result, ecs.flg_mov_pat,ecs.flg_execute,
                                                    ecs.flg_timeout, ecs.flg_result_notes, ecs.flg_first_execute, ecs.flg_chargeable
                                                    INTO l_flg_first_result, l_flg_mov_pat,l_flg_execute,
                                                    l_flg_timeout,l_flg_result_notes,l_flg_first_execute,l_flg_chargeable
                                                    FROM alert_default.exam_clin_serv ecs
                                                    JOIN alert_default.exam de ON de.id_exam = ecs.id_exam
                                                    WHERE de.id_content = l_a_exams(i)
                                                    AND de.flg_available = 'Y'
                                                    AND ecs.flg_type = l_a_flg_type(j)
                                                    AND ecs.id_software IN (" & i_software & ");
                        
                                                    INSERT INTO alert.exam_dep_clin_serv
                                                        (ID_EXAM_DEP_CLIN_SERV,
                                                         ID_EXAM,
                                                         ID_DEP_CLIN_SERV,
                                                         FLG_TYPE,
                                                         rank,
                                                         id_institution,
                                                         id_software,
                                                         FLG_FIRST_RESULT,
                                                         FLG_MOV_PAT,
                                                         FLG_EXECUTE,
                                                         FLG_TIMEOUT,
                                                         FLG_RESULT_NOTES,
                                                         FLG_FIRST_EXECUTE,
                                                         FLG_CHARGEABLE)
                                                    VALUES
                                                        (alert.seq_exam_dep_clin_serv.nextval,
                                                         l_id_exam,
                                                         NULL,
                                                         l_a_flg_type(j),
                                                         0,
                                                         " & i_institution & ",
                                                         " & i_software & ",
                                                         l_flg_first_result,
                                                         l_flg_mov_pat,
                                                         l_flg_execute,
                                                         l_flg_timeout,
                                                         l_flg_result_notes,
                                                         l_flg_first_execute,
                                                         l_flg_chargeable                                                                                                                 
                                                         );
                                                EXCEPTION
                                                    WHEN OTHERS THEN
                                                        continue;
                                                END;
                    
                                            ELSE
                    
                                                SELECT decs.flg_first_result, decs.flg_mov_pat,decs.flg_execute,
                                                       decs.flg_timeout, decs.flg_result_notes, decs.flg_first_execute, decs.flg_chargeable,
                                                       dps.id_dep_clin_serv BULK COLLECT
                                                INTO l_a_flg_first_result, l_a_flg_mov_pat, l_a_flg_execute,
                                                l_a_flg_timeout,l_a_flg_result_notes,l_a_flg_first_execute,l_a_flg_chargeable,
                                                l_a_dep_clin_serv
                                                FROM alert_default.exam_clin_serv decs
                                                JOIN alert_default.exam de ON de.id_exam = decs.id_exam
                                                JOIN alert_default.clinical_service dc ON dc.id_clinical_service = decs.id_clinical_service
                                                JOIN alert.clinical_service cs ON cs.id_content = dc.id_content
                                                                           AND cs.flg_available = 'Y'
                                                JOIN alert.dep_clin_serv dps ON dps.id_clinical_service = cs.id_clinical_service
                                                                         AND dps.flg_available = 'Y'
                                                JOIN department d ON d.id_department = dps.id_department
                                                WHERE de.id_content = l_a_exams(i)
                                                AND de.flg_available='Y'
                                                AND decs.flg_type IN (l_a_flg_type(j))
                                                AND decs.id_software IN (" & i_software & ")
                                                AND d.id_institution = " & i_institution & "
                                                AND d.id_software = " & i_software & ";
                    
                                                FOR k IN 1 .. l_a_dep_clin_serv.count()
                                                LOOP
                        
                                                    BEGIN
                                                        INSERT INTO alert.exam_dep_clin_serv
                                                                    (ID_EXAM_DEP_CLIN_SERV,
                                                                     ID_EXAM,
                                                                     ID_DEP_CLIN_SERV,
                                                                     FLG_TYPE,
                                                                     rank,
                                                                     id_institution,
                                                                     id_software,
                                                                     FLG_FIRST_RESULT,
                                                                     FLG_MOV_PAT,
                                                                     FLG_EXECUTE,
                                                                     FLG_TIMEOUT,
                                                                     FLG_RESULT_NOTES,
                                                                     FLG_FIRST_EXECUTE,
                                                                     FLG_CHARGEABLE)
                                                        VALUES
                                                                  (alert.seq_exam_dep_clin_serv.nextval,
                                                                   l_id_exam,
                                                                   l_a_dep_clin_serv(k),
                                                                   l_a_flg_type(j),
                                                                   0,
                                                                   " & i_institution & ",
                                                                   " & i_software & ",
                                                                   l_a_flg_first_result(k),
                                                                   l_a_flg_mov_pat(k),
                                                                   l_a_flg_execute(k),
                                                                   l_a_flg_timeout(k),
                                                                   l_a_flg_result_notes(k),
                                                                   l_a_flg_first_execute(k),
                                                                   l_a_flg_chargeable(k)                                                                                                                 
                                                                   );
                                                    EXCEPTION
                                                        WHEN OTHERS THEN
                                                            continue;
                                                    END;
                        
                                                END LOOP;
                    
                                            END IF;
                
                                        END LOOP;
            
                                    EXCEPTION
                                        WHEN OTHERS THEN
                                            continue;
                
                                    END;
        
                                END LOOP;

                             ELSIF   l_flg_type_indicator = 1 THEN
       
                                   FOR i IN 1 .. l_a_exams.count()
                                      LOOP
                                          BEGIN
                  
                                              SELECT e.id_exam
                                              INTO l_id_exam
                                              FROM alert.exam e
                                              WHERE e.id_content = l_a_exams(i)
                                              AND e.flg_available = 'Y';
                  
                                              SELECT DISTINCT decs.flg_type BULK COLLECT
                                              INTO l_a_flg_type
                                              FROM alert_default.exam_clin_serv decs
                                              JOIN alert_default.exam de ON de.id_exam = decs.id_exam
                                              WHERE de.id_content = l_a_exams(i);
                  
                                              FOR j IN 1 .. l_a_flg_type.count()
                                              LOOP
                      
                                                  IF (l_a_flg_type(j) = 'P')
                                                  THEN
                          
                                                      BEGIN
                                                    
                                                    SELECT ecs.flg_first_result, ecs.flg_mov_pat,ecs.flg_execute,
                                                    ecs.flg_timeout, ecs.flg_result_notes, ecs.flg_first_execute, ecs.flg_chargeable
                                                    INTO l_flg_first_result, l_flg_mov_pat,l_flg_execute,
                                                    l_flg_timeout,l_flg_result_notes,l_flg_first_execute,l_flg_chargeable
                                                    FROM alert_default.exam_clin_serv ecs
                                                    JOIN alert_default.exam de ON de.id_exam = ecs.id_exam
                                                    WHERE de.id_content = l_a_exams(i)
                                                    AND de.flg_available = 'Y'
                                                    AND ecs.flg_type = l_a_flg_type(j)
                                                    AND ecs.id_software IN (" & i_software & ");
                        
                                                    INSERT INTO alert.exam_dep_clin_serv
                                                        (ID_EXAM_DEP_CLIN_SERV,
                                                         ID_EXAM,
                                                         ID_DEP_CLIN_SERV,
                                                         FLG_TYPE,
                                                         rank,
                                                         id_institution,
                                                         id_software,
                                                         FLG_FIRST_RESULT,
                                                         FLG_MOV_PAT,
                                                         FLG_EXECUTE,
                                                         FLG_TIMEOUT,
                                                         FLG_RESULT_NOTES,
                                                         FLG_FIRST_EXECUTE,
                                                         FLG_CHARGEABLE)
                                                    VALUES
                                                        (alert.seq_exam_dep_clin_serv.nextval,
                                                         l_id_exam,
                                                         NULL,
                                                         l_a_flg_type(j),
                                                         0,
                                                         " & i_institution & ",
                                                         " & i_software & ",
                                                         l_flg_first_result,
                                                         l_flg_mov_pat,
                                                         l_flg_execute,
                                                         l_flg_timeout,
                                                         l_flg_result_notes,
                                                         l_flg_first_execute,
                                                         l_flg_chargeable                                                                                                                 
                                                         );
                                                EXCEPTION
                                                    WHEN OTHERS THEN
                                                        continue;
                                                END;
                          
                                                  ELSIF (l_a_flg_type(j) = 'M') THEN
                          
                                                      SELECT decs.flg_first_result, decs.flg_mov_pat,decs.flg_execute,
                                                       decs.flg_timeout, decs.flg_result_notes, decs.flg_first_execute, decs.flg_chargeable,
                                                       dps.id_dep_clin_serv BULK COLLECT
                                                INTO l_a_flg_first_result, l_a_flg_mov_pat, l_a_flg_execute,
                                                l_a_flg_timeout,l_a_flg_result_notes,l_a_flg_first_execute,l_a_flg_chargeable,
                                                l_a_dep_clin_serv
                                                FROM alert_default.exam_clin_serv decs
                                                JOIN alert_default.exam de ON de.id_exam = decs.id_exam
                                                JOIN alert_default.clinical_service dc ON dc.id_clinical_service = decs.id_clinical_service
                                                JOIN alert.clinical_service cs ON cs.id_content = dc.id_content
                                                                           AND cs.flg_available = 'Y'
                                                JOIN alert.dep_clin_serv dps ON dps.id_clinical_service = cs.id_clinical_service
                                                                         AND dps.flg_available = 'Y'
                                                JOIN department d ON d.id_department = dps.id_department
                                                WHERE de.id_content = l_a_exams(i)
                                                AND de.flg_available='Y'
                                                AND decs.flg_type IN (l_a_flg_type(j))
                                                AND decs.id_software IN (" & i_software & ")
                                                AND d.id_institution = " & i_institution & "
                                                AND d.id_software = " & i_software & ";
                    
                                                FOR k IN 1 .. l_a_dep_clin_serv.count()
                                                LOOP
                        
                                                    BEGIN
                                                        INSERT INTO alert.exam_dep_clin_serv
                                                                    (ID_EXAM_DEP_CLIN_SERV,
                                                                     ID_EXAM,
                                                                     ID_DEP_CLIN_SERV,
                                                                     FLG_TYPE,
                                                                     rank,
                                                                     id_institution,
                                                                     id_software,
                                                                     FLG_FIRST_RESULT,
                                                                     FLG_MOV_PAT,
                                                                     FLG_EXECUTE,
                                                                     FLG_TIMEOUT,
                                                                     FLG_RESULT_NOTES,
                                                                     FLG_FIRST_EXECUTE,
                                                                     FLG_CHARGEABLE)
                                                        VALUES
                                                                  (alert.seq_exam_dep_clin_serv.nextval,
                                                                   l_id_exam,
                                                                   l_a_dep_clin_serv(k),
                                                                   l_a_flg_type(j),
                                                                   0,
                                                                   " & i_institution & ",
                                                                   " & i_software & ",
                                                                   l_a_flg_first_result(k),
                                                                   l_a_flg_mov_pat(k),
                                                                   l_a_flg_execute(k),
                                                                   l_a_flg_timeout(k),
                                                                   l_a_flg_result_notes(k),
                                                                   l_a_flg_first_execute(k),
                                                                   l_a_flg_chargeable(k)                                                                                                                 
                                                                   );
                                                    EXCEPTION
                                                        WHEN OTHERS THEN
                                                            continue;
                                                    END;
                        
                                                END LOOP;
                                                
                                                  END IF;
                      
                                              END LOOP;
                  
                                          EXCEPTION
                                              WHEN OTHERS THEN
                                                  continue;
                      
                                          END;
              
                                      END LOOP;
              
                            ELSE
      
                                FOR i IN 1 .. l_a_exams.count()
                                          LOOP
                                              BEGIN
                      
                                                  SELECT e.id_exam
                                                  INTO   l_id_exam
                                                  FROM   alert.exam e
                                                  WHERE  e.id_content = l_a_exams(i)
                                                  AND    e.flg_available='Y';
                      
                                                  SELECT DISTINCT dcs.flg_type BULK COLLECT
                                                  INTO l_a_flg_type
                                                  FROM alert_default.exam_clin_serv dcs
                                                  JOIN alert_default.exam de ON de.id_exam=dcs.id_exam
                                                  WHERE de.id_content = l_a_exams(i);
                      
                                                  FOR j IN 1 .. l_a_flg_type.count()
                                                  LOOP
                          
                                                      IF (l_a_flg_type(j) = 'B')
                                                      THEN
                              
                                                          BEGIN
                                                            
                                                   SELECT ecs.flg_first_result, ecs.flg_mov_pat,ecs.flg_execute,
                                                    ecs.flg_timeout, ecs.flg_result_notes, ecs.flg_first_execute, ecs.flg_chargeable
                                                    INTO l_flg_first_result, l_flg_mov_pat,l_flg_execute,
                                                    l_flg_timeout,l_flg_result_notes,l_flg_first_execute,l_flg_chargeable
                                                    FROM alert_default.exam_clin_serv ecs
                                                    JOIN alert_default.exam de ON de.id_exam = ecs.id_exam
                                                    WHERE de.id_content = l_a_exams(i)
                                                    AND de.flg_available = 'Y'
                                                    AND ecs.flg_type = l_a_flg_type(j)
                                                    AND ecs.id_software IN (" & i_software & ");
                        
                                                    INSERT INTO alert.exam_dep_clin_serv
                                                        (ID_EXAM_DEP_CLIN_SERV,
                                                         ID_EXAM,
                                                         ID_DEP_CLIN_SERV,
                                                         FLG_TYPE,
                                                         rank,
                                                         id_institution,
                                                         id_software,
                                                         FLG_FIRST_RESULT,
                                                         FLG_MOV_PAT,
                                                         FLG_EXECUTE,
                                                         FLG_TIMEOUT,
                                                         FLG_RESULT_NOTES,
                                                         FLG_FIRST_EXECUTE,
                                                         FLG_CHARGEABLE)
                                                    VALUES
                                                        (alert.seq_exam_dep_clin_serv.nextval,
                                                         l_id_exam,
                                                         NULL,
                                                         l_a_flg_type(j),
                                                         0,
                                                         " & i_institution & ",
                                                         " & i_software & ",
                                                         l_flg_first_result,
                                                         l_flg_mov_pat,
                                                         l_flg_execute,
                                                         l_flg_timeout,
                                                         l_flg_result_notes,
                                                         l_flg_first_execute,
                                                         l_flg_chargeable                                                                                                                 
                                                         );
                                                EXCEPTION
                                                    WHEN OTHERS THEN
                                                        continue;
                                                        
                                                          END;
                              
                                                      ELSIF (l_a_flg_type(j) = 'A') THEN

                                                     SELECT decs.flg_first_result, decs.flg_mov_pat,decs.flg_execute,
                                                       decs.flg_timeout, decs.flg_result_notes, decs.flg_first_execute, decs.flg_chargeable,
                                                       dps.id_dep_clin_serv BULK COLLECT
                                                INTO l_a_flg_first_result, l_a_flg_mov_pat, l_a_flg_execute,
                                                l_a_flg_timeout,l_a_flg_result_notes,l_a_flg_first_execute,l_a_flg_chargeable,
                                                l_a_dep_clin_serv
                                                FROM alert_default.exam_clin_serv decs
                                                JOIN alert_default.exam de ON de.id_exam = decs.id_exam
                                                JOIN alert_default.clinical_service dc ON dc.id_clinical_service = decs.id_clinical_service
                                                JOIN alert.clinical_service cs ON cs.id_content = dc.id_content
                                                                           AND cs.flg_available = 'Y'
                                                JOIN alert.dep_clin_serv dps ON dps.id_clinical_service = cs.id_clinical_service
                                                                         AND dps.flg_available = 'Y'
                                                JOIN department d ON d.id_department = dps.id_department
                                                WHERE de.id_content = l_a_exams(i)
                                                AND de.flg_available='Y'
                                                AND decs.flg_type IN (l_a_flg_type(j))
                                                AND decs.id_software IN (" & i_software & ")
                                                AND d.id_institution = " & i_institution & "
                                                AND d.id_software = " & i_software & ";
                                                          
                                               FOR k IN 1 .. l_a_dep_clin_serv.count()
                                                          LOOP
                                  
                                                              BEGIN
                              
                                                       INSERT INTO alert.exam_dep_clin_serv
                                                                    (ID_EXAM_DEP_CLIN_SERV,
                                                                     ID_EXAM,
                                                                     ID_DEP_CLIN_SERV,
                                                                     FLG_TYPE,
                                                                     rank,
                                                                     id_institution,
                                                                     id_software,
                                                                     FLG_FIRST_RESULT,
                                                                     FLG_MOV_PAT,
                                                                     FLG_EXECUTE,
                                                                     FLG_TIMEOUT,
                                                                     FLG_RESULT_NOTES,
                                                                     FLG_FIRST_EXECUTE,
                                                                     FLG_CHARGEABLE)
                                                        VALUES
                                                                  (alert.seq_exam_dep_clin_serv.nextval,
                                                                   l_id_exam,
                                                                   l_a_dep_clin_serv(k),
                                                                   l_a_flg_type(j),
                                                                   0,
                                                                   " & i_institution & ",
                                                                   " & i_software & ",
                                                                   l_a_flg_first_result(k),
                                                                   l_a_flg_mov_pat(k),
                                                                   l_a_flg_execute(k),
                                                                   l_a_flg_timeout(k),
                                                                   l_a_flg_result_notes(k),
                                                                   l_a_flg_first_execute(k),
                                                                   l_a_flg_chargeable(k)                                                                                                                 
                                                                   );
                                                    EXCEPTION
                                                        WHEN OTHERS THEN
                                                            continue;
                                                              END;
                                  
                                                          END LOOP;
                              
                                                      END IF;
                          
                                                  END LOOP;
                      
                                              EXCEPTION
                                                  WHEN OTHERS THEN
                                                      continue;
                          
                                              END;
                  
                                          END LOOP;
                     
                            END IF;

                        END;"


    End Function


    Function SET_EXAM_ALERT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_set_exams() As exams_default, ByVal i_exam_type As String, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Try

            For i As Integer = 0 To i_set_exams.Count() - 1

                'Verificar se o exame existe na tabela alert.exam
                If Not CHECK_EXAM_EXISTANCE(i_institution, i_set_exams(i).id_content_exam, i_conn) Then

                    ''1 - Inserir o EXAME
                    Dim sql As String = ""

                    If i_set_exams(i).age_min < 0 And i_set_exams(i).age_max < 0 Then

                        sql = "Declare
                                         l_id_content_exam_cat alert.exam_cat.id_content%type;
                                         begin
                                         select ec.id_exam_cat
                                         into l_id_content_exam_cat
                                         from alert.exam_cat ec
                                         where ec.id_content= '" & i_set_exams(i).id_content_category & "'
                                         and ec.flg_available='Y';
                                         insert into alert.exam (ID_EXAM, CODE_EXAM, FLG_PAT_RESP, FLG_PAT_PREP, FLG_MOV_PAT, FLG_AVAILABLE, RANK, FLG_TYPE, GENDER, AGE_MIN, AGE_MAX, ID_EXAM_CAT, ID_CONTENT)
                                         values (alert.seq_exam.nextval, 'EXAM.CODE_EXAM.'||alert.seq_exam.nextval, 'N', 'N', 'Y', 'Y', 0, '" & i_exam_type & "', '" & i_set_exams(i).gender & "' , null, null , l_id_content_exam_cat,'" & i_set_exams(i).id_content_exam & "');
                                         end;"

                    ElseIf i_set_exams(i).age_min < 0 Then

                        sql = "declare
                                         l_id_content_exam_cat alert.exam_cat.id_content%type;
                                         begin
                                         select ec.id_exam_cat
                                         into l_id_content_exam_cat
                                         from alert.exam_cat ec
                                         where ec.id_content= '" & i_set_exams(i).id_content_category & "'
                                         and ec.flg_available='Y';
                                         insert into alert.exam (ID_EXAM, CODE_EXAM, FLG_PAT_RESP, FLG_PAT_PREP, FLG_MOV_PAT, FLG_AVAILABLE, RANK, FLG_TYPE, GENDER, AGE_MIN, AGE_MAX, ID_EXAM_CAT, ID_CONTENT)
                                         values (alert.seq_exam.nextval, 'EXAM.CODE_EXAM.'||alert.seq_exam.nextval, 'N', 'N', 'Y', 'Y', 0, '" & i_exam_type & "', '" & i_set_exams(i).gender & "' , null, " & i_set_exams(i).age_max & ", l_id_content_exam_cat,'" & i_set_exams(i).id_content_exam & "');
                                         end;"

                    ElseIf i_set_exams(i).age_max < 0 Then

                        sql = "declare
                                         l_id_content_exam_cat alert.exam_cat.id_content%type;
                                         begin
                                         select ec.id_exam_cat
                                         into l_id_content_exam_cat
                                         from alert.exam_cat ec
                                         where ec.id_content= '" & i_set_exams(i).id_content_category & "'
                                         and ec.flg_available='Y';
                                         insert into alert.exam (ID_EXAM, CODE_EXAM, FLG_PAT_RESP, FLG_PAT_PREP, FLG_MOV_PAT, FLG_AVAILABLE, RANK, FLG_TYPE, GENDER, AGE_MIN, AGE_MAX, ID_EXAM_CAT, ID_CONTENT)
                                         values (alert.seq_exam.nextval, 'EXAM.CODE_EXAM.'||alert.seq_exam.nextval, 'N', 'N', 'Y', 'Y', 0, '" & i_exam_type & "', '" & i_set_exams(i).gender & "' , " & i_set_exams(i).age_min & ", null, l_id_content_exam_cat,'" & i_set_exams(i).id_content_exam & "');
                                         end;"

                    Else

                        sql = "declare
                                         l_id_content_exam_cat alert.exam_cat.id_content%type;
                                         begin
                                         select ec.id_exam_cat
                                         into l_id_content_exam_cat
                                         from alert.exam_cat ec
                                         where ec.id_content= '" & i_set_exams(i).id_content_category & "'
                                         and ec.flg_available='Y';
                                         insert into alert.exam (ID_EXAM, CODE_EXAM, FLG_PAT_RESP, FLG_PAT_PREP, FLG_MOV_PAT, FLG_AVAILABLE, RANK, FLG_TYPE, GENDER, AGE_MIN, AGE_MAX, ID_EXAM_CAT, ID_CONTENT)
                                         values (alert.seq_exam.nextval, 'EXAM.CODE_EXAM.'||alert.seq_exam.nextval, 'N', 'N', 'Y', 'Y', 0, '" & i_exam_type & "', '" & i_set_exams(i).gender & "' , " & i_set_exams(i).age_min & ", " & i_set_exams(i).age_max & ", l_id_content_exam_cat,'" & i_set_exams(i).id_content_exam & "');
                                         end;"

                    End If

                    Dim cmd As New OracleCommand(sql, i_conn)

                    Try

                        cmd.CommandType = CommandType.Text
                        cmd.ExecuteNonQuery()
                        cmd.Dispose()

                    Catch ex As Exception

                        cmd.Dispose()
                        Return False

                    End Try

                    ''2 - Inserir a tradução do exame
                    ''2.1 - Obter code DE TRADUÇÃO do exame

                    Dim l_code_desc As String = ""

                    sql = "Select e.code_exam from alert.exam e
                            where e.flg_available='Y'
                            and   e.id_content='" & i_set_exams(i).id_content_exam & "'"

                    Dim cmd_get_code_trans As New OracleCommand(sql, i_conn)
                    Dim dr As OracleDataReader

                    Try

                        cmd_get_code_trans.CommandType = CommandType.Text
                        dr = cmd_get_code_trans.ExecuteReader()

                        While dr.Read()

                            l_code_desc = dr.Item(0)

                        End While

                    Catch ex As Exception

                        dr.Dispose()
                        dr.Close()
                        cmd.Dispose()
                        Return False

                    End Try

                    dr.Dispose()
                    dr.Close()
                    cmd.Dispose()

                    ''2.2 - Obter a tradução

                    Dim l_e_translation As String = i_set_exams(i).desc_exam

                    ''2.4 - Fazer INSERT

                    sql = "declare 

                                    l_desc clob; 

                                    l_id_lang integer := 0;

                            begin 

                                    l_id_lang:=" & l_id_language & ";

                                    select t.desc_lang_" & l_id_language & "
                                    into l_desc
                                    from alert_default.exam de
                                    join alert_default.translation t on t.code_translation=de.code_exam
                                    join alert.exam e on e.id_content=de.id_content
                                    where e.code_exam='" & l_code_desc & "';

                                    pk_translation.insert_into_translation( l_id_lang , '" & l_code_desc & "' , l_desc); 
                        end;"

                    Dim cmd_insert_trans As New OracleCommand(sql, i_conn)

                    Try


                        cmd_insert_trans.CommandType = CommandType.Text
                        cmd_insert_trans.ExecuteNonQuery()
                        cmd.Dispose()

                    Catch ex As Exception

                        cmd.Dispose()
                        Return False

                    End Try


                    'Existe na tabela de exames. Verificar se tem tradução para a língua da instituição
                ElseIf Not CHECK_EXAM_TRANSLATION_EXISTANCE(i_set_exams(i).id_content_exam, i_institution, i_conn) Then

                    ''2 - Inserir a tradução do exame

                    ''2.1 - Obter code DE TRADUÇÃO do exame

                    Dim l_code_desc As String = ""

                    Dim Sql = "Select e.code_exam from alert.exam e
                               where e.id_content ='" & i_set_exams(i).id_content_exam & "'"

                    Dim cmd_get_code_trans As New OracleCommand(Sql, i_conn)
                    cmd_get_code_trans.CommandType = CommandType.Text
                    Dim dr As OracleDataReader = cmd_get_code_trans.ExecuteReader()

                    While dr.Read()

                        l_code_desc = dr.Item(0)

                    End While

                    dr.Dispose()
                    dr.Close()
                    cmd_get_code_trans.Dispose()

                    ''2.2 - Obter a tradução 

                    Dim l_e_translation As String = i_set_exams(i).desc_exam


                    ''2.4 - Fazer INSERT  (TRANSFORMAR Em FUNÇÂO)
                    Sql = "declare 

                             l_desc clob; 

                             l_id_lang integer := 0;

                        begin 

                              l_id_lang:=" & l_id_language & ";

                            select t.desc_lang_" & l_id_language & "
                            into l_desc
                            from alert_default.exam de
                            join alert_default.translation t on t.code_translation=de.code_exam
                            join alert.exam e on e.id_content=de.id_content
                            where e.code_exam='" & l_code_desc & "';

                         pk_translation.insert_into_translation( l_id_lang , '" & l_code_desc & "' , l_desc); 
                        end;"

                    Dim cmd_insert_trans As New OracleCommand(Sql, i_conn)
                    cmd_insert_trans.CommandType = CommandType.Text
                    cmd_insert_trans.ExecuteNonQuery()
                    cmd_insert_trans.Dispose()

                    'Exame existe e tem tradução. É necessário garantir que está na categoria correta.
                Else

                    Try

                        If Not UPDATE_EXAM_CAT(i_set_exams(i).id_content_exam, i_set_exams(i).id_content_category, i_conn) Then

                            Return False

                        End If

                    Catch ex As Exception

                        Return False

                    End Try

                End If

                '4 - Inserir o registo na alert.exam_dep_clin_serv

            Next

        Catch ex As Exception

            Return False

        End Try

        '5 - correr o lucene?

        Return True

    End Function

    Function SET_EXAM_CAT_TRANSLATION(ByVal i_institution As Int64, ByVal i_exams As exams_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)
        Dim sql As String = "DECLARE

                                l_a_exams table_varchar := table_varchar('" & i_exams.id_content_category & "');"

        sql = sql & "   l_exam_desc alert_default.translation.desc_lang_1%TYPE;
                                l_exam_code alert.intervention.code_intervention%TYPE;

                            BEGIN

                                FOR i IN 1 .. l_a_exams.count()
                                LOOP
                                    BEGIN
        
                                        SELECT ec.code_exam_cat
                                        INTO l_exam_code
                                        FROM alert.exam_cat ec
                                        WHERE ec.id_content = l_a_exams(i)
                                        AND ec.flg_available = 'Y'
                                        AND pk_translation.get_translation(" & l_id_language & ", ec.code_exam_cat) IS NULL;
        
                                        IF l_exam_code IS NOT NULL
                                        THEN
            
                                            SELECT alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dec.code_exam_cat)
                                            INTO l_exam_desc
                                            FROM alert_default.exam_cat dec
                                            WHERE dec.id_content = l_a_exams(i)
                                            AND dec.flg_available = 'Y';
            
                                            SELECT ec.code_exam_cat
                                            INTO l_exam_code
                                            FROM alert.exam_cat ec
                                            WHERE ec.id_content = l_a_exams(i)
                                            AND ec.flg_available = 'Y';
            
                                            pk_translation.insert_into_translation(" & l_id_language & ", l_exam_code, l_exam_desc);
            
                                        END IF;
        
                                        l_exam_code := '';
        
                                    EXCEPTION
                                        WHEN OTHERS THEN
                                            continue;
                                    END;
                                END LOOP;

                            END;"

        Dim cmd_insert_exam_cat As New OracleCommand(sql, i_conn)

        Try
            cmd_insert_exam_cat.CommandType = CommandType.Text
            cmd_insert_exam_cat.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_exam_cat.Dispose()
            Return False
        End Try

        cmd_insert_exam_cat.Dispose()
        Return True

    End Function


    Function GET_DISTINCT_CATEGORIES(ByVal i_selected_default_analysis() As exams_default, ByVal i_conn As OracleConnection, ByRef i_Dr As OracleDataReader) As Boolean

        Dim sql As String = "Select distinct ec.id_content from alert.exam_cat ec
                                    where ec.flg_available = 'Y'
                                    and ec.id_content in ("


        For i As Integer = 0 To i_selected_default_analysis.Count() - 1

            If i < i_selected_default_analysis.Count() - 1 Then

                sql = sql & "'" & i_selected_default_analysis(i).id_content_category & "',"

            Else

                sql = sql & "'" & i_selected_default_analysis(i).id_content_category & "')"

            End If

        Next

        Dim cmd As New OracleCommand(sql, i_conn)

        Try

            cmd.CommandType = CommandType.Text
            i_Dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True

        Catch ex As Exception

            cmd.Dispose()
            Return False

        End Try

    End Function

    Function CHECK_RECORD_EXISTENCE(ByVal i_id_content_record As String, ByVal i_sql As String, ByVal i_conn As OracleConnection) As Boolean

        Dim l_total_records As Int16 = 0

        Dim sql As String = "Select count(*) from " & i_sql & " r
                                 where r.id_content='" & i_id_content_record & "'
                                 and r.flg_available='Y'"

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text
        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Try

            While dr.Read()

                l_total_records = dr.Item(0)

            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

            'Se l_total_records for maior que 0 significa que a análise já existe no ALERT

            If l_total_records > 0 Then

                Return True

            Else

                Return False

            End If

        Catch ex As Exception

            dr.Dispose()
            dr.Close()
            cmd.Dispose()
            Return False

        End Try

    End Function

    Function CHECK_RECORD_TRANSLATION_EXISTENCE(ByVal i_institution As Int64, ByVal id_content_record As String, ByVal i_sql As String, ByVal i_conn As OracleConnection) As Boolean

        Dim l_translation As String = ""

        Dim sql As String = "Select pk_translation.get_translation(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & "," & i_sql & " r
                             where r.id_content='" & id_content_record & "'
                             And r.flg_available='Y'"

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Try

            While dr.Read()

                l_translation = dr.Item(0)

            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        Catch ex As Exception

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

            Return False

        End Try

        Return True

    End Function

    Function GET_CODE_EXAM_CAT_ALERT(ByVal i_id_content_exam_cat As String, ByVal i_conn As OracleConnection) As String

        Dim sql As String = "Select ec.code_exam_cat from alert.exam_cat ec
                             where ec.id_content='" & i_id_content_exam_cat & "'
                             and ec.flg_available='Y'"

        Dim l_code As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_code = dr.Item(0)

        End While

        cmd.Dispose()
        dr.Dispose()
        dr.Close()

        Return l_code

    End Function

    Function GET_CODE_EXAM_CAT_DEFAULT(ByVal i_id_content_exam_cat As String, ByVal i_conn As OracleConnection) As String

        Dim sql As String = "Select ec.code_exam_cat from alert_default.exam_cat ec
                             where ec.id_content='" & i_id_content_exam_cat & "'
                             and ec.flg_available='Y'"

        Dim l_code As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_code = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_code

    End Function

    Function GET_ID_CAT_ALERT(ByVal i_id_content_exam_cat As String, ByVal i_conn As OracleConnection) As Int64

        Dim sql As String = "Select ec.id_exam_cat from alert.exam_cat ec
                             where ec.id_content='" & i_id_content_exam_cat & "'
                             and ec.flg_available='Y'"

        Dim l_id_alert As Int64 = 0

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_id_alert = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()
        Return l_id_alert

    End Function

    Function GET_CAT_RANK(ByVal i_id_content_exam_cat As String, ByVal i_conn As OracleConnection) As Int64

        Dim sql As String = "Select ec.rank from alert.exam_cat ec
                             where ec.id_content='" & i_id_content_exam_cat & "'
                             and ec.flg_available='Y'"

        Dim l_id_alert As Int64 = 0

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_id_alert = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()
        Return l_id_alert

    End Function

    Function GET_CAT_FLG_INTERFACE(ByVal i_id_content_exam_cat As String, ByVal i_conn As OracleConnection) As Char

        Dim sql As String = "Select ec.flg_interface from alert_DEFAULT.exam_cat ec
                             where ec.id_content='" & i_id_content_exam_cat & "'"

        Dim l_flg_interface As Char = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_flg_interface = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()
        Return l_flg_interface

    End Function

    Function SET_EXAM_CAT(ByVal i_institution As Int64, ByVal i_a_exams() As exams_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        '1 - Remover as categorias repetidas do array de entrada
        Dim l_a_distinct_ec() As String
        Dim dr_distinct_ec As OracleDataReader

        Try

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not GET_DISTINCT_CATEGORIES(i_a_exams, i_conn, dr_distinct_ec) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                dr_distinct_ec.Dispose()
                dr_distinct_ec.Close()
                Return False

            Else

                Dim l_index As Int64 = 0

                While dr_distinct_ec.Read()

                    ReDim Preserve l_a_distinct_ec(l_index)
                    l_a_distinct_ec(l_index) = dr_distinct_ec.Item(0)
                    l_index = l_index + 1

                End While

                dr_distinct_ec.Dispose()
                dr_distinct_ec.Close()

            End If

        Catch ex As Exception

            dr_distinct_ec.Dispose()
            dr_distinct_ec.Close()
            Return False

        End Try

        '2 - Processar as categorias de exames que já foram filtradas pelo bloco anterior
        Try

            'Ciclo que vai correr as categorias todas enviadas à função
#Disable Warning BC42104 ' Variable is used before it has been assigned a value
            For i As Integer = 0 To l_a_distinct_ec.Count() - 1
#Enable Warning BC42104 ' Variable is used before it has been assigned a value

                '' 1 - Verificar se existe Categoria pai
                Dim l_cat_parent As Int64 = 0
                Dim l_id_alert_cat_parent As Int64 = 0
                Dim l_rank As Int64 = 0
                Dim l_flg_interface As Char = ""

                Dim sql As String = "Select ec.parent_id from alert_default.Exam_Cat ec
                                     where ec.id_content='" & l_a_distinct_ec(i) & "'
                                     and ec.flg_available='Y'"

                Dim cmd As New OracleCommand(sql, i_conn)

                cmd.CommandType = CommandType.Text

                Dim dr As OracleDataReader = cmd.ExecuteReader()

                Try

                    While dr.Read()

                        l_cat_parent = dr.Item(0)

                    End While

                    dr.Dispose()
                    dr.Close()
                    cmd.Dispose()

                Catch ex As Exception

                    l_cat_parent = 0

                    dr.Dispose()
                    dr.Close()
                    cmd.Dispose()

                End Try

                If l_cat_parent > 0 Then 'Significa que existe Categoria Pai no default

                    '' 1.1 - Verificar se cat pai e tradução existem no alert. Se não existem, inserir.
                    ''1.1.1 - Verificar se existe no ALERT
                    sql = "Select ecp.id_content
                           from alert_default.exam_cat ec
                           join alert_default.exam_cat ecp
                           on ecp.id_exam_cat = ec.parent_id
                           where ec.id_content = '" & l_a_distinct_ec(i) & "'
                           and ec.flg_available='Y'"

                    Dim l_id_content_cat_parent As String = ""
                    Dim cmd_2 As New OracleCommand(sql, i_conn)
                    cmd_2.CommandType = CommandType.Text
                    Dim dr_2 As OracleDataReader = cmd_2.ExecuteReader()

                    While dr_2.Read()

                        l_id_content_cat_parent = dr_2.Item(0)

                    End While

                    dr_2.Dispose()
                    dr_2.Close()
                    cmd_2.Dispose()

                    If Not CHECK_RECORD_EXISTENCE(l_id_content_cat_parent, "alert.exam_cat", i_conn) Then 'Significa que Categoria Pai não existe no ALERT, é necessário inserir.

                        'INSERT EXAM_CAT_PARENT  -Criar função de inserção de categoria(Recursivo)? e função de inserção de tradução ( de tradução deve ir para o generall)
                        'Estrutura auxiliar para ser chamada na recursividade (apenas terá o  id_content da categoria pai)
                        Dim l_exam(0) As exams_default
                        l_exam(0).id_content_category = l_id_content_cat_parent

                        If Not SET_EXAM_CAT(i_institution, l_exam, i_conn) Then

                            Return False

                        End If

                        'Uma vez que foi adicionada uma nova categoria pai, sérá necessário atualizar o id alert da categoria pai das categorias filho
                        Dim sql_update_parents As String = "UPDATE alert.exam_cat ec
                                                            SET ec.parent_id = (Select ecp_n.id_exam_cat from alert.exam_cat ecp_n where ecp_n.id_content='" & l_id_content_cat_parent & "' and ecp_n.flg_available='Y')
                                                            WHERE ec.parent_id IN (SELECT ecp.id_exam_cat
                                                                               FROM alert.exam_cat ecp
                                                                               WHERE ecp.id_content = '" & l_id_content_cat_parent & "' and ec.flg_available='Y')"

                        Dim cmd_update_parents As New OracleCommand(sql_update_parents, i_conn)
                        cmd_update_parents.CommandType = CommandType.Text

                        cmd_update_parents.ExecuteNonQuery()

                        cmd_update_parents.Dispose()

                        '1.2 - Existe registo no ALERT, verificar se eciste tradução para a língua da isntituição
                    ElseIf Not CHECK_RECORD_TRANSLATION_EXISTENCE(i_institution, l_id_content_cat_parent, "r.code_exam_cat) from alert.exam_cat", i_conn) Then

                        ''Inserir tradução no ALERT
                        Dim l_code_cat_parent As String = GET_CODE_EXAM_CAT_ALERT(l_id_content_cat_parent, i_conn)
                        Dim l_exam_translation_default As String = db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, GET_CODE_EXAM_CAT_DEFAULT(l_id_content_cat_parent, i_conn), i_conn)

                        If Not db_access_general.SET_TRANSLATION(l_id_language, l_code_cat_parent, l_exam_translation_default, i_conn) Then

                            Return False

                        End If

                        '1.3 - Uma vez que existe no alert e existe tradução, verificar se tradução do alert é igual à do default
                    ElseIf Not db_access_general.CHECK_TRANSLATIONS(l_id_language, GET_CODE_EXAM_CAT_DEFAULT(l_id_content_cat_parent, i_conn), GET_CODE_EXAM_CAT_ALERT(l_id_content_cat_parent, i_conn), i_conn) Then

                        Dim l_code_cat_parent As String = GET_CODE_EXAM_CAT_ALERT(l_id_content_cat_parent, i_conn)
                        Dim l_exam_translation_default As String = db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, GET_CODE_EXAM_CAT_DEFAULT(l_id_content_cat_parent, i_conn), i_conn)

                        If Not db_access_general.SET_TRANSLATION(l_id_language, l_code_cat_parent, l_exam_translation_default, i_conn) Then

                            Return False

                        End If

                    End If

                    '' 1.2 - Se existir no alert, determinar id. (Neste ponto já vai sempre existir)
                    Try
                        l_id_alert_cat_parent = GET_ID_CAT_ALERT(l_id_content_cat_parent, i_conn)

                    Catch ex As Exception

                        Return False

                    End Try

                End If

                '2 - Verificar se categoria já existe no ALERT
                If Not CHECK_RECORD_EXISTENCE(l_a_distinct_ec(i), "alert.exam_cat", i_conn) Then

                    '2.1 - Não existe, Inserir.
                    '2.1.1 - Determinar RANK da categoria E flg_interface
                    Try

                        l_rank = GET_CAT_RANK(l_a_distinct_ec(i), i_conn)

                    Catch ex As Exception

                        l_rank = 0

                    End Try

                    '2.1.2 - Determinar flg_interface da categoria
                    Try

                        l_flg_interface = GET_CAT_FLG_INTERFACE(l_a_distinct_ec(i), i_conn)

                    Catch ex As Exception

                        l_flg_interface = "N"

                    End Try

                    '2.1.3 - Inserir Categoria
                    Dim sql_insert_cat As String

                    If l_id_alert_cat_parent = 0 Then

                        sql_insert_cat = "begin
                                      insert into alert.exam_cat (ID_EXAM_CAT, CODE_EXAM_CAT, FLG_AVAILABLE, FLG_LAB, ID_CONTENT, FLG_INTERFACE, RANK, PARENT_ID)
                                      values (alert.seq_exam_cat.nextval, 'EXAM_CAT.CODE_EXAM_CAT.' || alert.seq_exam_cat.nextval, 'Y', 'N', '" & l_a_distinct_ec(i) & "', '" & l_flg_interface & "', " & l_rank & ", null);
                                      end;"
                    Else

                        sql_insert_cat = "begin
                                      insert into alert.exam_cat (ID_EXAM_CAT, CODE_EXAM_CAT, FLG_AVAILABLE, FLG_LAB, ID_CONTENT, FLG_INTERFACE, RANK, PARENT_ID)
                                      values (alert.seq_exam_cat.nextval, 'EXAM_CAT.CODE_EXAM_CAT.' || alert.seq_exam_cat.nextval, 'Y', 'N', '" & l_a_distinct_ec(i) & "', '" & l_flg_interface & "', " & l_rank & ", " & l_id_alert_cat_parent & ");
                                      end;"

                    End If

                    Dim cmd_insert_cat As New OracleCommand(sql_insert_cat, i_conn)
                    cmd_insert_cat.CommandType = CommandType.Text

                    Try
                        cmd_insert_cat.ExecuteNonQuery()
                    Catch ex As Exception

                        cmd_insert_cat.Dispose()
                        Return False

                    End Try

                    cmd_insert_cat.Dispose()

                    Dim l_code_ec_default As String = GET_CODE_EXAM_CAT_DEFAULT(l_a_distinct_ec(i), i_conn)
                    Dim l_code_ec_alert As String = GET_CODE_EXAM_CAT_ALERT(l_a_distinct_ec(i), i_conn)

                    '2.1.4 Inserir translation
                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_ec_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_ec_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING CATEGORY TRANSLATION - LABS_API >> SET_TRANSLATION")
                        Return False

                    End If

                    '2.1.5 - Fazer update a todas as análises que utilizavam o id da categoria antiga com o id da nova categoria (alert.analysis_instit_soft)

                    Dim l_id_alert_category As Int64 = GET_ID_CAT_ALERT(l_a_distinct_ec(i), i_conn)

                    Dim sql_update_analysis_cat As String = "update alert.analysis_instit_soft ais 
                                                         set ais.id_exam_cat=" & l_id_alert_category & "
                                                         where ais.id_exam_cat in (select ec.id_exam_cat  from alert.exam_cat ec where ec.id_content='" & l_a_distinct_ec(i) & "')"

                    Dim cmd_update_analysis_cat As New OracleCommand(sql_update_analysis_cat, i_conn)
                    cmd_update_analysis_cat.CommandType = CommandType.Text

                    cmd_update_analysis_cat.ExecuteNonQuery()

                    cmd_update_analysis_cat.Dispose()

                    '2.2 - Uma vez que existe no ALERT, verificar se exsite tradução para a lingua da instituição
                ElseIf Not CHECK_RECORD_TRANSLATION_EXISTENCE(i_institution, l_a_distinct_ec(i), "r.code_exam_cat) from alert.exam_cat", i_conn) Then

                    Dim l_code_ec_default As String = GET_CODE_EXAM_CAT_DEFAULT(l_a_distinct_ec(i), i_conn)
                    Dim l_code_ec_alert As String = GET_CODE_EXAM_CAT_ALERT(l_a_distinct_ec(i), i_conn)

                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_ec_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_ec_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING EXAM_CAT TRANSLATION - LABS_API >> CHECK_RECORD_TRANSLATION_EXISTENCE >> SET_TRANSLATION " & l_id_language)
                        Return False

                    End If

                    '2.3 - Umvez que existe no alert e existe tradução, verificar se tradução do alert é igual à do default
                ElseIf Not db_access_general.CHECK_TRANSLATIONS(l_id_language, GET_CODE_EXAM_CAT_DEFAULT(l_a_distinct_ec(i), i_conn), GET_CODE_EXAM_CAT_ALERT(l_a_distinct_ec(i), i_conn), i_conn) Then

                    Dim l_code_ec_default As String = GET_CODE_EXAM_CAT_DEFAULT(l_a_distinct_ec(i), i_conn)
                    Dim l_code_ec_alert As String = GET_CODE_EXAM_CAT_ALERT(l_a_distinct_ec(i), i_conn)

                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_ec_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_ec_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING EXAM_CAT TRANSLATION - LABS_API >> CHECK_RECORD_TRANSLATION_EXISTENCE >> SET_TRANSLATION" & l_id_language)
                        Return False

                    End If

                End If
            Next

        Catch ex As Exception

            Return False

        End Try

        Return True

    End Function

    Function SET_EXAMS_DEP_CLIN_SERV_FREQ(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_a_exams As exams_alert_flg, ByVal i_dep_clin_serv As Int64, ByVal i_flg_type As Integer, ByVal i_conn As OracleConnection) As Boolean

        Dim sql As String = ""

        sql = "DECLARE

                                l_a_exams table_varchar := table_varchar(" & "'" & i_a_exams.id_content_exam & "'); "

        sql = sql & " l_id_exam           alert.exam.id_exam%TYPE;
                      l_id_insert_type    INTEGER :=" & i_flg_type & ";
                      l_id_dep_clin_serv  alert.dep_clin_serv.id_dep_clin_serv%type := " & i_dep_clin_serv & ";

                    BEGIN

                        FOR i IN 1 .. l_a_exams.count()
                        LOOP
    
                            SELECT e.id_exam
                            INTO l_id_exam
                            FROM alert.exam e
                            WHERE e.id_content = l_a_exams(i)
                            and e.flg_available='Y';
    
                      BEGIN
                            IF l_id_insert_type = 1
                            THEN       
                                
                                insert into alert.exam_dep_clin_serv (ID_EXAM_DEP_CLIN_SERV, ID_EXAM, ID_DEP_CLIN_SERV, FLG_TYPE, RANK, ID_INSTITUTION, ID_SOFTWARE, FLG_FIRST_RESULT, FLG_MOV_PAT, FLG_EXECUTE,      FLG_TIMEOUT, FLG_RESULT_NOTES, FLG_FIRST_EXECUTE)
                                values (alert.seq_exam_dep_clin_serv.nextval, l_id_exam, l_id_dep_clin_serv, 'M', 0, " & i_institution & ",  " & i_software & ", 'DTN', 'N', 'Y', 'N', 'N', 'DTN');

                            ELSIF l_id_insert_type = 2
                            THEN
        
                                insert into alert.exam_dep_clin_serv (ID_EXAM_DEP_CLIN_SERV, ID_EXAM, ID_DEP_CLIN_SERV, FLG_TYPE, RANK, ID_INSTITUTION, ID_SOFTWARE, FLG_FIRST_RESULT, FLG_MOV_PAT, FLG_EXECUTE,      FLG_TIMEOUT, FLG_RESULT_NOTES, FLG_FIRST_EXECUTE)
                                values (alert.seq_exam_dep_clin_serv.nextval, l_id_exam, l_id_dep_clin_serv, 'A', 0, " & i_institution & ",  " & i_software & ", 'DTN', 'N', 'Y', 'N', 'N', 'DTN');

                            ELSE
        
                                insert into alert.exam_dep_clin_serv (ID_EXAM_DEP_CLIN_SERV, ID_EXAM, ID_DEP_CLIN_SERV, FLG_TYPE, RANK, ID_INSTITUTION, ID_SOFTWARE, FLG_FIRST_RESULT, FLG_MOV_PAT, FLG_EXECUTE,      FLG_TIMEOUT, FLG_RESULT_NOTES, FLG_FIRST_EXECUTE)
                                values (alert.seq_exam_dep_clin_serv.nextval, l_id_exam, l_id_dep_clin_serv, 'M', 0, " & i_institution & ",  " & i_software & ", 'DTN', 'N', 'Y', 'N', 'N', 'DTN');

                                insert into alert.exam_dep_clin_serv (ID_EXAM_DEP_CLIN_SERV, ID_EXAM, ID_DEP_CLIN_SERV, FLG_TYPE, RANK, ID_INSTITUTION, ID_SOFTWARE, FLG_FIRST_RESULT, FLG_MOV_PAT, FLG_EXECUTE,      FLG_TIMEOUT, FLG_RESULT_NOTES, FLG_FIRST_EXECUTE)
                                values (alert.seq_exam_dep_clin_serv.nextval, l_id_exam, l_id_dep_clin_serv, 'A', 0, " & i_institution & ",  " & i_software & ", 'DTN', 'N', 'Y', 'N', 'N', 'DTN');

                            END IF;

                    EXCEPTION
                      WHEN DUP_VAL_ON_INDEX THEN
                        CONTINUE;
            
                        END;"

        Dim cmd_insert_exam_dep_clin_serv As New OracleCommand(sql, i_conn)

        Try
            cmd_insert_exam_dep_clin_serv.CommandType = CommandType.Text
            cmd_insert_exam_dep_clin_serv.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_exam_dep_clin_serv.Dispose()
            Return False
        End Try

        cmd_insert_exam_dep_clin_serv.Dispose()
        Return True

    End Function

End Class
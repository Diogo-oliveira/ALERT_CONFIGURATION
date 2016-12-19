Imports Oracle.DataAccess.Client
Public Class EXAMS_API

    ''Estrutura dos exames carregados do default
    Public Structure exams_default
        Public id_content_category As String
        Public id_content_exam As String
        Public flg_first_result As String
        Public flg_execute As String
        Public flg_timeout As String
        Public flg_result_notes As String
        Public flg_first_execute As String
        Public age_min As Integer
        Public age_max As Integer
        Public gender As String
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

        Return l_id_inst

    End Function

    Public Function GET_INSTITUTION(ByVal i_ID_INST As Int16, ByVal i_oradb As String) As String

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

    Public Function GET_SOFT_INST(ByVal i_ID_INST As Int16, ByVal i_oradb As String) As OracleDataReader

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

    Public Function GET_CLIN_SERV(ByVal i_ID_INST As Int16, ByVal i_ID_SOFT As Int16, ByVal i_oradb As String) As OracleDataReader

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

    Function GET_SELECTED_SOFT(ByVal i_index As Int16, ByVal i_inst As Int16, ByVal i_oradb As String) As Int16

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

        Return l_soft

    End Function

    Function GET_FREQ_EXAM(ByVal I_ID_SOFT As Int16, ByVal I_ID_DEP_CLIN_SERV As Int64, ByVal I_ID_INST As Int64, ByVal i_oradb As String) As OracleDataReader

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

        Dim sql As String = "Select s.id_exam,decode(i.id_market,
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
              T.desc_lang_19) from alert.exam_dep_clin_serv s
 join alert.exam e on e.id_exam=s.id_exam
 join translation t on t.code_translation=e.code_exam
 join alert.dep_clin_serv dps on dps.id_dep_clin_serv=s.id_dep_clin_serv
 join alert.department dep on dep.id_department=dps.id_department
 join institution i on i.id_institution=dep.id_institution
 join alert.exam_dep_clin_serv eds_P on eds_P.Id_Exam = s.id_exam and eds_P.Id_Institution = " & I_ID_INST & " and eds_P.Id_Software= " & I_ID_SOFT & " and eds_P.Flg_Type='P'
 where s.id_software=" & I_ID_SOFT & "
 and s.flg_type='M'
 and s.id_dep_clin_serv = " & I_ID_DEP_CLIN_SERV & "
 and e.flg_available='Y'
 and e.flg_type='I'
 order by 2 asc"

        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Return dr

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

    Function DELETE_EXAMS_DEP_CLIN_SERV(ByVal i_exam As Integer(), ByVal i_dep_clin_serv As Int64, ByVal i_oradb As String) As Boolean

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

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Function GET_EXAMS_CAT(ByVal i_id_inst As Int64, ByVal i_id_soft As Int64, ByVal i_oradb As String) As OracleDataReader

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

        Dim sql As String = "Select distinct (decode(i.id_market,
              1,
              tec.desc_lang_1,
              2,
              tec.desc_lang_2,
              3,
              tec.desc_lang_11,
              4,
              tec.desc_lang_5,
              5,
              tec.desc_lang_4,
              6,
              tec.desc_lang_3,
              7,
              tec.desc_lang_10,
              8,
              tec.desc_lang_7,
              9,
              tec.desc_lang_6,
              10,
              tec.desc_lang_9,
              12,
              tec.desc_lang_16,
              16,
              tec.desc_lang_17,
              17,
              tec.desc_lang_18,
              19,
              tec.desc_lang_19)),ec.id_exam_cat
              
              from alert.exam_dep_clin_serv d
 join alert.exam e on e.id_exam=d.id_exam
  join alert.exam_cat ec on ec.id_exam_cat=e.id_exam_cat
 join translation tec on tec.code_translation=ec.code_exam_cat
 join institution i on i.id_institution= " & i_id_inst & "
 where d.id_institution = " & i_id_inst & "
 and e.flg_type='I'
 and e.flg_available='Y' and ec.flg_available='Y'
 and d.id_software= " & i_id_soft & "
 and d.flg_type = 'P'
 order by 1 asc"

        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Return dr

    End Function

    Function GET_EXAMS(ByVal i_id_inst As Int64, ByVal i_id_soft As Int64, ByVal i_id_exam_cat As Int64, ByVal i_oradb As String) As OracleDataReader

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

        Dim sql As String = ""

        If i_id_exam_cat = 0 Then

            sql = " Select distinct (decode(i.id_market,
              1,
              te.desc_lang_1,
              2,
              te.desc_lang_2,
              3,
              te.desc_lang_11,
              4,
              te.desc_lang_5,
              5,
              te.desc_lang_4,
              6,
              te.desc_lang_3,
              7,
              te.desc_lang_10,
              8,
              te.desc_lang_7,
              9,
              te.desc_lang_6,
              10,
              te.desc_lang_9,
              12, 
              te.desc_lang_16,
              16,
              te.desc_lang_17,
              17,
              te.desc_lang_18,
              19,
              te.desc_lang_19)),ec.id_exam_cat,e.id_exam
              
              from alert.exam_dep_clin_serv d
 join alert.exam e on e.id_exam=d.id_exam
 join translation te on te.code_translation=e.code_exam
 join alert.exam_cat ec on ec.id_exam_cat=e.id_exam_cat
 join institution i on i.id_institution= " & i_id_inst & "
 where d.id_institution = " & i_id_inst & "
 and e.flg_type='I'
 and e.flg_available='Y' and ec.flg_available='Y'
 and d.id_software= " & i_id_soft & "
 and d.flg_type = 'P'
 order by 1 asc"

        Else

            sql = " Select distinct (decode(i.id_market,
              1,
              te.desc_lang_1,
              2,
              te.desc_lang_2,
              3,
              te.desc_lang_11,
              4,
              te.desc_lang_5,
              5,
              te.desc_lang_4,
              6,
              te.desc_lang_3,
              7,
              te.desc_lang_10,
              8,
              te.desc_lang_7,
              9,
              te.desc_lang_6,
              10,
              te.desc_lang_9,
              12,
              te.desc_lang_16,
              16,
              te.desc_lang_17,
              17,
              te.desc_lang_18,
              19,
              te.desc_lang_19)),ec.id_exam_cat,e.id_exam
              
              from alert.exam_dep_clin_serv d
 join alert.exam e on e.id_exam=d.id_exam
 join translation te on te.code_translation=e.code_exam
 join alert.exam_cat ec on ec.id_exam_cat=e.id_exam_cat
 join institution i on i.id_institution= " & i_id_inst & "
 where d.id_institution = " & i_id_inst & "
 and e.flg_type='I'
 and e.flg_available='Y' and ec.flg_available='Y'
 and d.id_software=" & i_id_soft & "
 and d.flg_type = 'P'
 and e.id_exam_cat = " & i_id_exam_cat & " 
 order by 1 asc"

        End If

        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Return dr

    End Function

    Function DELETE_EXAMS(ByVal i_exam As Int64(), ByVal i_institution As Int64, ByVal i_software As Int64, ByVal i_oradb As String) As Boolean

        Try

            Dim oradb As String = i_oradb

            Dim conn As New OracleConnection(oradb)

            conn.Open()

            For i As Integer = 0 To i_exam.Count() - 1

                Dim sql As String = "   DELETE from alert.exam_dep_clin_serv s
                                        where s.id_exam= " & i_exam(i) & "
                                        and (
                                        (s.id_institution= " & i_institution & " and s.flg_type='P' and s.id_software= " & i_software & " ) 
                                        or 
                                        ((s.id_institution is null or s.id_institution=" & i_institution & ") and s.flg_type='M' and s.id_software= " & i_software & " )
                                        )"

                Dim cmd As New OracleCommand(sql, conn)
                cmd.CommandType = CommandType.Text

                cmd.ExecuteNonQuery()

            Next

            conn.Close()

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Function GET_EXAMS_CAT_DEFAULT(ByVal i_version As String, ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_oradb As String) As OracleDataReader

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

        Dim sql As String = "Select distinct ec.id_content, 
       decode(v.id_market,
              1,
              tc.desc_lang_1,
              2,
              tc.desc_lang_2,
              3,
              tc.desc_lang_11,
              4,
              tc.desc_lang_5,
              5,
              tc.desc_lang_4,
              6,
              tc.desc_lang_3,
              7,
              tc.desc_lang_10,
              8,
              tc.desc_lang_7,
              9,
              tc.desc_lang_6,
              10,
              tc.desc_lang_9,
              12,
              tc.desc_lang_16,
              16,
              tc.desc_lang_17,
              17,
              tc.desc_lang_18,
              19,
              tc.desc_lang_19)
  From alert_default.exam e
  Join alert_default.exam_mrk_vrs v
    On v.id_exam = e.id_exam
    Join alert_default.exam_cat ec on ec.id_exam_cat=e.id_exam_cat
    Join alert_default.translation tc on tc.code_translation=ec.code_exam_cat
    Join alert_default.exam_clin_serv ecs on ecs.id_exam=e.id_exam 
    join institution i on i.id_market=v.id_market
 where i.id_institution= " & i_institution & "
   And v.version = '" & i_version & "'
   And e.flg_type='I'
   And e.flg_available='Y'
   And ecs.id_software= " & i_software & " 
   And ecs.flg_type='P'
   order by 2 asc"


        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Return dr

    End Function

    Function GET_DEFAULT_VERSIONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_oradb As String) As OracleDataReader

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

        Dim sql As String = "Select distinct v.version
  from alert_default.exam e
  join alert_default.exam_mrk_vrs v
    on v.id_exam = e.id_exam
  join alert_default.exam_clin_serv ecs
    on ecs.id_exam = e.id_exam
   and ecs.flg_type = 'P'
  join institution i
    on i.id_market = v.id_market
 where i.id_institution = 470
 order by 1 asc"

        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Return dr

    End Function


    Function GET_EXAMS_DEFAULT_BY_CAT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_version As String, ByVal i_id_cat As String, ByVal i_oradb As String) As OracleDataReader

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

        Dim sql As String = "Select ec.id_content, 
       decode(v.id_market,
              1,
              tc.desc_lang_1,
              2,
              tc.desc_lang_2,
              3,
              tc.desc_lang_11,
              4,
              tc.desc_lang_5,
              5,
              tc.desc_lang_4,
              6,
              tc.desc_lang_3,
              7,
              tc.desc_lang_10,
              8,
              tc.desc_lang_7,
              9,
              tc.desc_lang_6,
              10,
              tc.desc_lang_9,
             12,
              tc.desc_lang_16,
              16,
              tc.desc_lang_17,
              17,
              tc.desc_lang_18,
              19,
              tc.desc_lang_19), 
       e.id_content, 
       decode(v.id_market,
              1,
              te.desc_lang_1,
              2,
              te.desc_lang_2,
              3,
              te.desc_lang_11,
              4,
              te.desc_lang_5,
              5,
              te.desc_lang_4,
              6,
              te.desc_lang_3,
              7,
              te.desc_lang_10,
              8,
              te.desc_lang_7,
              9,
              te.desc_lang_6,
              10,
              te.desc_lang_9,
              12,
              te.desc_lang_16,
              16,
              te.desc_lang_17,
              17,
              te.desc_lang_18,
              19,
              te.desc_lang_19),
       ecs.flg_first_result, ecs.flg_execute, ecs.flg_timeout, ecs.flg_result_notes, ecs.flg_first_execute,
       e.age_min, e.age_max, e.gender
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
   and e.flg_type='I'
   and e.flg_available='Y'
   and ecs.id_software= " & i_software & "
   and ecs.flg_type='P'"

        If i_id_cat = "0" Then

            sql = sql & " order by 2 asc, 4 asc"

        Else

            sql = sql & " And ec.id_content = '" & i_id_cat & "'
                         order by 2 asc, 4 asc"
        End If

        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Return dr

    End Function

    Function CHECK_CATEGORY_EXISTANCE(ByVal i_id_content_cat, ByVal i_id_institution, ByVal i_oradb) As Boolean

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

        Dim sql As String = "Select count(*)
  from alert.exam_cat ec
  join translation t
    on t.code_translation = ec.code_exam_cat
  join institution i
    on i.id_institution = " & i_id_institution & "
 where ec.id_content = '" & i_id_content_cat & "'
   and ec.flg_available = 'Y'
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

            Dim l_cat_exist As Integer = 0

            While dr.Read()

                l_cat_exist = dr.Item(0)

            End While

            If l_cat_exist > 0 Then

                Return True

            Else

                Return False

            End If

        Catch ex As Exception

            Return False

        End Try

    End Function

    Function SET_EXAM_ALERT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_set_exams() As exams_default, ByVal i_oradb As String) As Boolean

        'Function insert() exam
        '1 - VEr se categoria já existe no lado do alert
        Dim oradb As String = i_oradb

        ' Dim conn As New OracleConnection(oradb)

        ' conn.Open()

        Try

            Dim conn As New OracleConnection(oradb)

            conn.Open()

            For i As Integer = 0 To i_set_exams.Count() - 1

                If Not CHECK_CATEGORY_EXISTANCE(i_set_exams(i).id_content_category, i_institution, i_oradb) Then


                    ''1 - Inserir a Categoria
                    Dim sql As String = "insert into alert.exam_cat (ID_EXAM_CAT, CODE_EXAM_CAT, ADW_LAST_UPDATE, FLG_AVAILABLE, FLG_LAB, ID_CONTENT, FLG_INTERFACE, RANK)
                                         values (alert.seq_exam_cat.nextval, 'EXAM_CAT.CODE_EXAM_CAT.'||alert.seq_exam_cat.nextval, null, 'Y', 'N','" & i_set_exams(i).id_content_category & "' , 'N', 0)"


                    Dim cmd As New OracleCommand(sql, conn)
                    cmd.CommandType = CommandType.Text

                    cmd.ExecuteNonQuery()

                    ''2 - Inserir a tradução da categoria

                    ''2.1 - Obter code DE TRADUÇÃO do exame

                    Dim l_code_desc As String = ""

                    sql = "Select ec.code_exam_cat from alert.exam_cat ec
                           where ec.id_content ='" & i_set_exams(i).id_content_category & "'"

                    Dim cmd_get_code_trans As New OracleCommand(sql, conn)
                    cmd_get_code_trans.CommandType = CommandType.Text
                    Dim dr As OracleDataReader = cmd_get_code_trans.ExecuteReader()

                    While dr.Read()

                        l_code_desc = dr.Item(0)

                    End While

                    ''2.2 - Obter a tradução (TRANSFORMAR Em FUNÇÂO)

                    Dim l_ec_translation As String = ""

                    sql = "Select decode(i.id_market,
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
              t.desc_lang_19) from alert_default.translation t
join alert_default.exam_cat ec on ec.code_exam_cat=t.code_translation
join institution i on i.id_institution= " & i_institution & " 
where ec.id_content='" & i_set_exams(i).id_content_category & "'"

                    '
                    Dim cmd_get_trans As New OracleCommand(sql, conn)
                    cmd_get_trans.CommandType = CommandType.Text
                    dr = cmd_get_trans.ExecuteReader()

                    While dr.Read()

                        l_ec_translation = dr.Item(0)

                    End While

                    ''2.3 - Obter ID da lingua da instituição  (TRANSFORMAR Em FUNÇÂO)

                    Dim l_id_lang As Integer = 0

                    sql = "Select decode(i.id_market,
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

                    Dim cmd_get_id_lang As New OracleCommand(sql, conn)
                    cmd_get_id_lang.CommandType = CommandType.Text
                    dr = cmd_get_id_lang.ExecuteReader()

                    While dr.Read()

                        l_id_lang = dr.Item(0)

                    End While

                    ''2.4 - Fazer INSERT  (TRANSFORMAR Em FUNÇÂO)
                    sql = "begin pk_translation.insert_into_translation( " & l_id_lang & " , '" & l_code_desc & "' , '" & l_ec_translation & "' ); end;"

                    Dim cmd_insert_trans As New OracleCommand(sql, conn)
                    cmd_insert_trans.CommandType = CommandType.Text

                    cmd_insert_trans.ExecuteNonQuery()

                End If

            Next

            conn.Close()
        Catch ex As Exception
            Return False
        End Try

        '2 - Se não existir, será necessário inserir exame e tradução
        '3 - Fazer o mesmo para a categoria
        '4 - Inserir o registo na alert.exam_dep_clin_serv
        '5 - correr o lucene?

        Return True

    End Function

End Class
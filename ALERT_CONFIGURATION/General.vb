Imports Oracle.DataAccess.Client
Public Class General

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

        conn.Close()
        conn.Dispose()

        Return l_soft

    End Function

    Function GET_DEFAULT_TRANSLATION(ByVal i_lang As Int16, ByVal i_code_translation As String, ByVal i_oradb As String) As String

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

        Dim sql As String = "select alert_default.pk_translation_default.get_translation_default(" & i_lang & ",'" & i_code_translation & "') from dual"

        Dim translation As String = ""

        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Try

            While dr.Read()

                translation = dr.Item(0)

            End While

        Catch ex As Exception

            Return "No available translation!"

        End Try

        Return translation

    End Function

    Function GET_ID_LANG(ByVal i_id_institution As Int64, ByVal i_oradb As String) As Int16

        Dim l_id_market As Int16 = 0

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()


        Dim sql As String = "Select i.id_market from institution i
                             where i.id_institution= " & i_id_institution

        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()


        While dr.Read()

            l_id_market = dr.Item(0)

        End While


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

End Class

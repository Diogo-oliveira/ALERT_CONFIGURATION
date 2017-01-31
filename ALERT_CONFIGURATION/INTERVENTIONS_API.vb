Imports Oracle.DataAccess.Client
Public Class INTERVENTIONS_API

    Dim db_access_general As New General

    Public Structure interventions_default
        Public id_content_category As String
        Public id_content_intervention As String
        Public desc_intervention As String
    End Structure

    Function GET_DEFAULT_VERSIONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT DISTINCT dim.version
                                FROM alert_default.intervention di
                                JOIN alert_default.translation dti ON dti.code_translation = di.code_intervention
                                JOIN alert_default.interv_int_cat diic ON diic.id_intervention = di.id_intervention
                                JOIN alert_default.interv_mrk_vrs dim ON dim.id_intervention = di.id_intervention
                                JOIN alert.interv_category aic ON aic.id_interv_category = diic.id_interv_category
                                JOIN translation t ON t.code_translation = aic.code_interv_category
                                JOIN institution i ON i.id_market = dim.id_market
                                WHERE di.flg_status = 'A'
                                AND diic.id_software IN (0, " & i_software & ")
                                AND i.id_institution =" & i_institution & "
                                ORDER BY 1 ASC"
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

    Function GET_INTERV_CATS_INST_SOFT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT DISTINCT ic.id_content, pk_translation.get_translation(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", ic.code_interv_category)
                                FROM alert.intervention i
                                JOIN alert.interv_int_cat iic ON iic.id_intervention = i.id_intervention
                                JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                                JOIN alert.interv_dep_clin_serv idcs ON idcs.id_intervention = i.id_intervention
                                WHERE i.flg_status = 'A'
                                AND iic.id_software IN (0, " & i_software & ")
                                AND iic.id_institution IN (0, " & i_institution & ")
                                AND iic.flg_add_remove = 'A'
                                AND ic.flg_available = 'Y'
                                AND pk_translation.get_translation(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", ic.code_interv_category) IS NOT NULL
                                AND idcs.id_institution IN (0, " & i_institution & ")
                                AND idcs.flg_type = 'P'
                                ORDER BY 2 ASC"
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

    Function GET_INTERV_CATS_DEFAULT(ByVal i_version As String, ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT DISTINCT ic.id_content, pk_translation.get_translation(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", ic.code_interv_category)
                                FROM alert_default.intervention di
                                JOIN alert_default.interv_int_cat diic ON diic.id_intervention = di.id_intervention
                                JOIN alert_default.interv_mrk_vrs dim ON dim.id_intervention = di.id_intervention
                                JOIN alert.interv_category ic ON ic.id_interv_category = diic.id_interv_category
                                WHERE di.flg_status = 'A'
                                AND diic.id_software IN (0, " & i_software & ")
                                AND ic.flg_available = 'Y'
                                AND pk_translation.get_translation(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", ic.code_interv_category) IS NOT NULL
                                AND dim.version = '" & i_version & "'
                                ORDER BY 2 ASC"
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

    Function GET_INTERVS_DEFAULT_BY_CAT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_version As String, ByVal i_id_cat As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT ic.id_content, di.id_content, alert_default.pk_translation_default.get_translation_default(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", di.code_intervention)
                                FROM alert_default.intervention di
                                JOIN alert_default.interv_int_cat diic ON diic.id_intervention = di.id_intervention
                                JOIN alert_default.interv_mrk_vrs dim ON dim.id_intervention = di.id_intervention
                                JOIN alert.interv_category ic ON ic.id_interv_category = diic.id_interv_category
                                WHERE di.flg_status = 'A'
                                AND diic.id_software IN (0, " & i_software & ")
                                AND ic.flg_available = 'Y'
                                --AND alert_default.pk_translation_default.get_translation_default(6, di.code_intervention) IS NOT NULL
                                AND dim.version = '" & i_version & "'"
        If i_id_cat = "0" Then
            sql = sql & " order by 3 asc"
        Else
            sql = sql & " and ic.id_content= '" & i_id_cat & "'
                          order by 3 asc"
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

End Class

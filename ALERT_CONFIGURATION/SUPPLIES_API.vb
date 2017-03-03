Imports Oracle.DataAccess.Client
Public Class SUPPLIES_API

    Dim db_access_general As New General

    Public Structure SUP_AREAS
        Public id_supply_area As Integer
        Public desc_supply_area As String
    End Structure

    Function GET_SUP_AREAS(ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT s.id_supply_area, pk_translation.get_translation(2, s.code_supply_area)
                             FROM alert.supply_area s"

        sql = sql & "  ORDER BY 2 ASC"

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

    Function GET_DEFAULT_VERSIONS(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_sup_area As Int16, ByVal i_sup_type As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = ""

        If i_sup_type = "ALL" Then

            sql = "SELECT DISTINCT (dmv.version)
                                FROM alert_default.supply_mrk_vrs dmv
                                JOIN alert_default.supply ds ON ds.id_supply = dmv.id_supply
                                JOIN alert_default.supply_soft_inst dssi ON dssi.id_supply = ds.id_supply
                                JOIN alert_default.supply_sup_area dssa ON dssa.id_supply_soft_inst = dssi.id_supply_soft_inst
                                JOIN alert_default.supply_type dst ON dst.id_supply_type = ds.id_supply_type
                                JOIN alert_core_data.ab_institution i ON i.id_ab_market = dmv.id_market
                                WHERE i.id_ab_institution = " & i_institution & "
                                AND ds.flg_available = 'Y'
                                AND dssi.id_software IN (0, " & i_software & ") --Há para software 0
                                AND dssa.id_supply_area = " & i_sup_area & "
                                AND dssa.flg_available = 'Y'
                                AND dst.flg_available = 'Y'
                                ORDER BY 1 ASC"
        Else

            sql = "SELECT DISTINCT (dmv.version)
                                FROM alert_default.supply_mrk_vrs dmv
                                JOIN alert_default.supply ds ON ds.id_supply = dmv.id_supply
                                JOIN alert_default.supply_soft_inst dssi ON dssi.id_supply = ds.id_supply
                                JOIN alert_default.supply_sup_area dssa ON dssa.id_supply_soft_inst = dssi.id_supply_soft_inst
                                JOIN alert_default.supply_type dst ON dst.id_supply_type = ds.id_supply_type
                                JOIN alert_core_data.ab_institution i ON i.id_ab_market = dmv.id_market
                                WHERE i.id_ab_institution = " & i_institution & "
                                AND ds.flg_available = 'Y'
                                AND ds.flg_type = '" & i_sup_type & "'
                                AND dssi.id_software IN (0, " & i_software & ") --Há para software 0
                                AND dssa.id_supply_area = " & i_sup_area & "
                                AND dssa.flg_available = 'Y'
                                AND dst.flg_available = 'Y'
                                ORDER BY 1 ASC"

        End If

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

    Function GET_SUPP_CATS_DEFAULT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_version As String, ByVal i_sup_area As Int16, ByVal i_flg_type As Integer, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)
        Dim sql As String = "SELECT DISTINCT dst.id_content, alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dst.code_supply_type)
                                FROM alert_default.supply_mrk_vrs dmv
                                JOIN alert_default.supply ds ON ds.id_supply = dmv.id_supply
                                JOIN alert_default.supply_soft_inst dssi ON dssi.id_supply = ds.id_supply
                                JOIN alert_default.supply_sup_area dssa ON dssa.id_supply_soft_inst = dssi.id_supply_soft_inst
                                JOIN alert_default.supply_type dst ON dst.id_supply_type = ds.id_supply_type
                                JOIN alert_core_data.ab_institution i ON i.id_ab_market = dmv.id_market
                                WHERE i.id_ab_institution = " & i_institution & "
                                AND dmv.version = '" & i_version & "'
                                AND ds.flg_available = 'Y'
                                AND ds.flg_type = '" & i_flg_type & "'
                                AND dssi.id_software IN (0, " & i_software & ")
                                AND dssa.id_supply_area = " & i_sup_area & "
                                AND dssa.flg_available = 'Y'
                                AND dst.flg_available = 'Y'
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

End Class

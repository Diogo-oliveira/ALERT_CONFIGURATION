Imports Oracle.DataAccess.Client
Public Class DISCHARGE_API

    Dim db_access_general As New General

    Public Structure DEFAULT_DISCAHRGE
        Public id_disch_reas_dest As Int64
        Public id_content As String
        Public desccription As String
        Public id_clinical_service As String
        Public type As String
    End Structure

    Public Structure DEFAULT_REASONS
        Public id_content As String
        Public desccription As String
    End Structure

    Public Structure DEFAULT_DISCH_PROFILE
        Public ID_PROFILE_TEMPLATE As Int64
        Public PROFILE_NAME As String
    End Structure

    Function GET_DEFAULT_VERSIONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT drd.version
                                FROM alert_default.discharge_reason dr
                                JOIN alert_default.discharge_reason_mrk_vrs drmv ON drmv.id_discharge_reason = dr.id_discharge_reason
                                JOIN alert_default.disch_reas_dest drd ON drd.id_discharge_reason = dr.id_discharge_reason
                                                                   AND drd.id_market = drmv.id_market
                                                                   AND drd.version = drmv.version
                                JOIN alert_default.profile_disch_reason pdr ON pdr.id_discharge_reason = dr.id_discharge_reason
                                JOIN institution i ON i.id_market = drmv.id_market
                                WHERE dr.flg_available = 'Y'
                                AND i.id_institution = " & i_institution & "
                                AND drd.flg_active = 'A'
                                AND drd.id_software_param = " & i_software & "
                                AND pdr.flg_available = 'Y'
                                ORDER BY 1 ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)
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

    Function GET_DEFAULT_REASONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_version As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        'Modificar o output. Passar apenas ID_CONTENT e DESCRITIVO. O Resto será chamado diretamente pela função responsável por incluir Reason e Dest na BD
        Dim sql As String = "SELECT DISTINCT dr.id_content,
                                                alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dr.code_discharge_reason)
                                               
                                FROM alert_default.discharge_reason dr
                                JOIN alert_default.discharge_reason_mrk_vrs drmv ON drmv.id_discharge_reason = dr.id_discharge_reason
                                JOIN alert_default.disch_reas_dest drd ON drd.id_discharge_reason = dr.id_discharge_reason
                                                                   AND drd.id_market = drmv.id_market
                                                                   AND drd.version = drmv.version
                                JOIN alert_default.profile_disch_reason pdr ON pdr.id_discharge_reason = dr.id_discharge_reason
                                JOIN INSTITUTION I ON I.id_market=DRMV.ID_MARKET
                                WHERE dr.flg_available = 'Y'
                                AND I.id_institution=" & i_institution & "
                                AND drmv.version = '" & i_version & "'
                                AND drd.flg_active = 'A'
                                AND drd.id_software_param = " & i_software & "
                                AND pdr.flg_available = 'Y'
                                order by 2 asc"

        Dim cmd As New OracleCommand(sql, Connection.conn)
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

    Function GET_DEFAULT_DESTINATIONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_reason As String, ByVal i_version As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT drd.id_disch_reas_dest,
                                                nvl(d.id_content, dr.id_content) AS id_content,
                                                nvl2(d.id_content, alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", d.code_discharge_dest),
                                                     alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dr.code_discharge_reason)) AS description,
                                                nvl(dcs.id_content, -1) AS clinical_service,
                                                nvl2(d.id_content, 'D', 'REASON') AS TYPE

                                FROM alert_default.discharge_reason dr
                                JOIN alert_default.discharge_reason_mrk_vrs drmv ON drmv.id_discharge_reason = dr.id_discharge_reason
                                JOIN alert_default.disch_reas_dest drd ON drd.id_discharge_reason = dr.id_discharge_reason
                                                                   AND drd.id_market = drmv.id_market
                                                                   AND drd.version = drmv.version
                                JOIN alert_default.profile_disch_reason pdr ON pdr.id_discharge_reason = dr.id_discharge_reason
                                LEFT JOIN alert_default.discharge_dest d ON d.id_discharge_dest = drd.id_discharge_dest
                                                                     AND d.flg_available = 'Y'
                                JOIN institution i ON i.id_market = drmv.id_market
                                LEFT JOIN alert_default.discharge_dest_mrk_vrs dv ON dv.id_discharge_dest = d.id_discharge_dest
                                                                              AND dv.id_market = i.id_market
                                                                              AND dv.version = drd.version
                                LEFT JOIN alert_default.clinical_service dcs ON dcs.id_clinical_service = drd.id_clinical_service

                                WHERE dr.flg_available = 'Y'
                                AND i.id_institution = " & i_institution & "
                                AND drmv.version = '" & i_version & "'
                                AND drd.flg_active = 'A'
                                AND drd.id_software_param = " & i_software & "
                                AND pdr.flg_available = 'Y'
                                AND dr.id_content = '" & i_reason & "'
                                ORDER BY 3 ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)

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

    Function GET_DEFAULT_PROFILE_DISCH_REASON(ByVal id_disch_reason As String, ByRef o_profile_templates As OracleDataReader) As Boolean

        Dim sql As String = "SELECT pdr.id_profile_template,PT.INTERN_NAME_TEMPL
                                FROM alert_default.discharge_reason dr
                                JOIN alert_default.profile_disch_reason pdr ON pdr.id_discharge_reason = dr.id_discharge_reason
                                JOIN alert.profile_template pt ON pt.id_profile_template = pdr.id_profile_template
                                WHERE dr.flg_available = 'Y'
                                AND dr.id_content = '" & id_disch_reason & "'
                                AND pdr.flg_available = 'Y'
                                ORDER BY 1 ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)

        Try
            cmd.CommandType = CommandType.Text
            o_profile_templates = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception
            cmd.Dispose()
            Return False
        End Try

    End Function

    Function CHECK_REASON(ByVal i_id_reason As String) As Boolean

        Dim sql As String = "SELECT COUNT(*)
                                FROM alert.discharge_reason dr
                                WHERE dr.flg_available = 'Y'
                                AND dr.id_content = '" & i_id_reason & "'"

        Dim cmd As New OracleCommand(sql, Connection.conn)

        Dim dr As OracleDataReader

        cmd.CommandType = CommandType.Text
        dr = cmd.ExecuteReader()
        cmd.Dispose()

        Dim l_total_record As Integer = 0

        While dr.Read()
            l_total_record = dr.Item(0)
        End While

        If l_total_record > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

End Class

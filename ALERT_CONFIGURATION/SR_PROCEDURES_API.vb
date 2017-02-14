Imports Oracle.DataAccess.Client

Public Class SR_PROCEDURES_API

    Function GET_DEFAULT_VERSIONS(ByVal i_institution As Int64, ByVal i_flg_type As Integer, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = ""

        If i_flg_type = 0 Then

            sql = "SELECT DISTINCT dv.version
                        FROM alert_default.sr_intervention di
                        JOIN alert_default.sr_interv_codification dc ON dc.flg_coding = di.flg_coding
                        JOIN alert_default.sr_intervention_mrk_vrs dv ON dv.id_sr_intervention = di.id_sr_intervention
                        join alert_core_data.ab_institution i on i.id_ab_market=dv.id_market
                        WHERE di.flg_status = 'A'
                        AND i.id_ab_institution= " & i_institution & "
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

End Class

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

End Class

Imports Oracle.DataAccess.Client

Public Class Translation_Updates
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim cmd_update As OracleCommand = Connection.conn.CreateCommand()

        cmd_update.CommandType = CommandType.Text

        cmd_update.CommandText = "
                                 PROCEDURE PR_TESTE(@param2 varchar2 OUT)
                                  
                                  BEGIN
                                        select t.desc_lang_6
                                        INTO @param2
                                        FROM alert.exam e
                                        JOIN translation t ON t.code_translation = e.code_exam
                                        WHERE t.desc_lang_6 is not null
                                        and rownum=1;
                                    END;
                                    /"

        'Dim l_id_language As Int16 = 6

        Dim l_string_vb As String = ""


        'cmd_update.Parameters.Add("param1", l_id_language)

        'cmd_update.Parameters.Add("param2", SqlDbType.NText, ParameterDirection.Output)
        'cmd_update.Parameters("param2").Direction = ParameterDirection.Output
        'cmd_update.Parameters("param2").Size = 5000
        'outparam = cmd_update.Parameters.Add("param2", OracleDbType.NClob, ParameterDirection.Output)

        Dim outparam As New OracleParameter()
        outparam.OracleDbType = OracleDbType.NClob
        outparam.Size = 100
        outparam.ParameterName = "param2"
        outparam.Direction = System.Data.ParameterDirection.Output


        cmd_update.Parameters.Add(outparam)

        cmd_update.ExecuteNonQuery()

        MsgBox(outparam.Value.ToString)

        ' MsgBox(cmd_update.Parameters(("param2").ToString))


    End Sub
End Class
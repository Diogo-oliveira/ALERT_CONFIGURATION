Imports Oracle.DataAccess.Client

Public Class Translation_Updates
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim cmd_update As OracleCommand = Connection.conn.CreateCommand()

        cmd_update.CommandType = CommandType.Text

        cmd_update.CommandText = "DECLARE

                                        @param2 alert.exam.id_exam%type;

                                    BEGIN

                                        --SELECT pk_translation.get_translation(:param1, e.code_exam)
                                        select e.id_exam
                                        INTO :param2
                                        FROM alert.exam e
                                        JOIN translation t ON t.code_translation = e.code_exam
                                        WHERE e.id_exam = 10166;                                                                             

                                    END;"

        Dim l_id_language As Int16 = 6
        Dim l_string_vb As String = ""


        cmd_update.Parameters.Add("param1", l_id_language)

        cmd_update.Parameters.Add("param2", SqlDbType.BigInt)
        cmd_update.Parameters("param2").Direction = ParameterDirection.Output

        cmd_update.ExecuteNonQuery()

        MsgBox(cmd_update.Parameters(("param2").ToString))


    End Sub
End Class
Imports Oracle.DataAccess.Client

Public Class Translation_Updates
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim cmd_update As OracleCommand = Connection.conn.CreateCommand()

        cmd_update.CommandType = CommandType.Text

        cmd_update.CommandText = "CREATE OR REPLACE PACKAGE BODY teste AS

                                        PROCEDURE get_desc(o_desc IN OUT t_cursor) IS
    
                                            v_cursor t_cursor;
    
                                        BEGIN
    
                                            OPEN v_cursor FOR
        
                                                SELECT t.desc_lang_6
                                                FROM alert.exam e
                                                JOIN translation t ON t.code_translation = e.code_exam
                                                WHERE t.desc_lang_6 IS NOT NULL
                                                AND rownum = 1;
            
                                           o_desc  := v_cursor;               
    
                                        END get_desc;
                                    END teste;
                                    "

        'Dim l_id_language As Int16 = 6

        Dim l_string_vb As String = ""


        'cmd_update.Parameters.Add("param1", l_id_language)

        'cmd_update.Parameters.Add("param2", SqlDbType.NText, ParameterDirection.Output)
        'cmd_update.Parameters("param2").Direction = ParameterDirection.Output
        'cmd_update.Parameters("param2").Size = 5000
        'outparam = cmd_update.Parameters.Add("param2", OracleDbType.NClob, ParameterDirection.Output)

        'Dim outparam As New OracleParameter()

        'outparam.OracleDbType = OracleDbType.Varchar2
        'outparam.Size = 100
        'outparam.ParameterName = "param2"
        'outparam.Direction = System.Data.ParameterDirection.Output


        'cmd_update.Parameters.Add(outparam)

        'cmd_update.ExecuteNonQuery()

        'MsgBox(outparam.Value)

        ' MsgBox(cmd_update.Parameters(("param2").ToString))


        Dim myCMD As New OracleCommand()
        myCMD.Connection = Connection.conn
        myCMD.CommandText = "teste.get_desc_X"
        myCMD.CommandType = CommandType.StoredProcedure
        myCMD.Parameters.Add(New OracleParameter("o_desc", OracleDbType.RefCursor)).Direction = ParameterDirection.Output


        myCMD.ExecuteNonQuery()

        'Dim MyDA As New OracleDataAdapter(myCMD)

        'Dim myReader As OracleDataReader

        MsgBox(myCMD.Parameters("o_desc").Value)





    End Sub
End Class
Imports Oracle.DataAccess.Client
Public Class SYS_CONFIGS

    Dim oradb As String = ""
    Dim search_sys_config As String = ""

    Private Sub SYS_CONFIGS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("QC4 v2.6.3.8.2")
        ComboBox1.Items.Add("QC4 v2.6.3.10.1")
        ComboBox1.Items.Add("QC4 v2.6.3.15")
        ComboBox1.Items.Add("QC4 v2.6.5.0")
        ComboBox1.Items.Add("QC4 v2.6.5.0.6")
        ComboBox1.Items.Add("QC4 v2.6.5.2.2")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        If ComboBox1.SelectedIndex = 0 Then

            oradb = "Data Source=QC4V26382;User Id=alert_config;Password=qcteam"

        ElseIf ComboBox1.SelectedIndex = 1 Then

            oradb = "Data Source=QC4V263101;User Id=alert_config;Password=qcteam"

        ElseIf ComboBox1.SelectedIndex = 2 Then

            oradb = "Data Source=QC4V26315;User Id=alert_config;Password=qcteam"

        ElseIf ComboBox1.SelectedIndex = 3 Then

            oradb = "Data Source=QC4V265;User Id=alert_config;Password=qcteam"

        ElseIf ComboBox1.SelectedIndex = 4 Then

            oradb = "Data Source=QC4V26506;User Id=alert_config;Password=qcteam"

        ElseIf ComboBox1.SelectedIndex = 5 Then

            oradb = "Data Source=QC4V26522;User Id=alert_config;Password=qcteam"

        ElseIf ComboBox1.SelectedIndex = 6 Then

            oradb = "Data Source=QC264;User Id=alert_config;Password=qcteam"

        End If

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Definir ligação à BD
        If ComboBox1.SelectedIndex >= 0 Then

            If TextBox1.Text = "" Then

                MsgBox("No SYS_CONFIG inserted!", vbCritical)

            Else

                Dim conn As New OracleConnection(oradb)

                conn.Open()

                search_sys_config = TextBox1.Text

                'Definir o comando a ser executado (EXECUTAR UMA FUNÇAO)
                Dim sql As String = "Select s.id_sys_config, s.value, s.desc_sys_config, s.id_institution, s.id_software, s.id_market   from sys_config s where upper(s.id_sys_config) like upper('%" & TextBox1.Text & "%') order by 1 asc, 6 asc, 5 asc, 4 asc, 2 asc"
                Dim cmd As New OracleCommand(sql, conn)
                cmd.CommandType = CommandType.Text

                Dim dr As OracleDataReader = cmd.ExecuteReader()

                Dim Table As New DataTable
                Table.Load(cmd.ExecuteReader)
                DataGridView1.DataSource = Table

                DataGridView1.Columns(0).Width = 200
                DataGridView1.Columns(2).Width = 500

                DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable

                DataGridView1.Columns(0).ReadOnly = True
                DataGridView1.Columns(2).ReadOnly = True
                DataGridView1.Columns(3).ReadOnly = True
                DataGridView1.Columns(4).ReadOnly = True
                DataGridView1.Columns(5).ReadOnly = True

                conn.Close()

                conn.Dispose()

                dr.Dispose()

                cmd.Dispose()

            End If

        Else
            MsgBox("No Version selected!", vbCritical)
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try

            Dim sql As String = "Select s.id_sys_config, s.value, s.desc_sys_config, s.id_institution, s.id_software, s.id_market   from sys_config s where upper(s.id_sys_config) like upper('%" & search_sys_config & "%') order by 1 asc, 6 asc, 5 asc, 4 asc, 2 asc"
            Dim conn As New OracleConnection(oradb)

            conn.Open()
            Dim cmd As New OracleCommand(sql, conn)
            cmd.CommandType = CommandType.Text

            Dim dr As OracleDataReader = cmd.ExecuteReader()
            Dim i As Int16 = 0

            While (dr.Read())

                If dr.Item(1) <> DataGridView1(1, i).Value Then

                    'Dim sql_update As String = "UPDATE SYS_CONFIG S SET S.Value = " & DataGridView1(1, i).Value & " WHERE S.ID_SYS_CONFIG = " & DataGridView1(0, i).Value & " AND S.ID_INSTITUTION= " & DataGridView1(3, i).Value & " AND S.ID_SOFTWARE= " & DataGridView1(4, i).Value & " AND S.ID_MARKET= " & DataGridView1(5, i).Value

                    Dim cmd_update As OracleCommand = conn.CreateCommand()

                    cmd_update.CommandType = CommandType.Text

                    cmd_update.CommandText = "UPDATE SYS_CONFIG S SET S.Value = UPPER(:param1) WHERE S.ID_SYS_CONFIG = :keyValue1 AND S.ID_INSTITUTION= :keyValue2 AND S.ID_SOFTWARE= :keyValue3 AND S.ID_MARKET= :keyValue4"
                    cmd_update.Parameters.Add("param1", DataGridView1(1, i).Value)
                    cmd_update.Parameters.Add("keyValue1", DataGridView1(0, i).Value)
                    cmd_update.Parameters.Add("keyValue2", DataGridView1(3, i).Value)
                    cmd_update.Parameters.Add("keyValue3", DataGridView1(4, i).Value)
                    cmd_update.Parameters.Add("keyValue4", DataGridView1(5, i).Value)

                    cmd_update.ExecuteNonQuery()

                    MsgBox("Record(s) Updated!", vbOKOnly)

                End If

                i += 1

            End While

            conn.Close()

            conn.Dispose()

            dr.Dispose()

            cmd.Dispose()

        Catch ex As Exception
            MsgBox("ERROR!", vbCritical)
        End Try

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim form1 As New Form1

        form1.Show()

        Me.Close()

    End Sub
End Class
Imports Oracle.DataAccess.Client
Public Class SYS_CONFIGS

    Dim search_sys_config As String = ""

    Dim db_access_general As New General

    Dim g_a_softwares() As Integer
    Dim g_selected_software As Integer = 0

    Dim g_a_markets() As Integer
    Dim g_selected_market As Integer = 0

    Private Sub SYS_CONFIGS_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.BackColor = Color.FromArgb(215, 215, 180)

        Me.Location = New Point(Form_location.x_position, Form_location.y_position)

        DataGridView1.BackgroundColor = Color.FromArgb(195, 195, 165)

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_SOFT_INST(0, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING SOFTWARES!", vbCritical)

        Else

            Dim l_index_soft As Integer = 0
            ReDim g_a_softwares(0)

            While dr.Read()

                ComboBox3.Items.Add(dr.Item(1))
                ReDim Preserve g_a_softwares(l_index_soft)
                g_a_softwares(l_index_soft) = dr.Item(0)
                l_index_soft = l_index_soft + 1

            End While

        End If

        dr.Dispose()
        dr.Close()

        Dim dr_market As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_MARKETS(dr_market) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING MARKETS!", vbCritical)

        Else

            Dim l_index_market As Integer = 0
            ReDim g_a_markets(0)

            While dr_market.Read()

                ComboBox1.Items.Add(dr_market.Item(1))
                ReDim Preserve g_a_markets(l_index_market)
                g_a_markets(l_index_market) = dr_market.Item(0)
                l_index_market = l_index_market + 1

            End While

        End If

        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

        dr_market.Dispose()
        dr_market.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox1.Text = "" Then

            DataGridView1.Columns.Clear()

        Else

            search_sys_config = TextBox1.Text

            'Definir o comando a ser executado (EXECUTAR UMA FUNÇAO)
            Dim sql As String = "Select s.id_sys_config, s.value, s.desc_sys_config, s.id_institution, s.id_software, s.id_market   from sys_config s where upper(s.id_sys_config) like upper('%" & TextBox1.Text & "%') order by 1 asc, 6 asc, 5 asc, 4 asc, 2 asc"
            Dim cmd As New OracleCommand(sql, Connection.conn)
            cmd.CommandType = CommandType.Text

            Dim dr As OracleDataReader = cmd.ExecuteReader()

            Dim Table As New DataTable
            Table.Load(cmd.ExecuteReader)
            DataGridView1.DataSource = Table

            DataGridView1.Columns(0).Width = 350
            DataGridView1.Columns(1).Width = 180
            DataGridView1.Columns(2).Width = 685

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

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try

            Dim sql As String = "Select s.id_sys_config, s.value, s.desc_sys_config, s.id_institution, s.id_software, s.id_market   from sys_config s where upper(s.id_sys_config) like upper('%" & search_sys_config & "%') order by 1 asc, 6 asc, 5 asc, 4 asc, 2 asc"

            Dim cmd As New OracleCommand(sql, Connection.conn)
            cmd.CommandType = CommandType.Text

            Dim dr As OracleDataReader = cmd.ExecuteReader()
            Dim i As Int16 = 0

            Dim l_records_changed As Boolean = False

            While (dr.Read())

                If dr.Item(1) <> DataGridView1(1, i).Value Then

                    Dim cmd_update As OracleCommand = Connection.conn.CreateCommand()

                    cmd_update.CommandType = CommandType.Text

                    cmd_update.CommandText = "UPDATE SYS_CONFIG S SET S.Value = UPPER(:param1) WHERE S.ID_SYS_CONFIG = :keyValue1 AND S.ID_INSTITUTION= :keyValue2 AND S.ID_SOFTWARE= :keyValue3 AND S.ID_MARKET= :keyValue4"
                    cmd_update.Parameters.Add("param1", DataGridView1(1, i).Value)
                    cmd_update.Parameters.Add("keyValue1", DataGridView1(0, i).Value)
                    cmd_update.Parameters.Add("keyValue2", DataGridView1(3, i).Value)
                    cmd_update.Parameters.Add("keyValue3", DataGridView1(4, i).Value)
                    cmd_update.Parameters.Add("keyValue4", DataGridView1(5, i).Value)

                    cmd_update.ExecuteNonQuery()

                    l_records_changed = True

                End If

                i += 1

            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

            If l_records_changed = True Then

                MsgBox("Record(s) updated.", vbInformation)

            End If

        Catch ex As Exception
            MsgBox("ERROR UPDATING SYSCONFIG!", vbCritical)
        End Try

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim form1 As New Form1

        form1.Show()

        Me.Close()

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        g_selected_software = g_a_softwares(ComboBox3.SelectedIndex)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        If ComboBox4.Text = "" Then

            MsgBox("Please select a SYSCONFIG!")

        ElseIf TextBox3.Text = "" Then

            MsgBox("Please insert a VALUE for the sysconfig!")

        ElseIf TextBox4.Text = "" Then

            MsgBox("Please insert an INSTITUTION!")

        ElseIf ComboBox3.Text = "" Then

            MsgBox("Please select a SOFTWARE!")

        ElseIf ComboBox1.Text = "" Then

            MsgBox("Please select a MARKET!")

        Else

            If Not db_access_general.SET_SYSCONFIG(ComboBox4.Text, TextBox3.Text, TextBox4.Text, g_selected_software, g_selected_market) Then

                MsgBox("ERROR INSERTING RECORD.", vbCritical)

            Else

                MsgBox("Record inserted.", vbInformation)

                search_sys_config = ComboBox4.Text
                TextBox1.Text = ComboBox4.Text

                'Definir o comando a ser executado (EXECUTAR UMA FUNÇAO)
                Dim sql As String = "Select s.id_sys_config, s.value, s.desc_sys_config, s.id_institution, s.id_software, s.id_market   from sys_config s where upper(s.id_sys_config) like upper('%" & TextBox1.Text & "%') order by 1 asc, 6 asc, 5 asc, 4 asc, 2 asc"
                Dim cmd As New OracleCommand(sql, Connection.conn)
                cmd.CommandType = CommandType.Text

                Dim dr As OracleDataReader = cmd.ExecuteReader()

                Dim Table As New DataTable
                Table.Load(cmd.ExecuteReader)
                DataGridView1.DataSource = Table

                DataGridView1.Columns(0).Width = 350
                DataGridView1.Columns(1).Width = 180
                DataGridView1.Columns(2).Width = 685

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

                dr.Dispose()
                dr.Close()
                cmd.Dispose()

            End If

        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim result As Integer = 0
        result = MsgBox("Selected SYSCONFIG will be deleted. Confirm?", MessageBoxButtons.YesNo)

        If (result = DialogResult.Yes) Then

            Dim sql As String = "DELETE 
                                    FROM sys_config sy
                                    WHERE (sy.id_sys_config, sy.value, sy.desc_sys_config, sy.id_institution, sy.id_software, sy.id_market) IN
                                          (SELECT tt.id_sys_config, tt.value, tt.desc_sys_config, tt.id_institution, tt.id_software, tt.id_market
                                           FROM (SELECT t.*, rownum rn
                                                 FROM (SELECT s.id_sys_config, s.value, s.desc_sys_config, s.id_institution, s.id_software, s.id_market
                                                       FROM sys_config s
                                                       WHERE upper(s.id_sys_config) LIKE upper('%" & search_sys_config & "%')
                                                       ORDER BY 1 ASC, 6 ASC, 5 ASC, 4 ASC, 2 ASC) t) tt
                                           WHERE rn = " & DataGridView1.CurrentRow.Index + 1 & ")"


            Dim cmd_delete_sysconfig As New OracleCommand(sql, Connection.conn)

            Try

                cmd_delete_sysconfig.CommandType = CommandType.Text
                cmd_delete_sysconfig.ExecuteNonQuery()
                cmd_delete_sysconfig.Dispose()

                MsgBox("Record deleted.", vbInformation)

                TextBox1.Text = search_sys_config

                'Definir o comando a ser executado (EXECUTAR UMA FUNÇAO)
                Dim sql_updated As String = "Select s.id_sys_config, s.value, s.desc_sys_config, s.id_institution, s.id_software, s.id_market   from sys_config s where upper(s.id_sys_config) like upper('%" & TextBox1.Text & "%') order by 1 asc, 6 asc, 5 asc, 4 asc, 2 asc"
                Dim cmd As New OracleCommand(sql_updated, Connection.conn)
                cmd.CommandType = CommandType.Text

                Dim dr As OracleDataReader = cmd.ExecuteReader()

                Dim Table As New DataTable
                Table.Load(cmd.ExecuteReader)
                DataGridView1.DataSource = Table

                DataGridView1.Columns(0).Width = 350
                DataGridView1.Columns(1).Width = 180
                DataGridView1.Columns(2).Width = 685

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

                dr.Dispose()
                dr.Close()
                cmd.Dispose()

            Catch ex As Exception

                MsgBox("ERROR DELETING SYSCONFIG", vbCritical)
                cmd_delete_sysconfig.Dispose()
            End Try

        End If

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        g_selected_market = g_a_markets(ComboBox1.SelectedIndex)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ComboBox4_Click(sender As Object, e As EventArgs) Handles ComboBox4.Click

        Dim sql As String = "Select DISTINCT s.id_sys_config  from sys_config s where upper(s.id_sys_config) like upper('%" & TextBox2.Text & "%') order by 1 asc"
        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        ComboBox4.Items.Clear()
        ComboBox4.SelectedText = ""
        While dr.Read()

            ComboBox4.Items.Add(dr.Item(0))

        End While

        dr.Dispose()
        dr.Close()

    End Sub

End Class
Imports Oracle.DataAccess.Client
Public Class SYS_CONFIGS

    Dim oradb As String
    Dim conn As New OracleConnection

    Dim search_sys_config As String = ""

    Private Sub SYS_CONFIGS_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            'Estabelecer ligação à BD

            conn.Open()

        Catch ex As Exception

            MsgBox("ERROR CONNECTING TO DATA BASE!", vbCritical)

        End Try

    End Sub

    Public Sub New(ByVal i_oradb As String)

        InitializeComponent()
        oradb = i_oradb
        conn = New OracleConnection(oradb)

    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox1.Text = "" Then

            DataGridView1.Columns.Clear()

        Else

            search_sys_config = TextBox1.Text

            'Definir o comando a ser executado (EXECUTAR UMA FUNÇAO)
            Dim sql As String = "Select s.id_sys_config, s.value, s.desc_sys_config, s.id_institution, s.id_software, s.id_market   from sys_config s where upper(s.id_sys_config) like upper('%" & TextBox1.Text & "%') order by 1 asc, 6 asc, 5 asc, 4 asc, 2 asc"
            Dim cmd As New OracleCommand(sql, conn)
            cmd.CommandType = CommandType.Text

            Dim dr As OracleDataReader = cmd.ExecuteReader()

            Dim Table As New DataTable
            Table.Load(cmd.ExecuteReader)
            DataGridView1.DataSource = Table

            DataGridView1.Columns(0).Width = 350
            DataGridView1.Columns(1).Width = 180
            DataGridView1.Columns(2).Width = 670

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

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        Catch ex As Exception
            MsgBox("ERROR!", vbCritical)
        End Try

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim form1 As New Form1
        form1.g_oradb = oradb
        form1.Show()

        Me.Close()

    End Sub

End Class
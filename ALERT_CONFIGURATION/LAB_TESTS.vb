Imports Oracle.DataAccess.Client
Public Class LAB_TESTS

    Dim db_access_general As New General
    Dim db_labs As New LABS_API
    Dim oradb As String = "Data Source=QC4V26522;User Id=alert_config;Password=qcteam"
    Dim l_selected_soft As Int16 = -1

    Private Sub LAB_TESTS_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dr As OracleDataReader = db_access_general.GET_ALL_INSTITUTIONS(oradb)

        Dim i As Integer = 0

        While dr.Read()

            ComboBox1.Items.Add(dr.Item(0))

        End While

        dr.Dispose()

        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Dim form1 As New Form1

        Me.Enabled = False

        Me.Dispose()

        form1.Show()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox1.Text <> "" Then

            ComboBox1.Text = db_access_general.GET_INSTITUTION(TextBox1.Text, oradb)

            ComboBox2.Items.Clear()
            ComboBox2.Text = ""


            Dim dr As OracleDataReader = db_access_general.GET_SOFT_INST(TextBox1.Text, oradb)

            Dim i As Integer = 0

            While dr.Read()

                ComboBox2.Items.Add(dr.Item(1))

            End While

        End If

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        l_selected_soft = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, TextBox1.Text, oradb)

        '1 - Fill Version combobox

        ComboBox3.Items.Clear()
        ComboBox3.Text = ""

        Cursor = Cursors.WaitCursor

        Try

            Dim dr_def_versions As OracleDataReader = db_labs.GET_DEFAULT_VERSIONS(TextBox1.Text, l_selected_soft, oradb)

            While dr_def_versions.Read()

                ComboBox3.Items.Add(dr_def_versions.Item(0))

            End While

        Catch ex As Exception

            MsgBox("ERROR LOADING DEFAULT VERSIONS -  ComboBox2_SelectedIndexChanged", MsgBoxStyle.Critical)

        End Try

        Cursor = Cursors.Arrow

    End Sub
End Class
Imports Oracle.DataAccess.Client

Public Class Form1

    Dim debug As New DEBUGGER

    Public g_oradb As String = ""

    Dim g_a_databases() As String

    Dim g_height_initial As Integer = 430
    Dim g_height_extended As Integer = 500

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Connection.conn.State = 0 Then

            Me.Text = "ALERT(R) ENVIRONMENTS CONFIGURATION"

            Me.Height = g_height_initial

        Else

            Me.Text = "ALERT(R) ENVIRONMENTS CONFIGURATION  ::::  Connected to " & Connection.db

            Me.Height = g_height_extended

        End If

        Me.BackColor = Color.FromArgb(215, 215, 180)

        Me.Location = New Point(Form_location.x_position, Form_location.y_position)

        TextBox1.Text = "sys"
        TextBox2.Text = "dbasysqc"
        ComboBox1.Text = "qc4v2701"

        If Connection.conn.State = 0 Then

            Button1.Visible = False
            Button2.Visible = False
            Button3.Visible = False
            Button4.Visible = False
            Button5.Visible = False
            Button6.Visible = False
            Button9.Visible = False
            Button10.Visible = False
            Button11.Visible = False
            Button12.Visible = False
            Button13.Visible = False
            Button14.Visible = False
            Button15.Visible = False
            Button16.Visible = False

            Dim db_list As New OracleDataSourceEnumerator()

            Dim l_index As Integer = 0
            Try

                ReDim g_a_databases(0)

                While True

                    ReDim Preserve g_a_databases(l_index)
                    g_a_databases(l_index) = db_list.GetDataSources(l_index).Item(0)

                    l_index = l_index + 1

                End While

            Catch ex As Exception

            End Try

            ''Forma de remover o último valor que gerou a exceção anterior
            ReDim Preserve g_a_databases(l_index - 1)

            Array.Sort(g_a_databases)

            For i As Integer = 0 To g_a_databases.Count() - 2

                ComboBox1.Items.Add(g_a_databases(i))

            Next

        Else

            TextBox1.Visible = False
            TextBox2.Visible = False
            ComboBox1.Visible = False
            Label1.Visible = False
            Label2.Visible = False
            Label3.Visible = False
            Button7.Visible = False
            Button8.Visible = False

            Button1.Visible = True
            Button2.Visible = True
            Button3.Visible = True
            Button4.Visible = True
            Button5.Visible = True
            Button6.Visible = True
            Button9.Visible = True
            Button10.Visible = True
            Button11.Visible = True
            Button12.Visible = True
            Button13.Visible = True
            Button14.Visible = True
            Button15.Visible = True
            Button16.Visible = True

        End If

        Me.CenterToScreen()

        'Criar diretório e ficheiro para o debugger       
        debug.CREATE_DEBUG_FOLDER()
        debug.CLEAN_DEBUG()
        debug.CREATE_DEBUG_FILE()



    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Connection.conn.Dispose()
        Connection.conn.Close()

        Application.Exit()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Dim FORM_ADD_IMAGING As New INSERT_IMAGING_EXAMS

        Me.Hide()

        FORM_ADD_IMAGING.ShowDialog()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Dim Form_SYS_CONFIG As New SYS_CONFIGS

        Me.Hide()

        Form_SYS_CONFIG.ShowDialog()

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click


    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Dim form_insert_other As New INSERT_OTHER_EXAM

        Me.Hide()

        form_insert_other.ShowDialog()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Dim FORM_LAB_TESTS As New LAB_TESTS

        Me.Hide()

        FORM_LAB_TESTS.ShowDialog()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Dim FORM_PROCEDURES As New Procedures

        Me.Hide()

        FORM_PROCEDURES.ShowDialog()

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        If TextBox1.Text = "" Then

            MsgBox("Please enter a Username")

        ElseIf TextBox2.Text = "" Then

            MsgBox("Please enter a Password")

        ElseIf ComboBox1.Text = "" Then

            MsgBox("Please select a Database")

        Else

            If TextBox1.Text.ToUpper = "SYS" Then

                g_oradb = "Data Source=" & ComboBox1.Text & ";User Id=" & TextBox1.Text & ";Password=" & TextBox2.Text & ";DBA Privilege=SYSDBA"

            Else

                g_oradb = "Data Source=" & ComboBox1.Text & ";User Id=" & TextBox1.Text & ";Password=" & TextBox2.Text & ""

            End If


            Connection.conn = New OracleConnection(g_oradb)

            Try

                Connection.conn.Open()

                Me.Height = g_height_extended

                TextBox1.Visible = False
                TextBox2.Visible = False
                ComboBox1.Visible = False
                Label1.Visible = False
                Label2.Visible = False
                Label3.Visible = False
                Button7.Visible = False
                Button8.Visible = False

                Button1.Visible = True
                Button2.Visible = True
                Button3.Visible = True
                Button4.Visible = True
                Button5.Visible = True
                Button6.Visible = True
                Button9.Visible = True
                Button10.Visible = True
                Button11.Visible = True
                Button12.Visible = True
                Button13.Visible = True
                Button14.Visible = True
                Button15.Visible = True
                Button16.Visible = True

                Connection.db = ComboBox1.Text

                Me.Text = "ALERT(R) ENVIRONMENTS CONFIGURATION  ::::  Connected to " & Connection.db

            Catch ex As Exception

                MsgBox("Error connecting to Database. Please check credentials.", vbCritical)

            End Try

        End If

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        Connection.conn.Dispose()
        Connection.conn.Close()

        Connection.db = ""

        Me.Text = "ALERT(R) ENVIRONMENTS CONFIGURATION"

        TextBox1.Visible = True
        TextBox2.Visible = True
        ComboBox1.Visible = True
        Label1.Visible = True
        Label2.Visible = True
        Label3.Visible = True
        Button7.Visible = True
        Button8.Visible = True

        Button1.Visible = False
        Button2.Visible = False
        Button3.Visible = False
        Button4.Visible = False
        Button5.Visible = False
        Button6.Visible = False
        Button9.Visible = False
        Button10.Visible = False
        Button11.Visible = False
        Button12.Visible = False
        Button13.Visible = False
        Button14.Visible = False
        Button15.Visible = False
        Button16.Visible = False

        Me.Height = g_height_initial

        Dim db_list As New OracleDataSourceEnumerator()

        Dim l_index As Integer = 0
        Try

            ReDim g_a_databases(0)

            While True

                ReDim Preserve g_a_databases(l_index)
                g_a_databases(l_index) = db_list.GetDataSources(l_index).Item(0)

                l_index = l_index + 1

            End While

        Catch ex As Exception

        End Try

        ''Forma de remover o último valor que gerou a exceção anterior
        ReDim Preserve g_a_databases(l_index - 1)

        Array.Sort(g_a_databases)

        For i As Integer = 0 To g_a_databases.Count() - 2

            ComboBox1.Items.Add(g_a_databases(i))

        Next

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Application.Exit()

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Dim FORM_SR_PROCEDURES As New SR_Procedures()

        Me.Hide()

        FORM_SR_PROCEDURES.Show()

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Dim FORM_SUPPLIES As New Supplies

        Me.Hide()

        FORM_SUPPLIES.ShowDialog()

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Dim FORM_TRANSLATION As New Translation_Updates

        Me.Hide()

        FORM_TRANSLATION.ShowDialog()

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

        ''MsgBox("AWAITING Development!", vbInformation)

        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Dim FORM_DISCHARGE As New DISCHARGE

        Me.Hide()

        FORM_DISCHARGE.ShowDialog()

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        MsgBox("AWAITING Development!", vbInformation)

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

        MsgBox("AWAITING Development!", vbInformation)

    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click

        MsgBox("AWAITING Development!", vbInformation)

    End Sub
End Class

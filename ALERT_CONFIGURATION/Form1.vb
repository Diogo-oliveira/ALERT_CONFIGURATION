﻿Imports Oracle.DataAccess.Client

Public Class Form1

    Public g_oradb As String = ""
    Dim conn As New OracleConnection

    Dim g_a_databases() As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If g_oradb = "" Then

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

            conn = New OracleConnection(g_oradb)
            conn.Open()

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

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        conn.Dispose()
        conn.Close()

        Application.Exit()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim FORM_ADD_IMAGING As New INSERT_IMAGING_EXAMS(g_oradb)

        Me.Hide()

        FORM_ADD_IMAGING.ShowDialog()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim Form_SYS_CONFIG As New SYS_CONFIGS(g_oradb)

        Me.Hide()

        Form_SYS_CONFIG.ShowDialog()

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click


    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim form_insert_other As New INSERT_OTHER_EXAM(g_oradb)

        Me.Hide()

        form_insert_other.ShowDialog()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim FORM_LAB_TESTS As New LAB_TESTS(g_oradb)

        Me.Hide()

        FORM_LAB_TESTS.ShowDialog()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim FORM_PROCEDURES As New Procedures(g_oradb)

        conn.Close()
        conn.Dispose()

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

            If ComboBox1.Text.ToUpper = "SYS" Then

                g_oradb = "Data Source=" & ComboBox1.Text & ";User Id=" & TextBox1.Text & ";Password=" & TextBox2.Text & "; DBA Privilege=SYSDBA"

            Else

                g_oradb = "Data Source=" & ComboBox1.Text & ";User Id=" & TextBox1.Text & ";Password=" & TextBox2.Text & ""

            End If

            Dim conn As New OracleConnection(g_oradb)
            Try

                conn.Open()

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

            Catch ex As Exception

                MsgBox("Error connecting to Database. Please check credentials.", vbCritical)

            End Try

        End If

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        conn.Dispose()
        conn.Close()

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

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Application.Exit()

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        Dim FOMR_SR_PROCEDURES As New SR_Procedures(g_oradb)

        Me.Hide()

        FOMR_SR_PROCEDURES.ShowDialog()


    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        Dim FORM_SUPPLIES As New Supplies(g_oradb)

        Me.Hide()

        FORM_SUPPLIES.ShowDialog()

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        MsgBox("Waiting Development!", vbInformation)

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

        MsgBox("Waiting Development!", vbInformation)

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        MsgBox("Waiting Development!", vbInformation)

    End Sub
End Class

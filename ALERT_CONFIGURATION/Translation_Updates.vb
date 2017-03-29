Imports Oracle.DataAccess.Client

Public Class Translation_Updates

    Dim translation As New Translation_API
    Dim db_access_general As New General

    '#######################################
    '#   Definição das variáves 
    '#   de área a atualizar
    Dim g_exams As String = "EXAMS"

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Cursor = Cursors.WaitCursor

        Dim l_column_width As Int64 = DataGridView1.Size.Width - 2 'Tentar evitar o scroll

        If (TextBox1.Text <> "" And ComboBox1.Text <> "") Then

            If (ComboBox2.Text = g_exams) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_EXAMS(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_exams & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                        If Not translation.DELETE_TEMP_TABLE() Then

                            MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                        End If

                    End If

                End If

            Else

                MsgBox("Please select an area to be updated.", vbInformation)

            End If

        Else

            MsgBox("Please select an institution.", vbInformation)

        End If
        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Cursor = Cursors.WaitCursor

        If TextBox1.Text <> "" Then

            ComboBox1.Text = db_access_general.GET_INSTITUTION(TextBox1.Text)

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Translation_Updates_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "TRANSLATION UPDATE  ::::  Connected to " & Connection.db

        Me.BackColor = Color.FromArgb(215, 215, 180)

        Me.Location = New Point(Form_location.x_position, Form_location.y_position)

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_ALL_INSTITUTIONS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING ALL INSTITUTIONS!", vbCritical)

        Else

            While dr.Read()

                ComboBox1.Items.Add(dr.Item(0))

            End While

        End If

        dr.Dispose()
        dr.Close()

        ComboBox2.Items.Add(g_exams)

        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Dim form1 As New Form1

        form1.Show()

        Me.Close()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        DataGridView1.DataSource = ""

    End Sub
End Class
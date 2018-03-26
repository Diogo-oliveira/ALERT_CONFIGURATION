Imports Oracle.DataAccess.Client
Public Class MEDICATION

    Dim db_access_general As New General
    Dim medication As New Medication_API

    Dim g_selected_soft As Int16 = -1
    Dim g_selected_index As Int64 = -1
    Dim g_id_product_supplier As String = ""
    Dim g_a_list_products() As String

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Cursor = Cursors.WaitCursor

        'Limpar arrays
        g_selected_soft = -1
        g_selected_index = -1
        g_id_product_supplier = ""
        ReDim g_a_list_products(0)

        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox6.Text = ""
        ComboBox7.Text = ""
        ComboBox8.Text = ""
        ComboBox9.Text = ""
        ComboBox10.Text = ""
        ComboBox11.Text = ""
        ComboBox12.Text = ""
        TextBox3.Text = ""

        If TextBox1.Text <> "" Then

            ComboBox1.Text = db_access_general.GET_INSTITUTION(TextBox1.Text)

            ComboBox2.Items.Clear()
            ComboBox2.Text = ""

            Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not db_access_general.GET_SOFT_INST(TextBox1.Text, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                MsgBox("ERROR GETTING SOFTWARES!", vbCritical)
            Else
                While dr.Read()
                    ComboBox2.Items.Add(dr.Item(1))
                End While
            End If

            dr.Dispose()
            dr.Close()

            g_id_product_supplier = medication.GET_PRODUCT_SUPPLIER(TextBox1.Text)

        End If

        Cursor = Cursors.Arrow
    End Sub

    Private Sub MEDICATION_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "MEDICATION  ::::  Connected to " & Connection.db

        Me.BackColor = Color.FromArgb(215, 215, 180)

        Me.Location = New Point(Form_location.x_position, Form_location.y_position)

        DataGridView1.BackgroundColor = Color.FromArgb(195, 195, 165)

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_ALL_INSTITUTIONS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING ALL INSTITUTIONS!")

        Else

            While dr.Read()
                ComboBox1.Items.Add(dr.Item(0))
            End While

        End If

        dr.Dispose()
        dr.Close()

        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

        ComboBox3.Items.Add("Y")
        ComboBox3.Items.Add("N")

        ComboBox4.Items.Add("IV")
        ComboBox4.Items.Add("Non-IV")

        ComboBox5.Items.Add("1 - External Prescription")
        ComboBox5.Items.Add("2 - Administer Here")
        ComboBox5.Items.Add("3 - Home Medication")

        ComboBox6.Items.Add("0")
        ComboBox6.Items.Add("1")
        ComboBox6.Items.Add("2")
        ComboBox6.Items.Add("3")
        ComboBox6.Items.Add("4")

        ComboBox7.Items.Add("Y")
        ComboBox7.Items.Add("N")

        ComboBox8.Items.Add("Y")
        ComboBox8.Items.Add("N")

        ComboBox9.Items.Add("Y")
        ComboBox9.Items.Add("N")

        ComboBox10.Items.Add("Y")
        ComboBox10.Items.Add("N")

        ComboBox11.Items.Add("Y")
        ComboBox11.Items.Add("N")

        ComboBox12.Items.Add("Y")
        ComboBox12.Items.Add("N")
        ComboBox12.Items.Add("U")

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        g_selected_soft = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, TextBox1.Text)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Cursor = Cursors.WaitCursor

        Dim l_column_width As Int64 = DataGridView1.Size.Width - 122 'Tentar evitar o scroll

        Dim dr As OracleDataReader
        If (ComboBox1.Text <> "") Then
            If (TextBox2.Text <> "") Then
                If Not medication.GET_LIST_PRODUCTS(TextBox1.Text, g_id_product_supplier, TextBox2.Text, dr) Then

                    MsgBox("Error getting list of products!", vbCritical)

                Else
                    DataGridView1.Columns.Clear()

                    Dim Table As New DataTable

                    Table.Load(dr)
                    DataGridView1.DataSource = Table

                    DataGridView1.Columns(0).Width = l_column_width
                    DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                    Dim l_dimension_list_products As Int64 = 0

                    For Each row As DataRow In Table.Rows

                        ReDim Preserve g_a_list_products(l_dimension_list_products)
                        g_a_list_products(l_dimension_list_products) = row.Item("ID_PRODUCT")
                        l_dimension_list_products = l_dimension_list_products + 1
                    Next
                    dr.Close()
                End If
            Else
                MsgBox("Please write a medication description! ", vbCritical)
            End If
        Else
            MsgBox("Please select an institution! ", vbCritical)
        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub dataGridView1_CellStateChanged(ByVal sender As Object,
    ByVal e As DataGridViewCellStateChangedEventArgs) Handles DataGridView1.CellStateChanged

        If g_selected_index <> e.Cell.RowIndex Then
            g_selected_index = e.Cell.RowIndex
        End If

        If g_selected_index > -1 Then
            Dim dr As OracleDataReader
            If Not medication.GET_PRODUCT_OPTIONS(TextBox1.Text, g_selected_soft, g_a_list_products(g_selected_index), g_id_product_supplier, ComboBox5.SelectedIndex + 1, dr) Then

                MsgBox("Error getting product options!", vbCritical)

            Else

                While dr.Read

                    Try
                        ComboBox3.Text = dr.Item(1)
                    Catch ex As Exception
                        ComboBox3.Text = ""
                    End Try
                    Try
                        ComboBox4.Text = dr.Item(3)
                    Catch ex As Exception
                        ComboBox4.Text = ""
                    End Try
                    Try
                        ComboBox6.Text = dr.Item(2)
                    Catch ex As Exception
                        ComboBox6.Text = ""
                    End Try
                    Try
                        ComboBox7.Text = dr.Item(4)
                    Catch ex As Exception
                        ComboBox7.Text = ""
                    End Try
                    Try
                        ComboBox8.Text = dr.Item(5)
                    Catch ex As Exception
                        ComboBox8.Text = ""
                    End Try
                    Try
                        ComboBox9.Text = dr.Item(6)
                    Catch ex As Exception
                        ComboBox9.Text = ""
                    End Try
                    Try
                        ComboBox10.Text = dr.Item(7)
                    Catch ex As Exception
                        ComboBox10.Text = ""
                    End Try
                    Try
                        ComboBox11.Text = dr.Item(8)
                    Catch ex As Exception
                        ComboBox11.Text = ""
                    End Try
                    Try
                        ComboBox12.Text = dr.Item(9)
                    Catch ex As Exception
                        ComboBox12.Text = ""
                    End Try
                    Try
                        TextBox3.Text = dr.Item(10)
                    Catch ex As Exception
                        TextBox3.Text = ""
                    End Try


                End While

                dr.Close()
            End If


        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim l_med_type As Int16

        If ComboBox4.Text = "IV" Then
            l_med_type = 2
        Else
            l_med_type = 1
        End If

        If Not medication.SET_PARAMETERS(TextBox1.Text, g_a_list_products(g_selected_index), g_id_product_supplier, ComboBox3.Text, ComboBox5.SelectedIndex + 1, ComboBox6.Text,
                                         l_med_type, ComboBox7.Text, ComboBox8.Text, ComboBox9.Text, ComboBox10.Text, ComboBox11.Text, ComboBox12.Text, TextBox3.Text) Then

            MsgBox("Error updating parameters!", vbCritical)
        End If

    End Sub
End Class
Imports Oracle.DataAccess.Client
Public Class MEDICATION

    Dim db_access_general As New General
    Dim medication As New Medication_API

    Dim g_selected_soft As Int16 = -1
    Dim g_selected_index As Int64 = -1
    Dim g_id_product_supplier As String = ""
    Dim g_a_list_products() As String
    Dim g_a_product_routes() As String
    Dim g_a_market_routes() As String
    Dim g_a_market_um() As Int64
    Dim g_a_product_um() As Int64

    Dim g_selection_aux As Boolean = False

    Function RESET_PRODUCT_PARAMETERS()

        DataGridView1.Columns.Clear()
        g_selected_index = -1
        ReDim g_a_list_products(0)
        ReDim g_a_product_routes(0)

        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1
        ComboBox6.SelectedIndex = -1
        ComboBox7.SelectedIndex = -1
        ComboBox8.SelectedIndex = -1
        ComboBox9.SelectedIndex = -1
        ComboBox10.SelectedIndex = -1
        ComboBox11.SelectedIndex = -1
        ComboBox12.SelectedIndex = -1

        TextBox3.Text = ""

        CheckedListBox2.Items.Clear()

        If CheckedListBox1.Items.Count() > 0 Then
            For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, False)
            Next
        End If

        Return True

    End Function

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

        Dim dr_routes_market As OracleDataReader
        If Not medication.GET_MARKET_ROUTES(TextBox1.Text, g_id_product_supplier, dr_routes_market) Then
            MsgBox("Error getting MARKET routes!", vbCritical)
        Else
            CheckedListBox1.Items.Clear()
            ReDim g_a_market_routes(0)
            Dim i As Integer = 0
            Dim l_aux As String = ""
            While dr_routes_market.Read
                CheckedListBox1.Items.Add(dr_routes_market.Item(1))

                ReDim Preserve g_a_market_routes(i)
                g_a_market_routes(i) = dr_routes_market(0)
                i = i + 1
            End While
        End If
        dr_routes_market.Close()

        DataGridView1.Columns.Clear()

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

        ComboBox13.Items.Add("Regular dose")
        ComboBox13.Items.Add("Dose to be dispensed")
        ComboBox13.SelectedIndex = 0

        ComboBox14.Items.Add("Regular dose")
        ComboBox14.Items.Add("Dose to be dispensed")
        ComboBox14.SelectedIndex = 0
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        g_selected_soft = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, TextBox1.Text)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Cursor = Cursors.WaitCursor

        Dim l_column_width As Int64 = DataGridView1.Size.Width - 120 'Tentar evitar o scroll

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

                    DataGridView1.ClearSelection()

                    CheckedListBox2.Items.Clear()
                    ReDim g_a_product_routes(0)

                    Dim dr_routes As OracleDataReader
                    Try
                        DataGridView1.Rows(0).Selected = True
                        If Not medication.GET_PRODUCT_ROUTES(TextBox1.Text, g_a_list_products(g_selected_index), g_id_product_supplier, dr_routes) Then
                            MsgBox("Error getting product routes!", vbCritical)
                        Else
                            CheckedListBox2.Items.Clear()
                            ReDim g_a_product_routes(0)
                            Dim i As Integer = 0
                            While dr_routes.Read
                                CheckedListBox2.Items.Add(dr_routes.Item(1))
                                If dr_routes.Item(2) = "Y" Then
                                    CheckedListBox2.SetItemChecked(i, True)
                                End If
                                ReDim Preserve g_a_product_routes(i)
                                g_a_product_routes(i) = dr_routes(0)
                                i = i + 1
                            End While
                        End If
                        dr_routes.Close()

                    Catch ex As Exception
                        MsgBox("No results found.", vbInformation)
                        g_selected_index = -1
                        If Not RESET_PRODUCT_PARAMETERS() Then
                            MsgBox("Error reseting product parameters.", vbCritical)
                        End If
                    End Try

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

            Dim dr_routes As OracleDataReader
            If Not medication.GET_PRODUCT_ROUTES(TextBox1.Text, g_a_list_products(g_selected_index), g_id_product_supplier, dr_routes) Then
                MsgBox("Error getting product routes!", vbCritical)
            Else
                CheckedListBox2.Items.Clear()
                ReDim g_a_product_routes(0)
                Dim i As Integer = 0
                While dr_routes.Read
                    CheckedListBox2.Items.Add(dr_routes.Item(1))
                    If dr_routes.Item(2) = "Y" Then
                        CheckedListBox2.SetItemChecked(i, True)
                    End If
                    ReDim Preserve g_a_product_routes(i)
                    g_a_product_routes(i) = dr_routes(0)
                    i = i + 1
                End While
            End If
            dr_routes.Close()

            If CheckedListBox1.Items.Count() > 0 Then
                For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                    CheckedListBox1.SetItemChecked(i, False)
                Next
            End If

            If g_a_market_routes.Count > 0 And g_a_product_routes.Count > 0 Then
                For i As Integer = 0 To g_a_market_routes.Count - 1
                    For j As Integer = 0 To g_a_product_routes.Count - 1
                        If g_a_market_routes(i) = g_a_product_routes(j) Then
                            CheckedListBox1.SetItemChecked(i, True)
                        End If
                    Next
                Next
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

        If g_selected_index > -1 Then
            If Not medication.SET_PARAMETERS(TextBox1.Text, g_a_list_products(g_selected_index), g_id_product_supplier, ComboBox3.Text, ComboBox5.SelectedIndex + 1, ComboBox6.Text,
                                         l_med_type, ComboBox7.Text, ComboBox8.Text, ComboBox9.Text, ComboBox10.Text, ComboBox11.Text, ComboBox12.Text, TextBox3.Text) Then
                MsgBox("Error updating parameters!", vbCritical)
            Else
                MsgBox("Record updated!", vbInformation)
            End If
        Else
            MsgBox("Please select a product!", vbCritical)
        End If

    End Sub


    Private Sub CheckedListBox2_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles CheckedListBox2.ItemCheck
        If e.NewValue = CheckState.Checked Then
            For i As Integer = 0 To Me.CheckedListBox2.Items.Count - 1 Step 1
                If i <> e.Index Then
                    Me.CheckedListBox2.SetItemChecked(i, False)
                End If
            Next i
        End If
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Me.Enabled = False
        Me.Dispose()
        Form1.Show()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        Dim l_a_product_routes() As String
        Dim l_index As Integer = 0

        If g_selected_index > -1 Then
            If CheckedListBox1.CheckedIndices.Count > 0 Then
                For Each indexChecked In CheckedListBox1.CheckedIndices
                    ReDim Preserve l_a_product_routes(l_index)
                    l_a_product_routes(l_index) = g_a_market_routes(indexChecked)
                    l_index = l_index + 1
                Next
                If Not medication.SET_PRODUCT_ROUTES(TextBox1.Text, g_a_list_products(g_selected_index), g_id_product_supplier, l_a_product_routes) Then
                    MsgBox("Error inserting procut routes!", vbCritical)
                Else
                    MsgBox("Record updated!", vbInformation)
                End If

                Dim dr_routes As OracleDataReader
                If Not medication.GET_PRODUCT_ROUTES(TextBox1.Text, g_a_list_products(g_selected_index), g_id_product_supplier, dr_routes) Then
                    MsgBox("Error getting product routes!", vbCritical)
                Else
                    CheckedListBox2.Items.Clear()
                    ReDim g_a_product_routes(0)
                    Dim i As Integer = 0
                    While dr_routes.Read
                        CheckedListBox2.Items.Add(dr_routes.Item(1))
                        If dr_routes.Item(2) = "Y" Then
                            CheckedListBox2.SetItemChecked(i, True)
                        End If
                        ReDim Preserve g_a_product_routes(i)
                        g_a_product_routes(i) = dr_routes(0)
                        i = i + 1
                    End While
                End If
                dr_routes.Close()

            Else
                MsgBox("Please select at least one route! Medication cannot be prescribed withour a route.", vbInformation)
            End If

        Else
            MsgBox("Please select a product!", vbCritical)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If g_selected_index > -1 Then
            If CheckedListBox2.CheckedIndices.Count > 0 Then
                Dim l_index As Integer = 0
                For Each indexChecked In CheckedListBox2.CheckedIndices
                    l_index = indexChecked
                Next
                If Not medication.SET_ROUTE_DEFAULT(TextBox1.Text, g_a_list_products(g_selected_index), g_id_product_supplier, g_a_product_routes(l_index)) Then
                    MsgBox("Error setting default route!", vbCritical)
                Else
                    MsgBox("Record updated!", vbInformation)
                End If
            Else
                MsgBox("Please select one default route.", vbInformation)
            End If
        Else
            MsgBox("Please select a product!", vbCritical)
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim l_a_product_um() As Int64
        Dim l_index As Integer = 0

        If ComboBox13.SelectedIndex >= 0 Then
            If CheckedListBox3.CheckedIndices.Count > 0 Then

                For Each indexChecked In CheckedListBox3.CheckedIndices
                    ReDim Preserve l_a_product_um(l_index)
                    l_a_product_um(l_index) = g_a_market_um(indexChecked)
                    l_index = l_index + 1
                Next

                If Not medication.SET_PRODUCT_UM(TextBox1.Text, g_a_list_products(g_selected_index), g_id_product_supplier, ComboBox13.SelectedIndex + 1, l_a_product_um) Then
                    MsgBox("Error inserting product unit measures!", vbCritical)
                Else
                    ComboBox14.SelectedIndex = ComboBox13.SelectedIndex

                    'OBTER AS UNIDADES DE MEDIDA DO PRODUTO
                    Dim dr_product_um As OracleDataReader
                    If Not medication.GET_PRODUCT_UM(TextBox1.Text, g_a_list_products(g_selected_index), g_id_product_supplier, ComboBox14.SelectedIndex + 1, dr_product_um) Then
                        MsgBox("Error getting product unit measures!", vbCritical)
                    Else
                        CheckedListBox4.Items.Clear()
                        ReDim g_a_product_um(0)
                        Dim i As Integer = 0
                        While dr_product_um.Read
                            CheckedListBox4.Items.Add(dr_product_um.Item(1))
                            If dr_product_um.Item(2) = "Y" Then
                                CheckedListBox4.SetItemChecked(i, True)
                            End If
                            ReDim Preserve g_a_product_um(i)
                            g_a_product_um(i) = dr_product_um(0)
                            i = i + 1
                        End While
                        dr_product_um.Close()
                    End If
                End If
                    Else
                MsgBox("Please select at least one route! Medication cannot be prescribed withour a route.", vbInformation)
            End If
        Else
            MsgBox("Please select a dose context!", vbCritical)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        If TextBox1.Text <> "" Then
            Cursor = Cursors.WaitCursor
            Dim dr_UM As OracleDataReader
            If Not medication.GET_MARKET_UM(TextBox1.Text, TextBox4.Text, dr_UM) Then
                MsgBox("Error getting MARKET UNIT MEASURES!", vbCritical)
            Else
                CheckedListBox3.Items.Clear()
                ReDim g_a_market_um(0)
                Dim i As Integer = 0
                While dr_UM.Read
                    CheckedListBox3.Items.Add(dr_UM.Item(1))

                    ReDim Preserve g_a_market_um(i)
                    g_a_market_um(i) = dr_UM(0)
                    i = i + 1
                End While
            End If
            dr_UM.Close()
            Cursor = Cursors.Arrow
        Else
            MsgBox("Please select an Institution.")
        End If

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Dim l_a_product_um(0) As Int64
        Dim l_index As Integer = 0

        For Each indexChecked In CheckedListBox4.CheckedIndices
            ReDim Preserve l_a_product_um(l_index)
            l_a_product_um(l_index) = g_a_product_um(indexChecked)
            l_index = l_index + 1
        Next
        If Not medication.DELETE_PRODUCT_UM(TextBox1.Text, g_a_list_products(g_selected_index), g_id_product_supplier, ComboBox14.SelectedIndex + 1, l_a_product_um) Then
            MsgBox("Error deleting product unit measures!", vbCritical)
        End If

        ''fazer refresh à grelha de unidades do produto
        ''atualizar array das unidades do produto
    End Sub
End Class
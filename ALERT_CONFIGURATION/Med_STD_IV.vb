Imports Oracle.DataAccess.Client
Public Class Med_STD_IV

    Dim db_access_general As New General
    Dim medication As New Medication_API

    Dim g_id_product As String = ""
    Dim g_id_product_supplier As String = ""
    Dim g_id_institution As Int64 = 0
    Dim g_id_software_index As Int16 = -1 ''index do software
    Dim g_selected_software As Int16 = -1
    Dim g_default_route As String = -1
    Dim g_id_market As Int16 = -1
    Dim g_id_std_presc_dir_item As Int64

    Dim g_a_med_set_instructions() As Medication_API.MED_SET_INSTRUCTIONS
    Dim g_a_frequencies() As Int64
    Dim g_a_product_um() As Int64
    Dim g_a_admin_methods() As Int64
    Dim g_a_admin_sites() As Int64
    Dim g_a_duration_um() As Int64

    Public Sub New(ByVal i_institution As Int64, ByVal i_software_index As Int16, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_default_route As String)

        InitializeComponent()
        g_id_product = i_id_product
        g_id_product_supplier = i_id_product_supplier
        g_id_institution = i_institution
        g_id_software_index = i_software_index
        g_default_route = i_default_route

    End Sub
    Function RESET_FREQUENCIES(ByVal i_prn As String)

        ComboBox3.Items.Clear()
        ComboBox3.SelectedIndex = -1

        Dim dr_freq As OracleDataReader
        ReDim g_a_frequencies(0)
        Dim l_index_freq As Int16 = 0
        If Not medication.GET_ALL_FREQS(g_id_institution, g_selected_software, i_prn, dr_freq) Then
            MsgBox("Error getting all frequencies")
        Else
            ComboBox3.Items.Add("")
            While dr_freq.Read()
                ComboBox3.Items.Add(dr_freq.Item(1))

                ReDim Preserve g_a_frequencies(l_index_freq)
                g_a_frequencies(l_index_freq) = dr_freq.Item(0)
                l_index_freq = l_index_freq + 1
            End While
        End If

    End Function


    Private Sub Med_STD_IV_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "MEDICATION - STANDARD INSTRUCTIONS ::::  Connected to " & Connection.db

        Me.BackColor = Color.FromArgb(215, 215, 180)

        TextBox2.Text = medication.GET_PRODUCT_DESC(g_id_institution, g_id_product, g_id_product_supplier)
        g_id_market = db_access_general.GET_INSTITUTION_MARKET(g_id_institution)

        Dim dr_soft As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_SOFT_INST(g_id_institution, dr_soft) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING SOFTWARES!", vbCritical)
        Else
            While dr_soft.Read()
                ComboBox2.Items.Add(dr_soft.Item(1))
            End While
        End If

        dr_soft.Dispose()
        dr_soft.Close()

        ComboBox2.SelectedIndex = g_id_software_index
        g_selected_software = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, g_id_institution)

        RESET_FREQUENCIES("N")

        ComboBox28.Items.Add("0 - ALL")
        ComboBox28.Items.Add("1 - External Prescription")
        ComboBox28.Items.Add("2 - Administer Here")
        ComboBox28.Items.Add("3 - Home Medication")

        ComboBox24.Items.Add("Y")
        ComboBox24.Items.Add("N")

        Dim l_dr_product_um As OracleDataReader
        Dim i As Integer = 0
        If Not medication.GET_PRODUCT_UM(g_id_institution, g_id_product, g_id_product_supplier, 1, l_dr_product_um) Then
            MsgBox("Error getting product unit measures!", vbCritical)
        Else

            ReDim g_a_product_um(0)

            While l_dr_product_um.Read()
                If i = 0 Then
                    ComboBox4.Items.Add("")
                    ComboBox6.Items.Add("")
                    ComboBox7.Items.Add("")
                    ComboBox8.Items.Add("")
                    ComboBox9.Items.Add("")
                    ComboBox10.Items.Add("")
                    ComboBox11.Items.Add("")
                End If
                ComboBox4.Items.Add(l_dr_product_um.Item(1))
                ComboBox6.Items.Add(l_dr_product_um.Item(1))
                ComboBox7.Items.Add(l_dr_product_um.Item(1))
                ComboBox8.Items.Add(l_dr_product_um.Item(1))
                ComboBox9.Items.Add(l_dr_product_um.Item(1))
                ComboBox10.Items.Add(l_dr_product_um.Item(1))
                ComboBox11.Items.Add(l_dr_product_um.Item(1))
                ReDim Preserve g_a_product_um(i)
                g_a_product_um(i) = l_dr_product_um(0)
                i = i + 1
            End While
            l_dr_product_um.Close()
        End If

        Dim l_dr_admin_method As OracleDataReader
        If Not medication.GET_ADMIN_METHOD_LIST(g_id_institution, g_default_route, g_id_product_supplier, l_dr_admin_method) Then
            MsgBox("Error getting list of administration methods!", vbCritical)
        End If

        ReDim g_a_admin_methods(0)
        i = 0
        While l_dr_admin_method.Read()
            ComboBox27.Items.Add(l_dr_admin_method.Item(1))
            ReDim Preserve g_a_admin_methods(i)
            g_a_admin_methods(i) = l_dr_admin_method(0)
            i = i + 1
        End While

        Dim l_dr_admin_sites As OracleDataReader
        If Not medication.GET_ADMIN_SITE_LIST(g_id_institution, g_default_route, g_id_product_supplier, l_dr_admin_sites) Then
            MsgBox("Error getting list of administration sotes!", vbCritical)
        End If

        ReDim g_a_admin_sites(0)
        Dim ii As Integer = 0
        While l_dr_admin_sites.Read()
            ComboBox26.Items.Add(l_dr_admin_sites.Item(1))
            ReDim Preserve g_a_admin_sites(ii)
            g_a_admin_sites(ii) = l_dr_admin_sites.Item(0)
            ii = ii + 1
        End While

        Dim l_dr_duration_um As OracleDataReader
        If Not medication.GET_DURATION_UM(g_id_institution, l_dr_duration_um) Then
            MsgBox("Error getting duration unit measures!", vbCritical)
        Else
            ComboBox5.Items.Add("")

            ReDim g_a_duration_um(0)
            Dim j As Integer = 0
            While l_dr_duration_um.Read()
                ComboBox5.Items.Add(l_dr_duration_um.Item(1))

                ReDim Preserve g_a_duration_um(j)
                g_a_duration_um(j) = l_dr_duration_um(0)
                j = j + 1
            End While
            l_dr_duration_um.Close()
        End If

        'RATES
        '1
        ComboBox18.Items.Add("")
        ComboBox18.Items.Add("mL/h")
        ComboBox18.Items.Add("Bolus")
        '2
        ComboBox17.Items.Add("")
        ComboBox17.Items.Add("mL/h")
        ComboBox17.Items.Add("Bolus")
        '3
        ComboBox16.Items.Add("")
        ComboBox16.Items.Add("mL/h")
        ComboBox16.Items.Add("Bolus")
        '4
        ComboBox15.Items.Add("")
        ComboBox15.Items.Add("mL/h")
        ComboBox15.Items.Add("Bolus")
        '5
        ComboBox14.Items.Add("")
        ComboBox14.Items.Add("mL/h")
        ComboBox14.Items.Add("Bolus")
        '6
        ComboBox13.Items.Add("")
        ComboBox13.Items.Add("mL/h")
        ComboBox13.Items.Add("Bolus")
        '7
        ComboBox12.Items.Add("")
        ComboBox12.Items.Add("mL/h")
        ComboBox12.Items.Add("Bolus")

        'infusiont times - Hours
        ComboBox38.Items.Add("hour(s)")
        ComboBox37.Items.Add("hour(s)")
        ComboBox36.Items.Add("hour(s)")
        ComboBox34.Items.Add("hour(s)")
        ComboBox39.Items.Add("hour(s)")
        ComboBox33.Items.Add("hour(s)")
        ComboBox32.Items.Add("hour(s)")
        ComboBox38.SelectedIndex = 0
        ComboBox37.SelectedIndex = 0
        ComboBox36.SelectedIndex = 0
        ComboBox34.SelectedIndex = 0
        ComboBox39.SelectedIndex = 0
        ComboBox33.SelectedIndex = 0
        ComboBox32.SelectedIndex = 0
        'infusiont times - Minutes
        ComboBox30.Items.Add("minute(s)")
        ComboBox23.Items.Add("minute(s)")
        ComboBox21.Items.Add("minute(s)")
        ComboBox20.Items.Add("minute(s)")
        ComboBox19.Items.Add("minute(s)")
        ComboBox22.Items.Add("minute(s)")
        ComboBox35.Items.Add("minute(s)")
        ComboBox30.SelectedIndex = 0
        ComboBox23.SelectedIndex = 0
        ComboBox21.SelectedIndex = 0
        ComboBox20.SelectedIndex = 0
        ComboBox19.SelectedIndex = 0
        ComboBox22.SelectedIndex = 0
        ComboBox35.SelectedIndex = 0
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox2.SelectedIndex > -1 Then
            Dim l_id_grant As Int64
            l_id_grant = medication.GET_ID_GRANT(g_id_institution, g_selected_software, "LNK_PRODUCT_STD_PRESC_DIR")

            'SE GRANT FOR = -1 ENTÃO É NECESSÁRIO CRIAR UM NOVO GRANT
            If l_id_grant = -1 Then
                If Not medication.SET_ID_GRANT(g_id_institution, g_selected_software, "LNK_PRODUCT_STD_PRESC_DIR") Then
                    MsgBox("Error creating ID_GRANT!", vbCritical)
                Else
                    l_id_grant = medication.GET_ID_GRANT(g_id_institution, g_selected_software, "LNK_PRODUCT_STD_PRESC_DIR")
                End If
            End If

            TextBox27.Text = l_id_grant

        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex > -1 And ComboBox2.Text <> "" Then

            ComboBox25.Items.Clear()

            g_selected_software = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, g_id_institution)

            Dim l_dr_sos As OracleDataReader
            If Not medication.GET_SOS_LIST(g_id_institution, g_selected_software, l_dr_sos) Then
                MsgBox("Error geting list of SOS reasons.", vbCritical)
            End If

            While l_dr_sos.Read()
                ComboBox25.Items.Add(l_dr_sos(1))
            End While

        End If
    End Sub

    Function RESET_MAIN_INSTRUCTIONS()
        'SOS
        ComboBox24.SelectedIndex = -1
        ComboBox25.SelectedIndex = -1
        TextBox24.Text = ""
        'ADMIN
        ComboBox26.SelectedIndex = -1
        ComboBox27.SelectedIndex = -1
        'NOTES
        TextBox36.Text = ""
        TextBox35.Text = ""
        'RANK
        TextBox26.Text = ""
        'COPY TO
        ComboBox29.SelectedIndex = -1
        ComboBox31.SelectedIndex = -1
        'GRANT
        TextBox27.Text = ""
        'FREQUÊNCIA/END_DATE
        ComboBox3.SelectedIndex = -1
        TextBox3.Text = ""
        ComboBox5.SelectedIndex = -1
        TextBox4.Text = ""
    End Function

    Function RESET_SET_INSTRUCTIONS()

        'doses
        TextBox1.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""

        ComboBox4.SelectedIndex = -1
        ComboBox6.SelectedIndex = -1
        ComboBox7.SelectedIndex = -1
        ComboBox8.SelectedIndex = -1
        ComboBox9.SelectedIndex = -1
        ComboBox10.SelectedIndex = -1
        ComboBox11.SelectedIndex = -1

        'rates
        TextBox17.Text = ""
        TextBox16.Text = ""
        TextBox15.Text = ""
        TextBox14.Text = ""
        TextBox13.Text = ""
        TextBox12.Text = ""
        TextBox11.Text = ""

        ComboBox18.SelectedIndex = -1
        ComboBox17.SelectedIndex = -1
        ComboBox16.SelectedIndex = -1
        ComboBox15.SelectedIndex = -1
        ComboBox14.SelectedIndex = -1
        ComboBox13.SelectedIndex = -1
        ComboBox12.SelectedIndex = -1

        'hours
        TextBox34.Text = ""
        TextBox33.Text = ""
        TextBox32.Text = ""
        TextBox29.Text = ""
        TextBox31.Text = ""
        TextBox28.Text = ""
        TextBox25.Text = ""

        'minutes
        TextBox23.Text = ""
        TextBox22.Text = ""
        TextBox20.Text = ""
        TextBox18.Text = ""
        TextBox30.Text = ""
        TextBox19.Text = ""
        TextBox21.Text = ""

    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        ComboBox1.SelectedIndex = -1

        RESET_MAIN_INSTRUCTIONS()

        RESET_SET_INSTRUCTIONS()
    End Sub

    Private Sub TextBox34_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox34.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox33_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox33.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox32_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox32.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox29_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox29.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox31_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox31.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox28_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox28.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox25_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox25.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox23_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox23.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox22_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox22.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox20_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox20.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox18_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox18.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox30_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox30.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox19_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox19.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox21_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox21.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub ComboBox24_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox24.SelectedIndexChanged
        If ComboBox24.Text = "N" Then
            RESET_FREQUENCIES("N")
        ElseIf ComboBox24.Text = "Y" Then
            RESET_FREQUENCIES("Y")
        Else
            RESET_FREQUENCIES("N")
        End If
    End Sub

    Private Sub ComboBox28_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox28.SelectedIndexChanged
        Cursor = Cursors.WaitCursor

        RESET_MAIN_INSTRUCTIONS()
        RESET_SET_INSTRUCTIONS()

        Dim dr_med_set_instruction As OracleDataReader
        ReDim g_a_med_set_instructions(0)
        ComboBox1.Items.Clear()
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not medication.GET_ALL_INSTRUCTIONS(g_id_institution, g_selected_software, g_id_product, g_id_product_supplier, ComboBox28.SelectedIndex, dr_med_set_instruction) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING LIST OF STANDARD INSTRUCTIONS!", vbCritical)
        Else
            Dim i As Integer = 0
            While dr_med_set_instruction.Read()
                ReDim Preserve g_a_med_set_instructions(i)
                g_a_med_set_instructions(i).id_product = dr_med_set_instruction.Item(0)
                g_a_med_set_instructions(i).id_std_presc_dir = dr_med_set_instruction.Item(1)
                g_a_med_set_instructions(i).rank = dr_med_set_instruction.Item(2)
                g_a_med_set_instructions(i).id_grant = dr_med_set_instruction.Item(3)
                g_a_med_set_instructions(i).market = dr_med_set_instruction.Item(4)
                g_a_med_set_instructions(i).market_desc = ""
                g_a_med_set_instructions(i).software = dr_med_set_instruction.Item(6)
                g_a_med_set_instructions(i).software_desc = dr_med_set_instruction.Item(7)
                g_a_med_set_instructions(i).id_pick_list = dr_med_set_instruction.Item(8)
                g_a_med_set_instructions(i).institution = dr_med_set_instruction.Item(9)

                ComboBox1.Items.Add(g_a_med_set_instructions(i).rank)

                i = i + 1

            End While
        End If

        Cursor = Cursors.Arrow
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        If ComboBox1.SelectedIndex > -1 Then

            Cursor = Cursors.WaitCursor

            RESET_MAIN_INSTRUCTIONS()
            RESET_SET_INSTRUCTIONS()

            TextBox27.Text = g_a_med_set_instructions(ComboBox1.SelectedIndex).id_grant

            Dim dr_std_presc_dir As OracleDataReader
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not medication.GET_STD_PRESC_DIR(g_id_institution, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_std_presc_dir, dr_std_presc_dir) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                MsgBox("ERROR GETTING STANDARD_PRESC_DIR!", vbCritical)
            Else
                While dr_std_presc_dir.Read()
                    ComboBox24.Text = dr_std_presc_dir.Item(1)
                    Try
                        TextBox24.Text = dr_std_presc_dir.Item(3)
                    Catch ex As Exception
                        TextBox24.Text = ""
                    End Try
                    Try
                        ComboBox26.Text = dr_std_presc_dir.Item(4)
                    Catch ex As Exception
                        ComboBox26.Text = ""
                    End Try
                    Try
                        ComboBox27.Text = dr_std_presc_dir.Item(5)
                    Catch ex As Exception
                        ComboBox27.Text = ""
                    End Try
                    Try
                        TextBox36.Text = dr_std_presc_dir.Item(6)
                    Catch ex As Exception
                        TextBox36.Text = ""
                    End Try
                    Try
                        TextBox35.Text = dr_std_presc_dir.Item(7)
                    Catch ex As Exception
                        TextBox35.Text = ""
                    End Try
                End While
                dr_std_presc_dir.Close()
            End If

            Dim dr_std_presc_dir_item As OracleDataReader
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not medication.GET_STD_PRESC_DIR_ITEM_IV(g_id_institution, g_id_product, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_pick_list, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_std_presc_dir, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_grant, g_a_med_set_instructions(ComboBox1.SelectedIndex).rank, dr_std_presc_dir_item) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                MsgBox("ERROR GETTING STANDARD_PRESC_DIR_ITEM!", vbCritical)
            Else
                While dr_std_presc_dir_item.Read()
                    'FREQUENCY
                    Try
                        ComboBox3.Text = dr_std_presc_dir_item.Item(6)
                    Catch ex As Exception
                        ComboBox3.Text = ""
                    End Try
                    'DURATION
                    Try
                        TextBox3.Text = dr_std_presc_dir_item.Item(1)
                    Catch ex As Exception
                        TextBox3.Text = ""
                    End Try
                    'DURATION UNIT MEASURE
                    Try
                        ComboBox5.Text = dr_std_presc_dir_item.Item(3)
                    Catch ex As Exception
                        ComboBox5.Text = ""
                    End Try
                    'EXECUTIONS
                    Try
                        TextBox4.Text = dr_std_presc_dir_item.Item(4)
                    Catch ex As Exception
                        TextBox4.Text = ""
                    End Try
                    g_id_std_presc_dir_item = dr_std_presc_dir_item.Item(8)
                End While
            End If

            '''continuar com a obtençao das doses
            '''será necessário criar nova função na api para ir ler STD_PRESC_DIR_ITEM_SEQ (utilizar o valor de g_id_std_presc_dir_item
            '''tratar do reset ao valor g_id_std_presc_dir_item
            Cursor = Cursors.Arrow
        End If

    End Sub
End Class
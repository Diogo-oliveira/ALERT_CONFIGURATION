Imports Oracle.DataAccess.Client
Public Class MED_STD_NON_IV

    Dim db_access_general As New General
    Dim medication As New Medication_API

    Dim g_id_product As String = ""
    Dim g_id_product_supplier As String = ""
    Dim g_id_institution As Int64 = 0
    Dim g_id_software_index As Int16 = -1 ''index do software
    Dim g_selected_software As Int16 = -1
    Dim g_default_route As String = -1
    Dim g_id_market As Int16 = -1

    Dim g_a_med_set_instructions() As Medication_API.MED_SET_INSTRUCTIONS
    Dim g_a_admin_methods() As Int64
    Dim g_a_admin_sites() As Int64
    Dim g_a_product_um() As Int64
    Dim g_a_duration_um() As Int64

    Public Sub New(ByVal i_institution As Int64, ByVal i_software_index As Int16, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_default_route As Int64)

        InitializeComponent()
        g_id_product = i_id_product
        g_id_product_supplier = i_id_product_supplier
        g_id_institution = i_institution
        g_id_software_index = i_software_index
        g_default_route = i_default_route

    End Sub

    Private Sub MED_STD_NON_IV_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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

        Dim dr_freq As OracleDataReader

        If Not medication.GET_ALL_FREQS(g_id_institution, g_selected_software, dr_freq) Then
            MsgBox("Error getting all frequencies")
        Else
            ComboBox3.Items.Add("")
            ComboBox11.Items.Add("")
            ComboBox20.Items.Add("")
            ComboBox17.Items.Add("")
            ComboBox14.Items.Add("")
            ComboBox23.Items.Add("")
            While dr_freq.Read()
                ComboBox3.Items.Add(dr_freq.Item(1))
                ComboBox8.Items.Add(dr_freq.Item(1))
                ComboBox11.Items.Add(dr_freq.Item(1))
                ComboBox20.Items.Add(dr_freq.Item(1))
                ComboBox17.Items.Add(dr_freq.Item(1))
                ComboBox14.Items.Add(dr_freq.Item(1))
                ComboBox23.Items.Add(dr_freq.Item(1))
            End While
        End If

        ComboBox28.Items.Add("0 - ALL")
        ComboBox28.Items.Add("1 - External Prescription")
        ComboBox28.Items.Add("2 - Administer Here")
        ComboBox28.Items.Add("3 - Home Medication")

        ComboBox24.Items.Add("Y")
        ComboBox24.Items.Add("N")

        Dim l_dr_market As OracleDataReader
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_MARKETS(l_dr_market) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            MsgBox("ERROR GETTING MARKETS!", vbCritical)
        Else

            While l_dr_market.Read()
                If l_dr_market.Item(0) = 0 Or l_dr_market.Item(0) = g_id_market Then
                    ComboBox30.Items.Add(l_dr_market.Item(1))
                End If
            End While
        End If

        Dim l_dr_product_um As OracleDataReader
        If Not medication.GET_PRODUCT_UM(g_id_institution, g_id_product, g_id_product_supplier, 1, l_dr_product_um) Then
            MsgBox("Error getting product unit measures!", vbCritical)
        Else

            ReDim g_a_product_um(0)
            Dim i As Integer = 0
            While l_dr_product_um.Read()
                If i = 0 Then
                    ComboBox4.Items.Add("")
                    ComboBox7.Items.Add("")
                    ComboBox10.Items.Add("")
                    ComboBox19.Items.Add("")
                    ComboBox16.Items.Add("")
                    ComboBox13.Items.Add("")
                    ComboBox22.Items.Add("")
                End If
                ComboBox4.Items.Add(l_dr_product_um.Item(1))
                ComboBox7.Items.Add(l_dr_product_um.Item(1))
                ComboBox10.Items.Add(l_dr_product_um.Item(1))
                ComboBox19.Items.Add(l_dr_product_um.Item(1))
                ComboBox16.Items.Add(l_dr_product_um.Item(1))
                ComboBox13.Items.Add(l_dr_product_um.Item(1))
                ComboBox22.Items.Add(l_dr_product_um.Item(1))
                ReDim Preserve g_a_product_um(i)
                g_a_product_um(i) = l_dr_product_um(0)
                i = i + 1
            End While
            l_dr_product_um.Close()
        End If

        Dim l_dr_duration_um As OracleDataReader
        If Not medication.GET_DURATION_UM(g_id_institution, l_dr_duration_um) Then
            MsgBox("Error getting duration unit measures!", vbCritical)
        Else
            ComboBox5.Items.Add("")
            ComboBox6.Items.Add("")
            ComboBox9.Items.Add("")
            ComboBox18.Items.Add("")
            ComboBox15.Items.Add("")
            ComboBox12.Items.Add("")
            ComboBox21.Items.Add("")
            ReDim g_a_duration_um(0)
            Dim i As Integer = 0
            While l_dr_duration_um.Read()
                ComboBox5.Items.Add(l_dr_duration_um.Item(1))
                ComboBox6.Items.Add(l_dr_duration_um.Item(1))
                ComboBox9.Items.Add(l_dr_duration_um.Item(1))
                ComboBox18.Items.Add(l_dr_duration_um.Item(1))
                ComboBox15.Items.Add(l_dr_duration_um.Item(1))
                ComboBox12.Items.Add(l_dr_duration_um.Item(1))
                ComboBox21.Items.Add(l_dr_duration_um.Item(1))
                ReDim Preserve g_a_duration_um(i)
                g_a_duration_um(i) = l_dr_duration_um(0)
                i = i + 1
            End While
            l_dr_duration_um.Close()
        End If


    End Sub

    Private Sub ComboBox28_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox28.SelectedIndexChanged
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
                g_a_med_set_instructions(i).market_desc = dr_med_set_instruction.Item(5)
                g_a_med_set_instructions(i).software = dr_med_set_instruction.Item(6)
                g_a_med_set_instructions(i).software_desc = dr_med_set_instruction.Item(7)
                g_a_med_set_instructions(i).id_pick_list = dr_med_set_instruction.Item(8)
                g_a_med_set_instructions(i).institution = dr_med_set_instruction.Item(9)

                ComboBox1.Items.Add(g_a_med_set_instructions(i).rank)

                i = i + 1

            End While
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox27.Text = g_a_med_set_instructions(ComboBox1.SelectedIndex).id_grant
        If g_a_med_set_instructions(ComboBox1.SelectedIndex).market = 0 Then
            ComboBox30.SelectedIndex = 0
        ElseIf g_a_med_set_instructions(ComboBox1.SelectedIndex).market = g_id_market Then
            ComboBox30.SelectedIndex = 1
        End If

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
                    ComboBox26 = dr_std_presc_dir.Item(4)
                Catch ex As Exception
                    ComboBox26.Text = ""
                End Try
                Try
                    ComboBox27 = dr_std_presc_dir.Item(5)
                Catch ex As Exception
                    ComboBox27.Text = ""
                End Try
                Try
                    TextBox5.Text = dr_std_presc_dir.Item(6)
                Catch ex As Exception
                    TextBox5.Text = ""
                End Try
                Try
                    TextBox25.Text = dr_std_presc_dir.Item(7)
                Catch ex As Exception
                    TextBox25.Text = ""
                End Try
            End While
            dr_std_presc_dir.Close()
        End If


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''CONTINUAR PARA TODOS OS CAMPOS
        ''CRAIR A ESTRUTURA OUS O ARRAYS PARA GUARDAR AS INSTRUÇÕES A SEREM GRAVADAS
        Dim dr_std_presc_dir_item As OracleDataReader
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not medication.GET_STD_PRESC_DIR_ITEM(g_id_institution, g_id_product, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_pick_list, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_std_presc_dir, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_grant, g_a_med_set_instructions(ComboBox1.SelectedIndex).rank, dr_std_presc_dir_item) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            MsgBox("ERROR GETTING STANDARD_PRESC_DIR_ITEM!", vbCritical)
        Else
            Dim i As Integer = 0
            While dr_std_presc_dir_item.Read()
                If i = 0 Then
                    'DOSE
                    Try
                        TextBox1.Text = dr_std_presc_dir_item.Item(6)
                    Catch ex As Exception
                        TextBox1.Text = ""
                    End Try
                    'DOSE UNIT MEASURE
                    Try
                        ComboBox4.Text = dr_std_presc_dir_item.Item(8)
                    Catch ex As Exception
                        ComboBox4.Text = ""
                    End Try
                    'FREQUENCY
                    Try
                        ComboBox3.Text = dr_std_presc_dir_item.Item(10)
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
                ElseIf i = 1 Then
                    'DOSE
                    Try
                        TextBox8.Text = dr_std_presc_dir_item.Item(6)
                    Catch ex As Exception
                        TextBox8.Text = ""
                    End Try
                    'DOSE UNIT MEASURE
                    Try
                        ComboBox7.Text = dr_std_presc_dir_item.Item(8)
                    Catch ex As Exception
                        ComboBox7.Text = ""
                    End Try
                    'FREQUENCY
                    Try
                        ComboBox8.Text = dr_std_presc_dir_item.Item(10)
                    Catch ex As Exception
                        ComboBox8.Text = ""
                    End Try
                    'DURATION
                    Try
                        TextBox7.Text = dr_std_presc_dir_item.Item(1)
                    Catch ex As Exception
                        TextBox7.Text = ""
                    End Try
                    'DURATION UNIT MEASURE
                    Try
                        ComboBox6.Text = dr_std_presc_dir_item.Item(3)
                    Catch ex As Exception
                        ComboBox6.Text = ""
                    End Try
                    'EXECUTIONS
                    Try
                        TextBox6.Text = dr_std_presc_dir_item.Item(4)
                    Catch ex As Exception
                        TextBox6.Text = ""
                    End Try
                ElseIf i = 2 Then
                    'DOSE
                    Try
                        TextBox11.Text = dr_std_presc_dir_item.Item(6)
                    Catch ex As Exception
                        TextBox11.Text = ""
                    End Try
                    'DOSE UNIT MEASURE
                    Try
                        ComboBox10.Text = dr_std_presc_dir_item.Item(8)
                    Catch ex As Exception
                        ComboBox10.Text = ""
                    End Try
                    'FREQUENCY
                    Try
                        ComboBox11.Text = dr_std_presc_dir_item.Item(10)
                    Catch ex As Exception
                        ComboBox11.Text = ""
                    End Try
                    'DURATION
                    Try
                        TextBox10.Text = dr_std_presc_dir_item.Item(1)
                    Catch ex As Exception
                        TextBox10.Text = ""
                    End Try
                    'DURATION UNIT MEASURE
                    Try
                        ComboBox9.Text = dr_std_presc_dir_item.Item(3)
                    Catch ex As Exception
                        ComboBox9.Text = ""
                    End Try
                    'EXECUTIONS
                    Try
                        TextBox9.Text = dr_std_presc_dir_item.Item(4)
                    Catch ex As Exception
                        TextBox9.Text = ""
                    End Try
                ElseIf i = 3 Then
                    'DOSE
                    Try
                        TextBox20.Text = dr_std_presc_dir_item.Item(6)
                    Catch ex As Exception
                        TextBox20.Text = ""
                    End Try
                    'DOSE UNIT MEASURE
                    Try
                        ComboBox19.Text = dr_std_presc_dir_item.Item(8)
                    Catch ex As Exception
                        ComboBox19.Text = ""
                    End Try
                    'FREQUENCY
                    Try
                        ComboBox20.Text = dr_std_presc_dir_item.Item(10)
                    Catch ex As Exception
                        ComboBox20.Text = ""
                    End Try
                    'DURATION
                    Try
                        TextBox19.Text = dr_std_presc_dir_item.Item(1)
                    Catch ex As Exception
                        TextBox19.Text = ""
                    End Try
                    'DURATION UNIT MEASURE
                    Try
                        ComboBox18.Text = dr_std_presc_dir_item.Item(3)
                    Catch ex As Exception
                        ComboBox18.Text = ""
                    End Try
                    'EXECUTIONS
                    Try
                        TextBox18.Text = dr_std_presc_dir_item.Item(4)
                    Catch ex As Exception
                        TextBox18.Text = ""
                    End Try
                ElseIf i = 4 Then
                    'DOSE
                    Try
                        TextBox17.Text = dr_std_presc_dir_item.Item(6)
                    Catch ex As Exception
                        TextBox17.Text = ""
                    End Try
                    'DOSE UNIT MEASURE
                    Try
                        ComboBox16.Text = dr_std_presc_dir_item.Item(8)
                    Catch ex As Exception
                        ComboBox16.Text = ""
                    End Try
                    'FREQUENCY
                    Try
                        ComboBox17.Text = dr_std_presc_dir_item.Item(10)
                    Catch ex As Exception
                        ComboBox17.Text = ""
                    End Try
                    'DURATION
                    Try
                        TextBox16.Text = dr_std_presc_dir_item.Item(1)
                    Catch ex As Exception
                        TextBox16.Text = ""
                    End Try
                    'DURATION UNIT MEASURE
                    Try
                        ComboBox15.Text = dr_std_presc_dir_item.Item(3)
                    Catch ex As Exception
                        ComboBox15.Text = ""
                    End Try
                    'EXECUTIONS
                    Try
                        TextBox15.Text = dr_std_presc_dir_item.Item(4)
                    Catch ex As Exception
                        TextBox15.Text = ""
                    End Try
                ElseIf i = 5 Then
                    'DOSE
                    Try
                        TextBox14.Text = dr_std_presc_dir_item.Item(6)
                    Catch ex As Exception
                        TextBox14.Text = ""
                    End Try
                    'DOSE UNIT MEASURE
                    Try
                        ComboBox13.Text = dr_std_presc_dir_item.Item(8)
                    Catch ex As Exception
                        ComboBox13.Text = ""
                    End Try
                    'FREQUENCY
                    Try
                        ComboBox14.Text = dr_std_presc_dir_item.Item(10)
                    Catch ex As Exception
                        ComboBox14.Text = ""
                    End Try
                    'DURATION
                    Try
                        TextBox13.Text = dr_std_presc_dir_item.Item(1)
                    Catch ex As Exception
                        TextBox13.Text = ""
                    End Try
                    'DURATION UNIT MEASURE
                    Try
                        ComboBox12.Text = dr_std_presc_dir_item.Item(3)
                    Catch ex As Exception
                        ComboBox12.Text = ""
                    End Try
                    'EXECUTIONS
                    Try
                        TextBox12.Text = dr_std_presc_dir_item.Item(4)
                    Catch ex As Exception
                        TextBox12.Text = ""
                    End Try
                ElseIf i = 6 Then
                    'DOSE
                    Try
                        TextBox23.Text = dr_std_presc_dir_item.Item(6)
                    Catch ex As Exception
                        TextBox23.Text = ""
                    End Try
                    'DOSE UNIT MEASURE
                    Try
                        ComboBox22.Text = dr_std_presc_dir_item.Item(8)
                    Catch ex As Exception
                        ComboBox22.Text = ""
                    End Try
                    'FREQUENCY
                    Try
                        ComboBox23.Text = dr_std_presc_dir_item.Item(10)
                    Catch ex As Exception
                        ComboBox23.Text = ""
                    End Try
                    'DURATION
                    Try
                        TextBox22.Text = dr_std_presc_dir_item.Item(1)
                    Catch ex As Exception
                        TextBox22.Text = ""
                    End Try
                    'DURATION UNIT MEASURE
                    Try
                        ComboBox21.Text = dr_std_presc_dir_item.Item(3)
                    Catch ex As Exception
                        ComboBox21.Text = ""
                    End Try
                    'EXECUTIONS
                    Try
                        TextBox21.Text = dr_std_presc_dir_item.Item(4)
                    Catch ex As Exception
                        TextBox21.Text = ""
                    End Try
                End If
                i = i + 1
            End While
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex > -1 And ComboBox2.Text <> "" Then
            g_selected_software = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, g_id_institution)

            Dim l_dr_sos As OracleDataReader
            If Not medication.GET_SOS_LIST(g_id_institution, g_selected_software, l_dr_sos) Then
                MsgBox("Error geting list of SOS reasons.", vbCritical)
            End If

            While l_dr_sos.Read()
                ComboBox25.Items.Add(l_dr_sos(1))
            End While

            Dim l_dr_admin_method As OracleDataReader
            If Not medication.GET_ADMIN_METHOD_LIST(g_id_institution, g_default_route, g_id_product_supplier, l_dr_admin_method) Then
                MsgBox("Error getting list of administration methods!", vbCritical)
            End If

            ReDim g_a_admin_methods(0)
            Dim i As Integer = 0
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

        End If
    End Sub

    Private Sub ComboBox25_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox25.SelectedIndexChanged
        TextBox24.Text = ""
    End Sub

    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged
        ComboBox25.SelectedIndex = -1
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Cursor = Cursors.WaitCursor
        Dim l_id_grant As Int64 = -1

        If ComboBox2.SelectedIndex < 0 Then
            MsgBox("Please select a software.", vbInformation)
        ElseIf ComboBox28.SelectedIndex < 0 Then
            MsgBox("Please select a type of prescription.", vbInformation)
        ElseIf ComboBox1.SelectedIndex < 0 And TextBox26.Text = "" Then
            MsgBox("Please select a rank.", vbInformation)
        Else
            'VERIFICAR SE EXISTE GRANT
            If TextBox27.Text = "" Then
                l_id_grant = medication.GET_ID_GRANT(g_id_institution, g_selected_software, "LNK_PRODUCT_STD_PRESC_DIR")
                'SE GRANT FOR = -1 ENTÃO É NECESSÁRIO CRIAR UM NOVO GRANT
                If Not medication.SET_ID_GRANT(g_id_institution, g_selected_software, "LNK_PRODUCT_STD_PRESC_DIR") Then
                    MsgBox("Error creating ID_GRANT!", vbCritical)
                Else
                    l_id_grant = medication.GET_ID_GRANT(g_id_institution, g_selected_software, "LNK_PRODUCT_STD_PRESC_DIR")
                End If
            Else
                'NESTE CASO JÁ EXISTIA INSTRUÇÃO. SERÁ FEITO UPDATE

                'VERIFICAR SE INSTRUÇÃO É UTILIZADA EM DIVERSAS PICK_LISTS/SOFTWARES
                'CASO SEJA É NECESSÁRIO PERGUNTAR SE SE FAZ UPDATE OU SE SE INSERE NOVA INSTRUÇÃO
                l_id_grant = TextBox27.Text
                Dim l_create_new As Integer = 0

                If medication.CHECK_DUP_INSTRUCTIONS(g_id_institution, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_std_presc_dir) > 1 Then
                    l_create_new = MsgBox("The current standard instruction is also being usaed for other softwares and/or type of prescriptions. Do you wish to create a new instruction just for the selected software and type of prescription? (Responding 'No' will result on the update of the current standard instruction)", MessageBoxButtons.YesNo)
                End If

                If l_create_new = 0 Or l_create_new = DialogResult.No Then
                    'UPDATE LNK_PRODUCT_STD_INSTRUCTION               
                    If TextBox26.Text <> "" Then
                        ''NESTE CASO É NECESSÁRIO FAZER UPDATE AO RANK
                        If Not medication.UPDATE_STD_PRESC_DIR(g_id_institution, g_id_product, g_id_product_supplier, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_std_presc_dir, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_grant, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_pick_list, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_std_presc_dir, TextBox26.Text) Then
                            MsgBox("Error updating instruction rank!", vbCritical)
                        End If
                    End If
                Else
                    ''CRIAÇÃO DE UMA NOVA INSTRUÇÃO
                    MsgBox("A new instruction will be created") ''APAGAR
                    ''criar novo id
                    Dim l_id_new_instruction As Int64 = medication.GET_NEW_STD_INSTRUCTION_ID(g_id_institution)
                    Dim l_flg_sos As String
                    Dim l_id_sos As Int16 = 19
                    Dim l_sos_condition As String = ""

                    If ComboBox24.Text = "" Then
                        l_flg_sos = "N"
                    Else
                        l_flg_sos = ComboBox24.Text
                    End If

                    If ComboBox24.Text = "Y" Then
                        l_id_sos = 18
                    End If

                    If TextBox24.Text <> "" Then
                        l_sos_condition = TextBox24.Text
                    ElseIf ComboBox25.Text <> "" Then
                        l_sos_condition = ComboBox25.Text
                    End If

                    Dim l_id_admin_site As String = "NULL"
                    If ComboBox26.SelectedIndex > -1 Then
                        l_id_admin_site = g_a_admin_sites(ComboBox26.SelectedIndex)
                    End If

                    Dim l_id_admin_method As String = "NULL"
                    If ComboBox27.SelectedIndex > -1 Then
                        l_id_admin_method = g_a_admin_methods(ComboBox27.SelectedIndex)
                    End If

                    If Not medication.CREATE_STD_INSTRUCTION(g_id_institution, l_id_new_instruction, l_flg_sos, l_id_sos, l_sos_condition, TextBox5.Text, TextBox25.Text, l_id_admin_site, l_id_admin_method) Then
                        MsgBox("Error creating standard instruction!", vbCritical)
                    End If
                    Dim l_rank As Int64 = 1
                    If TextBox26.Text <> "" Then
                        l_rank = TextBox26.Text
                    ElseIf combobox1.Text <> "" Then
                        l_rank = ComboBox1.Text
                    End If
                    If Not medication.UPDATE_STD_PRESC_DIR(g_id_institution, g_id_product, g_id_product_supplier, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_std_presc_dir, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_grant, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_pick_list, l_id_new_instruction, l_rank) Then
                        MsgBox("Error updating instruction id/rank!", vbCritical)
                    End If
                End If
            End If
        End If
        Cursor = Cursors.Arrow
    End Sub
End Class
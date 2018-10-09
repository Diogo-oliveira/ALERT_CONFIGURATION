
Imports Oracle.DataAccess.Client
Public Class Med_STD_Complex_IV

    Dim db_access_general As New General
    Dim medication As New Medication_API

    Dim g_id_product As String = ""
    Dim g_id_product_supplier As String = ""
    Dim g_id_institution As Int64 = 0
    Dim g_id_software_index As Int16 = -1 ''index do software
    Dim g_selected_software As Int16 = -1
    Dim g_default_route As String = -1
    Dim g_id_market As Int16 = -1
    Dim g_id_std_presc_dir_item As Int64 = -1
    Dim g_component_selected_index As Int64 = -1

    Dim g_a_med_set_instructions() As Medication_API.MED_SET_INSTRUCTIONS
    Dim g_a_frequencies() As Int64
    Dim g_a_product_um() As Int64
    Dim g_a_admin_methods() As Int64
    Dim g_a_admin_sites() As Int64
    Dim g_a_duration_um() As Int64
    Dim g_a_components_list() As String
    Dim g_a_doses_list() As Medication_API.MED_SET_DOSES
    Dim g_a_unit_measure() As Medication_API.UM_INFO

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

    Function CREATE_SET_INSTRUCTIONS(ByVal i_id_grant As Int64, ByVal i_id_pick_list As Int16, ByVal i_create_new As String) As Boolean
        Try
            ''CRIAÇÃO DE UMA NOVA INSTRUÇÃO
            ''criar novo id
            Dim l_id_new_instruction As Int64 = medication.GET_NEW_STD_INSTRUCTION_ID(g_id_institution)
            Dim l_id_new_std_presc_dir_item As Int64 = medication.GET_NEW_STD_PRESC_DIR_ITEM_ID(g_id_institution)
            Dim l_flg_sos As String
            Dim l_id_sos As Int16 = 19
            Dim l_sos_condition As String = ""

            'flg_sos
            If ComboBox24.Text = "" Then
                l_flg_sos = "N"
            Else
                l_flg_sos = ComboBox24.Text
            End If

            If ComboBox24.Text = "Y" Then
                l_id_sos = 18
            End If

            If TextBox24.Text <> "" Then
                l_sos_condition = "'" & TextBox24.Text & "'"
            ElseIf ComboBox25.Text <> "" Then
                l_sos_condition = "'" & ComboBox25.Text & "'"
            Else
                l_sos_condition = "NULL"
            End If

            Dim l_id_admin_site As String = "NULL"
            If ComboBox26.SelectedIndex > -1 Then
                l_id_admin_site = g_a_admin_sites(ComboBox26.SelectedIndex)
            End If

            Dim l_id_admin_method As String = "NULL"
            If ComboBox27.SelectedIndex > -1 Then
                l_id_admin_method = g_a_admin_methods(ComboBox27.SelectedIndex)
            End If

            If Not medication.CREATE_STD_INSTRUCTION(g_id_institution, l_id_new_instruction, l_flg_sos, l_id_sos, l_sos_condition, TextBox36.Text, TextBox35.Text, l_id_admin_site, l_id_admin_method) Then
                MsgBox("Error creating standard instruction!", vbCritical)
            End If

            Dim l_rank As Int64 = 1
            If TextBox26.Text <> "" Then
                l_rank = TextBox26.Text
            ElseIf ComboBox1.Text <> "" Then
                l_rank = ComboBox1.Text
            End If

            Dim l_previous_id_directions As Int64 = 0

            If i_create_new = "N" Then

                'VER MELHOR
                If Not medication.UPDATE_STD_PRESC_DIR(g_id_institution, g_id_product, g_id_product_supplier, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_std_presc_dir, i_id_grant, i_id_pick_list, l_id_new_instruction, l_rank, i_id_grant) Then
                    MsgBox("Error updating lnk_product_std_presc_dir!", vbCritical)
                End If

            Else

                If Not medication.CREATE_STD_PRESC_DIR(g_id_institution, g_id_product, g_id_product_supplier, l_id_new_instruction, i_id_grant, i_id_pick_list, l_rank) Then
                    MsgBox("Error creating new lnk_product_std_presc_dir!!", vbCritical)
                End If
            End If

            Dim l_id_frequency As String = ""
            If ComboBox3.SelectedIndex > 0 Then
                l_id_frequency = g_a_frequencies(ComboBox3.SelectedIndex - 1)
            End If

            Dim l_id_duration As String = ""
            If ComboBox5.SelectedIndex > 0 Then
                l_id_duration = g_a_duration_um(ComboBox5.SelectedIndex - 1)
            End If

            If Not medication.CREATE_STD_PRESC_DIR_ITEM_IV(g_id_institution, l_id_new_instruction, l_id_new_std_presc_dir_item, l_id_frequency, TextBox3.Text, l_id_duration, TextBox4.Text) Then
                MsgBox("ERROR CREATE_STD_PRESC_DIR_ITEM_IV!", vbCritical)
            End If

            Dim l_number_instructions_to_add As Int16 = CHECK_NUMBER_INSTRUCTIONS()
            If l_number_instructions_to_add > -1 Then
                Dim l_a_instructions() As String

                For i As Integer = 0 To l_number_instructions_to_add
                    GET_INSTRUCTIONS(i, l_a_instructions)

                    If Not medication.CREATE_STD_PRESC_DIR_ITEM_SEQ(g_id_institution, l_id_new_std_presc_dir_item, g_id_product_supplier, i + 1, l_a_instructions) Then
                        MsgBox("Error createing std_presc_dir_item!", vbCritical)
                    End If
                Next
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Function CHECK_NUMBER_INSTRUCTIONS() As Int16
        'VARIFICAR QUANTOS SETS DE INSTRUÇÕES DEVEM SER GRAVADOS
        Dim l_n_of_instruction As Int16 = -1
        '#1
        If (TextBox1.Text <> "" Or TextBox17.Text <> "" Or TextBox34.Text <> "" Or TextBox23.Text <> "" Or ComboBox18.SelectedIndex = 2) Then
            l_n_of_instruction = l_n_of_instruction + 1
        Else
            Return l_n_of_instruction
        End If
        '#2
        If (TextBox5.Text <> "" Or TextBox16.Text <> "" Or TextBox33.Text <> "" Or TextBox22.Text <> "" Or ComboBox17.SelectedIndex = 2) Then
            l_n_of_instruction = l_n_of_instruction + 1
        Else
            Return l_n_of_instruction
        End If
        '#3
        If (TextBox6.Text <> "" Or TextBox15.Text <> "" Or TextBox32.Text <> "" Or TextBox20.Text <> "" Or ComboBox16.SelectedIndex = 2) Then
            l_n_of_instruction = l_n_of_instruction + 1
        Else
            Return l_n_of_instruction
        End If
        '#4
        If (TextBox7.Text <> "" Or TextBox14.Text <> "" Or TextBox29.Text <> "" Or TextBox18.Text <> "" Or ComboBox15.SelectedIndex = 2) Then
            l_n_of_instruction = l_n_of_instruction + 1
        Else
            Return l_n_of_instruction
        End If
        '#5
        If (TextBox8.Text <> "" Or TextBox13.Text <> "" Or TextBox31.Text <> "" Or TextBox30.Text <> "" Or ComboBox14.SelectedIndex = 2) Then
            l_n_of_instruction = l_n_of_instruction + 1
        Else
            Return l_n_of_instruction
        End If
        '#6
        If (TextBox9.Text <> "" Or TextBox12.Text <> "" Or TextBox28.Text <> "" Or TextBox19.Text <> "" Or ComboBox13.SelectedIndex = 2) Then
            l_n_of_instruction = l_n_of_instruction + 1
        Else
            Return l_n_of_instruction
        End If
        '#7
        If (TextBox10.Text <> "" Or TextBox11.Text <> "" Or TextBox25.Text <> "" Or TextBox21.Text <> "" Or ComboBox12.SelectedIndex = 2) Then
            l_n_of_instruction = l_n_of_instruction + 1
        Else
            Return l_n_of_instruction
        End If

        Return l_n_of_instruction

    End Function

    Function GET_INSTRUCTIONS(ByVal i_index_instructions As Int16, ByRef o_array_instructions() As String) As Boolean

        ReDim o_array_instructions(6)

        Try
            If i_index_instructions = 0 Then
                'DOSE
                If TextBox1.Text <> "" Then
                    o_array_instructions(0) = (TextBox1.Text).Replace(",", ".")
                Else
                    o_array_instructions(0) = "NULL"
                End If
                'DOSE UNIT MEASURE
                If ComboBox4.Text <> "" Then
                    o_array_instructions(1) = g_a_product_um(ComboBox4.SelectedIndex - 1) ''-1 PORQUE A 1ª POSIÇÃO DA COMBOBOX É NULL
                Else
                    o_array_instructions(1) = "NULL"
                End If
                'RATE
                If TextBox17.Text <> "" And ComboBox18.SelectedIndex <> 2 Then 'INDEX 2 REFERE-SE AO BOLUS
                    o_array_instructions(2) = (TextBox17.Text).Replace(",", ".")
                Else
                    o_array_instructions(2) = "NULL"
                End If
                'RATE UNIT MEASURE
                If ComboBox18.SelectedIndex = 1 Then
                    o_array_instructions(3) = "10491"
                Else
                    o_array_instructions(3) = "NULL"
                End If
                'RATE UNIT MEASURE - BOLUS
                If ComboBox18.SelectedIndex = 2 Then
                    o_array_instructions(4) = "21"
                Else
                    o_array_instructions(4) = "9999"
                End If
                ''DURATION
                Dim l_duration As String = "NULL"
                Dim l_aux As Int64 = 0
                If TextBox34.Text <> "" Then
                    l_aux = TextBox34.Text * 60
                End If

                If TextBox23.Text <> "" And TextBox34.Text <> "" Then
                    l_aux = l_aux + TextBox23.Text
                ElseIf TextBox23.Text <> "" And TextBox34.Text = "" Then
                    l_aux = TextBox23.Text
                End If
                If l_aux > 0 Then
                    l_duration = l_aux
                End If
                o_array_instructions(5) = l_duration

                If l_duration = "NULL" Then
                    o_array_instructions(6) = "NULL"
                Else
                    o_array_instructions(6) = "10374"
                End If

            ElseIf i_index_instructions = 1 Then
                'DOSE
                If TextBox5.Text <> "" Then
                    o_array_instructions(0) = (TextBox5.Text).Replace(",", ".")
                Else
                    o_array_instructions(0) = "NULL"
                End If
                'DOSE UNIT MEASURE
                If ComboBox6.Text <> "" Then
                    o_array_instructions(1) = g_a_product_um(ComboBox6.SelectedIndex - 1) ''-1 PORQUE A 1ª POSIÇÃO DA COMBOBOX É NULL
                Else
                    o_array_instructions(1) = "NULL"
                End If
                'RATE
                If TextBox16.Text <> "" And ComboBox17.SelectedIndex <> 2 Then 'INDEX 2 REFERE-SE AO BOLUS
                    o_array_instructions(2) = (TextBox16.Text).Replace(",", ".")
                Else
                    o_array_instructions(2) = "NULL"
                End If
                'RATE UNIT MEASURE
                If ComboBox17.SelectedIndex = 1 Then
                    o_array_instructions(3) = "10491"
                Else
                    o_array_instructions(3) = "NULL"
                End If
                'RATE UNIT MEASURE - BOLUS
                If ComboBox17.SelectedIndex = 2 Then
                    o_array_instructions(4) = "21"
                Else
                    o_array_instructions(4) = "9999"
                End If
                ''DURATION
                Dim l_duration As String = "NULL"
                Dim l_aux As Int64 = 0
                If TextBox33.Text <> "" Then
                    l_aux = TextBox33.Text * 60
                End If

                If TextBox22.Text <> "" And TextBox33.Text <> "" Then
                    l_aux = l_aux + TextBox22.Text
                ElseIf TextBox22.Text <> "" And TextBox33.Text = "" Then
                    l_aux = TextBox22.Text
                End If
                If l_aux > 0 Then
                    l_duration = l_aux
                End If
                o_array_instructions(5) = l_duration

                If l_duration = "NULL" Then
                    o_array_instructions(6) = "NULL"
                Else
                    o_array_instructions(6) = "10374"
                End If

            ElseIf i_index_instructions = 2 Then
                'DOSE
                If TextBox6.Text <> "" Then
                    o_array_instructions(0) = (TextBox6.Text).Replace(",", ".")
                Else
                    o_array_instructions(0) = "NULL"
                End If
                'DOSE UNIT MEASURE
                If ComboBox7.Text <> "" Then
                    o_array_instructions(1) = g_a_product_um(ComboBox7.SelectedIndex - 1) ''-1 PORQUE A 1ª POSIÇÃO DA COMBOBOX É NULL
                Else
                    o_array_instructions(1) = "NULL"
                End If
                'RATE
                If TextBox15.Text <> "" And ComboBox16.SelectedIndex <> 2 Then 'INDEX 2 REFERE-SE AO BOLUS
                    o_array_instructions(2) = (TextBox15.Text).Replace(",", ".")
                Else
                    o_array_instructions(2) = "NULL"
                End If
                'RATE UNIT MEASURE
                If ComboBox16.SelectedIndex = 1 Then
                    o_array_instructions(3) = "10491"
                Else
                    o_array_instructions(3) = "NULL"
                End If
                'RATE UNIT MEASURE - BOLUS
                If ComboBox16.SelectedIndex = 2 Then
                    o_array_instructions(4) = "21"
                Else
                    o_array_instructions(4) = "9999"
                End If
                ''DURATION
                Dim l_duration As String = "NULL"
                Dim l_aux As Int64 = 0
                If TextBox32.Text <> "" Then
                    l_aux = TextBox32.Text * 60
                End If

                If TextBox20.Text <> "" And TextBox32.Text <> "" Then
                    l_aux = l_aux + TextBox20.Text
                ElseIf TextBox20.Text <> "" And TextBox32.Text = "" Then
                    l_aux = TextBox20.Text
                End If
                If l_aux > 0 Then
                    l_duration = l_aux
                End If
                o_array_instructions(5) = l_duration

                If l_duration = "NULL" Then
                    o_array_instructions(6) = "NULL"
                Else
                    o_array_instructions(6) = "10374"
                End If

            ElseIf i_index_instructions = 3 Then
                'DOSE
                If TextBox7.Text <> "" Then
                    o_array_instructions(0) = (TextBox7.Text).Replace(",", ".")
                Else
                    o_array_instructions(0) = "NULL"
                End If
                'DOSE UNIT MEASURE
                If ComboBox8.Text <> "" Then
                    o_array_instructions(1) = g_a_product_um(ComboBox8.SelectedIndex - 1) ''-1 PORQUE A 1ª POSIÇÃO DA COMBOBOX É NULL
                Else
                    o_array_instructions(1) = "NULL"
                End If
                'RATE
                If TextBox14.Text <> "" And ComboBox15.SelectedIndex <> 2 Then 'INDEX 2 REFERE-SE AO BOLUS
                    o_array_instructions(2) = (TextBox14.Text).Replace(",", ".")
                Else
                    o_array_instructions(2) = "NULL"
                End If
                'RATE UNIT MEASURE
                If ComboBox15.SelectedIndex = 1 Then
                    o_array_instructions(3) = "10491"
                Else
                    o_array_instructions(3) = "NULL"
                End If
                'RATE UNIT MEASURE - BOLUS
                If ComboBox15.SelectedIndex = 2 Then
                    o_array_instructions(4) = "21"
                Else
                    o_array_instructions(4) = "9999"
                End If
                ''DURATION
                Dim l_duration As String = "NULL"
                Dim l_aux As Int64 = 0
                If TextBox29.Text <> "" Then
                    l_aux = TextBox29.Text * 60
                End If

                If TextBox18.Text <> "" And TextBox29.Text <> "" Then
                    l_aux = l_aux + TextBox18.Text
                ElseIf TextBox18.Text <> "" And TextBox29.Text = "" Then
                    l_aux = TextBox18.Text
                End If
                If l_aux > 0 Then
                    l_duration = l_aux
                End If
                o_array_instructions(5) = l_duration

                If l_duration = "NULL" Then
                    o_array_instructions(6) = "NULL"
                Else
                    o_array_instructions(6) = "10374"
                End If

            ElseIf i_index_instructions = 4 Then
                'DOSE
                If TextBox8.Text <> "" Then
                    o_array_instructions(0) = (TextBox8.Text).Replace(",", ".")
                Else
                    o_array_instructions(0) = "NULL"
                End If
                'DOSE UNIT MEASURE
                If ComboBox9.Text <> "" Then
                    o_array_instructions(1) = g_a_product_um(ComboBox9.SelectedIndex - 1) ''-1 PORQUE A 1ª POSIÇÃO DA COMBOBOX É NULL
                Else
                    o_array_instructions(1) = "NULL"
                End If
                'RATE
                If TextBox13.Text <> "" And ComboBox14.SelectedIndex <> 2 Then 'INDEX 2 REFERE-SE AO BOLUS
                    o_array_instructions(2) = (TextBox13.Text).Replace(",", ".")
                Else
                    o_array_instructions(2) = "NULL"
                End If
                'RATE UNIT MEASURE
                If ComboBox14.SelectedIndex = 1 Then
                    o_array_instructions(3) = "10491"
                Else
                    o_array_instructions(3) = "NULL"
                End If
                'RATE UNIT MEASURE - BOLUS
                If ComboBox14.SelectedIndex = 2 Then
                    o_array_instructions(4) = "21"
                Else
                    o_array_instructions(4) = "9999"
                End If
                ''DURATION
                Dim l_duration As String = "NULL"
                Dim l_aux As Int64 = 0
                If TextBox31.Text <> "" Then
                    l_aux = TextBox31.Text * 60
                End If

                If TextBox30.Text <> "" And TextBox31.Text <> "" Then
                    l_aux = l_aux + TextBox30.Text
                ElseIf TextBox30.Text <> "" And TextBox31.Text = "" Then
                    l_aux = TextBox30.Text
                End If
                If l_aux > 0 Then
                    l_duration = l_aux
                End If
                o_array_instructions(5) = l_duration

                If l_duration = "NULL" Then
                    o_array_instructions(6) = "NULL"
                Else
                    o_array_instructions(6) = "10374"
                End If

            ElseIf i_index_instructions = 5 Then
                'DOSE
                If TextBox9.Text <> "" Then
                    o_array_instructions(0) = (TextBox9.Text).Replace(",", ".")
                Else
                    o_array_instructions(0) = "NULL"
                End If
                'DOSE UNIT MEASURE
                If ComboBox10.Text <> "" Then
                    o_array_instructions(1) = g_a_product_um(ComboBox10.SelectedIndex - 1) ''-1 PORQUE A 1ª POSIÇÃO DA COMBOBOX É NULL
                Else
                    o_array_instructions(1) = "NULL"
                End If
                'RATE
                If TextBox12.Text <> "" And ComboBox13.SelectedIndex <> 2 Then 'INDEX 2 REFERE-SE AO BOLUS
                    o_array_instructions(2) = (TextBox12.Text).Replace(",", ".")
                Else
                    o_array_instructions(2) = "NULL"
                End If
                'RATE UNIT MEASURE
                If ComboBox13.SelectedIndex = 1 Then
                    o_array_instructions(3) = "10491"
                Else
                    o_array_instructions(3) = "NULL"
                End If
                'RATE UNIT MEASURE - BOLUS
                If ComboBox13.SelectedIndex = 2 Then
                    o_array_instructions(4) = "21"
                Else
                    o_array_instructions(4) = "9999"
                End If
                ''DURATION
                Dim l_duration As String = "NULL"
                Dim l_aux As Int64 = 0
                If TextBox28.Text <> "" Then
                    l_aux = TextBox28.Text * 60
                End If

                If TextBox19.Text <> "" And TextBox28.Text <> "" Then
                    l_aux = l_aux + TextBox19.Text
                ElseIf TextBox19.Text <> "" And TextBox28.Text = "" Then
                    l_aux = TextBox19.Text
                End If
                If l_aux > 0 Then
                    l_duration = l_aux
                End If
                o_array_instructions(5) = l_duration

                If l_duration = "NULL" Then
                    o_array_instructions(6) = "NULL"
                Else
                    o_array_instructions(6) = "10374"
                End If

            ElseIf i_index_instructions = 6 Then
                'DOSE
                If TextBox10.Text <> "" Then
                    o_array_instructions(0) = (TextBox10.Text).Replace(",", ".")
                Else
                    o_array_instructions(0) = "NULL"
                End If
                'DOSE UNIT MEASURE
                If ComboBox11.Text <> "" Then
                    o_array_instructions(1) = g_a_product_um(ComboBox11.SelectedIndex - 1) ''-1 PORQUE A 1ª POSIÇÃO DA COMBOBOX É NULL
                Else
                    o_array_instructions(1) = "NULL"
                End If
                'RATE
                If TextBox11.Text <> "" And ComboBox12.SelectedIndex <> 2 Then 'INDEX 2 REFERE-SE AO BOLUS
                    o_array_instructions(2) = (TextBox11.Text).Replace(",", ".")
                Else
                    o_array_instructions(2) = "NULL"
                End If
                'RATE UNIT MEASURE
                If ComboBox12.SelectedIndex = 1 Then
                    o_array_instructions(3) = "10491"
                Else
                    o_array_instructions(3) = "NULL"
                End If
                'RATE UNIT MEASURE - BOLUS
                If ComboBox12.SelectedIndex = 2 Then
                    o_array_instructions(4) = "21"
                Else
                    o_array_instructions(4) = "9999"
                End If
                ''DURATION
                Dim l_duration As String = "NULL"
                Dim l_aux As Int64 = 0
                If TextBox25.Text <> "" Then
                    l_aux = TextBox25.Text * 60
                End If

                If TextBox21.Text <> "" And TextBox25.Text <> "" Then
                    l_aux = l_aux + TextBox21.Text
                ElseIf TextBox21.Text <> "" And TextBox25.Text = "" Then
                    l_aux = TextBox21.Text
                End If
                If l_aux > 0 Then
                    l_duration = l_aux
                End If
                o_array_instructions(5) = l_duration

                If l_duration = "NULL" Then
                    o_array_instructions(6) = "NULL"
                Else
                    o_array_instructions(6) = "10374"
                End If

                ''construir função de gravaçao da item_seq inserindo o array o_array_instructions
            End If



        Catch ex As Exception
            Return False
        End Try

        Return True

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
            ComboBox29.Items.Add("")
            While dr_soft.Read()
                ComboBox2.Items.Add(dr_soft.Item(1))
                ComboBox29.Items.Add(dr_soft.Item(1))
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

        ComboBox31.Items.Add("")
        ComboBox31.Items.Add("0 - ALL")
        ComboBox31.Items.Add("1 - External Prescription")
        ComboBox31.Items.Add("2 - Administer Here")
        ComboBox31.Items.Add("3 - Home Medication")

        ComboBox24.Items.Add("Y")
        ComboBox24.Items.Add("N")

        Dim l_dr_admin_method As OracleDataReader
        If Not medication.GET_ADMIN_METHOD_LIST(g_id_institution, g_default_route, g_id_product_supplier, l_dr_admin_method) Then
            MsgBox("Error getting list of administration methods!", vbCritical)
        End If

        Dim i As Integer = 0

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
            MsgBox("Error getting list of administration sites!", vbCritical)
        End If

        ReDim g_a_admin_sites(0)
        Dim ii As Integer = 0
        While l_dr_admin_sites.Read()
            ComboBox26.Items.Add(l_dr_admin_sites.Item(1))
            ReDim Preserve g_a_admin_sites(ii)
            g_a_admin_sites(ii) = l_dr_admin_sites.Item(0)
            ii = ii + 1
        End While

        Dim l_dr_components_list As OracleDataReader
        If Not medication.GET_COMPONENTS_LIST(g_id_institution, g_id_product, g_id_product_supplier, l_dr_components_list) Then
            MsgBox("Error getting list of components!", vbCritical)
        End If

        ReDim g_a_components_list(0)
        ReDim g_a_doses_list(0)
        Dim i_cl As Integer = 0
        While l_dr_components_list.Read()
            ComboBox40.Items.Add(l_dr_components_list.Item(1))
            ReDim Preserve g_a_components_list(i_cl)
            g_a_components_list(i_cl) = l_dr_components_list.Item(0)

            ReDim Preserve g_a_doses_list(i_cl)
            g_a_doses_list(i_cl).id_product_component = l_dr_components_list.Item(0)
            g_a_doses_list(i_cl).dose_value_1 = -1
            g_a_doses_list(i_cl).dose_value_2 = -1
            g_a_doses_list(i_cl).dose_value_3 = -1
            g_a_doses_list(i_cl).dose_value_4 = -1
            g_a_doses_list(i_cl).dose_value_5 = -1
            g_a_doses_list(i_cl).dose_value_6 = -1
            g_a_doses_list(i_cl).dose_value_7 = -1
            g_a_doses_list(i_cl).id_unit_dose_1 = -1
            g_a_doses_list(i_cl).id_unit_dose_2 = -1
            g_a_doses_list(i_cl).id_unit_dose_3 = -1
            g_a_doses_list(i_cl).id_unit_dose_4 = -1
            g_a_doses_list(i_cl).id_unit_dose_5 = -1
            g_a_doses_list(i_cl).id_unit_dose_6 = -1
            g_a_doses_list(i_cl).id_unit_dose_7 = -1
            g_a_doses_list(i_cl).desc_unit_dose_1 = ""
            g_a_doses_list(i_cl).desc_unit_dose_2 = ""
            g_a_doses_list(i_cl).desc_unit_dose_3 = ""
            g_a_doses_list(i_cl).desc_unit_dose_4 = ""
            g_a_doses_list(i_cl).desc_unit_dose_5 = ""
            g_a_doses_list(i_cl).desc_unit_dose_6 = ""
            g_a_doses_list(i_cl).desc_unit_dose_7 = ""
            g_a_doses_list(i_cl).flg_updated = "N"

            i_cl = i_cl + 1
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

        'ReDim g_a_doses_list(0)

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

        'std_presc_dir_item
        g_id_std_presc_dir_item = -1

    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        ComboBox1.SelectedIndex = -1

        TextBox27.Text = ""
        TextBox26.Text = ""

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
                        Try
                            g_id_std_presc_dir_item = dr_std_presc_dir_item.Item(8)
                        Catch ex As Exception
                            g_id_std_presc_dir_item = -1
                        End Try

                    End While
                End If

            Dim dr_std_presc_dir_admix_seq As OracleDataReader
            Dim l_index_seq As Int16 = 0
            If Not medication.GET_STD_PRESC_DIR_ADMIX_SEQ(g_id_institution, g_id_std_presc_dir_item, dr_std_presc_dir_admix_seq) Then
                MsgBox("Error getting standard admixture seq instructions.", vbCritical)
            Else
                While dr_std_presc_dir_admix_seq.Read()
                    If l_index_seq = 0 Then
                        ''rate value
                        Try
                            TextBox17.Text = dr_std_presc_dir_admix_seq.Item(3)
                        Catch ex As Exception
                            TextBox17.Text = ""
                        End Try
                        ''rate desc
                        Try
                            ComboBox18.Text = dr_std_presc_dir_admix_seq.Item(5)
                        Catch ex As Exception
                            ComboBox18.Text = ""
                        End Try
                        'hours
                        Try
                            TextBox34.Text = dr_std_presc_dir_admix_seq.Item(1)
                        Catch ex As Exception
                            TextBox34.Text = ""
                        End Try
                        'minutes
                        Try
                            TextBox23.Text = dr_std_presc_dir_admix_seq.Item(2)
                        Catch ex As Exception
                            TextBox23.Text = ""
                        End Try

                    ElseIf l_index_seq = 1 Then
                        ''rate value
                        Try
                            TextBox16.Text = dr_std_presc_dir_admix_seq.Item(3)
                        Catch ex As Exception
                            TextBox16.Text = ""
                        End Try
                        ''rate desc
                        Try
                            ComboBox17.Text = dr_std_presc_dir_admix_seq.Item(5)
                        Catch ex As Exception
                            ComboBox17.Text = ""
                        End Try
                        'hours
                        Try
                            TextBox33.Text = dr_std_presc_dir_admix_seq.Item(1)
                        Catch ex As Exception
                            TextBox33.Text = ""
                        End Try
                        'minutes
                        Try
                            TextBox22.Text = dr_std_presc_dir_admix_seq.Item(2)
                        Catch ex As Exception
                            TextBox22.Text = ""
                        End Try

                    ElseIf l_index_seq = 2 Then
                        ''rate value
                        Try
                            TextBox15.Text = dr_std_presc_dir_admix_seq.Item(3)
                        Catch ex As Exception
                            TextBox15.Text = ""
                        End Try
                        ''rate desc
                        Try
                            ComboBox16.Text = dr_std_presc_dir_admix_seq.Item(5)
                        Catch ex As Exception
                            ComboBox16.Text = ""
                        End Try
                        'hours
                        Try
                            TextBox32.Text = dr_std_presc_dir_admix_seq.Item(1)
                        Catch ex As Exception
                            TextBox32.Text = ""
                        End Try
                        'minutes
                        Try
                            TextBox20.Text = dr_std_presc_dir_admix_seq.Item(2)
                        Catch ex As Exception
                            TextBox20.Text = ""
                        End Try

                    ElseIf l_index_seq = 3 Then
                        ''rate value
                        Try
                            TextBox14.Text = dr_std_presc_dir_admix_seq.Item(3)
                        Catch ex As Exception
                            TextBox14.Text = ""
                        End Try
                        ''rate desc
                        Try
                            ComboBox15.Text = dr_std_presc_dir_admix_seq.Item(5)
                        Catch ex As Exception
                            ComboBox15.Text = ""
                        End Try
                        'hours
                        Try
                            TextBox29.Text = dr_std_presc_dir_admix_seq.Item(1)
                        Catch ex As Exception
                            TextBox29.Text = ""
                        End Try
                        'minutes
                        Try
                            TextBox18.Text = dr_std_presc_dir_admix_seq.Item(2)
                        Catch ex As Exception
                            TextBox18.Text = ""
                        End Try

                    ElseIf l_index_seq = 4 Then
                        ''rate value
                        Try
                            TextBox13.Text = dr_std_presc_dir_admix_seq.Item(3)
                        Catch ex As Exception
                            TextBox13.Text = ""
                        End Try
                        ''rate desc
                        Try
                            ComboBox14.Text = dr_std_presc_dir_admix_seq.Item(5)
                        Catch ex As Exception
                            ComboBox14.Text = ""
                        End Try
                        'hours
                        Try
                            TextBox31.Text = dr_std_presc_dir_admix_seq.Item(1)
                        Catch ex As Exception
                            TextBox31.Text = ""
                        End Try
                        'minutes
                        Try
                            TextBox30.Text = dr_std_presc_dir_admix_seq.Item(2)
                        Catch ex As Exception
                            TextBox30.Text = ""
                        End Try

                    ElseIf l_index_seq = 5 Then
                        ''rate value
                        Try
                            TextBox12.Text = dr_std_presc_dir_admix_seq.Item(3)
                        Catch ex As Exception
                            TextBox12.Text = ""
                        End Try
                        ''rate desc
                        Try
                            ComboBox13.Text = dr_std_presc_dir_admix_seq.Item(5)
                        Catch ex As Exception
                            ComboBox13.Text = ""
                        End Try
                        'hours
                        Try
                            TextBox28.Text = dr_std_presc_dir_admix_seq.Item(1)
                        Catch ex As Exception
                            TextBox28.Text = ""
                        End Try
                        'minutes
                        Try
                            TextBox19.Text = dr_std_presc_dir_admix_seq.Item(2)
                        Catch ex As Exception
                            TextBox19.Text = ""
                        End Try

                    ElseIf l_index_seq = 6 Then
                        ''rate value
                        Try
                            TextBox11.Text = dr_std_presc_dir_admix_seq.Item(3)
                        Catch ex As Exception
                            TextBox11.Text = ""
                        End Try
                        ''rate desc
                        Try
                            ComboBox12.Text = dr_std_presc_dir_admix_seq.Item(5)
                        Catch ex As Exception
                            ComboBox12.Text = ""
                        End Try
                        'hours
                        Try
                            TextBox25.Text = dr_std_presc_dir_admix_seq.Item(1)
                        Catch ex As Exception
                            TextBox25.Text = ""
                        End Try
                        'minutes
                        Try
                            TextBox21.Text = dr_std_presc_dir_admix_seq.Item(2)
                        Catch ex As Exception
                            TextBox21.Text = ""
                        End Try
                    End If
                    l_index_seq = l_index_seq + 1
                End While
            End If

            'Só obtém valores se for selecionado o componente
            If ComboBox40.SelectedIndex > -1 Then

                l_index_seq = 0
                Dim dr_std_presc_dir_item_seq As OracleDataReader
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                If Not medication.GET_STD_PRESC_DIR_ITEM_SEQ_COMPLEX(g_id_institution, g_id_std_presc_dir_item, g_a_components_list(ComboBox40.SelectedIndex), g_id_product_supplier, dr_std_presc_dir_item_seq) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                    MsgBox("ERROR GETTING STANDARD_PRESC_DIR_ITEM_seq!", vbCritical)
                Else
                    While dr_std_presc_dir_item_seq.Read()
                            If l_index_seq = 0 Then
                                ''dose value
                                Try
                                    TextBox1.Text = dr_std_presc_dir_item_seq.Item(0)
                                Catch ex As Exception
                                    TextBox1.Text = ""
                                End Try
                            ''dose desc
                            Try
                                ComboBox4.Text = dr_std_presc_dir_item_seq.Item(2)
                            Catch ex As Exception
                                ComboBox4.Text = ""
                            End Try

                        ElseIf l_index_seq = 1 Then
                                ''dose value
                                Try
                                    TextBox5.Text = dr_std_presc_dir_item_seq.Item(0)
                                Catch ex As Exception
                                    TextBox5.Text = ""
                                End Try
                            ''dose desc
                            Try
                                ComboBox6.Text = dr_std_presc_dir_item_seq.Item(2)
                            Catch ex As Exception
                                ComboBox6.Text = ""
                            End Try

                        ElseIf l_index_seq = 2 Then
                                ''dose value
                                Try
                                    TextBox6.Text = dr_std_presc_dir_item_seq.Item(0)
                                Catch ex As Exception
                                    TextBox6.Text = ""
                                End Try
                            ''dose desc
                            Try
                                ComboBox7.Text = dr_std_presc_dir_item_seq.Item(2)
                            Catch ex As Exception
                                ComboBox7.Text = ""
                            End Try

                        ElseIf l_index_seq = 3 Then
                                ''dose value
                                Try
                                    TextBox7.Text = dr_std_presc_dir_item_seq.Item(0)
                                Catch ex As Exception
                                    TextBox7.Text = ""
                                End Try
                            ''dose desc
                            Try
                                ComboBox8.Text = dr_std_presc_dir_item_seq.Item(2)
                            Catch ex As Exception
                                ComboBox8.Text = ""
                            End Try

                        ElseIf l_index_seq = 4 Then
                                ''dose value
                                Try
                                    TextBox8.Text = dr_std_presc_dir_item_seq.Item(0)
                                Catch ex As Exception
                                    TextBox8.Text = ""
                                End Try
                            ''dose desc
                            Try
                                ComboBox9.Text = dr_std_presc_dir_item_seq.Item(2)
                            Catch ex As Exception
                                ComboBox9.Text = ""
                            End Try

                        ElseIf l_index_seq = 5 Then
                                ''dose value
                                Try
                                    TextBox9.Text = dr_std_presc_dir_item_seq.Item(0)
                                Catch ex As Exception
                                    TextBox9.Text = ""
                                End Try
                            ''dose desc
                            Try
                                ComboBox10.Text = dr_std_presc_dir_item_seq.Item(2)
                            Catch ex As Exception
                                ComboBox10.Text = ""
                            End Try

                        ElseIf l_index_seq = 6 Then
                            ''dose value
                            Try
                                TextBox10.Text = dr_std_presc_dir_item_seq.Item(0)
                            Catch ex As Exception
                                TextBox10.Text = ""
                            End Try
                            ''dose desc
                            Try
                                ComboBox11.Text = dr_std_presc_dir_item_seq.Item(2)
                            Catch ex As Exception
                                ComboBox11.Text = ""
                            End Try
                        End If
                            l_index_seq = l_index_seq + 1
                        End While
                    End If
                End If
                Cursor = Cursors.Arrow
            End If
    End Sub

    Private Sub ComboBox18_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox18.SelectedIndexChanged
        If ComboBox18.SelectedIndex = 2 Or ComboBox18.SelectedIndex = 0 Then
            TextBox17.Text = ""
        End If
    End Sub

    Private Sub ComboBox17_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox17.SelectedIndexChanged
        If ComboBox17.SelectedIndex = 2 Or ComboBox17.SelectedIndex = 0 Then
            TextBox16.Text = ""
        End If
    End Sub

    Private Sub ComboBox16_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox16.SelectedIndexChanged
        If ComboBox16.SelectedIndex = 2 Or ComboBox16.SelectedIndex = 0 Then
            TextBox15.Text = ""
        End If
    End Sub

    Private Sub ComboBox15_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox15.SelectedIndexChanged
        If ComboBox15.SelectedIndex = 2 Or ComboBox15.SelectedIndex = 0 Then
            TextBox14.Text = ""
        End If
    End Sub

    Private Sub ComboBox14_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox14.SelectedIndexChanged
        If ComboBox14.SelectedIndex = 2 Or ComboBox14.SelectedIndex = 0 Then
            TextBox13.Text = ""
        End If
    End Sub

    Private Sub ComboBox13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox13.SelectedIndexChanged
        If ComboBox13.SelectedIndex = 2 Or ComboBox13.SelectedIndex = 0 Then
            TextBox12.Text = ""
        End If
    End Sub

    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged
        If ComboBox12.SelectedIndex = 2 Or ComboBox12.SelectedIndex = 0 Then
            TextBox11.Text = ""
        End If
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged
        If TextBox17.Text <> "" And ComboBox18.SelectedIndex <> 1 Then
            ComboBox18.SelectedIndex = 1
        End If
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
        If TextBox16.Text <> "" And ComboBox17.SelectedIndex <> 1 Then
            ComboBox17.SelectedIndex = 1
        End If
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        If TextBox15.Text <> "" And ComboBox16.SelectedIndex <> 1 Then
            ComboBox16.SelectedIndex = 1
        End If
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        If TextBox14.Text <> "" And ComboBox15.SelectedIndex <> 1 Then
            ComboBox15.SelectedIndex = 1
        End If
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        If TextBox13.Text <> "" And ComboBox14.SelectedIndex <> 1 Then
            ComboBox14.SelectedIndex = 1
        End If
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        If TextBox12.Text <> "" And ComboBox13.SelectedIndex <> 1 Then
            ComboBox13.SelectedIndex = 1
        End If
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        If TextBox11.Text <> "" And ComboBox12.SelectedIndex <> 1 Then
            ComboBox12.SelectedIndex = 1
        End If
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
            'VERIFICAR SE NÃO EXISTE GRANT
            If TextBox27.Text = "" Or ComboBox1.SelectedIndex < 0 Then
                l_id_grant = medication.GET_ID_GRANT(g_id_institution, g_selected_software, "LNK_PRODUCT_STD_PRESC_DIR")
                'SE GRANT FOR = -1 ENTÃO É NECESSÁRIO CRIAR UM NOVO GRANT
                If l_id_grant = -1 Then
                    If Not medication.SET_ID_GRANT(g_id_institution, g_selected_software, "LNK_PRODUCT_STD_PRESC_DIR") Then
                        MsgBox("Error creating ID_GRANT!", vbCritical)
                        Cursor = Cursors.Arrow
                        Exit Sub
                    Else
                        l_id_grant = medication.GET_ID_GRANT(g_id_institution, g_selected_software, "LNK_PRODUCT_STD_PRESC_DIR")
                    End If
                End If

                If Not CREATE_SET_INSTRUCTIONS(l_id_grant, ComboBox28.SelectedIndex, "Y") Then
                    MsgBox("Error creating new set of instructions", vbCritical)
                    Cursor = Cursors.Arrow
                    Exit Sub
                End If
            Else
                'NESTE CASO JÁ EXISTIA INSTRUÇÃO. SERÁ FEITO UPDATE

                'VERIFICAR SE INSTRUÇÃO É UTILIZADA EM DIVERSAS PICK_LISTS/SOFTWARES
                'CASO SEJA É NECESSÁRIO PERGUNTAR SE SE FAZ UPDATE OU SE SE INSERE NOVA INSTRUÇÃO
                l_id_grant = TextBox27.Text
                Dim l_create_new As Integer = 0

                If medication.CHECK_DUP_INSTRUCTIONS(g_id_institution, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_std_presc_dir) > 1 And ComboBox28.SelectedIndex > 0 Then
                    l_create_new = MsgBox("The current standard instruction is also being usaed for other softwares and/or type of prescriptions. Do you wish to create a new instruction just for the selected software and type of prescription? (Responding 'No' will result on the update of the current standard instruction)", MessageBoxButtons.YesNo)
                End If

                If l_create_new = 0 Or l_create_new = DialogResult.No Then
                    'UPDATE LNK_PRODUCT_STD_INSTRUCTION               
                    Dim l_rank As Int64 = 0
                    If TextBox26.Text <> "" Then
                        l_rank = TextBox26.Text
                    Else
                        l_rank = ComboBox1.Text
                    End If

                    ''NESTE CASO É NECESSÁRIO FAZER UPDATE AO RANK
                    If Not medication.UPDATE_STD_PRESC_DIR(g_id_institution, g_id_product, g_id_product_supplier, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_std_presc_dir, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_grant, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_pick_list, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_std_presc_dir, l_rank, l_id_grant) Then
                        MsgBox("Error updating instruction rank!", vbCritical)
                        Cursor = Cursors.Arrow
                        Exit Sub
                    End If

                    If Not CREATE_SET_INSTRUCTIONS(TextBox27.Text, ComboBox28.SelectedIndex, "N") Then
                        MsgBox("Error creating new set of instructions", vbCritical)
                        Cursor = Cursors.Arrow
                        Exit Sub
                    End If

                Else
                    If Not CREATE_SET_INSTRUCTIONS(TextBox27.Text, ComboBox28.SelectedIndex, "Y") Then
                        MsgBox("Error creating new set of instructions", vbCritical)
                        Cursor = Cursors.Arrow
                        Exit Sub
                    End If
                End If
            End If

            'VERIFICAR SE É PARA COPIAR AS INTRUÇÕES PARA UM SEGUNDO SOFTWARE/PICK_LIST
            If (ComboBox29.SelectedIndex > 0 And ComboBox31.SelectedIndex > 0) Then
                Dim l_id_software_copy As Int16 = db_access_general.GET_SELECTED_SOFT(ComboBox29.SelectedIndex - 1, g_id_institution)
                l_id_grant = medication.GET_ID_GRANT(g_id_institution, l_id_software_copy, "LNK_PRODUCT_STD_PRESC_DIR")
                'SE GRANT FOR = -1 ENTÃO É NECESSÁRIO CRIAR UM NOVO GRANT

                If l_id_grant = -1 Then
                    If Not medication.SET_ID_GRANT(g_id_institution, l_id_software_copy, "LNK_PRODUCT_STD_PRESC_DIR") Then
                        MsgBox("Error creating ID_GRANT!", vbCritical)
                        Cursor = Cursors.Arrow
                        Exit Sub
                    Else
                        l_id_grant = medication.GET_ID_GRANT(g_id_institution, l_id_software_copy, "LNK_PRODUCT_STD_PRESC_DIR")
                    End If
                End If
                If Not CREATE_SET_INSTRUCTIONS(l_id_grant, ComboBox31.SelectedIndex - 1, "Y") Then
                    MsgBox("Error creating new set of instructions", vbCritical)
                    Cursor = Cursors.Arrow
                    Exit Sub
                End If
            End If

            Dim dr_med_set_instruction As OracleDataReader
            ReDim g_a_med_set_instructions(0)
            ComboBox1.Items.Clear()
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not medication.GET_ALL_INSTRUCTIONS(g_id_institution, g_selected_software, g_id_product, g_id_product_supplier, ComboBox28.SelectedIndex, dr_med_set_instruction) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                MsgBox("ERROR GETTING LIST OF STANDARD INSTRUCTIONS!", vbCritical)
                Cursor = Cursors.Arrow
                Exit Sub
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

            MsgBox("Record inserted.", vbInformation)

        End If

        Cursor = Cursors.Arrow
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        If ComboBox1.SelectedIndex < 0 Then
            MsgBox("Please select a standard instruction from the RANK dropdown menu to be deleted.", vbInformation)
        Else
            If Not medication.DELETE_STD_INSTRUCTION(g_id_institution, g_id_product, g_id_product_supplier, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_std_presc_dir, g_a_med_set_instructions(ComboBox1.SelectedIndex).rank, TextBox27.Text, ComboBox28.SelectedIndex) Then
                MsgBox("Error deleteing standard instruction!", vbCritical)
            Else
                MsgBox("Record deleted.", vbInformation)

                TextBox27.Text = ""

                TextBox26.Text = ""
                ComboBox29.SelectedIndex = -1

                RESET_SET_INSTRUCTIONS()
                RESET_MAIN_INSTRUCTIONS()

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

            End If

        End If

    End Sub

    'DOSE 1
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." AndAlso Not e.KeyChar = "," Then
            e.Handled = True
        End If
    End Sub
    'RATE1
    Private Sub TextBox17_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox17.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." AndAlso Not e.KeyChar = "," Then
            e.Handled = True
        End If
    End Sub
    'DOSE2
    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." AndAlso Not e.KeyChar = "," Then
            e.Handled = True
        End If
    End Sub
    'RATE2
    Private Sub TextBox16_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox16.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." AndAlso Not e.KeyChar = "," Then
            e.Handled = True
        End If
    End Sub
    'DOSE3
    Private Sub TextBox6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox6.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." AndAlso Not e.KeyChar = "," Then
            e.Handled = True
        End If
    End Sub
    'RATE3
    Private Sub TextBox15_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox15.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." AndAlso Not e.KeyChar = "," Then
            e.Handled = True
        End If
    End Sub
    'DOSE4
    Private Sub TextBox7_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox7.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." AndAlso Not e.KeyChar = "," Then
            e.Handled = True
        End If
    End Sub
    'RATE4
    Private Sub TextBox14_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox14.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." AndAlso Not e.KeyChar = "," Then
            e.Handled = True
        End If
    End Sub
    'DOSE5
    Private Sub TextBox8_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox8.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." AndAlso Not e.KeyChar = "," Then
            e.Handled = True
        End If
    End Sub
    'RATE5
    Private Sub TextBox13_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox13.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." AndAlso Not e.KeyChar = "," Then
            e.Handled = True
        End If
    End Sub
    'DOSE6
    Private Sub TextBox9_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox9.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." AndAlso Not e.KeyChar = "," Then
            e.Handled = True
        End If
    End Sub
    'RATE6
    Private Sub TextBox12_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox12.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." AndAlso Not e.KeyChar = "," Then
            e.Handled = True
        End If
    End Sub
    'DOSE5
    Private Sub TextBox10_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox10.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." AndAlso Not e.KeyChar = "," Then
            e.Handled = True
        End If
    End Sub
    'RATE5
    Private Sub TextBox11_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox11.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." AndAlso Not e.KeyChar = "," Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub ComboBox40_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox40.SelectedIndexChanged

        ''Gravar para a estrutura de doses antes de limpar
        If g_component_selected_index > -1 Then
            If TextBox1.Text <> "" And ComboBox4.SelectedIndex > 0 Then
                g_a_doses_list(g_component_selected_index).dose_value_1 = TextBox1.Text
            Else
                g_a_doses_list(g_component_selected_index).dose_value_1 = -1
            End If
            If TextBox5.Text <> "" And ComboBox6.SelectedIndex > 0 Then
                g_a_doses_list(g_component_selected_index).dose_value_2 = TextBox5.Text
            Else
                g_a_doses_list(g_component_selected_index).dose_value_2 = -1
            End If
            If TextBox6.Text <> "" And ComboBox7.SelectedIndex > 0 Then
                g_a_doses_list(g_component_selected_index).dose_value_3 = TextBox6.Text
            Else
                g_a_doses_list(g_component_selected_index).dose_value_3 = -1
            End If
            If TextBox7.Text <> "" And ComboBox8.SelectedIndex > 0 Then
                g_a_doses_list(g_component_selected_index).dose_value_4 = TextBox7.Text
            Else
                g_a_doses_list(g_component_selected_index).dose_value_4 = -1
            End If
            If TextBox8.Text <> "" And ComboBox9.SelectedIndex > 0 Then
                g_a_doses_list(g_component_selected_index).dose_value_5 = TextBox8.Text
            Else
                g_a_doses_list(g_component_selected_index).dose_value_5 = -1
            End If
            If TextBox9.Text <> "" And ComboBox10.SelectedIndex > 0 Then
                g_a_doses_list(g_component_selected_index).dose_value_6 = TextBox9.Text
            Else
                g_a_doses_list(g_component_selected_index).dose_value_6 = -1
            End If
            If TextBox10.Text <> "" And ComboBox11.SelectedIndex > 0 Then
                g_a_doses_list(g_component_selected_index).dose_value_7 = TextBox10.Text
            Else
                g_a_doses_list(g_component_selected_index).dose_value_7 = -1
            End If

            If TextBox1.Text <> "" And ComboBox4.SelectedIndex > 0 Then
                MsgBox(g_a_unit_measure.Count)
                g_a_doses_list(g_component_selected_index).id_unit_dose_1 = g_a_unit_measure(ComboBox4.SelectedIndex - 1).id_unit_measure
                g_a_doses_list(g_component_selected_index).desc_unit_dose_1 = g_a_unit_measure(ComboBox4.SelectedIndex - 1).unit_measure_desc
            Else
                g_a_doses_list(g_component_selected_index).id_unit_dose_1 = -1
                g_a_doses_list(g_component_selected_index).desc_unit_dose_1 = ""
            End If
            If TextBox5.Text <> "" And ComboBox6.SelectedIndex > 0 Then
                g_a_doses_list(g_component_selected_index).id_unit_dose_2 = g_a_unit_measure(ComboBox6.SelectedIndex - 1).id_unit_measure
                g_a_doses_list(g_component_selected_index).desc_unit_dose_2 = g_a_unit_measure(ComboBox6.SelectedIndex - 1).unit_measure_desc
            Else
                g_a_doses_list(g_component_selected_index).id_unit_dose_2 = -1
                g_a_doses_list(g_component_selected_index).desc_unit_dose_2 = ""
            End If
            If TextBox6.Text <> "" And ComboBox7.SelectedIndex > 0 Then
                g_a_doses_list(g_component_selected_index).id_unit_dose_3 = g_a_unit_measure(ComboBox7.SelectedIndex - 1).id_unit_measure
                g_a_doses_list(g_component_selected_index).desc_unit_dose_3 = g_a_unit_measure(ComboBox7.SelectedIndex - 1).unit_measure_desc
            Else
                g_a_doses_list(g_component_selected_index).id_unit_dose_3 = -1
                g_a_doses_list(g_component_selected_index).desc_unit_dose_3 = ""
            End If
            If TextBox7.Text <> "" And ComboBox8.SelectedIndex > 0 Then
                g_a_doses_list(g_component_selected_index).id_unit_dose_4 = g_a_unit_measure(ComboBox8.SelectedIndex - 1).id_unit_measure
                g_a_doses_list(g_component_selected_index).desc_unit_dose_4 = g_a_unit_measure(ComboBox8.SelectedIndex - 1).unit_measure_desc
            Else
                g_a_doses_list(g_component_selected_index).id_unit_dose_4 = -1
                g_a_doses_list(g_component_selected_index).desc_unit_dose_4 = ""
            End If
            If TextBox8.Text <> "" And ComboBox9.SelectedIndex > 0 Then
                g_a_doses_list(g_component_selected_index).id_unit_dose_5 = g_a_unit_measure(ComboBox9.SelectedIndex - 1).id_unit_measure
                g_a_doses_list(g_component_selected_index).desc_unit_dose_5 = g_a_unit_measure(ComboBox9.SelectedIndex - 1).unit_measure_desc
            Else
                g_a_doses_list(g_component_selected_index).id_unit_dose_5 = -1
                g_a_doses_list(g_component_selected_index).desc_unit_dose_5 = ""
            End If
            If TextBox9.Text <> "" And ComboBox10.SelectedIndex > 0 Then
                g_a_doses_list(g_component_selected_index).id_unit_dose_6 = g_a_unit_measure(ComboBox10.SelectedIndex - 1).id_unit_measure
                g_a_doses_list(g_component_selected_index).desc_unit_dose_6 = g_a_unit_measure(ComboBox10.SelectedIndex - 1).unit_measure_desc
            Else
                g_a_doses_list(g_component_selected_index).id_unit_dose_6 = -1
                g_a_doses_list(g_component_selected_index).desc_unit_dose_6 = ""
            End If
            If TextBox10.Text <> "" And ComboBox11.SelectedIndex > 0 Then
                g_a_doses_list(g_component_selected_index).id_unit_dose_7 = g_a_unit_measure(ComboBox11.SelectedIndex - 1).id_unit_measure
                g_a_doses_list(g_component_selected_index).desc_unit_dose_7 = g_a_unit_measure(ComboBox11.SelectedIndex - 1).unit_measure_desc
            Else
                g_a_doses_list(g_component_selected_index).id_unit_dose_7 = -1
                g_a_doses_list(g_component_selected_index).desc_unit_dose_7 = ""
            End If

            g_a_doses_list(g_component_selected_index).flg_updated = "Y"

        End If

        ''LIMPAR
        ComboBox4.Items.Clear()
        ComboBox6.Items.Clear()
        ComboBox7.Items.Clear()
        ComboBox8.Items.Clear()
        ComboBox9.Items.Clear()
        ComboBox10.Items.Clear()
        ComboBox11.Items.Clear()

        'UNIDADES DE MEDIDA DO COMPONENTE SELECIONADO
        Dim l_dr_product_um As OracleDataReader
        Dim i As Integer = 0
        If Not medication.GET_PRODUCT_UM(g_id_institution, g_a_components_list(ComboBox40.SelectedIndex), g_id_product_supplier, 1, l_dr_product_um) Then
            MsgBox("Error getting product unit measures!", vbCritical)
        Else

            ReDim g_a_product_um(0)

            Dim l_index_um As Integer = 0
            ReDim g_a_unit_measure(0)

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

                ReDim Preserve g_a_unit_measure(l_index_um)
                g_a_unit_measure(l_index_um).id_unit_measure = l_dr_product_um.Item(0)
                g_a_unit_measure(l_index_um).unit_measure_desc = l_dr_product_um.Item(1)
                l_index_um = l_index_um + 1

                ReDim Preserve g_a_product_um(i)
                g_a_product_um(i) = l_dr_product_um(0)
                i = i + 1
            End While
            l_dr_product_um.Close()
        End If

        'DOSES DO COMPONENTE SELECIONADO
        Dim l_index_seq As Integer = 0
        TextBox1.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""

        ComboBox4.Text = ""
        ComboBox6.Text = ""
        ComboBox7.Text = ""
        ComboBox8.Text = ""
        ComboBox9.Text = ""
        ComboBox10.Text = ""
        ComboBox11.Text = ""

        If g_a_doses_list(ComboBox40.SelectedIndex).flg_updated = "Y" Then
            If g_a_doses_list(ComboBox40.SelectedIndex).dose_value_1 > -1 Then
                TextBox1.Text = g_a_doses_list(ComboBox40.SelectedIndex).dose_value_1
            End If
            If g_a_doses_list(ComboBox40.SelectedIndex).dose_value_2 > -1 Then
                TextBox5.Text = g_a_doses_list(ComboBox40.SelectedIndex).dose_value_2
            End If
            If g_a_doses_list(ComboBox40.SelectedIndex).dose_value_3 > -1 Then
                TextBox6.Text = g_a_doses_list(ComboBox40.SelectedIndex).dose_value_3
            End If
            If g_a_doses_list(ComboBox40.SelectedIndex).dose_value_4 > -1 Then
                TextBox7.Text = g_a_doses_list(ComboBox40.SelectedIndex).dose_value_4
            End If
            If g_a_doses_list(ComboBox40.SelectedIndex).dose_value_5 > -1 Then
                TextBox8.Text = g_a_doses_list(ComboBox40.SelectedIndex).dose_value_5
            End If
            If g_a_doses_list(ComboBox40.SelectedIndex).dose_value_6 > -1 Then
                TextBox9.Text = g_a_doses_list(ComboBox40.SelectedIndex).dose_value_6
            End If
            If g_a_doses_list(ComboBox40.SelectedIndex).dose_value_7 > -1 Then
                TextBox10.Text = g_a_doses_list(ComboBox40.SelectedIndex).dose_value_7
            End If

            If g_a_doses_list(ComboBox40.SelectedIndex).id_unit_dose_1 > -1 Then
                ComboBox4.Text = g_a_doses_list(ComboBox40.SelectedIndex).desc_unit_dose_1
            End If
            If g_a_doses_list(ComboBox40.SelectedIndex).id_unit_dose_2 > -1 Then
                ComboBox6.Text = g_a_doses_list(ComboBox40.SelectedIndex).desc_unit_dose_2
            End If
            If g_a_doses_list(ComboBox40.SelectedIndex).id_unit_dose_3 > -1 Then
                ComboBox7.Text = g_a_doses_list(ComboBox40.SelectedIndex).desc_unit_dose_3
            End If
            If g_a_doses_list(ComboBox40.SelectedIndex).id_unit_dose_4 > -1 Then
                ComboBox8.Text = g_a_doses_list(ComboBox40.SelectedIndex).desc_unit_dose_4
            End If
            If g_a_doses_list(ComboBox40.SelectedIndex).id_unit_dose_5 > -1 Then
                ComboBox9.Text = g_a_doses_list(ComboBox40.SelectedIndex).desc_unit_dose_5
            End If
            If g_a_doses_list(ComboBox40.SelectedIndex).id_unit_dose_6 > -1 Then
                ComboBox10.Text = g_a_doses_list(ComboBox40.SelectedIndex).desc_unit_dose_6
            End If
            If g_a_doses_list(ComboBox40.SelectedIndex).id_unit_dose_7 > -1 Then
                ComboBox11.Text = g_a_doses_list(ComboBox40.SelectedIndex).desc_unit_dose_7
            End If

        Else
            Dim dr_std_presc_dir_item_seq As OracleDataReader
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not medication.GET_STD_PRESC_DIR_ITEM_SEQ_COMPLEX(g_id_institution, g_id_std_presc_dir_item, g_a_components_list(ComboBox40.SelectedIndex), g_id_product_supplier, dr_std_presc_dir_item_seq) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                MsgBox("ERROR GETTING STANDARD_PRESC_DIR_ITEM_seq!", vbCritical)
            Else
                While dr_std_presc_dir_item_seq.Read()
                If l_index_seq = 0 Then
                    ''dose value
                    Try
                        TextBox1.Text = dr_std_presc_dir_item_seq.Item(0)
                    Catch ex As Exception
                        TextBox1.Text = ""
                    End Try
                        ''dose desc
                        Try
                            ComboBox4.Text = dr_std_presc_dir_item_seq.Item(2)
                        Catch ex As Exception
                            ComboBox4.Text = ""
                        End Try

                ElseIf l_index_seq = 1 Then
                    ''dose value
                    Try
                        TextBox5.Text = dr_std_presc_dir_item_seq.Item(0)
                    Catch ex As Exception
                        TextBox5.Text = ""
                    End Try
                    ''dose desc
                    Try
                            ComboBox6.Text = dr_std_presc_dir_item_seq.Item(2)
                        Catch ex As Exception
                            ComboBox6.Text = ""
                        End Try

                ElseIf l_index_seq = 2 Then
                    ''dose value
                    Try
                        TextBox6.Text = dr_std_presc_dir_item_seq.Item(0)
                    Catch ex As Exception
                        TextBox6.Text = ""
                    End Try
                    ''dose desc
                    Try
                            ComboBox7.Text = dr_std_presc_dir_item_seq.Item(2)
                        Catch ex As Exception
                            ComboBox7.Text = ""
                        End Try

                ElseIf l_index_seq = 3 Then
                    ''dose value
                    Try
                        TextBox7.Text = dr_std_presc_dir_item_seq.Item(0)
                    Catch ex As Exception
                        TextBox7.Text = ""
                    End Try
                    ''dose desc
                    Try
                            ComboBox8.Text = dr_std_presc_dir_item_seq.Item(2)
                        Catch ex As Exception
                            ComboBox8.Text = ""
                        End Try

                ElseIf l_index_seq = 4 Then
                    ''dose value
                    Try
                        TextBox8.Text = dr_std_presc_dir_item_seq.Item(0)
                    Catch ex As Exception
                        TextBox8.Text = ""
                    End Try
                    ''dose desc
                    Try
                            ComboBox9.Text = dr_std_presc_dir_item_seq.Item(2)
                        Catch ex As Exception
                            ComboBox9.Text = ""
                        End Try

                ElseIf l_index_seq = 5 Then
                    ''dose value
                    Try
                        TextBox9.Text = dr_std_presc_dir_item_seq.Item(0)
                    Catch ex As Exception
                        TextBox9.Text = ""
                    End Try
                    ''dose desc
                    Try
                            ComboBox10.Text = dr_std_presc_dir_item_seq.Item(2)
                        Catch ex As Exception
                            ComboBox10.Text = ""
                        End Try

                ElseIf l_index_seq = 6 Then
                    ''dose value
                    Try
                        TextBox10.Text = dr_std_presc_dir_item_seq.Item(0)
                    Catch ex As Exception
                        TextBox10.Text = ""
                    End Try
                    ''dose desc
                    Try
                            ComboBox11.Text = dr_std_presc_dir_item_seq.Item(2)
                        Catch ex As Exception
                            ComboBox11.Text = ""
                        End Try
                End If
                l_index_seq = l_index_seq + 1
            End While
        End If
        End If

        g_component_selected_index = ComboBox40.SelectedIndex

    End Sub

End Class
﻿Imports Oracle.DataAccess.Client
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
    Dim g_a_frequencies() As Int64

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
        ComboBox8.Items.Clear()
        ComboBox11.Items.Clear()
        ComboBox20.Items.Clear()
        ComboBox17.Items.Clear()
        ComboBox14.Items.Clear()
        ComboBox23.Items.Clear()

        ComboBox3.SelectedIndex = -1
        ComboBox8.SelectedIndex = -1
        ComboBox11.SelectedIndex = -1
        ComboBox20.SelectedIndex = -1
        ComboBox17.SelectedIndex = -1
        ComboBox14.SelectedIndex = -1
        ComboBox23.SelectedIndex = -1

        Dim dr_freq As OracleDataReader
        ReDim g_a_frequencies(0)
        Dim l_index_freq As Int16 = 0
        If Not medication.GET_ALL_FREQS(g_id_institution, g_selected_software, i_prn, dr_freq) Then
            MsgBox("Error getting all frequencies")
        Else
            ComboBox3.Items.Add("")
            ComboBox8.Items.Add("")
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

                ReDim Preserve g_a_frequencies(l_index_freq)
                g_a_frequencies(l_index_freq) = dr_freq.Item(0)
                l_index_freq = l_index_freq + 1
            End While
        End If
        dr_freq.Dispose()
        dr_freq.Close()
    End Function

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

        ComboBox24.Items.Add("Y")
        ComboBox24.Items.Add("N")

        ComboBox31.Items.Add("")
        ComboBox31.Items.Add("0 - ALL")
        ComboBox31.Items.Add("1 - External Prescription")
        ComboBox31.Items.Add("2 - Administer Here")
        ComboBox31.Items.Add("3 - Home Medication")

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
        End If

        l_dr_product_um.Dispose()
        l_dr_product_um.Close()

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
        End If

        l_dr_duration_um.Dispose()
        l_dr_duration_um.Close()

        Dim l_dr_admin_method As OracleDataReader
        If Not medication.GET_ADMIN_METHOD_LIST(g_id_institution, g_default_route, g_id_product_supplier, l_dr_admin_method) Then
            MsgBox("Error getting list of administration methods!", vbCritical)
        End If

        ReDim g_a_admin_methods(0)
        Dim i_admin_method As Integer = 0
        While l_dr_admin_method.Read()
            ComboBox27.Items.Add(l_dr_admin_method.Item(1))
            ReDim Preserve g_a_admin_methods(i_admin_method)
            g_a_admin_methods(i_admin_method) = l_dr_admin_method(0)
            i_admin_method = i_admin_method + 1
        End While
        l_dr_admin_method.Dispose()
        l_dr_admin_method.Close()

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
        l_dr_admin_sites.Dispose()
        l_dr_admin_sites.Close()

    End Sub

    Function CHECK_NUMBER_INSTRUCTIONS() As Int16
        'VARIFICAR QUANTOS SETS DE INSTRUÇÕES DEVEM SER GRAVADOS
        Dim l_n_of_instruction As Int16 = -1
        '#1
        If (TextBox1.Text <> "" Or ComboBox4.Text <> "" Or ComboBox3.Text <> "" Or TextBox3.Text <> "" Or ComboBox5.Text <> "" Or TextBox4.Text <> "") Then
            l_n_of_instruction = l_n_of_instruction + 1
        Else
            Return l_n_of_instruction
        End If
        '#2
        If (TextBox8.Text <> "" Or ComboBox7.Text <> "" Or ComboBox8.Text <> "" Or TextBox7.Text <> "" Or ComboBox6.Text <> "" Or TextBox6.Text <> "") Then
            l_n_of_instruction = l_n_of_instruction + 1
        Else
            Return l_n_of_instruction
        End If
        '#3
        If (TextBox11.Text <> "" Or ComboBox10.Text <> "" Or ComboBox11.Text <> "" Or TextBox10.Text <> "" Or ComboBox9.Text <> "" Or TextBox9.Text <> "") Then
            l_n_of_instruction = l_n_of_instruction + 1
        Else
            Return l_n_of_instruction
        End If
        '#4
        If (TextBox20.Text <> "" Or ComboBox19.Text <> "" Or ComboBox20.Text <> "" Or TextBox19.Text <> "" Or ComboBox18.Text <> "" Or TextBox18.Text <> "") Then
            l_n_of_instruction = l_n_of_instruction + 1
        Else
            Return l_n_of_instruction
        End If
        '#5
        If (TextBox17.Text <> "" Or ComboBox16.Text <> "" Or ComboBox17.Text <> "" Or TextBox16.Text <> "" Or ComboBox15.Text <> "" Or TextBox15.Text <> "") Then
            l_n_of_instruction = l_n_of_instruction + 1
        Else
            Return l_n_of_instruction
        End If
        '#6
        If (TextBox14.Text <> "" Or ComboBox13.Text <> "" Or ComboBox14.Text <> "" Or TextBox13.Text <> "" Or ComboBox12.Text <> "" Or TextBox12.Text <> "") Then
            l_n_of_instruction = l_n_of_instruction + 1
        Else
            Return l_n_of_instruction
        End If
        '#7
        If (TextBox23.Text <> "" Or ComboBox22.Text <> "" Or ComboBox23.Text <> "" Or TextBox22.Text <> "" Or ComboBox21.Text <> "" Or TextBox21.Text <> "") Then
            l_n_of_instruction = l_n_of_instruction + 1
        Else
            Return l_n_of_instruction
        End If

        Return l_n_of_instruction

    End Function

    Function GET_INSTRUCTIONS(ByVal i_index_instructions As Int16, ByRef o_array_instructions() As String) As Boolean
        ReDim o_array_instructions(5)

        Try
            If i_index_instructions = 0 Then
                If TextBox1.Text <> "" Then
                    o_array_instructions(0) = TextBox1.Text
                Else
                    o_array_instructions(0) = "NULL"
                End If
                If ComboBox4.Text <> "" Then
                    o_array_instructions(1) = g_a_product_um(ComboBox4.SelectedIndex - 1) ''+1 PORQUE A 1ª POSIÇÃO DA COMBOBOX É NULL
                Else
                    o_array_instructions(1) = "NULL"
                End If
                If ComboBox3.Text <> "" Then
                    o_array_instructions(2) = g_a_frequencies(ComboBox3.SelectedIndex - 1)
                Else
                    o_array_instructions(2) = "NULL"
                End If
                If TextBox3.Text <> "" Then
                    o_array_instructions(3) = TextBox3.Text
                Else
                    o_array_instructions(3) = "NULL"
                End If
                If ComboBox5.Text <> "" Then
                    o_array_instructions(4) = g_a_duration_um(ComboBox5.SelectedIndex - 1)
                Else
                    o_array_instructions(4) = "NULL"
                End If
                If TextBox4.Text <> "" Then
                    o_array_instructions(5) = TextBox4.Text
                Else
                    o_array_instructions(5) = "NULL"
                End If

            ElseIf i_index_instructions = 1 Then
                If TextBox8.Text <> "" Then
                    o_array_instructions(0) = TextBox8.Text
                Else
                    o_array_instructions(0) = "NULL"
                End If
                If ComboBox7.Text <> "" Then
                    o_array_instructions(1) = g_a_product_um(ComboBox7.SelectedIndex - 1) ''+1 PORQUE A 1ª POSIÇÃO DA COMBOBOX É NULL
                Else
                    o_array_instructions(1) = "NULL"
                End If
                If ComboBox8.Text <> "" Then
                    o_array_instructions(2) = g_a_frequencies(ComboBox8.SelectedIndex - 1)
                Else
                    o_array_instructions(2) = "NULL"
                End If
                If TextBox7.Text <> "" Then
                    o_array_instructions(3) = TextBox7.Text
                Else
                    o_array_instructions(3) = "NULL"
                End If
                If ComboBox6.Text <> "" Then
                    o_array_instructions(4) = g_a_duration_um(ComboBox6.SelectedIndex - 1)
                Else
                    o_array_instructions(4) = "NULL"
                End If
                If TextBox6.Text <> "" Then
                    o_array_instructions(5) = TextBox6.Text
                Else
                    o_array_instructions(5) = "NULL"
                End If

            ElseIf i_index_instructions = 2 Then
                If TextBox11.Text <> "" Then
                    o_array_instructions(0) = TextBox11.Text
                Else
                    o_array_instructions(0) = "NULL"
                End If
                If ComboBox10.Text <> "" Then
                    o_array_instructions(1) = g_a_product_um(ComboBox10.SelectedIndex - 1) ''+1 PORQUE A 1ª POSIÇÃO DA COMBOBOX É NULL
                Else
                    o_array_instructions(1) = "NULL"
                End If
                If ComboBox11.Text <> "" Then
                    o_array_instructions(2) = g_a_frequencies(ComboBox11.SelectedIndex - 1)
                Else
                    o_array_instructions(2) = "NULL"
                End If
                If TextBox10.Text <> "" Then
                    o_array_instructions(3) = TextBox10.Text
                Else
                    o_array_instructions(3) = "NULL"
                End If
                If ComboBox9.Text <> "" Then
                    o_array_instructions(4) = g_a_duration_um(ComboBox9.SelectedIndex - 1)
                Else
                    o_array_instructions(4) = "NULL"
                End If
                If TextBox9.Text <> "" Then
                    o_array_instructions(5) = TextBox9.Text
                Else
                    o_array_instructions(5) = "NULL"
                End If

            ElseIf i_index_instructions = 3 Then
                If TextBox20.Text <> "" Then
                    o_array_instructions(0) = TextBox20.Text
                Else
                    o_array_instructions(0) = "NULL"
                End If
                If ComboBox19.Text <> "" Then
                    o_array_instructions(1) = g_a_product_um(ComboBox19.SelectedIndex - 1) ''+1 PORQUE A 1ª POSIÇÃO DA COMBOBOX É NULL
                Else
                    o_array_instructions(1) = "NULL"
                End If
                If ComboBox20.Text <> "" Then
                    o_array_instructions(2) = g_a_frequencies(ComboBox20.SelectedIndex - 1)
                Else
                    o_array_instructions(2) = "NULL"
                End If
                If TextBox19.Text <> "" Then
                    o_array_instructions(3) = TextBox19.Text
                Else
                    o_array_instructions(3) = "NULL"
                End If
                If ComboBox18.Text <> "" Then
                    o_array_instructions(4) = g_a_duration_um(ComboBox18.SelectedIndex - 1)
                Else
                    o_array_instructions(4) = "NULL"
                End If
                If TextBox18.Text <> "" Then
                    o_array_instructions(5) = TextBox18.Text
                Else
                    o_array_instructions(5) = "NULL"
                End If

            ElseIf i_index_instructions = 4 Then
                If TextBox17.Text <> "" Then
                    o_array_instructions(0) = TextBox17.Text
                Else
                    o_array_instructions(0) = "NULL"
                End If
                If ComboBox16.Text <> "" Then
                    o_array_instructions(1) = g_a_product_um(ComboBox16.SelectedIndex - 1) ''+1 PORQUE A 1ª POSIÇÃO DA COMBOBOX É NULL
                Else
                    o_array_instructions(1) = "NULL"
                End If
                If ComboBox17.Text <> "" Then
                    o_array_instructions(2) = g_a_frequencies(ComboBox17.SelectedIndex - 1)
                Else
                    o_array_instructions(2) = "NULL"
                End If
                If TextBox16.Text <> "" Then
                    o_array_instructions(3) = TextBox16.Text
                Else
                    o_array_instructions(3) = "NULL"
                End If
                If ComboBox15.Text <> "" Then
                    o_array_instructions(4) = g_a_duration_um(ComboBox15.SelectedIndex - 1)
                Else
                    o_array_instructions(4) = "NULL"
                End If
                If TextBox15.Text <> "" Then
                    o_array_instructions(5) = TextBox15.Text
                Else
                    o_array_instructions(5) = "NULL"
                End If

            ElseIf i_index_instructions = 5 Then
                If TextBox14.Text <> "" Then
                    o_array_instructions(0) = TextBox14.Text
                Else
                    o_array_instructions(0) = "NULL"
                End If
                If ComboBox13.Text <> "" Then
                    o_array_instructions(1) = g_a_product_um(ComboBox13.SelectedIndex - 1) ''+1 PORQUE A 1ª POSIÇÃO DA COMBOBOX É NULL
                Else
                    o_array_instructions(1) = "NULL"
                End If
                If ComboBox14.Text <> "" Then
                    o_array_instructions(2) = g_a_frequencies(ComboBox14.SelectedIndex - 1)
                Else
                    o_array_instructions(2) = "NULL"
                End If
                If TextBox13.Text <> "" Then
                    o_array_instructions(3) = TextBox13.Text
                Else
                    o_array_instructions(3) = "NULL"
                End If
                If ComboBox12.Text <> "" Then
                    o_array_instructions(4) = g_a_duration_um(ComboBox12.SelectedIndex - 1)
                Else
                    o_array_instructions(4) = "NULL"
                End If
                If TextBox12.Text <> "" Then
                    o_array_instructions(5) = TextBox12.Text
                Else
                    o_array_instructions(5) = "NULL"
                End If

            ElseIf i_index_instructions = 6 Then
                If TextBox23.Text <> "" Then
                    o_array_instructions(0) = TextBox23.Text
                Else
                    o_array_instructions(0) = "NULL"
                End If
                If ComboBox22.Text <> "" Then
                    o_array_instructions(1) = g_a_product_um(ComboBox22.SelectedIndex - 1) ''+1 PORQUE A 1ª POSIÇÃO DA COMBOBOX É NULL
                Else
                    o_array_instructions(1) = "NULL"
                End If
                If ComboBox23.Text <> "" Then
                    o_array_instructions(2) = g_a_frequencies(ComboBox23.SelectedIndex - 1)
                Else
                    o_array_instructions(2) = "NULL"
                End If
                If TextBox22.Text <> "" Then
                    o_array_instructions(3) = TextBox22.Text
                Else
                    o_array_instructions(3) = "NULL"
                End If
                If ComboBox21.Text <> "" Then
                    o_array_instructions(4) = g_a_duration_um(ComboBox21.SelectedIndex - 1)
                Else
                    o_array_instructions(4) = "NULL"
                End If
                If TextBox21.Text <> "" Then
                    o_array_instructions(5) = TextBox21.Text
                Else
                    o_array_instructions(5) = "NULL"
                End If
            End If

        Catch ex As Exception
            Return False
        End Try

        Return True

    End Function

    Function CREATE_SET_INSTRUCTIONS(ByVal i_id_grant As Int64, ByVal i_id_pick_list As Int16, ByVal i_create_new As String) As Boolean
        Try
            ''CRIAÇÃO DE UMA NOVA INSTRUÇÃO
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

            If Not medication.CREATE_STD_INSTRUCTION(g_id_institution, l_id_new_instruction, l_flg_sos, l_id_sos, l_sos_condition, TextBox5.Text, TextBox25.Text, l_id_admin_site, l_id_admin_method) Then
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
                If Not medication.UPDATE_STD_PRESC_DIR(g_id_institution, g_id_product, g_id_product_supplier, g_a_med_set_instructions(ComboBox1.SelectedIndex - 1).id_std_presc_dir, i_id_grant, i_id_pick_list, l_id_new_instruction, l_rank, i_id_grant) Then
                    MsgBox("Error updating lnk_product_std_presc_dir!", vbCritical)
                End If

            Else
                If Not medication.CREATE_STD_PRESC_DIR(g_id_institution, g_id_product, g_id_product_supplier, l_id_new_instruction, i_id_grant, i_id_pick_list, l_rank) Then
                    MsgBox("Error creating new lnk_product_std_presc_dir!!", vbCritical)
                End If
            End If

            Dim l_number_instructions_to_add As Int16 = CHECK_NUMBER_INSTRUCTIONS()
            If l_number_instructions_to_add > -1 Then
                Dim l_a_instructions() As String

                For i As Integer = 0 To l_number_instructions_to_add
                    GET_INSTRUCTIONS(i, l_a_instructions)
                    If Not medication.CREATE_STD_PRESC_DIR_ITEM(g_id_institution, l_id_new_instruction, i + 1, l_a_instructions) Then
                        MsgBox("Error creating standard prescription direction item!", vbCritical)
                    End If
                Next
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Function RESET_MAIN_INSTRUCTIONS()
        ComboBox24.SelectedIndex = -1
        ComboBox25.SelectedIndex = -1
        TextBox24.Text = ""
        ComboBox26.SelectedIndex = -1
        ComboBox27.SelectedIndex = -1
        TextBox5.Text = ""
        TextBox25.Text = ""
        TextBox26.Text = ""
        ComboBox29.SelectedIndex = -1
        ComboBox31.SelectedIndex = -1
        TextBox27.Text = ""
    End Function

    Function RESET_SET_INSTRUCTIONS()
        ''SETS DE INSTRUÇÕES
        TextBox1.Text = ""
        ComboBox4.Text = ""
        ComboBox3.Text = ""
        TextBox3.Text = ""
        ComboBox5.Text = ""
        TextBox4.Text = ""
        '#2
        TextBox8.Text = ""
        ComboBox7.Text = ""
        ComboBox8.Text = ""
        TextBox7.Text = ""
        ComboBox6.Text = ""
        TextBox6.Text = ""
        '#3
        TextBox11.Text = ""
        ComboBox10.Text = ""
        ComboBox11.Text = ""
        TextBox10.Text = ""
        ComboBox9.Text = ""
        TextBox9.Text = ""
        '#4
        TextBox20.Text = ""
        ComboBox19.Text = ""
        ComboBox20.Text = ""
        TextBox19.Text = ""
        ComboBox18.Text = ""
        TextBox18.Text = ""
        '#5
        TextBox17.Text = ""
        ComboBox16.Text = ""
        ComboBox17.Text = ""
        TextBox16.Text = ""
        ComboBox15.Text = ""
        TextBox15.Text = ""
        '#6
        TextBox14.Text = ""
        ComboBox13.Text = ""
        ComboBox14.Text = ""
        TextBox13.Text = ""
        ComboBox12.Text = ""
        TextBox12.Text = ""
        '#7
        TextBox23.Text = ""
        ComboBox22.Text = ""
        ComboBox23.Text = ""
        TextBox22.Text = ""
        ComboBox21.Text = ""
        TextBox21.Text = ""
    End Function

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
            ComboBox1.Items.Add("")
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
        dr_med_set_instruction.Dispose()
        dr_med_set_instruction.Close()

        Cursor = Cursors.Arrow
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        If ComboBox1.SelectedIndex > 0 Then ''estava -1

            Cursor = Cursors.WaitCursor

            RESET_MAIN_INSTRUCTIONS()
            RESET_SET_INSTRUCTIONS()

            TextBox27.Text = g_a_med_set_instructions(ComboBox1.SelectedIndex - 1).id_grant

            Dim dr_std_presc_dir As OracleDataReader
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not medication.GET_STD_PRESC_DIR(g_id_institution, g_a_med_set_instructions(ComboBox1.SelectedIndex - 1).id_std_presc_dir, dr_std_presc_dir) Then
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
            End If

            dr_std_presc_dir.Dispose()
            dr_std_presc_dir.Close()

            Dim dr_std_presc_dir_item As OracleDataReader
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not medication.GET_STD_PRESC_DIR_ITEM(g_id_institution, g_id_product, g_a_med_set_instructions(ComboBox1.SelectedIndex - 1).id_pick_list, g_a_med_set_instructions(ComboBox1.SelectedIndex - 1).id_std_presc_dir, g_a_med_set_instructions(ComboBox1.SelectedIndex - 1).id_grant, g_a_med_set_instructions(ComboBox1.SelectedIndex - 1).rank, dr_std_presc_dir_item) Then
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
            dr_std_presc_dir_item.Dispose()
            dr_std_presc_dir_item.Close()

            Cursor = Cursors.Arrow
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
            l_dr_sos.Dispose()
            l_dr_sos.Close()

        End If

        If ComboBox28.SelectedIndex > -1 Then

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
                ComboBox1.Items.Add("")
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
            dr_med_set_instruction.Dispose()
            dr_med_set_instruction.Close()
        End If
    End Sub

    Private Sub ComboBox25_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox25.SelectedIndexChanged
        If ComboBox25.SelectedIndex > -1 Then
            ComboBox24.Text = "Y"
            TextBox24.Text = ""
        End If
    End Sub

    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged
        If TextBox24.Text <> "" Then
            ComboBox25.SelectedIndex = -1
            ComboBox24.Text = "Y"
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Cursor = Cursors.WaitCursor
        Dim l_id_grant As Int64 = -1

        If ComboBox2.SelectedIndex < 0 Then
            MsgBox("Please select a software.", vbInformation)
        ElseIf ComboBox28.SelectedIndex < 0 Then
            MsgBox("Please select a type of prescription.", vbInformation)
        ElseIf ComboBox1.SelectedIndex < 1 And TextBox26.Text = "" Then ''estava 0
            MsgBox("Please select a rank.", vbInformation)
        ElseIf (TextBox3.Text = "" And
        ComboBox5.Text <> "") Then
            MsgBox("Please set a duration for instruction #1.", vbInformation)
        ElseIf (TextBox3.Text <> "" And
        ComboBox5.Text = "") Then
            MsgBox("Please select a duration unit measure for instruction #1.", vbInformation)
        ElseIf (TextBox7.Text = "" And
        ComboBox6.Text <> "") Then
            MsgBox("Please set a duration for instruction #2.", vbInformation)
        ElseIf (TextBox7.Text <> "" And
        ComboBox6.Text = "") Then
            MsgBox("Please select a duration unit measure for instruction #2.", vbInformation)
        ElseIf (TextBox10.Text = "" And
        ComboBox9.Text <> "") Then
            MsgBox("Please set a duration for instruction #3.", vbInformation)
        ElseIf (TextBox10.Text <> "" And
        ComboBox9.Text = "") Then
            MsgBox("Please select a duration unit measure for instruction #3.", vbInformation)
        ElseIf (TextBox19.Text = "" And
        ComboBox18.Text <> "") Then
            MsgBox("Please set a duration for instruction #4.", vbInformation)
        ElseIf (TextBox19.Text <> "" And
        ComboBox18.Text = "") Then
            MsgBox("Please select a duration unit measure for instruction #4.", vbInformation)
        ElseIf (TextBox16.Text = "" And
        ComboBox15.Text <> "") Then
            MsgBox("Please set a duration for instruction #5.", vbInformation)
        ElseIf (TextBox16.Text <> "" And
        ComboBox15.Text = "") Then
            MsgBox("Please select a duration unit measure for instruction #5.", vbInformation)
        ElseIf (TextBox13.Text = "" And
        ComboBox12.Text <> "") Then
            MsgBox("Please set a duration for instruction #6.", vbInformation)
        ElseIf (TextBox13.Text <> "" And
        ComboBox12.Text = "") Then
            MsgBox("Please select a duration unit measure for instruction #6.", vbInformation)
        ElseIf (TextBox22.Text = "" And
        ComboBox21.Text <> "") Then
            MsgBox("Please set a duration for instruction #7.", vbInformation)
        ElseIf (TextBox22.Text <> "" And
        ComboBox21.Text = "") Then
            MsgBox("Please select a duration unit measure for instruction #7.", vbInformation)
        ElseIf ((ComboBox2.text = ComboBox29.text) And (ComboBox28.text = ComboBox31.text)) Then
            MsgBox("The software and type of prescription to copy to cannot be the same as the software and type of prescription of origin. Please review the information.", vbInformation)
        Else
            'VERIFICAR SE NÃO EXISTE GRANT
            If TextBox27.Text = "" Or ComboBox1.SelectedIndex < 1 Then ''estava 0
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

                If medication.CHECK_DUP_INSTRUCTIONS(g_id_institution, g_a_med_set_instructions(ComboBox1.SelectedIndex - 1).id_std_presc_dir) > 1 And ComboBox28.SelectedIndex > 0 Then
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
                    If Not medication.UPDATE_STD_PRESC_DIR(g_id_institution, g_id_product, g_id_product_supplier, g_a_med_set_instructions(ComboBox1.SelectedIndex - 1).id_std_presc_dir, g_a_med_set_instructions(ComboBox1.SelectedIndex - 1).id_grant, g_a_med_set_instructions(ComboBox1.SelectedIndex - 1).id_pick_list, g_a_med_set_instructions(ComboBox1.SelectedIndex - 1).id_std_presc_dir, l_rank, l_id_grant) Then
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
                ComboBox1.Items.Add("")
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
            dr_med_set_instruction.Dispose()
            dr_med_set_instruction.Close()

            MsgBox("Record inserted.", vbInformation)
            End If
            Cursor = Cursors.Arrow
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        TextBox27.Text = ""

        ComboBox1.SelectedIndex = -1
        TextBox26.Text = ""
        ComboBox29.SelectedIndex = -1

        RESET_MAIN_INSTRUCTIONS()

        RESET_SET_INSTRUCTIONS()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        If ComboBox1.SelectedIndex < 1 Then
            MsgBox("Please select a standard instruction from the RANK dropdown menu to be deleted.", vbInformation)
        Else
            If Not medication.DELETE_STD_INSTRUCTION(g_id_institution, g_id_product, g_id_product_supplier, g_a_med_set_instructions(ComboBox1.SelectedIndex - 1).id_std_presc_dir, g_a_med_set_instructions(ComboBox1.SelectedIndex - 1).rank, TextBox27.Text, ComboBox28.SelectedIndex) Then
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
                    ComboBox1.Items.Add("")
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
                dr_med_set_instruction.Dispose()
                dr_med_set_instruction.Close()

            End If
        End If

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text <> "" Then
            TextBox3.Text = ""
            ComboBox5.Text = ""
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If TextBox6.Text <> "" Then
            TextBox7.Text = ""
            ComboBox6.Text = ""
        End If
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        If TextBox9.Text <> "" Then
            TextBox10.Text = ""
            ComboBox9.Text = ""
        End If
    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged
        If TextBox18.Text <> "" Then
            TextBox19.Text = ""
            ComboBox18.Text = ""
        End If
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        If TextBox15.Text <> "" Then
            TextBox16.Text = ""
            ComboBox15.Text = ""
        End If
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        If TextBox12.Text <> "" Then
            TextBox13.Text = ""
            ComboBox12.Text = ""
        End If
    End Sub

    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles TextBox21.TextChanged
        If TextBox21.Text <> "" Then
            TextBox22.Text = ""
            ComboBox21.Text = ""
        End If
    End Sub

    Private Sub ComboBox24_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox24.SelectedIndexChanged
        If ComboBox24.Text = "N" Then
            ComboBox25.SelectedIndex = -1
            TextBox24.Text = ""

            RESET_FREQUENCIES("N")
        ElseIf ComboBox24.Text = "Y" Then
            RESET_FREQUENCIES("Y")
        Else
            RESET_FREQUENCIES("N")
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text <> "" Then
            TextBox4.Text = ""
        End If
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        If ComboBox5.Text <> "" Then
            TextBox4.Text = ""
        End If
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If TextBox7.Text <> "" Then
            TextBox6.Text = ""
        End If
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        If ComboBox6.Text <> "" Then
            TextBox6.Text = ""
        End If
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        If TextBox10.Text <> "" Then
            TextBox9.Text = ""
        End If
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        If ComboBox9.Text <> "" Then
            TextBox9.Text = ""
        End If
    End Sub

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged
        If TextBox19.Text <> "" Then
            TextBox18.Text = ""
        End If
    End Sub

    Private Sub ComboBox18_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox18.SelectedIndexChanged
        If ComboBox18.Text <> "" Then
            TextBox18.Text = ""
        End If
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
        If TextBox16.Text <> "" Then
            TextBox15.Text = ""
        End If
    End Sub

    Private Sub ComboBox15_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox15.SelectedIndexChanged
        If ComboBox15.Text <> "" Then
            TextBox15.Text = ""
        End If
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        If TextBox13.Text <> "" Then
            TextBox12.Text = ""
        End If
    End Sub

    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged
        If ComboBox12.Text <> "" Then
            TextBox12.Text = ""
        End If
    End Sub

    Private Sub TextBox22_TextChanged(sender As Object, e As EventArgs) Handles TextBox22.TextChanged
        If TextBox22.Text <> "" Then
            TextBox21.Text = ""
        End If
    End Sub

    Private Sub ComboBox21_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox21.SelectedIndexChanged
        If ComboBox21.Text <> "" Then
            TextBox21.Text = ""
        End If
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
End Class
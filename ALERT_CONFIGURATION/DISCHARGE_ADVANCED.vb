﻿Imports Oracle.DataAccess.Client

Public Class DISCHARGE_ADVANCED

    Dim db_access_general As New General
    Dim db_discharge As New DISCHARGE_API
    Dim db_clin_serv As New CLINICAL_SERVICE_API

    'Variável que guarda o sotware selecionado
    Dim g_selected_soft As Int16 = -1

    'Array que vai guardar as REASONS caregadas do default
    Dim g_a_loaded_reasons_default() As DISCHARGE_API.DEFAULT_REASONS

    'Array que vai guardar as DESTINATIONS caregadas do default
    Dim g_a_loaded_destinations_default() As DISCHARGE_API.DEFAULT_REASONS

    'Array de profile templates disponíveis 
    Public Structure PROFILE_TEMPLATE
        Public ID_PROFILE_TEMPLATE As Integer
        Public PROFILE_NAME As String
        Public FLG_TYPE As String
    End Structure

    Dim g_a_profile_templates() As PROFILE_TEMPLATE

    'Array que vai guardar os ecrãs possíveis para configurar uma reason
    Dim g_a_screens() As String

    'Array com os clinical services da instituiçã/software
    Dim g_a_clin_serv_inst() As String

    'Array que guarda o tipo de profissionais a apresentar na lista
    Dim g_a_prof_types(5) As String

    'Estrutura que vai guardar as flags possíveis para perfis, e guardar se já existem perfis selecionados desse tipo
    Public Structure PROFILE_TEMPLATE_TYPE
        Public TYPE As String
        Public IS_TYPE_SELECTED As Boolean
    End Structure

    Dim g_a_profile_types() As PROFILE_TEMPLATE_TYPE

    'Array que vai concatenar os tipos de perfis distintos selecionados
    Dim g_a_prof_types_concat() As String

    Function reset_default_reasons()

        ReDim g_a_loaded_reasons_default(0)
        ComboBox3.Items.Clear()

        Dim dr_new As OracleDataReader

        If Not db_discharge.GET_ALL_DEFAULT_REASONS(TextBox1.Text, dr_new) Then

            MsgBox("ERROR GETING DEFAULT DISCHARGE REASONS.", vbCritical)

        Else

            Dim l_index_reason_default As Integer = 0
            ReDim g_a_loaded_reasons_default(0)

            While dr_new.Read()

                ReDim Preserve g_a_loaded_reasons_default(l_index_reason_default)
                g_a_loaded_reasons_default(l_index_reason_default).id_content = dr_new.Item(0)
                g_a_loaded_reasons_default(l_index_reason_default).desccription = dr_new.Item(1)
                l_index_reason_default = l_index_reason_default + 1

                ComboBox3.Items.Add(dr_new.Item(1) & "  -  [" & dr_new.Item(0) & "]")

            End While

        End If

    End Function

    Function reset_default_destinations()

        ReDim g_a_loaded_destinations_default(0)
        CheckedListBox1.Items.Clear()

        Dim dr_new As OracleDataReader

        If Not db_discharge.GET_ALL_DEFAULT_REASONS(TextBox1.Text, dr_new) Then

            MsgBox("ERROR GETING DEFAULT DISCHARGE DESTINATIONS.", vbCritical)

        Else

            Dim l_index_destinations_default As Integer = 0
            ReDim g_a_loaded_destinations_default(0)

            While dr_new.Read()

                ReDim Preserve g_a_loaded_destinations_default(l_index_destinations_default)
                g_a_loaded_destinations_default(l_index_destinations_default).id_content = dr_new.Item(0)
                g_a_loaded_destinations_default(l_index_destinations_default).desccription = dr_new.Item(1)
                l_index_destinations_default = l_index_destinations_default + 1

                CheckedListBox1.Items.Add(dr_new.Item(1) & "  -  [" & dr_new.Item(0) & "]")

            End While

        End If

    End Function

    Function reset_clin_serv()

        ReDim g_a_clin_serv_inst(0)
        ComboBox6.Items.Clear()

        Dim dr_new As OracleDataReader

        If Not db_clin_serv.GET_ALL_CLIN_SERV_INST(TextBox1.Text, g_selected_soft, dr_new) Then

            MsgBox("ERROR GETING CLINICAL SERVICES FROM INSTITUTION.", vbCritical)

        Else

            Dim l_index As Integer = 0
            ReDim g_a_clin_serv_inst(0)

            ComboBox6.Items.Add("None")

            While dr_new.Read()

                ReDim Preserve g_a_clin_serv_inst(l_index)
                g_a_clin_serv_inst(l_index) = dr_new.Item(0)
                l_index = l_index + 1

                ComboBox6.Items.Add(dr_new.Item(1) & "  -  [" & dr_new.Item(0) & "]")

            End While

        End If

    End Function

    Function reset_profile_types()

        For i As Integer = 0 To g_a_profile_types.Count() - 1

            g_a_profile_types(i).IS_TYPE_SELECTED = False

        Next

    End Function

    'Obter a lista de tipos de perfis selecionada
    Function check_profile_types(ByVal i_checked_profiles() As PROFILE_TEMPLATE, ByRef i_profile_types() As PROFILE_TEMPLATE_TYPE)

        Dim l_profile_type As String = ""

        For i As Integer = 0 To i_checked_profiles.Count() - 1

            l_profile_type = db_access_general.GET_PROFILE_TYPE(i_checked_profiles(i).ID_PROFILE_TEMPLATE)

            For j As Integer = 0 To i_profile_types.Count() - 1

                If i_profile_types(j).TYPE = l_profile_type Then

                    i_profile_types(j).IS_TYPE_SELECTED = True

                End If

            Next
        Next

    End Function

    'Concatenar os tipos numa só string para introduzir na discharge_reason
    Function concatentate_profiles(ByVal i_profiles() As PROFILE_TEMPLATE_TYPE, ByRef o_concatenated As String)

        o_concatenated = ""

        For i As Integer = 0 To i_profiles.Count() - 1

            If i_profiles(i).IS_TYPE_SELECTED = True Then

                o_concatenated = o_concatenated & i_profiles(i).TYPE

            End If

        Next

    End Function

    Function check_rank_integrity(ByVal i_rank As String) As Boolean

        'Código para ver se rank introduzido está correto
        Dim l_correct_rank As Boolean = True
        If i_rank <> "" Then

            '48 - 57  = Ascii codes for numbers
            For i As Integer = 0 To TextBox2.Text.Length() - 1

                If Asc(TextBox2.Text.Chars(i)) < 48 Or Asc(TextBox2.Text.Chars(i)) > 57 Then

                    l_correct_rank = False

                End If

            Next

        End If

        Return l_correct_rank

    End Function

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Cursor = Cursors.Arrow

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

        Else

            ComboBox1.Text = ""
            ComboBox2.Items.Clear()
            ComboBox2.Text = ""

        End If

        Cursor = Cursors.Arrow
    End Sub

    Private Sub DISCHARGE_ADVANCED_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "DISCHARGE  ::::  Connected to " & Connection.db

        Me.BackColor = Color.FromArgb(215, 215, 180)

        Me.Location = New Point(Form_location.x_position, Form_location.y_position)

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_ALL_INSTITUTIONS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING ALL INSTITUTIONS!")

        Else

            While dr.Read()

                ComboBox1.Items.Add(dr.Item(0))
                ComboBox7.Items.Add(dr.Item(0))

            End While

        End If

        'Tipos de profissionais
        'All
        'Physician
        'Nurse
        'Administrative
        'Other

        g_a_prof_types(0) = "-1"
        g_a_prof_types(1) = "D"
        g_a_prof_types(2) = "N"
        g_a_prof_types(3) = "A"
        g_a_prof_types(4) = "-2"

        ComboBox4.Items.Add("All")
        ComboBox4.Items.Add("Physician")
        ComboBox4.Items.Add("Nurse")
        ComboBox4.Items.Add("Administrative")
        ComboBox4.Items.Add("Other")

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_discharge.GET_REASON_SCREENS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING REASON SCREENS!")

        Else

            Dim l_dim_screens As Integer = 0
            ReDim g_a_screens(l_dim_screens)

            While dr.Read()

                ComboBox5.Items.Add(dr.Item(0))
                ReDim Preserve g_a_screens(l_dim_screens)
                g_a_screens(l_dim_screens) = dr.Item(0)
                l_dim_screens = l_dim_screens + 1

            End While

        End If

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_PROFILE_TYPES(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING PROFILE TYPES!")

        Else

            Dim l_dim_prof_types As Integer = 0
            ReDim g_a_profile_types(l_dim_prof_types)

            While dr.Read()

                ReDim Preserve g_a_profile_types(l_dim_prof_types)
                g_a_profile_types(l_dim_prof_types).TYPE = dr.Item(0)
                g_a_profile_types(l_dim_prof_types).IS_TYPE_SELECTED = False
                l_dim_prof_types = l_dim_prof_types + 1

            End While

        End If

        dr.Dispose()
        dr.Close()

        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        g_selected_soft = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, TextBox1.Text)

        'Obter as Reasons que existem no default (Mesmo as que estão not available)
        reset_default_reasons()

        ReDim g_a_profile_templates(0)
        ComboBox4.SelectedIndex = -1
        CheckedListBox2.Items.Clear()

        reset_clin_serv()

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        reset_default_destinations()

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_discharge.GET_DEFAULT_SCREEN(g_a_loaded_reasons_default(ComboBox3.SelectedIndex).id_content, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING DEFAULT REASON SCREEN!")

        Else

            While dr.Read()

                TextBox3.Text = dr.Item(0)

            End While

        End If

        dr.Dispose()
        dr.Close()

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        If ComboBox4.SelectedIndex > -1 Then


            ReDim g_a_profile_templates(0)
            CheckedListBox2.Items.Clear()

            Dim dr_new As OracleDataReader

            If Not db_access_general.GET_PROFILES(g_selected_soft, g_a_prof_types(ComboBox4.SelectedIndex), dr_new) Then

                MsgBox("ERROR GETING PROFILE TEMPLATES.", vbCritical)

            Else

                Dim l_index_prof_templates As Integer = 0
                ReDim g_a_profile_templates(0)

                While dr_new.Read()

                    ReDim Preserve g_a_profile_templates(l_index_prof_templates)
                    g_a_profile_templates(l_index_prof_templates).ID_PROFILE_TEMPLATE = dr_new.Item(0)
                    g_a_profile_templates(l_index_prof_templates).PROFILE_NAME = dr_new.Item(1)
                    g_a_profile_templates(l_index_prof_templates).FLG_TYPE = dr_new.Item(2)
                    l_index_prof_templates = l_index_prof_templates + 1

                    CheckedListBox2.Items.Add(dr_new.Item(0) & "  -  " & dr_new.Item(1))

                End While

            End If

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        TextBox1.Text = db_access_general.GET_INSTITUTION_ID(ComboBox1.SelectedIndex)

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""
        g_selected_soft = -1

        ReDim g_a_profile_templates(0)
        ComboBox4.SelectedIndex = -1
        CheckedListBox2.Items.Clear()

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

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If CheckedListBox2.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox2.Items.Count - 1

                CheckedListBox2.SetItemChecked(i, False)

            Next

        End If

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.MouseMove

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

        Cursor = Cursors.Arrow

        If TextBox4.Text <> "" Then

            ComboBox7.Text = db_access_general.GET_INSTITUTION(TextBox4.Text)

        Else

            ComboBox7.Text = ""

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Lista de Reasons
        If ComboBox3.SelectedIndex > -1 Then

            ''Lista de profissionais
            If CheckedListBox2.CheckedItems.Count() > 0 Then

                'Verificar se foi inserido rank para a Reason
                If TextBox2.Text <> "" Then

                    'Verificar integridade do rank inserido
                    If check_rank_integrity(TextBox2.Text) Then

                        'ARRAY QUE VAI GUARDAR OS PROFILE TEMPLATES SELECTIONADOS PELO UTILIZADOR
                        Dim l_a_selected_profiles_default() As PROFILE_TEMPLATE

                        ReDim l_a_selected_profiles_default(0)
                        Dim l_dim_selected_profiles = 0

                        For Each indexChecked In CheckedListBox2.CheckedIndices

                            ReDim Preserve l_a_selected_profiles_default(l_dim_selected_profiles)
                            l_a_selected_profiles_default(l_dim_selected_profiles).ID_PROFILE_TEMPLATE = g_a_profile_templates(indexChecked).ID_PROFILE_TEMPLATE
                            l_a_selected_profiles_default(l_dim_selected_profiles).PROFILE_NAME = g_a_profile_templates(indexChecked).PROFILE_NAME
                            l_a_selected_profiles_default(l_dim_selected_profiles).FLG_TYPE = g_a_profile_templates(indexChecked).FLG_TYPE
                            l_dim_selected_profiles = l_dim_selected_profiles + 1

                        Next

                        reset_profile_types()

                        check_profile_types(l_a_selected_profiles_default, g_a_profile_types)

                        Dim l_concatenated_distinct_profiles As String = ""

                        concatentate_profiles(g_a_profile_types, l_concatenated_distinct_profiles)
                        '-------------------------------------------------------------------------------

                        '2 - Verificar se existe Reason no ALERT (e respetiva tradução), caso não exista, inserir.
                        If Not db_discharge.CHECK_REASON(g_a_loaded_reasons_default(ComboBox3.SelectedIndex).id_content) Then

                            MsgBox("Não Existe")

                            If Not db_discharge.SET_MANUAL_REASON(TextBox1.Text, g_a_loaded_reasons_default(ComboBox3.SelectedIndex).id_content, l_concatenated_distinct_profiles, TextBox2.Text, ComboBox5.Text) Then

                                MsgBox("ERROR INSERTING DISCHARGE REASON!", vbCritical)

                            End If

                        ElseIf Not db_discharge.CHECK_REASON_translation(TextBox1.Text, g_a_loaded_reasons_default(ComboBox3.SelectedIndex).id_content) Then

                            If Not db_discharge.SET_REASON_TRANSLATION(TextBox1.Text, g_a_loaded_reasons_default(ComboBox3.SelectedIndex).id_content) Then

                                MsgBox("ERROR INSERTING DISCHARGE REASON TRANSLATION!", vbCritical)

                            End If

                            'Fazer Update()
                            If Not db_discharge.UPDATE_REASON(g_a_loaded_reasons_default(ComboBox3.SelectedIndex).id_content, l_concatenated_distinct_profiles, TextBox2.Text, ComboBox5.Text) Then

                                MsgBox("ERROR UPDATING DISCHARGE REASON!", vbCritical)

                            End If

                        Else

                            'FAZER  UPDATE
                            If Not db_discharge.UPDATE_REASON(g_a_loaded_reasons_default(ComboBox3.SelectedIndex).id_content, l_concatenated_distinct_profiles, TextBox2.Text, ComboBox5.Text) Then

                                MsgBox("ERROR UPDATING DISCHARGE REASON!", vbCritical)

                            End If

                        End If

                    Else
                        MsgBox("Please select a valid rank for the discharge reason.")
                    End If
                Else
                        MsgBox("Please set a rank for the discharge reason.")
                End If
            Else
                MsgBox("Please select, at least, one Profile.")
            End If
        Else
            MsgBox("Please select a discharge reason.")
        End If


        ''APAGAR (TESTE À UPDATE REASON)
        'Dim l_string(3) As String

        'l_string(0) = "D"
        'l_string(1) = "A"
        'l_string(2) = "S"
        'l_string(3) = "M"

        'If Not db_discharge.UPDATE_REASON("TMP39.259", l_string, 1, "DispositionCreateStep2LWBS.swf") Then

        'MsgBox("EROO")

        'End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        If CheckedListBox2.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox2.Items.Count - 1

                CheckedListBox2.SetItemChecked(i, True)

            Next

        End If

    End Sub

End Class
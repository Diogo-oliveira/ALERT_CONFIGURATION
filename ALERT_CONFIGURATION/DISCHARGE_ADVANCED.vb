Imports Oracle.DataAccess.Client

Public Class DISCHARGE_ADVANCED

    Dim db_access_general As New General
    Dim db_discharge As New DISCHARGE_API
    Dim db_clin_serv As New CLINICAL_SERVICE_API

    'Variável que guarda o sotware selecionado
    Dim g_selected_soft As Int16 = -1

    'Array que vai guardar as REASONS caregadas do default
    Dim g_a_loaded_reasons_default() As DISCHARGE_API.DEFAULT_REASONS
    'Array que vai guardar as REASONS caregadas do default
    Dim g_a_loaded_reasons_alert() As DISCHARGE_API.DEFAULT_DISCAHRGE

    'Array que vai guardar as DESTINATIONS caregadas do default
    Dim g_a_loaded_destinations_default() As DISCHARGE_API.DEFAULT_DISCAHRGE
    'Array que vai guardar as DESTINATIONS caregadas do alert
    Dim g_a_loaded_destinations_alert() As DISCHARGE_API.DEFAULT_DISCAHRGE

    'Array que vai gaurdar o tipo de episode types
    Dim g_a_loaded_eips_types() As Integer

    'Array de profile templates disponíveis 
    Public Structure PROFILE_TEMPLATE
        Public ID_PROFILE_TEMPLATE As Int64
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

    'Array que vai guardar os discharge_flash_files disponíveis
    Dim g_a_discharge_flash_files() As Integer

    Function reset_default_reasons()

        ReDim g_a_loaded_reasons_default(0)
        ComboBox3.SelectedIndex = -1
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

        dr_new.Dispose()
        dr_new.Close()

    End Function

    Function reset_alert_reasons()

        ReDim g_a_loaded_reasons_alert(0)
        ComboBox10.SelectedIndex = -1
        ComboBox10.Items.Clear()

        ComboBox13.SelectedIndex = -1
        ComboBox13.Items.Clear()

        Dim dr_new As OracleDataReader

        If Not db_discharge.GET_REASONS_ALERT(TextBox1.Text, dr_new) Then

            MsgBox("ERROR GETING DISCHARGE REASONS FROM ALERT.", vbCritical)

        Else

            Dim l_index_reason As Integer = 0
            ReDim g_a_loaded_reasons_alert(0)

            While dr_new.Read()

                ReDim Preserve g_a_loaded_reasons_alert(l_index_reason)
                g_a_loaded_reasons_alert(l_index_reason).id_content = dr_new.Item(0)
                g_a_loaded_reasons_alert(l_index_reason).description = dr_new.Item(1)
                l_index_reason = l_index_reason + 1

                ComboBox10.Items.Add(dr_new.Item(1) & "  -  [" & dr_new.Item(0) & "]")
                ComboBox13.Items.Add(dr_new.Item(1) & "  -  [" & dr_new.Item(0) & "]")

            End While

        End If

        dr_new.Dispose()
        dr_new.Close()

    End Function

    Function reset_alert_destinations()

        ReDim g_a_loaded_destinations_alert(0)
        ComboBox14.SelectedIndex = -1
        ComboBox14.Items.Clear()

        Dim dr_new As OracleDataReader

        If Not db_discharge.GET_DESTINATIONS_ALERT(TextBox1.Text, dr_new) Then

            MsgBox("ERROR GETING DISCHARGE DESTINATIONS FROM ALERT.", vbCritical)

        Else

            Dim l_index_destination As Integer = 0
            ReDim g_a_loaded_destinations_alert(0)

            While dr_new.Read()

                ReDim Preserve g_a_loaded_destinations_alert(l_index_destination)
                g_a_loaded_destinations_alert(l_index_destination).id_content = dr_new.Item(0)
                g_a_loaded_destinations_alert(l_index_destination).description = dr_new.Item(1)
                l_index_destination = l_index_destination + 1

                ComboBox14.Items.Add(dr_new.Item(1) & "  -  [" & dr_new.Item(0) & "]")

            End While

        End If

        dr_new.Dispose()
        dr_new.Close()

    End Function

    Function reset_default_destinations()

        ReDim g_a_loaded_destinations_default(0)
        ComboBox9.Items.Clear()
        ComboBox9.SelectedIndex = -1

        Dim dr_new As OracleDataReader

        If Not db_discharge.GET_ALL_DEFAULT_DESTINATIONS(TextBox1.Text, dr_new) Then

            MsgBox("ERROR GETING DEFAULT DISCHARGE DESTINATIONS.", vbCritical)

        Else

            Dim l_index_destinations_default As Integer = 0
            ReDim g_a_loaded_destinations_default(0)

            While dr_new.Read()

                ReDim Preserve g_a_loaded_destinations_default(l_index_destinations_default)
                g_a_loaded_destinations_default(l_index_destinations_default).id_content = dr_new.Item(0)
                g_a_loaded_destinations_default(l_index_destinations_default).description = dr_new.Item(1)
                l_index_destinations_default = l_index_destinations_default + 1

                ComboBox9.Items.Add(dr_new.Item(1) & "  -  [" & dr_new.Item(0) & "]")

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
            For i As Integer = 0 To i_rank.Length() - 1

                If Asc(i_rank.Chars(i)) < 48 Or Asc(i_rank.Chars(i)) > 57 Then

                    l_correct_rank = False

                End If

            Next

        End If

        Return l_correct_rank

    End Function

    Function clear_discharge_reason_box()

        'Reason
        ComboBox3.Items.Clear()
        ComboBox3.SelectedIndex = -1

        'Rank
        TextBox2.Text = ""

        'Default Screen
        TextBox3.Text = ""

        'Chosen Screen
        ComboBox5.SelectedIndex = -1

        'Type of discharge
        For i As Integer = 0 To CheckedListBox3.Items.Count - 1

            CheckedListBox3.SetItemChecked(i, False)

        Next

    End Function

    Function clear_discharge_destination_box()

        'Destination
        ComboBox9.Items.Clear()
        ComboBox9.SelectedIndex = -1

        'Rank
        TextBox5.Text = ""

        'Type of discharge
        For i As Integer = 0 To CheckedListBox4.Items.Count - 1

            CheckedListBox4.SetItemChecked(i, False)

        Next

    End Function

    Function clear_prof_disch_reas_box()

        'REASON
        ComboBox10.Items.Clear()
        ComboBox10.SelectedIndex = -1

        'FLAG DEFAULT
        ComboBox12.SelectedIndex = -1

        'Rank
        TextBox6.Text = ""

        'PROFESSIONAL
        ComboBox4.SelectedIndex = -1

        'DISCHARGE TYPE
        ComboBox11.SelectedIndex = -1

        'LIST OF PROFESIONALS
        For i As Integer = 0 To CheckedListBox2.Items.Count - 1

            CheckedListBox2.SetItemChecked(i, False)

        Next

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

            clear_discharge_reason_box()
            clear_discharge_destination_box()
            clear_prof_disch_reas_box()

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
        CheckedListBox1.BackColor = Color.FromArgb(195, 195, 165)
        CheckedListBox2.BackColor = Color.FromArgb(195, 195, 165)
        CheckedListBox3.BackColor = Color.FromArgb(195, 195, 165)
        CheckedListBox4.BackColor = Color.FromArgb(195, 195, 165)

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
        If Not db_discharge.GET_ALL_REASON_SCREENS(dr) Then
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

        ''DEFINIR OS TYPE OF DISCHARGE DIPONÍVEIS (Reasons)
        CheckedListBox3.Items.Add("Medical")
        CheckedListBox3.Items.Add("Nursing")
        CheckedListBox3.Items.Add("Administrative")
        CheckedListBox3.Items.Add("Social")
        CheckedListBox3.Items.Add("Triage")

        ''DEFINIR OS TYPE OF DISCHARGE DIPONÍVEIS (Destinations)
        CheckedListBox4.Items.Add("Medical")
        CheckedListBox4.Items.Add("Nursing")
        CheckedListBox4.Items.Add("Administrative")
        CheckedListBox4.Items.Add("Social")
        CheckedListBox4.Items.Add("Triage")

        'Obter os ecrãs de discharge que vão estar disponíveis para a tabela profile_disch_reas
        'São estes os ecrãs que de facto são apresentados na aplicação
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_discharge.GET_DISCH_FLASH_FILES(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING DISCHARGE FLASH FILES!")

        Else

            Dim l_dim_DISCH_FILES As Integer = 0
            ReDim g_a_discharge_flash_files(l_dim_DISCH_FILES)

            While dr.Read()

                ReDim Preserve g_a_discharge_flash_files(l_dim_DISCH_FILES)
                g_a_discharge_flash_files(l_dim_DISCH_FILES) = dr.Item(0)

                ComboBox11.Items.Add(dr.Item(1))

                l_dim_DISCH_FILES = l_dim_DISCH_FILES + 1

            End While

        End If

        'FLAG DEFAULT DO PROFILE_DISCH_REASON
        ComboBox12.Items.Add("Y")
        ComboBox12.Items.Add("N")

        'FLAG DEFAULT DO DISCH_REAS_DEST
        ComboBox15.Items.Add("Y")
        ComboBox15.Items.Add("N")

        'FLAG DISCHARGE DIAGNOSIS
        ComboBox8.Items.Add("Y")
        ComboBox8.Items.Add("N")

        'Preenchimento de Episode Types
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_EPIS_TYPES(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING EPISODE TYPES!")

        Else

            Dim l_dim_epis_types As Integer = 0
            ReDim g_a_loaded_eips_types(l_dim_epis_types)

            While dr.Read()

                ReDim Preserve g_a_loaded_eips_types(l_dim_epis_types)
                g_a_loaded_eips_types(l_dim_epis_types) = dr.Item(0)

                ComboBox16.Items.Add(dr.Item(0) & " - " & dr.Item(1))

                l_dim_epis_types = l_dim_epis_types + 1

            End While

        End If

        'Tipo de MCDTs que têm que estar executados ao dar discharge
        CheckedListBox1.Items.Add("Analysis")
        CheckedListBox1.Items.Add("Drugs")
        CheckedListBox1.Items.Add("Interventions")
        CheckedListBox1.Items.Add("Exams")
        CheckedListBox1.Items.Add("Continuous Medication")

        'OVERALL RESPONSABILITY REAS_DEST
        ComboBox17.Items.Add("Y")
        ComboBox17.Items.Add("N")

        ''AUTOMATIC CANCELATION OF PRESCRIPTIONS IN REAS_DEST
        ComboBox18.Items.Add("Y")
        ComboBox18.Items.Add("N")

        dr.Dispose()
        dr.Close()

        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        g_selected_soft = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, TextBox1.Text)

        clear_discharge_reason_box()
        clear_discharge_destination_box()
        clear_prof_disch_reas_box()

        'Obter as Reasons que existem no default (Mesmo as que estão not available)
        reset_default_reasons()

        'Obter as Destinations que existem no default (Mesmo as que estão not available)
        reset_default_destinations()

        'Obter as Reasons que estão available no ALERT
        reset_alert_reasons()

        'Obter as Destinations que estão available no ALERT
        reset_alert_destinations()

        ReDim g_a_profile_templates(0)
        ComboBox4.SelectedIndex = -1
        CheckedListBox2.Items.Clear()

        reset_clin_serv()

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        'Este código terá que passar para o quadro da REAS_DEST????

        'reset_default_destinations()

        ComboBox5.SelectedIndex = -1

        If ComboBox3.SelectedIndex > -1 Then

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

        End If

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

        clear_discharge_reason_box()
        clear_discharge_destination_box()
        clear_prof_disch_reas_box()

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


        If ComboBox13.SelectedIndex > -1 Then
            If ComboBox15.SelectedIndex > -1 Then
                'Verificar se foi inserido rank para a Reason
                If TextBox7.Text <> "" Then
                    'Verificar integridade do rank inserido
                    If check_rank_integrity(TextBox7.Text) Then

                        'TESTE DA FUNÇÃO COM DEP A -1
                        Dim O_DR As Int64
                        MsgBox(db_clin_serv.GET_DEP_CLIN_SERV(470, 11, -1, "TMP31.34", O_DR))
                        'FIM TESTE

                    Else
                        MsgBox("Please set a valid rank for this record.")
                    End If
                Else
                    MsgBox("Please set a rank for this record.")
                End If
            Else
                MsgBox("Please define if the record should be set as default.")
            End If
        Else
            MsgBox("Please select a discharge reason.")
        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        If CheckedListBox2.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox2.Items.Count - 1

                CheckedListBox2.SetItemChecked(i, True)

            Next

        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        'Lista de Reasons
        If ComboBox3.SelectedIndex > -1 Then
            'Verificar se foi inserido rank para a Reason
            If TextBox2.Text <> "" Then
                'Verificar integridade do rank inserido
                If check_rank_integrity(TextBox2.Text) Then
                    'verificar se foi escolido um ecrã
                    If ComboBox5.SelectedIndex > -1 Then
                        If CheckedListBox3.CheckedItems.Count() > 0 Then

                            '1 - TRATAR DO TIPO DE DISCHARGE
                            Dim l_selected_reas_disch_types As String = ""

                            For Each indexChecked In CheckedListBox3.CheckedIndices

                                If indexChecked = 0 Then
                                    'DOCTOR
                                    l_selected_reas_disch_types = l_selected_reas_disch_types & "D"

                                ElseIf indexChecked = 1 Then
                                    'NURSING
                                    l_selected_reas_disch_types = l_selected_reas_disch_types & "N"

                                ElseIf indexChecked = 2 Then
                                    'ADMINISTRATIVE
                                    l_selected_reas_disch_types = l_selected_reas_disch_types & "A"

                                ElseIf indexChecked = 3 Then
                                    'SOCIAL
                                    l_selected_reas_disch_types = l_selected_reas_disch_types & "S"

                                ElseIf indexChecked = 4 Then
                                    'TRIAGE
                                    l_selected_reas_disch_types = l_selected_reas_disch_types & "M"

                                End If

                            Next
                            ''-------------------------------------------------------------------------------

                            ''2 - Verificar se existe Reason no ALERT (e respetiva tradução), caso não exista, inserir.
                            If Not db_discharge.CHECK_REASON(g_a_loaded_reasons_default(ComboBox3.SelectedIndex).id_content) Then

                                If Not db_discharge.SET_MANUAL_REASON(TextBox1.Text, g_a_loaded_reasons_default(ComboBox3.SelectedIndex).id_content, l_selected_reas_disch_types, TextBox2.Text, ComboBox5.Text) Then

                                    MsgBox("ERROR INSERTING DISCHARGE REASON!", vbCritical)

                                End If

                            ElseIf Not db_discharge.CHECK_REASON_translation(TextBox1.Text, g_a_loaded_reasons_default(ComboBox3.SelectedIndex).id_content) Then

                                'Fazer Update()
                                If Not db_discharge.UPDATE_REASON(g_a_loaded_reasons_default(ComboBox3.SelectedIndex).id_content, l_selected_reas_disch_types, TextBox2.Text, ComboBox5.Text) Then

                                    MsgBox("ERROR UPDATING DISCHARGE REASON!", vbCritical)

                                End If

                                If Not db_discharge.SET_REASON_TRANSLATION(TextBox1.Text, g_a_loaded_reasons_default(ComboBox3.SelectedIndex).id_content) Then

                                    MsgBox("ERROR INSERTING DISCHARGE REASON TRANSLATION!", vbCritical)

                                End If

                            Else

                                'FAZER  UPDATE
                                If Not db_discharge.UPDATE_REASON(g_a_loaded_reasons_default(ComboBox3.SelectedIndex).id_content, l_selected_reas_disch_types, TextBox2.Text, ComboBox5.Text) Then

                                    MsgBox("ERROR UPDATING DISCHARGE REASON!", vbCritical)

                                End If

                            End If

                            MsgBox("Record correctly inserted.", vbInformation)

                            ''-------------------------------------------------------------------------------
                        Else
                                    MsgBox("Please select, at least, one discharge type.")
                        End If
                    Else
                            MsgBox("Please select a discharge screen.")
                End If
            Else
                MsgBox("Please select a valid rank for the discharge reason.")
            End If
        Else
            MsgBox("Please set a rank for the discharge reason.")
        End If
        Else
            MsgBox("Please select a discharge reason.")
        End If

    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        'Lista de DESTINATIONS
        If ComboBox9.SelectedIndex > -1 Then
            'Verificar se foi inserido rank para a Destinations
            If TextBox5.Text <> "" Then
                'Verificar integridade do rank inserido
                If check_rank_integrity(TextBox5.Text) Then
                    'vERIFICAR SE FOI ESCOLHIDO PELO MENOS UM TYPE DE DISCHARGE
                    If CheckedListBox4.CheckedItems.Count() > 0 Then

                        '1 - TRATAR DO TIPO DE DISCHARGE
                        Dim l_selected_dest_disch_types As String = ""

                        For Each indexChecked In CheckedListBox4.CheckedIndices

                            If indexChecked = 0 Then
                                'DOCTOR
                                l_selected_dest_disch_types = l_selected_dest_disch_types & "D"

                            ElseIf indexChecked = 1 Then
                                'NURSING
                                l_selected_dest_disch_types = l_selected_dest_disch_types & "N"

                            ElseIf indexChecked = 2 Then
                                'ADMINISTRATIVE
                                l_selected_dest_disch_types = l_selected_dest_disch_types & "A"

                            ElseIf indexChecked = 3 Then
                                'SOCIAL
                                l_selected_dest_disch_types = l_selected_dest_disch_types & "S"

                            ElseIf indexChecked = 4 Then
                                'TRIAGE
                                l_selected_dest_disch_types = l_selected_dest_disch_types & "M"

                            End If

                        Next
                        ''-------------------------------------------------------------------------------

                        ''2 - Verificar se existe DESTINATION no ALERT (e respetiva tradução), caso não exista, inserir.
                        If Not db_discharge.CHECK_DESTINATION(g_a_loaded_destinations_default(ComboBox9.SelectedIndex).id_content) Then

                            If Not db_discharge.SET_MANUAL_DESTINATION(TextBox1.Text, g_a_loaded_destinations_default(ComboBox9.SelectedIndex).id_content, TextBox5.Text, l_selected_dest_disch_types) Then

                                MsgBox("ERROR INSERTING DISCHARGE DESTINATION!", vbCritical)

                            End If

                        ElseIf Not db_discharge.CHECK_DESTINATION_TRANSLATION(TextBox1.Text, g_a_loaded_destinations_default(ComboBox9.SelectedIndex).id_content) Then

                            'Fazer Update()
                            If Not db_discharge.UPDATE_DESTINATION(g_a_loaded_destinations_default(ComboBox9.SelectedIndex).id_content, TextBox5.Text, l_selected_dest_disch_types) Then

                                MsgBox("ERROR UPDATING DISCHARGE DESTINATION!", vbCritical)

                            End If

                            If Not db_discharge.SET_DESTINATION_TRANSLATION(TextBox1.Text, g_a_loaded_destinations_default(ComboBox9.SelectedIndex)) Then

                                MsgBox("ERROR INSERTING DISCHARGE DESTINATION TRANSLATION!", vbCritical)

                            End If

                        Else

                            'FAZER  UPDATE
                            If Not db_discharge.UPDATE_DESTINATION(g_a_loaded_destinations_default(ComboBox9.SelectedIndex).id_content, TextBox5.Text, l_selected_dest_disch_types) Then

                                MsgBox("ERROR UPDATING DISCHARGE DESTINATION!", vbCritical)

                            End If

                        End If

                        MsgBox("Record correctly inserted.", vbInformation)

                        reset_alert_destinations()

                        ''-------------------------------------------------------------------------------
                    Else
                        MsgBox("Please select, at least, one discharge type.")
                    End If
                Else
                    MsgBox("Please select a valid rank for the discharge destination.")
                End If
            Else
                MsgBox("Please set a rank for the discharge destination.")
            End If
        Else
            MsgBox("Please select a discharge destination.")
        End If

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        'Lista de REASONS
        If ComboBox10.SelectedIndex > -1 Then
            'vERIFICAR SE FOI ESCOLHIDO PELO MENOS UM TYPE DE DISCHARGE
            If ComboBox11.SelectedIndex > -1 Then
                'Verificar se exstem profile templates selecionados
                If CheckedListBox2.CheckedItems.Count() > 0 Then
                    'Verificar integridade do rank inserido
                    If check_rank_integrity(TextBox6.Text) Then
                        If ComboBox12.SelectedIndex > -1 Then

                            '1 - TRATAR DO TIPO DE DISCHARGE
                            Dim l_selected_discharge_file As Integer = g_a_discharge_flash_files(ComboBox11.SelectedIndex)
                            ''-------------------------------------------------------------------------------

                            '2 - CORRER CADA PROFILE TEMLATE, OBTER A SUA FLAG_ACCESS E INSERIR NA PROFILE_DISCH_REASON
                            For Each indexChecked In CheckedListBox2.CheckedIndices

                                Dim l_flg_acces As String

                                l_flg_acces = db_access_general.GET_PROFILE_TYPE(g_a_profile_templates(indexChecked).ID_PROFILE_TEMPLATE)

                                If Not db_discharge.CHECK_PROF_DISCH_REASON(TextBox1.Text, g_a_loaded_reasons_alert(ComboBox10.SelectedIndex).id_content, g_a_profile_templates(indexChecked).ID_PROFILE_TEMPLATE) Then

                                    ''INSERT()
                                    If Not db_discharge.SET_MANUAL_PROFILE_DISCH_REASON(TextBox1.Text, g_a_loaded_reasons_alert(ComboBox10.SelectedIndex).id_content, g_a_profile_templates(indexChecked).ID_PROFILE_TEMPLATE, l_selected_discharge_file, l_flg_acces, TextBox6.Text, ComboBox12.Text) Then

                                        MsgBox("ERROR INSERTING PROFILE_DISCHARGE_REASON!", vbCritical)

                                    End If

                                Else
                                    'UPDATE()
                                    If Not db_discharge.UPDATE_PROF_DISCH_REASON(TextBox1.Text, g_a_loaded_reasons_alert(ComboBox10.SelectedIndex).id_content, g_a_profile_templates(indexChecked).ID_PROFILE_TEMPLATE, l_selected_discharge_file, l_flg_acces, TextBox6.Text, ComboBox12.Text) Then

                                        MsgBox("ERROR UPDATING PROFILE_DISCHARGE_REASON!", vbCritical)

                                    End If

                                End If

                            Next

                            MsgBox("Record correctly inserted.", vbInformation)

                            ''-------------------------------------------------------------------------------
                        Else
                            MsgBox("Please select a default status.")
                        End If
                    Else
                            MsgBox("Please select a valid rank for the discharge reason.")
                    End If
                Else
                        MsgBox("Please select, at least, one profile template.")
                End If
            Else
                    MsgBox("Please select a discharge type.")
            End If
        Else
            MsgBox("Please select a discharge reason.")
        End If

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        DISCHARGE.Show()
        Me.Close()

    End Sub

End Class
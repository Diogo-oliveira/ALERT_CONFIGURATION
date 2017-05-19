'TO DO:
' Ver as reasons que estão associadas ao clinical service => ver se todas as reasons devem ter de facto associação na dest_reason
'Nota: Exsitem registos na disch-reas_dest sem dest e com clinical service
'1- Criar uma clasee CLinical Service 
'1.1 - Criar uma função que verifique se o clinical service está disponível no alert - OK
'1.2 - Criar uma função para verificar se o dep_clin_serv existe - OK
'1.3 - Criar uma função para inserir clinical services - OK
'2 - Adpatar a função de versão, reason e dest para mostrar as reasons sem dest mas com clinical service

'3-Pensar numa função para devolver os profissionais associados a cada reason/dest

Imports Oracle.DataAccess.Client
Public Class DISCHARGE

    Dim db_access_general As New General
    Dim db_discharge As New DISCHARGE_API
    Dim db_clin_serv As New CLINICAL_SERVICE_API

    'Variável que guarda o sotware selecionado
    Dim g_selected_soft As Int16 = -1

    'Array que vai guardar as REASONS caregadas do default
    Dim g_a_loaded_reasons_default() As DISCHARGE_API.DEFAULT_REASONS

    'Array que vai guardar as REASONS caregadas do ALERT
    Dim g_a_loaded_reasons_alert() As DISCHARGE_API.DEFAULT_REASONS

    'Array que vai guardar as DESTINATIONS caregadas do default
    Dim g_a_loaded_destinations_default() As DISCHARGE_API.DEFAULT_DISCAHRGE

    'Array que vai guardar as DESTINATIONS caregadas do ALERT
    Dim g_a_loaded_destinations_alert() As DISCHARGE_API.DEFAULT_DISCAHRGE

    'ARRAY QUE VAI GUARDAR OS PROFILE TEMPLATES DA REASON SELECIONADA VINDOS DO DEFAULT
    Dim g_a_loaded_profiles_default() As DISCHARGE_API.DEFAULT_DISCH_PROFILE

    'ARRAY QUE VAI GUARDAR OS PROFILE TEMPLATES DA REASON SELECIONADA DO ALERT
    Dim g_a_loaded_profiles_alert() As DISCHARGE_API.DEFAULT_DISCH_PROFILE

    'Array que vai guardar os id_content dos grupos do default
    Dim g_a_loaded_instr_group() As DISCHARGE_API.DEFAULT_INSTR

    'Array que vai guardar os id_content dos grupos do ALERT
    Dim g_a_loaded_instr_group_alert() As DISCHARGE_API.DEFAULT_INSTR

    'Array que vai guardar os id_content DAS INSTRUCTIONS do DEFAULT
    Dim g_a_loaded_instr() As DISCHARGE_API.DEFAULT_INSTR

    'Array que vai guardar os id_content DAS INSTRUCTIONS do ALERT
    Dim g_a_loaded_instr_alert() As DISCHARGE_API.DEFAULT_INSTR

    Function RESET_REASONS_ALERT()

        ReDim g_a_loaded_reasons_alert(0)
        ComboBox7.Items.Clear()

        Dim dr_reason_alert As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_discharge.GET_REASONS(TextBox1.Text, g_selected_soft, dr_reason_alert) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING DISCHARGE REASONS FROM ALERT!", vbCritical)

        Else

            Dim l_dim_reas_alert As Integer = 0
            While dr_reason_alert.Read()

                ComboBox7.Items.Add(dr_reason_alert.Item(1) & "  -  [" & dr_reason_alert.Item(0) & "]")
                ReDim Preserve g_a_loaded_reasons_alert(l_dim_reas_alert)

                g_a_loaded_reasons_alert(l_dim_reas_alert).id_content = dr_reason_alert.Item(0)

                Try
                    g_a_loaded_reasons_alert(l_dim_reas_alert).desccription = dr_reason_alert.Item(1)
                Catch ex As Exception
                    g_a_loaded_reasons_alert(l_dim_reas_alert).desccription = ""
                End Try

                l_dim_reas_alert = l_dim_reas_alert + 1

            End While

        End If

        dr_reason_alert.Dispose()
        dr_reason_alert.Close()

    End Function

    Function RESET_DESTINATIONS_ALERT()

        ReDim g_a_loaded_destinations_alert(0)
        CheckedListBox4.Items.Clear()

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_discharge.GET_DESTINATIONS(TextBox1.Text, g_selected_soft, g_a_loaded_reasons_alert(ComboBox7.SelectedIndex).id_content, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING DISCHARGE DESTINATIONS FROM ALERT!", vbCritical)

        Else

            Dim l_dim_dest_alert As Integer = 0
            While dr.Read()

                'Só populo os 3 primeiros campos
                CheckedListBox4.Items.Add(dr.Item(2) & "  -  [" & dr.Item(1) & "]")

                ReDim Preserve g_a_loaded_destinations_alert(l_dim_dest_alert)

                g_a_loaded_destinations_alert(l_dim_dest_alert).id_disch_reas_dest = dr.Item(0)
                g_a_loaded_destinations_alert(l_dim_dest_alert).id_content = dr.Item(1)

                Try
                    g_a_loaded_destinations_alert(l_dim_dest_alert).description = dr.Item(2)
                Catch ex As Exception
                    g_a_loaded_destinations_alert(l_dim_dest_alert).description = ""
                End Try


                l_dim_dest_alert = l_dim_dest_alert + 1

            End While

        End If

        dr.Dispose()
        dr.Close()

    End Function

    Function RESET_PROFILES_ALERT()

        Dim dr As OracleDataReader

        ReDim g_a_loaded_profiles_alert(0)
        CheckedListBox5.Items.Clear()

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_discharge.GET_PROFILE_DISCH_REASON(TextBox1.Text, g_selected_soft, g_a_loaded_reasons_alert(ComboBox7.SelectedIndex).id_content, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING DISCHARGE PROFILES FROM ALERT!", vbCritical)

        Else

            Dim l_dim_prof_alert As Integer = 0
            While dr.Read()

                'Só populo os 3 primeiros campos
                CheckedListBox5.Items.Add(dr.Item(1) & " - " & dr.Item(2))

                ReDim Preserve g_a_loaded_profiles_alert(l_dim_prof_alert)

                g_a_loaded_profiles_alert(l_dim_prof_alert).ID_PROFILE_DISCH_REASON = dr.Item(0)
                g_a_loaded_profiles_alert(l_dim_prof_alert).ID_PROFILE_TEMPLATE = dr.Item(1)
                g_a_loaded_profiles_alert(l_dim_prof_alert).PROFILE_NAME = dr.Item(2)

                l_dim_prof_alert = l_dim_prof_alert + 1

            End While

        End If

        dr.Dispose()
        dr.Close()

    End Function

    Function RESET_GROUP_INSTR_ALERT()

        Dim dr As OracleDataReader

        ReDim g_a_loaded_instr_group_alert(0)
        ComboBox8.Items.Clear()

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_discharge.GET_ALERT_INSTR_GROUP(TextBox1.Text, g_selected_soft, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING DISCHARGE INSTRUCTIONS GROUPS FROM ALERT!", vbCritical)

        Else

            Dim l_dim_group As Integer = 0
            While dr.Read()

                ComboBox8.Items.Add(dr.Item(1) & " - " & dr.Item(0))

                ReDim Preserve g_a_loaded_instr_group_alert(l_dim_group)

                g_a_loaded_instr_group_alert(l_dim_group).ID_CONTENT = dr.Item(0)
                g_a_loaded_instr_group_alert(l_dim_group).DESCRIPTION = dr.Item(1)

                l_dim_group = l_dim_group + 1

            End While

        End If

        dr.Dispose()
        dr.Close()

    End Function
    Function RESET_INSTR_ALERT()

        Dim dr As OracleDataReader

        ReDim g_a_loaded_instr_alert(0)
        CheckedListBox6.Items.Clear()

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_discharge.GET_ALERT_INSTR(TextBox1.Text, g_selected_soft, g_a_loaded_instr_group_alert(ComboBox8.SelectedIndex).ID_CONTENT, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING DISCHARGE INSTRUCTIONS FROM ALERT!", vbCritical)

        Else

            Dim l_dim_instr As Integer = 0
            While dr.Read()

                CheckedListBox6.Items.Add(dr.Item(1) & " - " & dr.Item(0))

                ReDim Preserve g_a_loaded_instr_alert(l_dim_instr)

                g_a_loaded_instr_alert(l_dim_instr).ID_CONTENT = dr.Item(0)
                g_a_loaded_instr_alert(l_dim_instr).DESCRIPTION = dr.Item(1)

                l_dim_instr = l_dim_instr + 1

            End While

        End If

        dr.Dispose()
        dr.Close()

    End Function

    Private Sub DISCHARGE_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "DISCHARGE  ::::  Connected to " & Connection.db

        Me.BackColor = Color.FromArgb(215, 215, 180)

        Me.Location = New Point(Form_location.x_position, Form_location.y_position)

        CheckedListBox2.BackColor = Color.FromArgb(195, 195, 165)
        CheckedListBox1.BackColor = Color.FromArgb(195, 195, 165)
        CheckedListBox3.BackColor = Color.FromArgb(195, 195, 165)
        CheckedListBox4.BackColor = Color.FromArgb(195, 195, 165)
        CheckedListBox5.BackColor = Color.FromArgb(195, 195, 165)
        CheckedListBox6.BackColor = Color.FromArgb(195, 195, 165)

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

        CheckedListBox1.HorizontalScrollbar = True

        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Cursor = Cursors.Arrow

        ComboBox2.Items.Clear()
        g_selected_soft = -1
        ComboBox3.Items.Clear()
        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()
        ReDim g_a_loaded_instr_group(0)

        ComboBox4.Items.Clear()
        ReDim g_a_loaded_reasons_default(0)

        CheckedListBox1.Items.Clear()
        ReDim g_a_loaded_destinations_default(0)

        CheckedListBox2.Items.Clear()
        ReDim g_a_loaded_profiles_default(0)

        ReDim g_a_loaded_instr(0)
        CheckedListBox3.Items.Clear()

        ReDim g_a_loaded_reasons_alert(0)
        ComboBox7.Items.Clear()

        ReDim g_a_loaded_destinations_alert(0)
        CheckedListBox4.Items.Clear()

        ReDim g_a_loaded_profiles_alert(0)
        CheckedListBox5.Items.Clear()

        ComboBox8.Items.Clear()
        ReDim g_a_loaded_instr_group_alert(0)

        CheckedListBox6.Items.Clear()
        ReDim g_a_loaded_instr_alert(0)

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

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        TextBox1.Text = db_access_general.GET_INSTITUTION_ID(ComboBox1.SelectedIndex)

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""
        g_selected_soft = -1

        ComboBox3.Items.Clear()
        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()
        ReDim g_a_loaded_instr_group(0)

        ComboBox4.Items.Clear()
        ReDim g_a_loaded_reasons_default(0)

        CheckedListBox1.Items.Clear()
        ReDim g_a_loaded_destinations_default(0)

        CheckedListBox2.Items.Clear()
        ReDim g_a_loaded_profiles_default(0)

        ReDim g_a_loaded_instr(0)
        CheckedListBox3.Items.Clear()

        ReDim g_a_loaded_reasons_alert(0)
        ComboBox7.Items.Clear()

        ReDim g_a_loaded_destinations_alert(0)
        CheckedListBox4.Items.Clear()

        ReDim g_a_loaded_profiles_alert(0)
        CheckedListBox5.Items.Clear()

        ComboBox8.Items.Clear()
        ReDim g_a_loaded_instr_group_alert(0)

        CheckedListBox6.Items.Clear()
        ReDim g_a_loaded_instr_alert(0)

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

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Me.Enabled = False
        Me.Dispose()
        Form1.Show()

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        CheckedListBox1.Items.Clear()
        ReDim g_a_loaded_destinations_default(0)

        CheckedListBox2.Items.Clear()
        ReDim g_a_loaded_profiles_default(0)

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_discharge.GET_DEFAULT_DESTINATIONS(TextBox1.Text, g_selected_soft, g_a_loaded_reasons_default(ComboBox4.SelectedIndex).id_content, ComboBox3.Text, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING DEFAULT DESTINATIONS!!", vbCritical)

        Else

            Dim l_index_destinations_default As Integer = 0

            While dr.Read()
                ReDim Preserve g_a_loaded_destinations_default(l_index_destinations_default)

                g_a_loaded_destinations_default(l_index_destinations_default).id_disch_reas_dest = dr.Item(0)
                g_a_loaded_destinations_default(l_index_destinations_default).id_content = dr.Item(1)
                Try
                    g_a_loaded_destinations_default(l_index_destinations_default).description = dr.Item(2)
                Catch ex As Exception
                    g_a_loaded_destinations_default(l_index_destinations_default).description = ""
                End Try
                g_a_loaded_destinations_default(l_index_destinations_default).id_clinical_service = dr.Item(3)
                    g_a_loaded_destinations_default(l_index_destinations_default).type = dr.Item(4)

                'Verificar se Reason/Destintion está assocaido a um clinical service
                If g_a_loaded_destinations_default(l_index_destinations_default).id_clinical_service = "-1" Then

                    CheckedListBox1.Items.Add(dr.Item(2) & "  -  [" & dr.Item(1) & "]")

                Else

                    CheckedListBox1.Items.Add(dr.Item(2) & "  -  [" & dr.Item(1) & "]   (Clinical Service: " & db_clin_serv.GET_CLIN_SERV_DESC(TextBox1.Text, dr.Item(3)) & ")")

                End If

                l_index_destinations_default = l_index_destinations_default + 1

            End While

        End If

        dr.Dispose()
        dr.Close()


        ReDim g_a_loaded_profiles_default(0)
        Dim l_dimension_profiles As Integer = 0
        Dim dr_profile As OracleDataReader


        If Not db_discharge.GET_DEFAULT_PROFILE_DISCH_REASON(g_selected_soft, g_a_loaded_reasons_default(ComboBox4.SelectedIndex).id_content, dr_profile) Then

            MsgBox("ERROR GETTING DEFAULT DISCHARGE PROFILES!", vbCritical)

        Else

            While dr_profile.Read()

                CheckedListBox2.Items.Add(dr_profile.Item(1) & " - " & dr_profile.Item(2))
                CheckedListBox2.SetItemChecked(l_dimension_profiles, True)

                ReDim Preserve g_a_loaded_profiles_default(l_dimension_profiles)
                g_a_loaded_profiles_default(l_dimension_profiles).ID_PROFILE_DISCH_REASON = dr_profile.Item(0)
                g_a_loaded_profiles_default(l_dimension_profiles).ID_PROFILE_TEMPLATE = dr_profile.Item(1)
                g_a_loaded_profiles_default(l_dimension_profiles).PROFILE_NAME = dr_profile.Item(2)

                l_dimension_profiles = l_dimension_profiles + 1

            End While

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        g_selected_soft = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, TextBox1.Text)

        ComboBox3.Items.Clear()

        ComboBox4.Items.Clear()
        ReDim g_a_loaded_reasons_default(0)

        CheckedListBox1.Items.Clear()
        ReDim g_a_loaded_destinations_default(0)

        CheckedListBox2.Items.Clear()
        ReDim g_a_loaded_profiles_default(0)

        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()
        ReDim g_a_loaded_instr_group(0)

        ReDim g_a_loaded_instr(0)
        CheckedListBox3.Items.Clear()

        ReDim g_a_loaded_destinations_alert(0)
        CheckedListBox4.Items.Clear()

        ReDim g_a_loaded_profiles_alert(0)
        CheckedListBox5.Items.Clear()

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_discharge.GET_DEFAULT_VERSIONS(TextBox1.Text, g_selected_soft, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING DEFAULT DISCHARGE VERSIONS!", vbCritical)

        Else

            While dr.Read()

                ComboBox3.Items.Add(dr.Item(0))

            End While

        End If

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_discharge.GET_DISCH_INSTR_VERSIONS(TextBox1.Text, g_selected_soft, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING DEFAULT DISCHARGE INSTRUCTIONS VERSIONS!", vbCritical)

        Else

            While dr.Read()

                ComboBox5.Items.Add(dr.Item(0))

            End While

        End If

        RESET_REASONS_ALERT()

        RESET_GROUP_INSTR_ALERT()

        CheckedListBox6.Items.Clear()
        ReDim g_a_loaded_instr_alert(0)

        dr.Dispose()
        dr.Close()

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        ComboBox4.Items.Clear()
        ReDim g_a_loaded_reasons_default(0)

        CheckedListBox1.Items.Clear()
        ReDim g_a_loaded_destinations_default(0)

        CheckedListBox2.Items.Clear()
        ReDim g_a_loaded_profiles_default(0)

        If ComboBox3.Text <> "" Then

            Dim dr_new As OracleDataReader

            If Not db_discharge.GET_DEFAULT_REASONS(TextBox1.Text, g_selected_soft, ComboBox3.Text, dr_new) Then

                MsgBox("ERROR GETING DEFAULT DISCHARGE REASONS.", vbCritical)

            Else

                Dim l_index_reason_default As Integer = 0
                ReDim g_a_loaded_reasons_default(0)

                While dr_new.Read()

                    ReDim Preserve g_a_loaded_reasons_default(l_index_reason_default)
                    g_a_loaded_reasons_default(l_index_reason_default).id_content = dr_new.Item(0)
                    g_a_loaded_reasons_default(l_index_reason_default).desccription = dr_new.Item(1)
                    l_index_reason_default = l_index_reason_default + 1

                    ComboBox4.Items.Add(dr_new.Item(1) & "  -  [" & dr_new.Item(0) & "]")

                End While

            End If

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Cursor = Cursors.WaitCursor

        'Lista de Reasons
        If ComboBox4.SelectedIndex > -1 Then

            'Lista de Destinations
            If CheckedListBox1.CheckedIndices.Count() > 0 Then

                'Lista de Perfis
                If CheckedListBox2.CheckedItems.Count() > 0 Then

                    '1 - Verificar se existe Reason no ALERT (e respetiva tradução), caso não exista, inserir.
                    If Not db_discharge.CHECK_REASON(g_a_loaded_reasons_default(ComboBox4.SelectedIndex).id_content) Then

                        If Not db_discharge.SET_REASON(TextBox1.Text, g_a_loaded_reasons_default(ComboBox4.SelectedIndex).id_content) Then

                            MsgBox("ERROR INSERTING DISCHARGE REASON!", vbCritical)

                        End If

                    ElseIf Not db_discharge.CHECK_REASON_translation(TextBox1.TEXT, g_a_loaded_reasons_default(ComboBox4.SelectedIndex).id_content) Then

                        If Not db_discharge.SET_REASON_TRANSLATION(TextBox1.Text, g_a_loaded_reasons_default(ComboBox4.SelectedIndex).id_content) Then

                            MsgBox("ERROR INSERTING DISCHARGE REASON TRANSLATION!", vbCritical)

                        End If

                    End If

                    '2 - Verificar se existe Destination no ALERT (e respetiva tradução), caso não exista, inserir.

                    Dim l_a_checked_destinations(0) As DISCHARGE_API.DEFAULT_DISCAHRGE
                    Dim l_dimension_check_dest As Integer = 0

                    For Each indexChecked In CheckedListBox1.CheckedIndices

                        ReDim Preserve l_a_checked_destinations(l_dimension_check_dest)

                        l_a_checked_destinations(l_dimension_check_dest).id_disch_reas_dest = g_a_loaded_destinations_default(indexChecked).id_disch_reas_dest
                        l_a_checked_destinations(l_dimension_check_dest).id_content = g_a_loaded_destinations_default(indexChecked).id_content
                        l_a_checked_destinations(l_dimension_check_dest).description = g_a_loaded_destinations_default(indexChecked).description
                        l_a_checked_destinations(l_dimension_check_dest).id_clinical_service = g_a_loaded_destinations_default(indexChecked).id_clinical_service
                        l_a_checked_destinations(l_dimension_check_dest).type = g_a_loaded_destinations_default(indexChecked).type

                        l_dimension_check_dest = l_dimension_check_dest + 1

                    Next

                    For i As Integer = 0 To l_a_checked_destinations.Count() - 1

                        If (l_a_checked_destinations(i).type = "D") Then

                            If Not db_discharge.CHECK_DESTINATION(l_a_checked_destinations(i).id_content) Then

                                If Not db_discharge.SET_DESTINATION(TextBox1.Text, l_a_checked_destinations(i)) Then

                                    MsgBox("ERROR INSERTING DISCHARGE DESTINATION!", vbCritical)

                                End If

                            ElseIf Not db_discharge.CHECK_DESTINATION_TRANSLATION(TextBox1.Text, L_a_checked_destinations(i).id_content) Then

                                If Not db_discharge.SET_DESTINATION_TRANSLATION(TextBox1.Text, l_a_checked_destinations(i)) Then

                                    MsgBox("ERROR INSERTING DISCHARGE DESTINATION TRANSLATION!", vbCritical)

                                End If

                            End If

                        End If

                    Next

                    '3 - Gravar os profiles_discharge selecionados
                    'ARRAY QUE VAI GUARDAR OS PROFILE TEMPLATES SELECTIONADOS PELO UTILIZADOR
                    Dim l_a_selected_profiles_default() As DISCHARGE_API.DEFAULT_DISCH_PROFILE

                    ReDim l_a_selected_profiles_default(0)
                    Dim l_dim_selected_profiles = 0

                    For Each indexChecked In CheckedListBox2.CheckedIndices

                        ReDim Preserve l_a_selected_profiles_default(l_dim_selected_profiles)
                        l_a_selected_profiles_default(l_dim_selected_profiles).ID_PROFILE_DISCH_REASON = g_a_loaded_profiles_default(indexChecked).ID_PROFILE_DISCH_REASON
                        l_a_selected_profiles_default(l_dim_selected_profiles).ID_PROFILE_TEMPLATE = g_a_loaded_profiles_default(indexChecked).ID_PROFILE_TEMPLATE
                        l_a_selected_profiles_default(l_dim_selected_profiles).PROFILE_NAME = g_a_loaded_profiles_default(indexChecked).PROFILE_NAME
                        l_dim_selected_profiles = l_dim_selected_profiles + 1

                    Next

                    If Not db_discharge.SET_PROFILE_DISCH_REASON(TextBox1.Text, l_a_selected_profiles_default) Then

                        MsgBox("ERROR INSERTING PROFILE_DISCHARGE_REASON!", vbCritical)

                    End If

                    '4 - Gravar os DISCH_REAS_DEST
                    Dim l_a_selected_reas_dest() As DISCHARGE_API.DEFAULT_DISCAHRGE

                    ReDim l_a_selected_reas_dest(0)
                    Dim l_dim_selected_reas_dest = 0

                    For Each indexChecked In CheckedListBox1.CheckedIndices

                        ReDim Preserve l_a_selected_reas_dest(l_dim_selected_reas_dest)

                        l_a_selected_reas_dest(l_dim_selected_reas_dest).id_disch_reas_dest = g_a_loaded_destinations_default(indexChecked).id_disch_reas_dest
                        l_a_selected_reas_dest(l_dim_selected_reas_dest).id_content = g_a_loaded_destinations_default(indexChecked).id_content
                        l_a_selected_reas_dest(l_dim_selected_reas_dest).description = g_a_loaded_destinations_default(indexChecked).description
                        l_a_selected_reas_dest(l_dim_selected_reas_dest).id_clinical_service = g_a_loaded_destinations_default(indexChecked).id_clinical_service
                        l_a_selected_reas_dest(l_dim_selected_reas_dest).type = g_a_loaded_destinations_default(indexChecked).type

                        l_dim_selected_reas_dest = l_dim_selected_reas_dest + 1

                    Next

                    If Not db_discharge.SET_DISCH_REAS_DEST(TextBox1.Text, g_selected_soft, g_a_loaded_reasons_default(ComboBox4.SelectedIndex).id_content, l_a_selected_reas_dest) Then

                        MsgBox("ERROR INSERTING DISCH_REAS_DEST!", vbCritical)

                    Else

                        MsgBox("Records correctly inserted.", vbInformation)

                    End If

                    RESET_REASONS_ALERT()

                    ReDim g_a_loaded_destinations_alert(0)
                    CheckedListBox4.Items.Clear()

                    ReDim g_a_loaded_profiles_alert(0)
                    CheckedListBox5.Items.Clear()


                Else

                    MsgBox("Please select, at least, one Profile.")

                End If

            Else

                MsgBox("Please select, at least, one Discharge Destination.")

            End If

        Else

            MsgBox("Please select a Discharge Reason.")

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

        If CheckedListBox1.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox1.Items.Count - 1

                CheckedListBox1.SetItemChecked(i, False)

            Next

        End If

    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click

        If CheckedListBox1.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox1.Items.Count - 1

                CheckedListBox1.SetItemChecked(i, True)

            Next

        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        If CheckedListBox2.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox2.Items.Count - 1

                CheckedListBox2.SetItemChecked(i, True)

            Next

        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If CheckedListBox2.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox2.Items.Count - 1

                CheckedListBox2.SetItemChecked(i, False)

            Next

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Cursor = Cursors.WaitCursor

        If CheckedListBox3.CheckedItems.Count() > 0 Then

            Dim l_a_selected_instr() As DISCHARGE_API.DEFAULT_INSTR

            ReDim l_a_selected_instr(0)
            Dim l_dim_selected_instr = 0

            For Each indexChecked In CheckedListBox3.CheckedIndices

                ReDim Preserve l_a_selected_instr(l_dim_selected_instr)
                l_a_selected_instr(l_dim_selected_instr).ID_CONTENT = g_a_loaded_instr(indexChecked).ID_CONTENT
                l_a_selected_instr(l_dim_selected_instr).DESCRIPTION = g_a_loaded_instr(indexChecked).DESCRIPTION

                l_dim_selected_instr = l_dim_selected_instr + 1

            Next

            If Not db_discharge.SET_DISCH_INSTRUCTION(TextBox1.Text, g_selected_soft, g_a_loaded_instr_group(ComboBox6.SelectedIndex).ID_CONTENT, l_a_selected_instr) Then

                MsgBox("ERROR SETTING DISCHARGE INSTRUCTION!", vbCritical)

            Else

                RESET_GROUP_INSTR_ALERT()

                ReDim g_a_loaded_instr_alert(0)
                CheckedListBox6.Items.Clear()

                MsgBox("Discharge instructions correctly inserted.", vbInformation)

            End If

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If CheckedListBox3.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox3.Items.Count - 1

                CheckedListBox3.SetItemChecked(i, False)

            Next

        End If

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        MsgBox("Attention: You are about to perform changes that may not have been conceptualized to work with ALERT® systems. These changes may cause the malfunction of the system. Please proceed carefully. ", vbExclamation)

        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Dim FORM_ADV_DISCH As New DISCHARGE_ADVANCED
        FORM_ADV_DISCH.ShowDialog()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        If CheckedListBox3.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox3.Items.Count - 1

                CheckedListBox3.SetItemChecked(i, True)

            Next

        End If

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        ComboBox6.Items.Clear()
        ReDim g_a_loaded_instr_group(0)
        ReDim g_a_loaded_instr(0)
        CheckedListBox3.Items.Clear()

        CheckedListBox3.Items.Clear()

        If ComboBox5.Text <> "" Then

            Dim dr_new As OracleDataReader

            If Not db_discharge.GET_DEFAULT_INSTR_GROUP(TextBox1.Text, g_selected_soft, ComboBox5.Text, dr_new) Then

                MsgBox("ERROR GETING DEFAULT DISCHARGE INSTRUCTIONS GROUPS!", vbCritical)

            Else

                Dim l_index_groups As Integer = 0
                ReDim g_a_loaded_instr_group(0)

                While dr_new.Read()

                    ReDim Preserve g_a_loaded_instr_group(l_index_groups)
                    g_a_loaded_instr_group(l_index_groups).ID_CONTENT = dr_new.Item(0)
                    g_a_loaded_instr_group(l_index_groups).DESCRIPTION = dr_new.Item(1)
                    l_index_groups = l_index_groups + 1

                    ComboBox6.Items.Add(dr_new.Item(1) & "  -  [" & dr_new.Item(0) & "]")

                End While

            End If

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        ReDim g_a_loaded_instr(0)
        CheckedListBox3.Items.Clear()
        Dim l_dimension_instr As Integer = 0
        Dim dr_instr As OracleDataReader

        If Not db_discharge.GET_DEFAULT_INSTR_TITLES(TextBox1.Text, g_selected_soft, ComboBox5.Text, g_a_loaded_instr_group(ComboBox6.SelectedIndex).ID_CONTENT, dr_instr) Then

            MsgBox("ERROR GETTING DISCHARGE INSTRUCTIONS!", vbCritical)

        Else

            While dr_instr.Read()

                CheckedListBox3.Items.Add(dr_instr.Item(1) & " - " & dr_instr.Item(0))
                CheckedListBox3.SetItemChecked(l_dimension_instr, True)

                ReDim Preserve g_a_loaded_instr(l_dimension_instr)
                g_a_loaded_instr(l_dimension_instr).ID_CONTENT = dr_instr.Item(0)
                g_a_loaded_instr(l_dimension_instr).DESCRIPTION = dr_instr.Item(1)

                l_dimension_instr = l_dimension_instr + 1

            End While

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub CheckedListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox3.DoubleClick

        If CheckedListBox3.Items.Count() > 0 And CheckedListBox3.SelectedIndex > -1 Then

            Form_location.x_position = Me.Location.X
            Form_location.y_position = Me.Location.Y

            Dim l_instr_description As String

            If Not db_discharge.GET_DEFAULT_INSTR(TextBox1.Text, g_a_loaded_instr(CheckedListBox3.SelectedIndex).ID_CONTENT, l_instr_description) Then

                MsgBox("ERROR GETTING DISCHARGE INSTRUCTIONS!", vbCritical)

            End If

            Dim show_instr As New CONTENT_DISPLAY(g_a_loaded_instr(CheckedListBox3.SelectedIndex).DESCRIPTION, l_instr_description)

            show_instr.ShowDialog()

        End If

    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged

        ReDim g_a_loaded_destinations_alert(0)
        CheckedListBox4.Items.Clear()

        ReDim g_a_loaded_profiles_alert(0)
        CheckedListBox5.Items.Clear()

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_discharge.GET_DESTINATIONS(TextBox1.Text, g_selected_soft, g_a_loaded_reasons_alert(ComboBox7.SelectedIndex).id_content, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING DISCHARGE DESTINATIONS FROM ALERT!", vbCritical)

        Else

            Dim l_dim_dest_alert As Integer = 0
            While dr.Read()

                'Só populo os 3 primeiros campos
                CheckedListBox4.Items.Add(dr.Item(2) & "  -  [" & dr.Item(1) & "]")

                ReDim Preserve g_a_loaded_destinations_alert(l_dim_dest_alert)

                g_a_loaded_destinations_alert(l_dim_dest_alert).id_disch_reas_dest = dr.Item(0)

                Try
                    g_a_loaded_destinations_alert(l_dim_dest_alert).id_content = dr.Item(1)

                Catch ex As Exception

                    g_a_loaded_destinations_alert(l_dim_dest_alert).id_content = ""

                End Try

                Try
                    g_a_loaded_destinations_alert(l_dim_dest_alert).description = dr.Item(2)
                Catch ex As Exception
                    g_a_loaded_destinations_alert(l_dim_dest_alert).description = ""
                End Try


                l_dim_dest_alert = l_dim_dest_alert + 1

            End While

        End If

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_discharge.GET_PROFILE_DISCH_REASON(TextBox1.Text, g_selected_soft, g_a_loaded_reasons_alert(ComboBox7.SelectedIndex).id_content, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING DISCHARGE PROFILES FROM ALERT!", vbCritical)

        Else

            Dim l_dim_prof_alert As Integer = 0
            While dr.Read()

                'Só populo os 3 primeiros campos
                CheckedListBox5.Items.Add(dr.Item(1) & " - " & dr.Item(2))

                ReDim Preserve g_a_loaded_profiles_alert(l_dim_prof_alert)

                g_a_loaded_profiles_alert(l_dim_prof_alert).ID_PROFILE_DISCH_REASON = dr.Item(0)
                g_a_loaded_profiles_alert(l_dim_prof_alert).ID_PROFILE_TEMPLATE = dr.Item(1)
                g_a_loaded_profiles_alert(l_dim_prof_alert).PROFILE_NAME = dr.Item(1)

                l_dim_prof_alert = l_dim_prof_alert + 1

            End While

        End If

        dr.Dispose()
        dr.Close()


    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        Cursor = Cursors.WaitCursor

        'Se a Lista de Destinations estiver preenchida
        If CheckedListBox4.Items.Count() > 0 Then

            'Se for selecionado pelo menos uma destination
            If CheckedListBox4.CheckedItems.Count() > 0 Then

                Dim l_a_check_disch(0) As DISCHARGE_API.DEFAULT_DISCAHRGE

                Dim l_dim_check_disch As Integer = 0

                For Each indexChecked In CheckedListBox4.CheckedIndices

                    ReDim Preserve l_a_check_disch(l_dim_check_disch)
                    l_a_check_disch(l_dim_check_disch).id_disch_reas_dest = g_a_loaded_destinations_alert(indexChecked).id_disch_reas_dest

                    l_dim_check_disch = l_dim_check_disch + 1

                Next

                For i As Integer = 0 To l_a_check_disch.Count() - 1

                    If Not db_discharge.DELETE_DISCH_REAS_DEST(l_a_check_disch(i).id_disch_reas_dest) Then

                        MsgBox("ERROR DELETING FROM DISCH_REAS_DEST!", vbCritical)

                    End If

                Next

                '  Dim dr As OracleDataReader

                'A - Limpar Arrays e boxes
                'A.1 Se a lista total de Destination foi apagada => Limpar as duas boxes. Atualizar a box de Reasons
                If CheckedListBox4.CheckedItems.Count() = g_a_loaded_destinations_alert.Count() Then

                    RESET_REASONS_ALERT()

                    ReDim g_a_loaded_destinations_alert(0)
                    CheckedListBox4.Items.Clear()

                    ReDim g_a_loaded_profiles_alert(0)
                    CheckedListBox5.Items.Clear()

                    'A2 - Atualizar apenas a box de destinations
                Else

                    RESET_DESTINATIONS_ALERT()

                End If

                MsgBox("Record(s) deleted.", vbInformation)

            Else

                MsgBox("No records selected to be deleted.")

            End If

        Else

            MsgBox("No records selected to be deleted.")

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click

        Cursor = Cursors.WaitCursor

        'Lista de Profiles preenchida
        If CheckedListBox5.Items.Count() > 0 Then

            If CheckedListBox5.CheckedItems.Count() > 0 Then

                Dim l_a_check_prof_disch_reas(0) As DISCHARGE_API.DEFAULT_DISCH_PROFILE

                Dim l_dim_check_prof As Integer = 0

                For Each indexChecked In CheckedListBox5.CheckedIndices

                    ReDim Preserve l_a_check_prof_disch_reas(l_dim_check_prof)
                    l_a_check_prof_disch_reas(l_dim_check_prof).ID_PROFILE_DISCH_REASON = g_a_loaded_profiles_alert(indexChecked).ID_PROFILE_DISCH_REASON

                    l_dim_check_prof = l_dim_check_prof + 1
                Next


                For i As Integer = 0 To l_a_check_prof_disch_reas.Count() - 1

                    If Not db_discharge.DELETE_PROF_DISCH_REAS(l_a_check_prof_disch_reas(i).ID_PROFILE_DISCH_REASON) Then

                        MsgBox("ERROR DELETING FROM PROFILE_DISCH_REASON!", vbCritical)

                    End If

                Next

                'A - Limpar Arrays e boxes
                'A.1 Se a lista total de Perfis foi apagada => Limpar as duas boxes. Atualizar a box de Reasons
                If CheckedListBox5.CheckedItems.Count() = g_a_loaded_profiles_alert.Count() Then

                    RESET_REASONS_ALERT()

                    ReDim g_a_loaded_destinations_alert(0)
                    CheckedListBox4.Items.Clear()

                    ReDim g_a_loaded_profiles_alert(0)
                    CheckedListBox5.Items.Clear()

                    'A2 - Atualizar apenas a box de Perfis
                Else

                    RESET_PROFILES_ALERT()

                End If

                MsgBox("Record(s) deleted.", vbInformation)

            Else

                MsgBox("No records selected to be deleted.")

            End If

        Else

            MsgBox("No records selected to be deleted.")

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        If CheckedListBox4.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox4.Items.Count - 1

                CheckedListBox4.SetItemChecked(i, True)

            Next

        End If

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        If CheckedListBox4.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox4.Items.Count - 1

                CheckedListBox4.SetItemChecked(i, False)

            Next

        End If

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        If CheckedListBox5.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox5.Items.Count - 1

                CheckedListBox5.SetItemChecked(i, False)

            Next

        End If

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

        If CheckedListBox5.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox5.Items.Count - 1

                CheckedListBox5.SetItemChecked(i, True)

            Next

        End If

    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged

        Cursor = Cursors.WaitCursor


        RESET_INSTR_ALERT()

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click

        Cursor = Cursors.WaitCursor

        'Lista de Instructions preenchida
        If CheckedListBox6.Items.Count() > 0 Then

            If CheckedListBox6.CheckedItems.Count() > 0 Then

                Dim l_a_check_instr(0) As DISCHARGE_API.DEFAULT_INSTR

                Dim l_dim_check_instr As Integer = 0

                For Each indexChecked In CheckedListBox6.CheckedIndices

                    ReDim Preserve l_a_check_instr(l_dim_check_instr)
                    l_a_check_instr(l_dim_check_instr).ID_CONTENT = g_a_loaded_instr_alert(indexChecked).ID_CONTENT

                    l_dim_check_instr = l_dim_check_instr + 1
                Next


                For i As Integer = 0 To l_a_check_instr.Count() - 1

                    If Not db_discharge.DELETE_DISCH_INSTR_REL(TextBox1.Text, g_selected_soft, g_a_loaded_instr_group_alert(ComboBox8.SelectedIndex).ID_CONTENT, l_a_check_instr(i).ID_CONTENT) Then

                        MsgBox("ERROR DELETING FROM DISCHARGE_INSTR_RELATION!", vbCritical)

                    End If

                Next

                'A - Limpar Arrays e boxes
                'A.1 Se a lista total de Perfis foi apagada => Limpar OS GRUPOS E AS INSTRUÇÕES
                If CheckedListBox6.CheckedItems.Count() = g_a_loaded_instr_alert.Count() Then

                    RESET_GROUP_INSTR_ALERT()

                    ReDim g_a_loaded_instr_alert(0)
                    CheckedListBox6.Items.Clear()

                    'A2 - Atualizar apenas a box de Instruções
                Else

                    RESET_INSTR_ALERT()

                End If

                MsgBox("Record(s) deleted.", vbInformation)

            Else

                MsgBox("No records selected to be deleted.")

            End If

        Else

            MsgBox("No records selected to be deleted.")

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click

        If CheckedListBox6.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox6.Items.Count - 1

                CheckedListBox6.SetItemChecked(i, True)

            Next

        End If

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        If CheckedListBox6.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox6.Items.Count - 1

                CheckedListBox6.SetItemChecked(i, False)

            Next

        End If

    End Sub

    Private Sub CheckedListBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox6.DoubleClick

        If CheckedListBox6.Items.Count() > 0 And CheckedListBox6.SelectedIndex > -1 Then

            Form_location.x_position = Me.Location.X
            Form_location.y_position = Me.Location.Y

            Dim l_instr_description As String

            'FAZER FUNCÇÂO PARA OBTER TEXTO DO ALERT
            If Not db_discharge.GET_ALERT_INSTR(TextBox1.Text, g_a_loaded_instr_alert(CheckedListBox6.SelectedIndex).ID_CONTENT, l_instr_description) Then

                MsgBox("ERROR GETTING DISCHARGE INSTRUCTIONS!", vbCritical)

            End If

            Dim show_instr As New CONTENT_DISPLAY(g_a_loaded_instr_alert(CheckedListBox6.SelectedIndex).DESCRIPTION, l_instr_description)

            show_instr.ShowDialog()

        End If

    End Sub

End Class
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

    'Array que vai guardar as DESTINATIONS caregadas do default
    Dim g_a_loaded_destinations_default() As DISCHARGE_API.DEFAULT_DISCAHRGE

    'ARRAY QUE VAI GUARDAR OS PROFILE TEMPLATES DA REASON SELECIONADA
    Dim g_a_loaded_profiles_default() As DISCHARGE_API.DEFAULT_DISCH_PROFILE

    Private Sub DISCHARGE_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "DISCHARGE  ::::  Connected to " & Connection.db

        Me.BackColor = Color.FromArgb(215, 215, 180)

        Me.Location = New Point(Form_location.x_position, Form_location.y_position)

        CheckedListBox2.BackColor = Color.FromArgb(195, 195, 165)
        CheckedListBox1.BackColor = Color.FromArgb(195, 195, 165)
        CheckedListBox3.BackColor = Color.FromArgb(195, 195, 165)

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

        ComboBox4.Items.Clear()
        ReDim g_a_loaded_reasons_default(0)

        CheckedListBox1.Items.Clear()
        ReDim g_a_loaded_destinations_default(0)

        CheckedListBox2.Items.Clear()
        ReDim g_a_loaded_profiles_default(0)

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

        ComboBox4.Items.Clear()
        ReDim g_a_loaded_reasons_default(0)

        CheckedListBox1.Items.Clear()
        ReDim g_a_loaded_destinations_default(0)

        CheckedListBox2.Items.Clear()
        ReDim g_a_loaded_profiles_default(0)

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

        Dim form1 As New Form1

        form1.Show()

        Me.Close()

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
                g_a_loaded_destinations_default(l_index_destinations_default).desccription = dr.Item(2)
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


        If Not db_discharge.GET_DEFAULT_PROFILE_DISCH_REASON(g_a_loaded_reasons_default(ComboBox4.SelectedIndex).id_content, dr_profile) Then

            MsgBox("ERROR GETTING DISCHARGE PROFILES!", vbCritical)

        Else

            While dr_profile.Read()

                CheckedListBox2.Items.Add(dr_profile.Item(0) & " - " & dr_profile.Item(1))
                CheckedListBox2.SetItemChecked(l_dimension_profiles, True)

                ReDim Preserve g_a_loaded_profiles_default(l_dimension_profiles)
                g_a_loaded_profiles_default(l_dimension_profiles).ID_PROFILE_TEMPLATE = dr_profile.Item(0)
                g_a_loaded_profiles_default(l_dimension_profiles).PROFILE_NAME = dr_profile.Item(1)

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

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_discharge.GET_DEFAULT_VERSIONS(TextBox1.Text, g_selected_soft, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING SOFTWARES!", vbCritical)

        Else

            While dr.Read()

                ComboBox3.Items.Add(dr.Item(0))

            End While

        End If

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

        Dim l_has_parent As Boolean = False
        Dim l_id_parent As Int64 = -1

        If Not db_clin_serv.CHECK_CLIN_SERV(ComboBox1.Text) Then

            ''Verificação Parents
            If db_clin_serv.CHECK_HAS_PARENT(470, ComboBox1.Text) Then

                MsgBox("HAS PARENTS")

                l_has_parent = True
                Dim dr As OracleDataReader

                If Not db_clin_serv.GET_PARENT(470, ComboBox1.Text, dr) Then

                    MsgBox("ERROR 1")

                Else

                    While dr.Read()

                        If Not db_clin_serv.CHECK_CLIN_SERV(dr.Item(0)) Then

                            If Not db_clin_serv.SET_CLIN_SERV(dr.Item(0)) Then
                                MsgBox("ERROR 2")
                            Else
                                If Not db_clin_serv.SET_CLIN_SERV_TRANSLATION(470, dr.Item(0)) Then
                                    MsgBox("ERROR 3")
                                End If
                            End If

                        End If

                        If Not db_clin_serv.GET_ID_ALERT(dr.Item(0), l_id_parent) Then

                            MsgBox("ERROR 3.1")

                        Else

                            MsgBox(l_id_parent)

                        End If

                    End While

                End If

            End If

            'Inserção Clinical Services
            If Not db_clin_serv.SET_CLIN_SERV(ComboBox1.Text) Then

                MsgBox("ERROR 4")

            Else

                If Not db_clin_serv.SET_CLIN_SERV_TRANSLATION(470, ComboBox1.Text) Then

                    MsgBox("ERROR 5")

                End If

            End If

            If l_has_parent = True Then

                If Not db_clin_serv.SET_PARENT(ComboBox1.Text, l_id_parent) Then

                    MsgBox("ERROR 5.1")

                End If

            End If

        ElseIf Not db_clin_serv.CHECK_CLIN_SERV_TRANSLATION(470, ComboBox1.Text) Then

            If Not db_clin_serv.SET_CLIN_SERV_TRANSLATION(470, ComboBox1.Text) Then

                MsgBox("ERROR 6")

            End If

        End If

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

End Class
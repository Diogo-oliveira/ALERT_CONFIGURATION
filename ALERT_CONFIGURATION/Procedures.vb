﻿Imports Oracle.DataAccess.Client
Public Class Procedures
    Dim db_access_general As New General
    Dim oradb As String = "Data Source=QC4V265;User Id=alert_config;Password=qcteam"
    Dim conn As New OracleConnection(oradb)
    Dim db_intervention As New INTERVENTIONS_API

    Dim g_selected_soft As Int16 = -1
    ''Array que vai guardar os dep_clin_serv da instituição
    Dim g_a_dep_clin_serv_inst() As Int64
    Dim g_id_dep_clin_serv As Int64 = 0 'Variavel que vai guardar o id do dep_clin_serv_selecionado

    'Array que vai guardar as categorias disponíveis no ALERT
    Dim g_a_interv_cats_alert() As String

    Dim g_a_loaded_categories_default() As String ' Array que vai guardar os id_contents das categorias carregadas do default
    Dim g_selected_category As String = ""

    Dim g_a_loaded_interventions_default() As INTERVENTIONS_API.interventions_default 'Array que vai guardar os id_contents das análises carregadas do default
    Dim g_a_selected_default_interventions() As INTERVENTIONS_API.interventions_default

    Dim g_index_selected_intervention_from_default As Integer = 0 ''Variavel utilizada no botão de adicionar à box da direita (CHECKBOX 1)

    'Array que vai guardar as análises carregadas do ALERT
    Dim g_a_intervs_alert() As INTERVENTIONS_API.interventions_default
    Dim g_dimension_intervs_alert As Int64 = 0


    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Dim form1 As New Form1

        Me.Enabled = False

        Me.Dispose()

        form1.Show()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Cursor = Cursors.WaitCursor

        'Limpar arrays
        ' g_selected_room = -1
        'ReDim g_a_loaded_rooms(0)
        g_selected_soft = -1
        ReDim g_a_dep_clin_serv_inst(0)
        g_id_dep_clin_serv = 0
        ReDim g_a_loaded_categories_default(0)
        g_selected_category = -1
        ReDim g_a_loaded_interventions_default(0)
        ReDim g_a_selected_default_interventions(0)
        g_index_selected_intervention_from_default = 0
        'ReDim g_a_lab_cats_alert(0)
        ReDim g_a_interv_cats_alert(0)
        g_dimension_intervs_alert = 0
        'ReDim g_a_labs_for_clinical_service(0)
        'g_dimension_labs_cs = 0
        'ReDim g_a_selected_labs_delete_cs(0)

        'Limpar a seleção de quarto
        'g_selected_room = -1

        If TextBox1.Text <> "" Then

            ComboBox1.Text = db_access_general.GET_INSTITUTION(TextBox1.Text, conn)

            ComboBox2.Items.Clear()
            ComboBox2.Text = ""

            Dim dr As OracleDataReader

            'ReDim g_a_loaded_rooms(0)
            'Dim i_index_room As Int32 = 0


#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not db_access_general.GET_SOFT_INST(TextBox1.Text, conn, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                MsgBox("ERROR GETTING SOFTWARES!", vbCritical)

            Else

                While dr.Read()

                    ComboBox2.Items.Add(dr.Item(1))

                End While

                ComboBox3.Text = ""
                ComboBox3.Items.Clear()

                ComboBox4.Text = ""
                ComboBox4.Items.Clear()

                CheckedListBox2.Items.Clear()

                CheckedListBox1.Items.Clear()

                ComboBox5.Text = ""
                ComboBox5.Items.Clear()
                CheckedListBox3.Items.Clear()

                ComboBox6.Text = ""
                ComboBox6.Items.Clear()
                CheckedListBox4.Items.Clear()

                g_selected_category = ""

            End If

            dr.Dispose()
            dr.Close()

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Procedures_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            'Estabelecer ligação à BD
            conn.Open()

        Catch ex As Exception

            MsgBox("ERROR CONNECTING TO DATA BASE!", vbCritical)

        End Try

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_ALL_INSTITUTIONS(conn, dr) Then
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
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Cursor = Cursors.WaitCursor

        'Limpar arrays
        'g_selected_room = -1
        'ReDim g_a_loaded_rooms(0)
        g_selected_soft = -1
        ReDim g_a_dep_clin_serv_inst(0)
        g_id_dep_clin_serv = 0
        ReDim g_a_loaded_categories_default(0)
        g_selected_category = -1
        ReDim g_a_loaded_interventions_default(0)
        ReDim g_a_selected_default_interventions(0)
        g_index_selected_intervention_from_default = 0
        ReDim g_a_interv_cats_alert(0)
        ReDim g_a_interv_cats_alert(0)
        g_dimension_intervs_alert = 0
        'ReDim g_a_labs_for_clinical_service(0)
        'g_dimension_labs_cs = 0
        'ReDim g_a_selected_labs_delete_cs(0)

        TextBox1.Text = db_access_general.GET_INSTITUTION_ID(ComboBox1.SelectedIndex, conn)

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""

        ' g_selected_room = -1

        '  ReDim g_a_loaded_rooms(0)
        'Dim i_index_room As Int32 = 0

        Dim dr As OracleDataReader

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""

        If Not db_access_general.GET_SOFT_INST(TextBox1.Text, conn, dr) Then

            MsgBox("ERROR GETTING SOFTWARES!", vbCritical)

        Else

            While dr.Read()

                ComboBox2.Items.Add(dr.Item(1))

            End While

        End If

        dr.Dispose()
        dr.Close()

        ComboBox3.Text = ""
        ComboBox3.Items.Clear()

        ComboBox4.Text = ""
        ComboBox4.Items.Clear()

        CheckedListBox2.Items.Clear()

        CheckedListBox1.Items.Clear()

        ComboBox5.Text = ""
        ComboBox5.Items.Clear()
        CheckedListBox3.Items.Clear()

        ComboBox6.Text = ""
        ComboBox6.Items.Clear()
        CheckedListBox4.Items.Clear()

        Cursor = Cursors.Arrow
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        'Limpar arrays
        g_selected_soft = -1
        ReDim g_a_dep_clin_serv_inst(0)
        g_id_dep_clin_serv = 0
        ReDim g_a_loaded_categories_default(0)
        g_selected_category = -1
        ReDim g_a_loaded_interventions_default(0)
        ReDim g_a_selected_default_interventions(0)
        g_index_selected_intervention_from_default = 0
        'ReDim g_a_lab_cats_alert(0)
        ReDim g_a_interv_cats_alert(0)
        g_dimension_intervs_alert = 0
        'ReDim g_a_labs_for_clinical_service(0)
        'g_dimension_labs_cs = 0
        'ReDim g_a_selected_labs_delete_cs(0)

        CheckedListBox1.Items.Clear()
        CheckedListBox2.Items.Clear()
        CheckedListBox3.Items.Clear()
        CheckedListBox4.Items.Clear()

        ComboBox3.Items.Clear()
        ComboBox3.Text = ""
        ComboBox4.Items.Clear()
        ComboBox4.Text = ""
        ComboBox5.Items.Clear()
        ComboBox5.Text = ""
        ComboBox6.Items.Clear()
        ComboBox6.Text = ""

        g_selected_soft = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, TextBox1.Text, conn)

        '1 - Fill Version combobox

        Dim dr_def_versions As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_intervention.GET_DEFAULT_VERSIONS(TextBox1.Text, g_selected_soft, conn, dr_def_versions) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR LOADING DEFAULT VERSIONS -  ComboBox2_SelectedIndexChanged", MsgBoxStyle.Critical)

        Else

            While dr_def_versions.Read()

                ComboBox3.Items.Add(dr_def_versions.Item(0))

            End While

        End If

        dr_def_versions.Dispose()
        dr_def_versions.Close()

        ''''''''''''''''''''''
        'Box de categorias na instituição/software

        Dim dr_exam_cat As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_intervention.GET_INTERV_CATS_INST_SOFT(TextBox1.Text, g_selected_soft, conn, dr_exam_cat) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR LOADING LAB CATEGORIES FROM INSTITUTION!", vbCritical)

        Else

            ComboBox5.Items.Add("ALL")

            ReDim g_a_interv_cats_alert(0)
            g_a_interv_cats_alert(0) = 0

            Dim l_index As Int16 = 1

            While dr_exam_cat.Read()

                ComboBox5.Items.Add(dr_exam_cat.Item(1))
                ReDim Preserve g_a_interv_cats_alert(l_index)
                g_a_interv_cats_alert(l_index) = dr_exam_cat.Item(0)
                l_index = l_index + 1

            End While

        End If

        dr_exam_cat.Dispose()
        dr_exam_cat.Close()

        'Preencher os Clinical Services

        Dim dr_clin_serv As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_CLIN_SERV(TextBox1.Text, g_selected_soft, conn, dr_clin_serv) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING CLINICAL SERVICES!")

        Else

            Dim i As Integer = 0

            Dim l_index_dep_clin_serv As Integer = 0
            ReDim g_a_dep_clin_serv_inst(l_index_dep_clin_serv)

            While dr_clin_serv.Read()

                ComboBox6.Items.Add(dr_clin_serv.Item(0))

                ReDim Preserve g_a_dep_clin_serv_inst(l_index_dep_clin_serv)
                g_a_dep_clin_serv_inst(l_index_dep_clin_serv) = dr_clin_serv.Item(1)
                l_index_dep_clin_serv = l_index_dep_clin_serv + 1

            End While

        End If

        dr_clin_serv.Dispose()
        dr_clin_serv.Close()

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        'Limpar arrays
        ReDim g_a_loaded_categories_default(0)
        g_selected_category = -1
        ReDim g_a_loaded_interventions_default(0)
        ReDim g_a_selected_default_interventions(0)
        g_index_selected_intervention_from_default = 0

        ComboBox4.Items.Clear()
        ComboBox4.Text = ""

        CheckedListBox2.Items.Clear()

        'Determinar as categorias disponíveis para a versão escolhida
        'Array g_a_loaded_categories_default vai gaurdar os ids de todas as categorias

        ReDim g_a_loaded_categories_default(0)
        Dim l_index_loaded_categories As Int16 = 0

        Dim dr_lab_cat_def As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_intervention.GET_INTERV_CATS_DEFAULT(ComboBox3.Text, TextBox1.Text, g_selected_soft, conn, dr_lab_cat_def) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR LOADING DEFAULT LAB CATEGORIS -  ComboBox3_SelectedIndexChanged", MsgBoxStyle.Critical)

        Else

            ComboBox4.Items.Add("ALL")

            While dr_lab_cat_def.Read()

                ComboBox4.Items.Add(dr_lab_cat_def.Item(1))
                g_a_loaded_categories_default(l_index_loaded_categories) = dr_lab_cat_def.Item(0)
                l_index_loaded_categories = l_index_loaded_categories + 1
                ReDim Preserve g_a_loaded_categories_default(l_index_loaded_categories)

            End While

        End If

        dr_lab_cat_def.Dispose()
        dr_lab_cat_def.Close()

        CheckedListBox1.Items.Clear()

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

        '1 - Determinar o id_content da categoria selecionada
        If ComboBox4.SelectedIndex = 0 Then
            g_selected_category = 0
        Else
            g_selected_category = g_a_loaded_categories_default(ComboBox4.SelectedIndex - 1)
        End If

        Cursor = Cursors.WaitCursor
        CheckedListBox2.Items.Clear()
        ''2 - Carregar a grelha de análises por categoria
        ''e    
        ''3 - Criar estrutura com os elementos das análises carregados
        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_intervention.GET_INTERVS_DEFAULT_BY_CAT(TextBox1.Text, g_selected_soft, ComboBox3.SelectedItem.ToString, g_selected_category, conn, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING LAB TESTS BY CATEGORY >> ComboBox4_SelectedIndexChanged")

        Else
            ReDim g_a_loaded_interventions_default(0) ''Limpar estrutura
            Dim l_dimension_array_loaded_interventions As Int64 = 0

            While dr.Read()

                CheckedListBox2.Items.Add(dr.Item(2))

                ReDim Preserve g_a_loaded_interventions_default(l_dimension_array_loaded_interventions)

                g_a_loaded_interventions_default(l_dimension_array_loaded_interventions).id_content_category = dr.Item(0)
                g_a_loaded_interventions_default(l_dimension_array_loaded_interventions).id_content_intervention = dr.Item(1)
                g_a_loaded_interventions_default(l_dimension_array_loaded_interventions).desc_intervention = dr.Item(2)

                l_dimension_array_loaded_interventions = l_dimension_array_loaded_interventions + 1

            End While
        End If
        dr.Dispose()
        dr.Close()

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If CheckedListBox2.Items.Count() > 0 Then
            For i As Integer = 0 To CheckedListBox2.Items.Count - 1
                CheckedListBox2.SetItemChecked(i, True)
            Next
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        If CheckedListBox2.Items.Count() > 0 Then
            For i As Integer = 0 To CheckedListBox2.Items.Count - 1
                CheckedListBox2.SetItemChecked(i, False)
            Next
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Cursor = Cursors.WaitCursor
        For Each indexChecked In CheckedListBox2.CheckedIndices
            'If para verificar se já está incluido na checkbox da direita

            Dim l_record_already_selected As Boolean = False

            Dim j As Integer = 0

            For j = 0 To CheckedListBox1.Items.Count() - 1

                If (g_a_loaded_interventions_default(indexChecked.ToString()).id_content_intervention = g_a_selected_default_interventions(j).id_content_intervention And g_a_loaded_interventions_default(indexChecked.ToString()).id_content_category = g_a_selected_default_interventions(j).id_content_category) Then

                    l_record_already_selected = True
                    Exit For

                End If

            Next

            If l_record_already_selected = False Then

                ReDim Preserve g_a_selected_default_interventions(g_index_selected_intervention_from_default)

                g_a_selected_default_interventions(g_index_selected_intervention_from_default).id_content_category = g_a_loaded_interventions_default(indexChecked.ToString()).id_content_category
                g_a_selected_default_interventions(g_index_selected_intervention_from_default).id_content_intervention = g_a_loaded_interventions_default(indexChecked.ToString()).id_content_intervention
                g_a_selected_default_interventions(g_index_selected_intervention_from_default).desc_intervention = g_a_loaded_interventions_default(indexChecked.ToString()).desc_intervention

                CheckedListBox1.Items.Add((g_a_selected_default_interventions(g_index_selected_intervention_from_default).desc_intervention))

                CheckedListBox1.SetItemChecked((CheckedListBox1.Items.Count() - 1), True)

                g_index_selected_intervention_from_default = g_index_selected_intervention_from_default + 1

            End If

        Next
        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click

        If CheckedListBox1.Items.Count() > 0 Then
            For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, True)
            Next
        End If

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

        If CheckedListBox1.Items.Count() > 0 Then
            For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, False)
            Next
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Cursor = Cursors.WaitCursor

        'Se foram escolhidas análises do default para serem gravadas
        If CheckedListBox1.Items.Count() > 0 Then

            Dim l_a_checked_intervs() As INTERVENTIONS_API.interventions_default
            Dim l_index As Integer = 0

            For Each indexChecked In CheckedListBox1.CheckedIndices

                ReDim Preserve l_a_checked_intervs(l_index)

                l_a_checked_intervs(l_index).id_content_intervention = g_a_selected_default_interventions(indexChecked).id_content_intervention
                l_a_checked_intervs(l_index).id_content_category = g_a_selected_default_interventions(indexChecked).id_content_category
                l_a_checked_intervs(l_index).desc_intervention = g_a_selected_default_interventions(indexChecked).desc_intervention

                l_index = l_index + 1

            Next

            If db_intervention.SET_INTERVENTIONS(TextBox1.Text, l_a_checked_intervs, conn) Then
                If db_intervention.SET_INTERVS_TRANSLATION(TextBox1.Text, l_a_checked_intervs, conn) Then
                    If db_intervention.SET_INTERV_INT_CAT(TextBox1.Text, g_selected_soft, l_a_checked_intervs, conn) Then
                        If db_intervention.SET_INTERV_DEP_CLIN_SERV(TextBox1.Text, g_selected_soft, l_a_checked_intervs, conn) Then

                            MsgBox("Record(s) successfully inserted.", vbInformation)

                            '1 - Processo Limpeza
                            '1.1 - Limpar a box de análises a gravar no alert
                            CheckedListBox1.Items.Clear()

                            '1.2 - Remover o check das análises do default
                            For i As Integer = 0 To CheckedListBox2.Items.Count - 1

                                CheckedListBox2.SetItemChecked(i, False)

                            Next

                            '1.3 - Limpar g_a_selected_default_analysis (Array de analises do default selecionadas pelo utilizador)
                            ReDim g_a_selected_default_interventions(0)
                            g_index_selected_intervention_from_default = 0

                            '1.4 - Limpar a caixa de categorias de análises do ALERT
                            ComboBox5.Items.Clear()
                            ComboBox5.SelectedItem = ""

                            'Obter a nova lista de categorias do ALERT (foi atualizada por causa do último INSERT)
                            Dim dr_exam_cat As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                            If Not db_intervention.GET_INTERV_CATS_INST_SOFT(TextBox1.Text, g_selected_soft, conn, dr_exam_cat) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                                MsgBox("ERROR LOADING LAB CATEGORIES FROM INSTITUTION!", vbCritical)

                            Else

                                ComboBox5.Items.Add("ALL")

                                ReDim g_a_interv_cats_alert(0)
                                g_a_interv_cats_alert(0) = 0

                                Dim l_index_ec As Int16 = 1

                                While dr_exam_cat.Read()

                                    ComboBox5.Items.Add(dr_exam_cat.Item(1))
                                    ReDim Preserve g_a_interv_cats_alert(l_index_ec)
                                    g_a_interv_cats_alert(l_index_ec) = dr_exam_cat.Item(0)
                                    l_index_ec = l_index_ec + 1

                                End While

                            End If

                            dr_exam_cat.Dispose()
                            dr_exam_cat.Close()

                            '1.5 - Limpar as análises do ALERT apresentadas na BOX 3
                            'Isto porque podem ter sido adicionadas análises à categoria selecionada
                            CheckedListBox3.Items.Clear()

                            ReDim g_a_intervs_alert(0)
                            g_dimension_intervs_alert = 0
                        Else

                            MsgBox("ERROR INSERTING INTERV_DEP_CLIN_SERV!", vbCritical)

                        End If

                    Else

                        MsgBox("ERROR INSERTING INTERV_INT_CATS!", vbCritical)

                    End If

                Else

                        MsgBox("ERROR INSERTING INTERVENTIONS TRANSLATIONS!", vbCritical)

                End If

            Else
                    MsgBox("ERROR INSERTING INTERVENTIONS!", vbCritical)
            End If

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        CheckedListBox3.Items.Clear()

        Dim dr_intervs As OracleDataReader

        Dim g_selected_category_alert As String = ""

        g_selected_category_alert = g_a_interv_cats_alert(ComboBox5.SelectedIndex)

        g_dimension_intervs_alert = 0
        ReDim g_a_intervs_alert(g_dimension_intervs_alert)

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_intervention.GET_INTERVS_INST_SOFT(TextBox1.Text, g_selected_soft, g_selected_category_alert, conn, dr_intervs) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING LAB EXAMS FROM INSTITUTION!", MsgBoxStyle.Critical)

        Else

            While dr_intervs.Read()

                g_a_intervs_alert(g_dimension_intervs_alert).id_content_category = dr_intervs.Item(0)
                g_a_intervs_alert(g_dimension_intervs_alert).id_content_intervention = dr_intervs.Item(1)
                g_a_intervs_alert(g_dimension_intervs_alert).desc_intervention = dr_intervs.Item(2)

                g_dimension_intervs_alert = g_dimension_intervs_alert + 1
                ReDim Preserve g_a_intervs_alert(g_dimension_intervs_alert)

                CheckedListBox3.Items.Add((dr_intervs.Item(2)))

            End While

        End If

        dr_intervs.Dispose()
        dr_intervs.Close()

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        If CheckedListBox3.Items.Count() > 0 Then
            For i As Integer = 0 To CheckedListBox3.Items.Count - 1
                CheckedListBox3.SetItemChecked(i, True)
            Next
        End If

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        If CheckedListBox3.Items.Count() > 0 Then
            For i As Integer = 0 To CheckedListBox3.Items.Count - 1
                CheckedListBox3.SetItemChecked(i, False)
            Next
        End If

    End Sub
End Class
Imports Oracle.DataAccess.Client
Public Class SR_Procedures

    Dim db_access_general As New General
    Dim db_sr_procedure As New SR_PROCEDURES_API
    Dim oradb As String
    Dim conn As New OracleConnection

    Dim g_selected_soft As Int16 = -1
    ''Array que vai guardar os dep_clin_serv da instituição
    Dim g_a_dep_clin_serv_inst() As Int64
    Dim g_id_dep_clin_serv As Int64 = 0 'Variavel que vai guardar o id do dep_clin_serv_selecionado

    Dim g_codification As String = ""

    'Array que vai guardar os dados dos procedimentos carregadas do default
    Dim g_a_loaded_interventions_default() As SR_PROCEDURES_API.sr_interventions_default
    Dim g_a_selected_default_interventions() As SR_PROCEDURES_API.sr_interventions_default
    Dim g_index_selected_intervention_from_default As Integer = 0

    'Array que vai guardar as análises carregadas do ALERT
    Dim g_a_intervs_alert() As SR_PROCEDURES_API.sr_interventions_default
    Dim g_dimension_intervs_alert As Int64 = 0

    Public Sub New(ByVal i_oradb As String)

        InitializeComponent()
        oradb = i_oradb
        conn = New OracleConnection(oradb)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Cursor = Cursors.WaitCursor

        'Limpar arrays
        'g_selected_soft = -1
        'ReDim g_a_dep_clin_serv_inst(0)
        'g_id_dep_clin_serv = 0
        'ReDim g_a_loaded_categories_default(0)
        'g_selected_category = -1
        'ReDim g_a_loaded_interventions_default(0)
        'ReDim g_a_selected_default_interventions(0)
        'g_index_selected_intervention_from_default = 0
        'ReDim g_a_interv_cats_alert(0)
        'ReDim g_a_interv_cats_alert(0)
        'g_dimension_intervs_alert = 0
        'ReDim g_a_intervs_for_clinical_service(0)
        'g_dimension_intervs_cs = 0
        'ReDim g_a_selected_intervs_delete_cs(0)

        g_codification = ""


        If TextBox1.Text <> "" Then

            ComboBox1.Text = db_access_general.GET_INSTITUTION(TextBox1.Text, conn)

            ComboBox3.Text = ""
            ComboBox3.Items.Clear()

            CheckedListBox2.Items.Clear()

            CheckedListBox1.Items.Clear()

            CheckedListBox3.Items.Clear()

            ComboBox6.Text = ""
            ComboBox6.Items.Clear()
            CheckedListBox4.Items.Clear()

            ComboBox2.Items.Clear()
            ComboBox2.Text = ""

            '1 - Fill Version combobox

            Dim dr_def_versions As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not db_sr_procedure.GET_DEFAULT_VERSIONS(TextBox1.Text, conn, dr_def_versions) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                MsgBox("ERROR LOADING DEFAULT VERSIONS -  ComboBox2_SelectedIndexChanged", MsgBoxStyle.Critical)

            Else

                While dr_def_versions.Read()

                    ComboBox3.Items.Add(dr_def_versions.Item(0))

                End While

            End If

            dr_def_versions.Dispose()
            dr_def_versions.Close()

            '2 - Preencher os Clinical Services (Aqui será sempre o software ORIS)

            Dim dr_clin_serv As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not db_access_general.GET_CLIN_SERV(TextBox1.Text, 2, conn, dr_clin_serv) Then
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


            '''''''''''''''''''''''''''''''''''''''''''
            Dim dr_intervs As OracleDataReader

            g_dimension_intervs_alert = 0
            ReDim g_a_intervs_alert(g_dimension_intervs_alert)

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not db_sr_procedure.GET_INTERVS_DEP_CLIN_SERV(TextBox1.Text, 0, conn, dr_intervs) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                MsgBox("ERROR GETTING SR_INTERVENTIONS FROM INSTITUTION!", MsgBoxStyle.Critical)

            Else

                While dr_intervs.Read()

                    g_a_intervs_alert(g_dimension_intervs_alert).id_content_intervention = dr_intervs.Item(0)
                    g_a_intervs_alert(g_dimension_intervs_alert).desc_intervention = dr_intervs.Item(1)

                    g_dimension_intervs_alert = g_dimension_intervs_alert + 1
                    ReDim Preserve g_a_intervs_alert(g_dimension_intervs_alert)

                    CheckedListBox3.Items.Add((dr_intervs.Item(1)))

                End While

            End If

            dr_intervs.Dispose()
            dr_intervs.Close()

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub SR_Procedures_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Cursor = Cursors.WaitCursor

        'Limpar arrays
        g_selected_soft = -1
        ReDim g_a_dep_clin_serv_inst(0)
        g_id_dep_clin_serv = 0
        'ReDim g_a_loaded_categories_default(0)
        'g_selected_category = -1
        'ReDim g_a_loaded_interventions_default(0)
        'ReDim g_a_selected_default_interventions(0)
        'g_index_selected_intervention_from_default = 0
        'ReDim g_a_interv_cats_alert(0)
        'ReDim g_a_interv_cats_alert(0)
        'g_dimension_intervs_alert = 0
        'ReDim g_a_intervs_for_clinical_service(0)
        'g_dimension_intervs_cs = 0
        'ReDim g_a_selected_intervs_delete_cs(0)
        g_codification = ""

        TextBox1.Text = db_access_general.GET_INSTITUTION_ID(ComboBox1.SelectedIndex, conn)

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""

        ComboBox3.Text = ""
        ComboBox3.Items.Clear()

        CheckedListBox2.Items.Clear()

        CheckedListBox1.Items.Clear()

        CheckedListBox3.Items.Clear()

        ComboBox6.Text = ""
        ComboBox6.Items.Clear()
        CheckedListBox4.Items.Clear()

        ''''''''''''''''''''''''''''''''''''

        '1 - Fill Version combobox

        Dim dr_def_versions As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_sr_procedure.GET_DEFAULT_VERSIONS(TextBox1.Text, conn, dr_def_versions) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR LOADING DEFAULT VERSIONS -  ComboBox2_SelectedIndexChanged", MsgBoxStyle.Critical)

        Else

            While dr_def_versions.Read()

                ComboBox3.Items.Add(dr_def_versions.Item(0))

            End While

        End If

        dr_def_versions.Dispose()
        dr_def_versions.Close()

        '2 - Preencher os Clinical Services (Aqui será sempre o software ORIS)

        Dim dr_clin_serv As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_CLIN_SERV(TextBox1.Text, 2, conn, dr_clin_serv) Then
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

        Dim dr_intervs As OracleDataReader

        g_dimension_intervs_alert = 0
        ReDim g_a_intervs_alert(g_dimension_intervs_alert)

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_sr_procedure.GET_INTERVS_DEP_CLIN_SERV(TextBox1.Text, 0, conn, dr_intervs) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING SR_INTERVENTIONS FROM INSTITUTION!", MsgBoxStyle.Critical)

        Else

            While dr_intervs.Read()

                g_a_intervs_alert(g_dimension_intervs_alert).id_content_intervention = dr_intervs.Item(0)
                g_a_intervs_alert(g_dimension_intervs_alert).desc_intervention = dr_intervs.Item(1)

                g_dimension_intervs_alert = g_dimension_intervs_alert + 1
                ReDim Preserve g_a_intervs_alert(g_dimension_intervs_alert)

                CheckedListBox3.Items.Add((dr_intervs.Item(1)))

            End While

        End If

        dr_intervs.Dispose()
        dr_intervs.Close()

        Cursor = Cursors.Arrow
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""

        'Limpar arrays
        'ReDim g_a_loaded_categories_default(0)
        'g_selected_category = -1
        'ReDim g_a_loaded_interventions_default(0)
        'ReDim g_a_selected_default_interventions(0)
        'g_index_selected_intervention_from_default = 0

        g_codification = ""

        CheckedListBox2.Items.Clear()

        'Determinar as categorias disponíveis para a versão escolhida
        'Array g_a_loaded_categories_default vai gaurdar os ids de todas as categorias

        'ReDim g_a_loaded_categories_default(0)
        'Dim l_index_loaded_categories As Int16 = 0

        Dim dr_codification As OracleDataReader


        If db_sr_procedure.GET_DEFAULT_CODIFICATION(TextBox1.Text, ComboBox3.Text, conn, dr_codification) Then

            While dr_codification.Read()

                ComboBox2.Items.Add(dr_codification.Item(0))

            End While

        Else

            MsgBox("ERROR LOADING CODIFICATION!", vbCritical)

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        g_codification = ComboBox2.SelectedItem

        Cursor = Cursors.WaitCursor

        CheckedListBox2.Items.Clear()
        ''2 - Carregar a grelha de Surgical Interventions
        ''e    
        ''3 - Criar estrutura com os elementos das análises carregados
        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_sr_procedure.GET_DEFAULT_SR_INTERVENTIONS(TextBox1.Text, ComboBox3.SelectedItem.ToString, g_codification, conn, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING SURGICAL INTERVENTIONS FROM DEFAULT!", vbCritical)

        Else
            ReDim g_a_loaded_interventions_default(0) ''Limpar estrutura
            Dim l_dimension_array_loaded_interventions As Int64 = 0

            While dr.Read()

                CheckedListBox2.Items.Add(dr.Item(1))

                ReDim Preserve g_a_loaded_interventions_default(l_dimension_array_loaded_interventions)

                g_a_loaded_interventions_default(l_dimension_array_loaded_interventions).id_content_intervention = dr.Item(0)
                g_a_loaded_interventions_default(l_dimension_array_loaded_interventions).desc_intervention = dr.Item(1)

                l_dimension_array_loaded_interventions = l_dimension_array_loaded_interventions + 1

            End While
        End If
        dr.Dispose()
        dr.Close()

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Cursor = Cursors.WaitCursor
        For Each indexChecked In CheckedListBox2.CheckedIndices
            'If para verificar se já está incluido na checkbox da direita

            Dim l_record_already_selected As Boolean = False

            Dim j As Integer = 0

            For j = 0 To CheckedListBox1.Items.Count() - 1

                If (g_a_loaded_interventions_default(indexChecked.ToString()).id_content_intervention = g_a_selected_default_interventions(j).id_content_intervention) Then

                    l_record_already_selected = True
                    Exit For

                End If

            Next

            If l_record_already_selected = False Then

                ReDim Preserve g_a_selected_default_interventions(g_index_selected_intervention_from_default)

                g_a_selected_default_interventions(g_index_selected_intervention_from_default).id_content_intervention = g_a_loaded_interventions_default(indexChecked.ToString()).id_content_intervention
                g_a_selected_default_interventions(g_index_selected_intervention_from_default).desc_intervention = g_a_loaded_interventions_default(indexChecked.ToString()).desc_intervention

                CheckedListBox1.Items.Add((g_a_selected_default_interventions(g_index_selected_intervention_from_default).desc_intervention))

                CheckedListBox1.SetItemChecked((CheckedListBox1.Items.Count() - 1), True)

                g_index_selected_intervention_from_default = g_index_selected_intervention_from_default + 1

            End If

        Next
        Cursor = Cursors.Arrow
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Cursor = Cursors.WaitCursor

        'Se foram escolhidas interventions do default para serem gravadas
        If CheckedListBox1.Items.Count() > 0 Then

            Dim l_a_checked_intervs() As SR_PROCEDURES_API.sr_interventions_default
            Dim l_index As Integer = 0

            For Each indexChecked In CheckedListBox1.CheckedIndices

                ReDim Preserve l_a_checked_intervs(l_index)

                'Só interessa passar o id_content
                l_a_checked_intervs(l_index).id_content_intervention = g_a_selected_default_interventions(indexChecked).id_content_intervention

                l_index = l_index + 1

            Next

            If db_sr_procedure.SET_SR_INTERVENTIONS(TextBox1.Text, g_codification, l_a_checked_intervs, conn) Then
                If db_sr_procedure.SET_SR_INTERVS_TRANSLATION(TextBox1.Text, l_a_checked_intervs, conn) Then
                    If db_sr_procedure.SET_SR_INTERV_DEP_CLIN_SERV(TextBox1.Text, l_a_checked_intervs, conn) Then
                        'If db_intervention.SET_DEFAULT_INTERV_DEP_CLIN_SERV(TextBox1.Text, g_selected_soft, l_a_checked_intervs, g_procedure_type, conn) Then

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

                        'Obter a nova lista de categorias do ALERT (foi atualizada por causa do último INSERT)
                        Dim dr_exam_cat As OracleDataReader

                        'If Not db_intervention.GET_INTERV_CATS_INST_SOFT(TextBox1.Text, g_selected_soft, g_procedure_type, conn, dr_exam_cat) Then


                        'MsgBox("ERROR LOADING LAB CATEGORIES FROM INSTITUTION!", vbCritical)

                        'Else

                        'ComboBox5.Items.Add("ALL")

                        'ReDim g_a_interv_cats_alert(0)
                        'g_a_interv_cats_alert(0) = 0

                        'Dim l_index_ec As Int16 = 1

                        'While dr_exam_cat.Read()

                        'ComboBox5.Items.Add(dr_exam_cat.Item(1))
                        ' ReDim Preserve g_a_interv_cats_alert(l_index_ec)
                        'g_a_interv_cats_alert(l_index_ec) = dr_exam_cat.Item(0)
                        'l_index_ec = l_index_ec + 1

                        'End While

                        'End If

                        'dr_exam_cat.Dispose()
                        'dr_exam_cat.Close()

                        '1.5 - Limpar as análises do ALERT apresentadas na BOX 3
                        'Isto porque podem ter sido adicionadas análises à categoria selecionada
                        'CheckedListBox3.Items.Clear()

                        'ReDim g_a_intervs_alert(0)
                        'g_dimension_intervs_alert = 0
                        'Else

                        'MsgBox("ERROR INSERTING INTERV_DEP_CLIN_SERV!", vbCritical)
                        '
                        ' End If

                    Else

                        MsgBox("ERROR INSERTING SR_INTERV_DEP_CLIN_SERV!!", vbCritical)

                    End If

                Else

                        MsgBox("ERROR INSERTING SR_INTERVENTIONS TRANSLATIONS!", vbCritical)

                End If

            Else
                    MsgBox("ERROR INSERTING SR_INTERVENTIONS!", vbCritical)
            End If

        End If

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

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        'Array que vai guardar as sr_interventions 
        Dim l_a_sr_interventions_to_delete() As SR_PROCEDURES_API.sr_interventions_default
        Dim l_index_sr_intervs_to_delete As Int64 = 0

        If CheckedListBox3.CheckedIndices.Count > 0 Then

            Cursor = Cursors.WaitCursor

            Dim l_sucess As Boolean = True

            'Ciclo para correr todos os registos do ALERT marcados com o check
            For Each indexChecked In CheckedListBox3.CheckedIndices

                ReDim Preserve l_a_sr_interventions_to_delete(l_index_sr_intervs_to_delete)
                l_a_sr_interventions_to_delete(l_index_sr_intervs_to_delete).id_content_intervention = g_a_intervs_alert(indexChecked).id_content_intervention
                l_index_sr_intervs_to_delete = l_index_sr_intervs_to_delete + 1

            Next

            If Not db_sr_procedure.DELETE_SR_INTERV_DEP_CLIN_SERV(TextBox1.Text, l_a_sr_interventions_to_delete, 0, conn) Then

                l_sucess = False

            Else

                CheckedListBox3.Items.Clear()
                Dim dr_intervs As OracleDataReader

                g_dimension_intervs_alert = 0
                ReDim g_a_intervs_alert(g_dimension_intervs_alert)

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                If Not db_sr_procedure.GET_INTERVS_DEP_CLIN_SERV(TextBox1.Text, 0, conn, dr_intervs) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                    MsgBox("ERROR GETTING SR_INTERVENTIONS FROM INSTITUTION!", MsgBoxStyle.Critical)

                Else

                    While dr_intervs.Read()

                        g_a_intervs_alert(g_dimension_intervs_alert).id_content_intervention = dr_intervs.Item(0)
                        g_a_intervs_alert(g_dimension_intervs_alert).desc_intervention = dr_intervs.Item(1)

                        g_dimension_intervs_alert = g_dimension_intervs_alert + 1
                        ReDim Preserve g_a_intervs_alert(g_dimension_intervs_alert)

                        CheckedListBox3.Items.Add((dr_intervs.Item(1)))

                    End While

                End If

                dr_intervs.Dispose()
                dr_intervs.Close()

                ''Limpar o array de sr_interventions do Clinical Service
                'Limpar arrays - TRATAR DISTO!!!!
                'ReDim g_a_intervs_for_clinical_service(0)
                'ReDim g_a_intervs_alert(0)

                'g_dimension_intervs_cs = 0
                'g_dimension_intervs_alert = 0


                ''APAGAR da grelah de favoritos (já foi apagado anteriormente)
                ''4 - Limpar a box 

                '                CheckedListBox4.Items.Clear()

                '                '5 - Determinar os exames disponíveis como mais frequentes para esse dep_clin_serv
                '                Dim dr_delete As OracleDataReader

                '#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                '                If Not db_intervention.GET_FREQ_INTERVS(TextBox1.Text, g_selected_soft, g_procedure_type, g_id_dep_clin_serv, conn, dr_delete) Then
                '#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                '                    MsgBox("ERROR GETTING INTERVENTIONS_DEP_CLIN_SERV.", vbCritical)

                '                Else

                '                    Dim i As Integer = 0

                '                    '6 - Ler cursor e popular o campo
                '                    While dr_delete.Read()

                '                        CheckedListBox4.Items.Add(dr_delete.Item(1))

                '                        ReDim Preserve g_a_intervs_for_clinical_service(g_dimension_intervs_cs)


                '                        g_a_intervs_for_clinical_service(g_dimension_intervs_cs).id_content_intervention = dr_delete.Item(0)
                '                        g_a_intervs_for_clinical_service(g_dimension_intervs_cs).desc_intervention = dr_delete.Item(1)
                '                        g_a_intervs_for_clinical_service(g_dimension_intervs_cs).flg_new = "N"

                '                        g_dimension_intervs_cs = g_dimension_intervs_cs + 1

                '                    End While

                '                End If

                '                dr_delete.Dispose()
                '                dr_delete.Close()

            End If

            If l_sucess = True Then

                MsgBox("Records Deleted.")

            Else

                MsgBox("ERROR DELETING SR_INTERVENTIONS.", vbCritical)

            End If


            Cursor = Cursors.Arrow

            End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        conn.Close()
        conn.Dispose()

        Dim form1 As New Form1()
        form1.g_oradb = oradb

        Me.Enabled = False

        Me.Dispose()

        form1.Show()
    End Sub
End Class
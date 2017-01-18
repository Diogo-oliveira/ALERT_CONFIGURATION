Imports Oracle.DataAccess.Client
Public Class LAB_TESTS

    Dim db_access_general As New General
    Dim db_labs As New LABS_API
    Dim oradb As String = "Data Source=QC4V26522;User Id=alert_config;Password=qcteam"
    Dim l_selected_soft As Int16 = -1

    Dim l_loaded_categories_default() As String ' Array que vai guardar os id_contents das categorias carregadas do default
    Dim l_selected_category As String = ""

    Dim l_loaded_analysis_default() As LABS_API.analysis_default 'Array que vai guardar os id_contents das análises carregadas do default
    Dim l_selected_default_analysis() As LABS_API.analysis_default 'Array que vai guardar os id_contents das análises selecionadas do default

    Dim l_index_selected_analysis_from_default As Integer = 0 ''Variavel utilizada no botão de adicionar à box da direita (CHECKBOX 1)

    'Ligação à BD
    Dim conn_labs As New OracleConnection(oradb)

    'Array que vai guardar os quartos da instituição
    Dim l_loaded_rooms() As Int64

    'Variavel que guarda o id_room a ser inserido
    Dim l_selected_room As Int64 = -1

    'Array que vai guardar as categorias disponíveis no ALERT
    Dim l_lab_cats_alert() As String

    'Array que vai guardar as análises carregadas do ALERT
    Dim l_labs_alert() As LABS_API.analysis_alert

    'Array que vai guardar as análises selecionadas do ALERT
    Dim l_labs_selected_from_alert() As LABS_API.analysis_alert

    Dim l_index_selected_labs_from_alert As Integer = 0 ''Variavel utilizada no botão de adicionar à box da direita (CHECKBOX 4 - do alert para o clinical service)


    Dim a_labs_for_clinical_service() As LABS_API.analysis_alert_flg 'Array que vai guardar os exames do ALERT e os exames que existem no clinical service. A flag irá indicar se é oou não para introduzir na categoria
    Dim l_dimension_labs_cs As Integer = 0

    ''Array que vai guardar os dep_clin_serv da instituição
    Dim a_dep_clin_serv_inst() As Int64

    Dim l_id_dep_clin_serv As Int64 = 0 'Variavel que vai guardar o id do dep_clin_serv_selecionado

    Dim l_selected_labs() As String ' Array para remover análises do alert

    Private Sub LAB_TESTS_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            'Estabelecer ligação à BD
            conn_labs.Open()

        Catch ex As Exception

            MsgBox("Error Connecting to DataBase!", vbCritical)

        End Try

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_ALL_INSTITUTIONS(conn_labs, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING ALL INSTITUTIONS")

        Else

            While dr.Read()

                ComboBox1.Items.Add(dr.Item(0))

            End While

        End If

        dr.Dispose()
        dr.Close()

        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        conn_labs.FlushCache()
        conn_labs.Dispose()
        conn_labs.Close()

        Dim form1 As New Form1

        Me.Enabled = False

        Me.Dispose()

        form1.Show()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Cursor = Cursors.WaitCursor

        l_selected_room = -1

        If TextBox1.Text <> "" Then

            ComboBox1.Text = db_access_general.GET_INSTITUTION(TextBox1.Text, conn_labs)

            ComboBox2.Items.Clear()
            ComboBox2.Text = ""

            Dim dr As OracleDataReader

            ReDim l_loaded_rooms(0)
            Dim i_index_room As Int32 = 0


#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not db_access_general.GET_SOFT_INST(TextBox1.Text, conn_labs, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                MsgBox("ERROR GETTING SOFTWARES", vbCritical)

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

                l_selected_category = ""

            End If

            If Not db_access_general.GET_LAB_ROOMS(TextBox1.Text, conn_labs, dr) Then

                MsgBox("ERROR GETTING LAB ROOMS", vbCritical)

            Else

                ComboBox7.Items.Clear()

                While dr.Read()


                    ReDim Preserve l_loaded_rooms(i_index_room)
                    ComboBox7.Items.Add(dr.Item(0))
                    l_loaded_rooms(i_index_room) = dr.Item(1)
                    i_index_room = i_index_room + 1

                End While

            End If

            dr.Dispose()
            dr.Close()

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

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
        ComboBox7.Items.Clear()
        ComboBox7.Text = ""

        l_selected_soft = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, TextBox1.Text, conn_labs)

        '1 - Fill Version combobox

        Dim dr_def_versions As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_labs.GET_DEFAULT_VERSIONS(TextBox1.Text, l_selected_soft, conn_labs, dr_def_versions) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR LOADING DEFAULT VERSIONS -  ComboBox2_SelectedIndexChanged", MsgBoxStyle.Critical)

        Else

            While dr_def_versions.Read()

                ComboBox3.Items.Add(dr_def_versions.Item(0))

            End While

            dr_def_versions.Dispose()
            dr_def_versions.Close()

        End If

        ''''''''''''''''''''''
        'Box de categorias na instituição/software

        Dim dr_exam_cat As OracleDataReader

        If Not db_labs.GET_LAB_CATS_INST_SOFT(TextBox1.Text, l_selected_soft, conn_labs, dr_exam_cat) Then

            MsgBox("ERROR LOADING LAB CATEGORIES FROM INSTITUTION", vbCritical)
            dr_exam_cat.Dispose()
            dr_exam_cat.Close()

        Else

            ComboBox5.Items.Add("ALL")

            ReDim l_lab_cats_alert(0)
            l_lab_cats_alert(0) = 0

            Dim l_index As Int16 = 1

            While dr_exam_cat.Read()

                ComboBox5.Items.Add(dr_exam_cat.Item(1))
                ReDim Preserve l_lab_cats_alert(l_index)
                l_lab_cats_alert(l_index) = dr_exam_cat.Item(0)
                l_index = l_index + 1

            End While

        End If

        dr_exam_cat.Dispose()
        dr_exam_cat.Close()

        ''''''''''''''''''''''
        'Preencher os Clinical Services

        Dim dr_clin_serv As OracleDataReader

        If Not db_access_general.GET_CLIN_SERV(TextBox1.Text, l_selected_soft, conn_labs, dr_clin_serv) Then

            MsgBox("ERROR GETTING CLINICAL SERVICES")
            dr_clin_serv.Dispose()
            dr_clin_serv.Close()

        Else

            Dim i As Integer = 0

            Dim l_index_dep_clin_serv As Integer = 0
            ReDim a_dep_clin_serv_inst(l_index_dep_clin_serv)

            While dr_clin_serv.Read()

                ComboBox6.Items.Add(dr_clin_serv.Item(0))

                ReDim Preserve a_dep_clin_serv_inst(l_index_dep_clin_serv)
                a_dep_clin_serv_inst(l_index_dep_clin_serv) = dr_clin_serv.Item(1)
                l_index_dep_clin_serv = l_index_dep_clin_serv + 1

            End While

        End If

        dr_clin_serv.Dispose()
        dr_clin_serv.Close()

        '''''''''''''''''''''''''''''''''''
        Dim dr As OracleDataReader

        If Not db_access_general.GET_LAB_ROOMS(TextBox1.Text, conn_labs, dr) Then

            MsgBox("ERROR GETTING LAB ROOMS", vbCritical)

            dr.Dispose()
            dr.Close()

        Else

            While dr.Read()

                ComboBox7.Items.Add(dr.Item(0))

            End While

        End If

        dr.Dispose()
        dr.Close()

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        ComboBox4.Items.Clear()
        ComboBox4.Text = ""

        CheckedListBox2.Items.Clear()

        'Determinar as categorias disponíveis para a versão escolhida
        'Array l_loaded_categories_default vai gaurdar os ids de todas as categorias

        ReDim l_loaded_categories_default(0)
        Dim l_index_loaded_categories As Int16 = 0

        Dim dr_lab_cat_def As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_labs.GET_LAB_CATS_DEFAULT(ComboBox3.Text, TextBox1.Text, l_selected_soft, conn_labs, dr_lab_cat_def) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR LOADING DEFAULT LAB CATEGORIS -  ComboBox3_SelectedIndexChanged", MsgBoxStyle.Critical)

        Else

            ComboBox4.Items.Add("ALL")

            While dr_lab_cat_def.Read()

                ComboBox4.Items.Add(dr_lab_cat_def.Item(1))
                l_loaded_categories_default(l_index_loaded_categories) = dr_lab_cat_def.Item(0)
                l_index_loaded_categories = l_index_loaded_categories + 1
                ReDim Preserve l_loaded_categories_default(l_index_loaded_categories)

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

            l_selected_category = 0

        Else

            l_selected_category = l_loaded_categories_default(ComboBox4.SelectedIndex - 1)

        End If

        Cursor = Cursors.WaitCursor

        CheckedListBox2.Items.Clear()

        ''2 - Carregar a grelha de análises por categoria
        ''e    
        ''3 - Criar estrutura com os elementos das análises carregados

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_labs.GET_LABS_DEFAULT_BY_CAT(TextBox1.Text, l_selected_soft, ComboBox3.SelectedItem.ToString, l_selected_category, conn_labs, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING LAB TESTS BY CATEGORY >> ComboBox4_SelectedIndexChanged")
            dr.Dispose()
            dr.Close()

        Else

            ReDim l_loaded_analysis_default(0) ''Limpar estrutura
            Dim l_dimension_array_loaded_analysis As Int64 = 0

            While dr.Read()

                CheckedListBox2.Items.Add(dr.Item(1) & " - [" & dr.Item(3) & "]")

                ReDim Preserve l_loaded_analysis_default(l_dimension_array_loaded_analysis)

                l_loaded_analysis_default(l_dimension_array_loaded_analysis).id_content_category = dr.Item(6)
                l_loaded_analysis_default(l_dimension_array_loaded_analysis).id_content_analysis_sample_type = dr.Item(0)
                l_loaded_analysis_default(l_dimension_array_loaded_analysis).desc_analysis_sample_type = dr.Item(1)
                l_loaded_analysis_default(l_dimension_array_loaded_analysis).id_content_sample_recipient = dr.Item(2)
                l_loaded_analysis_default(l_dimension_array_loaded_analysis).desc_analysis_sample_recipient = dr.Item(3)
                l_loaded_analysis_default(l_dimension_array_loaded_analysis).id_content_analysis = dr.Item(4)
                l_loaded_analysis_default(l_dimension_array_loaded_analysis).id_content_sample_type = dr.Item(5)

                l_dimension_array_loaded_analysis = l_dimension_array_loaded_analysis + 1

            End While

            dr.Dispose()
            dr.Close()

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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        For Each indexChecked In CheckedListBox2.CheckedIndices

            'If para verificar se já está incluido na checkbox da direita

            Dim l_record_already_selected As Boolean = False

            Dim j As Integer = 0

            For j = 0 To CheckedListBox1.Items.Count() - 1

                If (l_loaded_analysis_default(indexChecked.ToString()).id_content_analysis_sample_type = l_selected_default_analysis(j).id_content_analysis_sample_type) Then

                    l_record_already_selected = True
                    Exit For

                End If

            Next

            If l_record_already_selected = False Then

                ReDim Preserve l_selected_default_analysis(l_index_selected_analysis_from_default)

                l_selected_default_analysis(l_index_selected_analysis_from_default).id_content_analysis = l_loaded_analysis_default(indexChecked.ToString()).id_content_analysis
                l_selected_default_analysis(l_index_selected_analysis_from_default).id_content_analysis_sample_type = l_loaded_analysis_default(indexChecked.ToString()).id_content_analysis_sample_type
                l_selected_default_analysis(l_index_selected_analysis_from_default).id_content_category = l_loaded_analysis_default(indexChecked.ToString()).id_content_category
                l_selected_default_analysis(l_index_selected_analysis_from_default).id_content_sample_recipient = l_loaded_analysis_default(indexChecked.ToString()).id_content_sample_recipient
                l_selected_default_analysis(l_index_selected_analysis_from_default).id_content_sample_type = l_loaded_analysis_default(indexChecked.ToString()).id_content_sample_type
                l_selected_default_analysis(l_index_selected_analysis_from_default).desc_analysis_sample_type = l_loaded_analysis_default(indexChecked.ToString()).desc_analysis_sample_type
                l_selected_default_analysis(l_index_selected_analysis_from_default).desc_analysis_sample_recipient = l_loaded_analysis_default(indexChecked.ToString()).desc_analysis_sample_recipient

                CheckedListBox1.Items.Add((l_selected_default_analysis(l_index_selected_analysis_from_default).desc_analysis_sample_type & " [" & l_selected_default_analysis(l_index_selected_analysis_from_default).desc_analysis_sample_recipient & "]"))

                CheckedListBox1.SetItemChecked((CheckedListBox1.Items.Count() - 1), True)

                l_index_selected_analysis_from_default = l_index_selected_analysis_from_default + 1

            End If

        Next

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Cursor = Cursors.WaitCursor

        If ComboBox7.SelectedIndex >= 0 Then

            If CheckedListBox1.Items.Count() > 0 Then

                Dim l_checked_labs() As LABS_API.analysis_default
                Dim l_index As Integer = 0

                For Each indexChecked In CheckedListBox1.CheckedIndices

                    ReDim Preserve l_checked_labs(l_index)

                    l_checked_labs(l_index).id_content_category = l_selected_default_analysis(indexChecked).id_content_category
                    l_checked_labs(l_index).id_content_analysis = l_selected_default_analysis(indexChecked).id_content_analysis
                    l_checked_labs(l_index).id_content_sample_type = l_selected_default_analysis(indexChecked).id_content_sample_type
                    l_checked_labs(l_index).id_content_analysis_sample_type = l_selected_default_analysis(indexChecked).id_content_analysis_sample_type
                    l_checked_labs(l_index).id_content_sample_recipient = l_selected_default_analysis(indexChecked).id_content_sample_recipient
                    l_checked_labs(l_index).desc_analysis_sample_type = l_selected_default_analysis(indexChecked).desc_analysis_sample_type
                    l_checked_labs(l_index).desc_analysis_sample_recipient = l_selected_default_analysis(indexChecked).desc_analysis_sample_recipient

                    l_index = l_index + 1

                Next

                If db_labs.SET_EXAM_CAT(TextBox1.Text, l_checked_labs, conn_labs) Then

                    If db_labs.SET_SAMPLE_TYPE(TextBox1.Text, l_checked_labs, conn_labs) Then

                        If db_labs.SET_ANALYSIS(TextBox1.Text, l_checked_labs, conn_labs) Then

                            If db_labs.SET_ANALYSIS_SAMPLE_TYPE(TextBox1.Text, l_checked_labs, conn_labs) Then

                                If db_labs.SET_PARAMETER(TextBox1.Text, l_selected_soft, l_checked_labs, conn_labs) Then

                                    If db_labs.SET_PARAM(TextBox1.Text, l_selected_soft, l_checked_labs, conn_labs) Then

                                        If db_labs.SET_SAMPLE_RECIPIENT(TextBox1.Text, l_checked_labs, conn_labs) Then

                                            If db_labs.SET_ANALYSIS_INST_SOFT(TextBox1.Text, l_selected_soft, l_checked_labs, conn_labs) Then

                                                If db_labs.SET_ANALYSIS_INST_RECIPIENT(TextBox1.Text, l_selected_soft, l_checked_labs, conn_labs) Then

                                                    If db_labs.SET_ANALYSIS_ROOM(TextBox1.Text, l_selected_soft, l_selected_room, l_checked_labs, conn_labs) Then

                                                        MsgBox("LABORATORIAL EXAM(S) SUCCESSFULLY INSERTED.", vbInformation)

                                                        CheckedListBox1.Items.Clear()

                                                        For i As Integer = 0 To CheckedListBox2.Items.Count - 1

                                                            CheckedListBox2.SetItemChecked(i, False)

                                                        Next

                                                        ComboBox5.Items.Clear()
                                                        ComboBox5.SelectedItem = ""

                                                        Dim dr_exam_cat As OracleDataReader

                                                        If Not db_labs.GET_LAB_CATS_INST_SOFT(TextBox1.Text, l_selected_soft, conn_labs, dr_exam_cat) Then

                                                            MsgBox("ERROR LOADING LAB CATEGORIES FROM INSTITUTION", vbCritical)
                                                            dr_exam_cat.Dispose()
                                                            dr_exam_cat.Close()

                                                        Else

                                                            ComboBox5.Items.Add("ALL")

                                                            ReDim l_lab_cats_alert(0)
                                                            l_lab_cats_alert(0) = 0

                                                            Dim l_index_ec As Int16 = 1

                                                            While dr_exam_cat.Read()

                                                                ComboBox5.Items.Add(dr_exam_cat.Item(1))
                                                                ReDim Preserve l_lab_cats_alert(l_index_ec)
                                                                l_lab_cats_alert(l_index_ec) = dr_exam_cat.Item(0)
                                                                l_index_ec = l_index_ec + 1

                                                            End While

                                                        End If

                                                        dr_exam_cat.Dispose()
                                                        dr_exam_cat.Close()

                                                        CheckedListBox3.Items.Clear()

                                                        ReDim l_selected_default_analysis(0)
                                                        l_index_selected_analysis_from_default = 0

                                                    End If

                                                End If

                                            End If

                                        End If

                                    End If

                                End If

                            End If

                        End If

                    End If

                End If

            End If

        Else

            MsgBox("No Room selected!", vbCritical)

        End If



        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        TextBox1.Text = db_access_general.GET_INSTITUTION_ID(ComboBox1.SelectedIndex, conn_labs)

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""

        l_selected_room = -1

        Dim dr As OracleDataReader

        If Not db_access_general.GET_LAB_ROOMS(TextBox1.Text, conn_labs, dr) Then

            MsgBox("ERROR GETTING LAB ROOMS", vbCritical)

            dr.Dispose()
            dr.Close()

        Else

            ComboBox7.Items.Clear()

            While dr.Read()

                ComboBox7.Items.Add(dr.Item(0))

            End While

        End If

        dr.Dispose()

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""

        If Not db_access_general.GET_SOFT_INST(TextBox1.Text, conn_labs, dr) Then

            MsgBox("ERROR GETTING SOFTWARES", vbCritical)

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

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged

        l_selected_room = l_loaded_rooms(ComboBox7.SelectedIndex)

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        CheckedListBox3.Items.Clear()

        Dim dr_labs As OracleDataReader

        Dim l_selected_category As String = ""

        l_selected_category = l_lab_cats_alert(ComboBox5.SelectedIndex)

        Dim l_index As Integer = 0

        ReDim l_labs_alert(l_index)

        If Not db_labs.GET_LABS_INST_SOFT(TextBox1.Text, l_selected_soft, l_selected_category, conn_labs, dr_labs) Then

            MsgBox("ERROR GETTING LAB EXAMS FROM INSTITUTION", MsgBoxStyle.Critical)

        Else

            While dr_labs.Read()

                l_labs_alert(l_index).id_content_analysis_sample_type = dr_labs.Item(0)
                l_labs_alert(l_index).desc_analysis_sample_type = dr_labs.Item(1)
                l_labs_alert(l_index).desc_analysis_sample_recipient = dr_labs.Item(2)
                l_index = l_index + 1
                ReDim Preserve l_labs_alert(l_index)

                CheckedListBox3.Items.Add((dr_labs.Item(1)) & " - [" & dr_labs.Item(2) & "]")

            End While

            dr_labs.Dispose()
            dr_labs.Close()

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        'Ciclo para correr todos os exames selecionados na caixa da esquerda (Por Categoria)
        For Each indexChecked In CheckedListBox3.CheckedIndices

            'If para verificar se já está incluido na checkbox da direita

            Dim l_record_already_selected As Boolean = False

            Dim j As Integer = 0

            For j = 0 To CheckedListBox4.Items.Count() - 1

                If (l_labs_alert(indexChecked.ToString()).id_content_analysis_sample_type = l_labs_selected_from_alert(j).id_content_analysis_sample_type) Then

                    l_record_already_selected = True
                    Exit For

                End If

            Next

            If l_record_already_selected = False Then

                ReDim Preserve l_labs_selected_from_alert(l_index_selected_labs_from_alert)

                l_labs_selected_from_alert(l_index_selected_labs_from_alert).id_content_analysis_sample_type = l_labs_alert(indexChecked.ToString()).id_content_analysis_sample_type
                l_labs_selected_from_alert(l_index_selected_labs_from_alert).desc_analysis_sample_type = l_labs_alert(indexChecked.ToString()).desc_analysis_sample_type
                l_labs_selected_from_alert(l_index_selected_labs_from_alert).desc_analysis_sample_recipient = l_labs_alert(indexChecked.ToString()).desc_analysis_sample_recipient

                ReDim Preserve a_labs_for_clinical_service(l_dimension_labs_cs)

                a_labs_for_clinical_service(l_dimension_labs_cs).id_content_analysis_sample_type = l_labs_alert(indexChecked.ToString()).id_content_analysis_sample_type
                a_labs_for_clinical_service(l_dimension_labs_cs).desc_analysis_sample_type = l_labs_alert(indexChecked.ToString()).desc_analysis_sample_recipient
                a_labs_for_clinical_service(l_dimension_labs_cs).desc_analysis_sample_recipient = l_labs_alert(indexChecked.ToString()).desc_analysis_sample_recipient
                a_labs_for_clinical_service(l_dimension_labs_cs).flg_new = "Y"

                l_dimension_labs_cs = l_dimension_labs_cs + 1

                CheckedListBox4.Items.Add(l_labs_selected_from_alert(l_index_selected_labs_from_alert).desc_analysis_sample_type & " - [" & l_labs_selected_from_alert(l_index_selected_labs_from_alert).desc_analysis_sample_recipient & "]")
                CheckedListBox4.SetItemChecked((CheckedListBox4.Items.Count() - 1), True)

                l_index_selected_labs_from_alert = l_index_selected_labs_from_alert + 1

            End If

        Next

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

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        Cursor = Cursors.WaitCursor

        Dim result As Integer = 0
        Dim l_sucess As Boolean = True

        If (CheckedListBox3.CheckedIndices.Count = CheckedListBox3.Items.Count()) Then

            result = MsgBox("All records from the chosen category will be deleted! Confirm?", MessageBoxButtons.YesNo)

        End If

        If (result = DialogResult.Yes Or CheckedListBox3.CheckedIndices.Count < CheckedListBox3.Items.Count()) Then

            Dim indexChecked As Integer

            Dim total_selected_labs As Integer = 0

            For Each indexChecked In CheckedListBox3.CheckedIndices

                total_selected_labs = total_selected_labs + 1

            Next

            ''1 - Determinar ID_Content_ast dos registos selecionados - Feito
            ''2 - Criar função para colocar a not available na analysis_inst_soft - Feito

            For Each indexChecked In CheckedListBox3.CheckedIndices

                If Not db_labs.DELETE_ANALYSIS_INST_SOFT(TextBox1.Text, l_selected_soft, l_labs_alert(indexChecked).id_content_analysis_sample_type, conn_labs) Then

                    l_sucess = False

                End If
            Next

            ''3 - Refresh à grelha
            ''3.1 - Se estão a ser apagados todos os registos de uma categoria:
            If CheckedListBox3.CheckedIndices.Count = CheckedListBox3.Items.Count() Then

                CheckedListBox3.Items.Clear()
                ComboBox5.Items.Clear()
                ComboBox5.Text = ""

                Dim dr_exam_cat As OracleDataReader

                If Not db_labs.GET_LAB_CATS_INST_SOFT(TextBox1.Text, l_selected_soft, conn_labs, dr_exam_cat) Then

                    MsgBox("ERROR LOADING LAB CATEGORIES FROM INSTITUTION", vbCritical)
                    dr_exam_cat.Dispose()
                    dr_exam_cat.Close()

                Else

                    ComboBox5.Items.Add("ALL")

                    ReDim l_lab_cats_alert(0)
                    l_lab_cats_alert(0) = 0

                    Dim l_index As Int16 = 1

                    While dr_exam_cat.Read()

                        ComboBox5.Items.Add(dr_exam_cat.Item(1))
                        ReDim Preserve l_lab_cats_alert(l_index)
                        l_lab_cats_alert(l_index) = dr_exam_cat.Item(0)
                        l_index = l_index + 1

                    End While

                End If

                dr_exam_cat.Dispose()
                dr_exam_cat.Close()

                'Limpar arrays
                ReDim a_labs_for_clinical_service(0)
                ReDim l_labs_selected_from_alert(0)

                l_dimension_labs_cs = 0
                l_index_selected_labs_from_alert = 0

            Else '3.2 - Eliminar apenas os registos selecionados

                CheckedListBox3.Items.Clear()

                Dim dr_labs As OracleDataReader

                Dim l_selected_category As String = ""

                l_selected_category = l_lab_cats_alert(ComboBox5.SelectedIndex)

                Dim l_index As Integer = 0

                ReDim l_labs_alert(l_index)

                If Not db_labs.GET_LABS_INST_SOFT(TextBox1.Text, l_selected_soft, l_selected_category, conn_labs, dr_labs) Then

                    MsgBox("ERROR GETTING LAB EXAMS FROM INSTITUTION", MsgBoxStyle.Critical)

                Else

                    While dr_labs.Read()

                        l_labs_alert(l_index).id_content_analysis_sample_type = dr_labs.Item(0)
                        l_labs_alert(l_index).desc_analysis_sample_type = dr_labs.Item(1)
                        l_labs_alert(l_index).desc_analysis_sample_recipient = dr_labs.Item(2)
                        l_index = l_index + 1
                        ReDim Preserve l_labs_alert(l_index)

                        CheckedListBox3.Items.Add((dr_labs.Item(1)) & " - [" & dr_labs.Item(2) & "]")

                    End While

                    dr_labs.Dispose()
                    dr_labs.Close()

                    'Limpar arrays
                    ReDim a_labs_for_clinical_service(0)
                    ReDim l_labs_selected_from_alert(0)

                    l_dimension_labs_cs = 0
                    l_index_selected_labs_from_alert = 0

                End If

            End If

            ''4 - Mensagem de sucesso no final de todos os registos. (modificar mensagem de erro para surgir apenas uma vez.
            If l_sucess = False Then

                MsgBox("ERROR DELETING ANALYSIS_INST_SOFT", vbCritical)

            Else

                MsgBox("RECORDS SUCCESSFULLY DELETED.", vbInformation)

            End If

        End If

        ''APAGAR da grelah de favoritos

        '4 - Limpar a box e os arrays
        ReDim Preserve a_labs_for_clinical_service(0)
        l_dimension_labs_cs = 0

        ReDim l_labs_selected_from_alert(0)
        l_index_selected_labs_from_alert = 0

        CheckedListBox4.Items.Clear()

        '5 - Determinar os exames disponíveis como mais frequentes para esse dep_clin_serv
        Dim dr_delete As OracleDataReader

        If Not db_labs.GET_ANALYSIS_DEP_CLIN_SERV(TextBox1.Text, l_selected_soft, l_id_dep_clin_serv, dr_delete, conn_labs) Then

            MsgBox("ERROR GETTING ANALYSIS_DEP_CLIN_SERV.", vbCritical)

        Else

            Dim i As Integer = 0

            '6 - Ler cursor e popular o campo
            While dr_delete.Read()

                CheckedListBox4.Items.Add(dr_delete.Item(1) & " - [" & dr_delete.Item(2) & "]")

                ReDim Preserve a_labs_for_clinical_service(l_dimension_labs_cs)

                a_labs_for_clinical_service(l_dimension_labs_cs).id_content_analysis_sample_type = dr_delete.Item(0)
                a_labs_for_clinical_service(l_dimension_labs_cs).desc_analysis_sample_type = dr_delete.Item(1)
                a_labs_for_clinical_service(l_dimension_labs_cs).desc_analysis_sample_recipient = dr_delete.Item(2)
                a_labs_for_clinical_service(l_dimension_labs_cs).flg_new = "N"

                l_dimension_labs_cs = l_dimension_labs_cs + 1

                ReDim Preserve l_labs_selected_from_alert(l_index_selected_labs_from_alert)

                l_labs_selected_from_alert(l_index_selected_labs_from_alert).id_content_analysis_sample_type = dr_delete.Item(0)
                l_labs_selected_from_alert(l_index_selected_labs_from_alert).desc_analysis_sample_type = dr_delete.Item(1)
                l_labs_selected_from_alert(l_index_selected_labs_from_alert).desc_analysis_sample_recipient = dr_delete.Item(2)

                l_index_selected_labs_from_alert = l_index_selected_labs_from_alert + 1

            End While

        End If

        dr_delete.Dispose()
        dr_delete.Close()

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged

        Dim l_unsaved_records As Boolean = False
        Dim l_sucess As Boolean = True

        Dim l_first_time As Boolean = False 'Variavel para determinar se é a primeira vez que se está a colocar o Clinical Service

        '1 - DEterminar o dep_clin_serv_selecionado
        Dim l_id_dep_clin_serv_aux As Int64 = a_dep_clin_serv_inst(ComboBox6.SelectedIndex)

        '2 - Determinar se existem registos a serem guardados
        If (l_dimension_labs_cs > 0 And l_id_dep_clin_serv <> 0) Then

            For j As Int16 = 0 To a_labs_for_clinical_service.Count() - 1

                If a_labs_for_clinical_service(j).flg_new = "Y" Then

                    l_unsaved_records = True
                    Exit For

                End If

            Next

        End If

        If (l_id_dep_clin_serv = 0) Then

            l_id_dep_clin_serv = l_id_dep_clin_serv_aux
            l_first_time = True

        End If

        '3 Caso existam, gravar.
        If l_unsaved_records = True Then

            Dim result As Integer = 0

            result = MsgBox("There are unsaved records. Do you wish to save them?", vbYesNo)

            If (result = DialogResult.Yes) Then

                For j As Int16 = 0 To a_labs_for_clinical_service.Count() - 1

                    If (a_labs_for_clinical_service(j).flg_new = "Y") Then

                        If Not db_labs.SET_ANALYSIS_DEP_CLIN_SERV(l_selected_soft, l_id_dep_clin_serv, a_labs_for_clinical_service(j).id_content_analysis_sample_type, conn_labs) Then

                            l_sucess = False

                        End If

                    End If

                Next

                If l_sucess = False Then

                    MsgBox("ERROR INSERTING ANALYSIS AS FREQUENT - ComboBox6_SelectedIndexChanged", vbCritical)

                Else

                    MsgBox("Selected record(s) saved.", vbInformation)
                    CheckedListBox4.Items.Clear()

                End If

            End If

        End If

        If (l_first_time = False) Then

            '4 - Limpar a box e os arrays
            ReDim Preserve a_labs_for_clinical_service(0)
            l_dimension_labs_cs = 0

            ReDim l_labs_selected_from_alert(0)
            l_index_selected_labs_from_alert = 0

            CheckedListBox4.Items.Clear()

        End If

        '5 - Determinar os exames disponíveis como mais frequentes para esse dep_clin_serv
        Dim dr As OracleDataReader

        If Not db_labs.GET_ANALYSIS_DEP_CLIN_SERV(TextBox1.Text, l_selected_soft, l_id_dep_clin_serv_aux, dr, conn_labs) Then

            MsgBox("ERROR GETTING ANALYSIS_DEP_CLIN_SERV.", vbCritical)

        Else

            l_id_dep_clin_serv = l_id_dep_clin_serv_aux

            Dim i As Integer = 0

            '6 - Ler cursor e popular o campo
            While dr.Read()

                CheckedListBox4.Items.Add(dr.Item(1) & " - [" & dr.Item(2) & "]")

                ReDim Preserve a_labs_for_clinical_service(l_dimension_labs_cs)

                a_labs_for_clinical_service(l_dimension_labs_cs).id_content_analysis_sample_type = dr.Item(0)
                a_labs_for_clinical_service(l_dimension_labs_cs).desc_analysis_sample_type = dr.Item(1)
                a_labs_for_clinical_service(l_dimension_labs_cs).desc_analysis_sample_recipient = dr.Item(2)
                a_labs_for_clinical_service(l_dimension_labs_cs).flg_new = "N"

                l_dimension_labs_cs = l_dimension_labs_cs + 1

                ReDim Preserve l_labs_selected_from_alert(l_index_selected_labs_from_alert)

                l_labs_selected_from_alert(l_index_selected_labs_from_alert).id_content_analysis_sample_type = dr.Item(0)
                l_labs_selected_from_alert(l_index_selected_labs_from_alert).desc_analysis_sample_type = dr.Item(1)
                l_labs_selected_from_alert(l_index_selected_labs_from_alert).desc_analysis_sample_recipient = dr.Item(2)

                l_index_selected_labs_from_alert = l_index_selected_labs_from_alert + 1

            End While

        End If

        dr.Dispose()
        dr.Close()

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        If CheckedListBox4.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox4.Items.Count - 1

                CheckedListBox4.SetItemChecked(i, True)

            Next

        End If

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

        If CheckedListBox4.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox4.Items.Count - 1

                CheckedListBox4.SetItemChecked(i, False)

            Next

        End If

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        Cursor = Cursors.WaitCursor

        If CheckedListBox4.CheckedIndices.Count() > 0 Then

            Dim i As Integer = 0

            Dim indexChecked As Integer

            Dim total_selected_labs As Integer = 0

            For Each indexChecked In CheckedListBox4.CheckedIndices

                total_selected_labs = total_selected_labs + 1

            Next

            ReDim l_selected_labs(total_selected_labs - 1)

            For Each indexChecked In CheckedListBox4.CheckedIndices

                Dim dr As OracleDataReader

                If Not db_labs.GET_ANALYSIS_DEP_CLIN_SERV(TextBox1.Text, l_selected_soft, l_id_dep_clin_serv, dr, conn_labs) Then

                    MsgBox("ERROR GETTING ANALYSIS_DEP_CLIN_SERV.", vbCritical)

                Else

                    Dim i_index As Integer = 0

                    While dr.Read()

                        If i_index = indexChecked.ToString() Then

                            l_selected_labs(i) = dr.Item(0)

                        End If

                        i_index = i_index + 1

                    End While

                    i = i + 1

                End If

                dr.Dispose()
                dr.Close()

            Next

            Dim l_sucess As Boolean = True

            For ii As Integer = 0 To l_selected_labs.Count() - 1

                If Not db_labs.DELETE_ANALYSIS_DEP_CLIN_SERV(l_selected_soft, l_id_dep_clin_serv, l_selected_labs(ii), conn_labs) Then

                    l_sucess = False

                End If

            Next

            ReDim Preserve a_labs_for_clinical_service(0)
                l_dimension_labs_cs = 0

                ReDim l_labs_selected_from_alert(0)
                l_index_selected_labs_from_alert = 0

                CheckedListBox4.Items.Clear()

                Dim dr_new As OracleDataReader

                If db_labs.GET_ANALYSIS_DEP_CLIN_SERV(TextBox1.Text, l_selected_soft, l_id_dep_clin_serv, dr_new, conn_labs) Then

                    Dim i_new As Integer = 0

                    While dr_new.Read()

                        CheckedListBox4.Items.Add(dr_new.Item(1) & " - [" & dr_new.Item(2) & "]")

                        'Bloco para repopular os arrays
                        ReDim Preserve a_labs_for_clinical_service(l_dimension_labs_cs)
                        a_labs_for_clinical_service(l_dimension_labs_cs).id_content_analysis_sample_type = dr_new.Item(0)
                        a_labs_for_clinical_service(l_dimension_labs_cs).desc_analysis_sample_type = dr_new.Item(1)
                        a_labs_for_clinical_service(l_dimension_labs_cs).desc_analysis_sample_recipient = dr_new.Item(2)
                        a_labs_for_clinical_service(l_dimension_labs_cs).flg_new = "N"

                        l_dimension_labs_cs = l_dimension_labs_cs + 1

                        ReDim Preserve l_labs_selected_from_alert(l_index_selected_labs_from_alert)

                        l_labs_selected_from_alert(l_index_selected_labs_from_alert).id_content_analysis_sample_type = dr_new.Item(0)
                        l_labs_selected_from_alert(l_index_selected_labs_from_alert).desc_analysis_sample_type = dr_new.Item(1)
                        l_labs_selected_from_alert(l_index_selected_labs_from_alert).desc_analysis_sample_recipient = dr_new.Item(2)

                        l_index_selected_labs_from_alert = l_index_selected_labs_from_alert + 1
                        'Fim bloco

                    End While

                Else

                    MsgBox("ERROR!")

                End If


            If l_sucess = True Then

                MsgBox("Record(s) Deleted", vbInformation)

            Else

                MsgBox("ERROR DELETING LABORATORIAl EXAMS", vbCritical)

            End If

            Else

                MsgBox("No selected laboratorial exams!", vbCritical)

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Cursor = Cursors.WaitCursor

        If ComboBox6.SelectedItem = "" Then

            MsgBox("No clincial Service selected", vbCritical)

        Else

            Dim l_id_dep_clin_serv As Int64 = a_dep_clin_serv_inst(ComboBox6.SelectedIndex)

            Dim l_sucess As Boolean = True

            If CheckedListBox4.Items.Count() > 0 Then

                For Each indexChecked In CheckedListBox4.CheckedIndices

                    If (a_labs_for_clinical_service(indexChecked).flg_new = "Y") Then

                        If Not db_labs.SET_ANALYSIS_DEP_CLIN_SERV(l_selected_soft, l_id_dep_clin_serv, a_labs_for_clinical_service(indexChecked).id_content_analysis_sample_type, conn_labs) Then

                            l_sucess = False

                        End If

                    End If

                Next

                If (l_sucess = True) Then

                    MsgBox("Selected record(s) saved.", vbInformation)

                    CheckedListBox4.Items.Clear()
                Else

                    MsgBox("ERROR SAVING EXAMS AS FAVORITE. Button8_Click", vbCritical)

                End If

                ReDim l_labs_selected_from_alert(0)
                l_index_selected_labs_from_alert = 0

                ReDim a_labs_for_clinical_service(0)
                l_dimension_labs_cs = 0

                For ii As Integer = 0 To CheckedListBox3.Items.Count - 1

                    CheckedListBox3.SetItemChecked(ii, False)

                Next

            Else

                MsgBox("No records selected!", vbInformation)

            End If

        End If

        Dim dr As OracleDataReader

        If Not db_labs.GET_ANALYSIS_DEP_CLIN_SERV(TextBox1.Text, l_selected_soft, l_id_dep_clin_serv, dr, conn_labs) Then

            MsgBox("ERROR GETTING ANALYSIS_DEP_CLIN_SERV", vbCritical)

        Else

            Dim i As Integer = 0

            While dr.Read()

                CheckedListBox4.Items.Add(dr.Item(1) & " - [" & dr.Item(2) & "]")

                ReDim Preserve a_labs_for_clinical_service(l_dimension_labs_cs)
                a_labs_for_clinical_service(l_dimension_labs_cs).id_content_analysis_sample_type = dr.Item(0)
                a_labs_for_clinical_service(l_dimension_labs_cs).desc_analysis_sample_recipient = dr.Item(1)
                a_labs_for_clinical_service(l_dimension_labs_cs).desc_analysis_sample_recipient = dr.Item(2)
                a_labs_for_clinical_service(l_dimension_labs_cs).flg_new = "N"

                l_dimension_labs_cs = l_dimension_labs_cs + 1

                ReDim Preserve l_labs_selected_from_alert(l_index_selected_labs_from_alert)

                l_labs_selected_from_alert(l_index_selected_labs_from_alert).id_content_analysis_sample_type = dr.Item(0)
                l_labs_selected_from_alert(l_index_selected_labs_from_alert).desc_analysis_sample_type = dr.Item(1)
                l_labs_selected_from_alert(l_index_selected_labs_from_alert).desc_analysis_sample_recipient = dr.Item(2)

                l_index_selected_labs_from_alert = l_index_selected_labs_from_alert + 1

            End While

        End If

        Cursor = Cursors.Arrow

    End Sub
End Class
Imports Oracle.DataAccess.Client

Public Class Supplies

    Dim groupbox As myGroupBox

    Dim db_access_general As New General
    Dim db_supplies As New SUPPLIES_API
    Dim g_selected_soft As Int16 = -1

    Dim oradb As String
    Dim conn As New OracleConnection

    'Array que vai guardar as SUPPLY_AREAS disponíveis

    Dim g_a_SUP_AREAS() As SUPPLIES_API.SUP_AREAS

    'Bloco para definir os tipos de supplies
    Dim g_activity_desc As String = "Activity Theraphist Supplies"
    Dim g_activity_flag As String = "M"

    Dim g_implants_desc As String = "Implants"
    Dim g_implants_flag As String = "P"

    Dim g_kits_desc As String = "Kits"
    Dim g_kits_flag As String = "K"

    Dim g_sets_desc As String = "Sets"
    Dim g_sets_flag As String = "S"

    Dim g_supplies_desc As String = "Supplies"
    Dim g_supplies_flag As String = "I"

    Dim g_surgical_desc As String = "Surgical Equipments"
    Dim g_surgical_flag As String = "E"

    Dim g_selected_category As String = ""
    Dim g_type_supply_alert As String = ""

    Dim g_a_loaded_categories_default() As String ' Array que vai guardar os id_contents das categorias carregadas do default
    Dim g_selected_supplycategory As String = ""

    'Array que vai guardar os id_contents dos supplies caregados do default
    Dim g_a_loaded_supplies_default() As SUPPLIES_API.supplies_default

    'Array que vai guardar os id_contents dos supplies escolhidos do default
    Dim g_a_selected_supplies_default() As SUPPLIES_API.supplies_default

    Dim g_index_selected_supplies_from_default As Integer = 0 ''Variavel utilizada no botão de adicionar à box da direita (CHECKBOX 1)

    'Array que vai guardar as categorias disponíveis no ALERT
    Dim g_a_supp_cats_alert() As String

    'Array que vai guardar os supplies carregadas do ALERT
    Dim g_a_supps_alert() As SUPPLIES_API.supplies_default
    Dim g_dimension_supp_alert As Int64 = 0

    Dim g_selected_supplycategory_alert As String = ""


    Public Sub New(ByVal i_oradb As String)

        InitializeComponent()
        oradb = i_oradb
        conn = New OracleConnection(oradb)

    End Sub

    Private Sub Supplies_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.BackColor = Color.FromArgb(215, 215, 180)
        CheckedListBox2.BackColor = Color.FromArgb(195, 195, 165)
        CheckedListBox1.BackColor = Color.FromArgb(195, 195, 165)
        CheckedListBox3.BackColor = Color.FromArgb(195, 195, 165)

        ' GroupBox1.bo

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

        ''Popular SUP_AREAS
        Dim dr_sup_areas As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_supplies.GET_SUP_AREAS(conn, dr_sup_areas) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING SUPPLY AREAS!", vbCritical)

        End If

        Dim l_index_sup_area As Integer = 0
        ReDim g_a_SUP_AREAS(0)

        While dr_sup_areas.Read()

            ComboBox8.Items.Add(dr_sup_areas.Item(1))
            ComboBox6.Items.Add(dr_sup_areas.Item(1))

            ReDim Preserve g_a_SUP_AREAS(l_index_sup_area)
            g_a_SUP_AREAS(l_index_sup_area).id_supply_area = dr_sup_areas.Item(0)
            g_a_SUP_AREAS(l_index_sup_area).desc_supply_area = dr_sup_areas.Item(1)
            l_index_sup_area = l_index_sup_area + 1

        End While

        dr_sup_areas.Dispose()
        dr_sup_areas.Close()

        ''Popular tipos de supplies
        ComboBox7.Items.Add("ALL")
        ComboBox7.Items.Add(g_activity_desc)
        ComboBox7.Items.Add(g_implants_desc)
        ComboBox7.Items.Add(g_kits_desc)
        ComboBox7.Items.Add(g_sets_desc)
        ComboBox7.Items.Add(g_supplies_desc)
        ComboBox7.Items.Add(g_surgical_desc)

        ComboBox9.Items.Add("ALL")
        ComboBox9.Items.Add(g_activity_desc)
        ComboBox9.Items.Add(g_implants_desc)
        ComboBox9.Items.Add(g_kits_desc)
        ComboBox9.Items.Add(g_sets_desc)
        ComboBox9.Items.Add(g_supplies_desc)
        ComboBox9.Items.Add(g_surgical_desc)

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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Cursor = Cursors.WaitCursor

        'Limpar arrays
        g_selected_soft = -1
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

        'Limpar Versão
        ComboBox3.Text = ""
        ComboBox3.Items.Clear()

        'Limpar Seleção da box de Supply Area
        ComboBox8.SelectedIndex = -1

        'Limpar Array de categorias
        ReDim g_a_loaded_categories_default(0)

        ComboBox4.Items.Clear()
        ComboBox4.Text = ""

        'Limpar categoria selecionada
        g_selected_category = ""
        ComboBox7.SelectedIndex = -1

        g_selected_supplycategory = ""

        'Limpar Box de Supplies
        CheckedListBox2.Items.Clear()

        'Bloco do ALERT
        ComboBox5.Items.Clear()
        ComboBox5.Text = ""
        ReDim g_a_supp_cats_alert(0)
        ComboBox9.SelectedIndex = -1
        ComboBox9.Text = ""
        ComboBox6.SelectedIndex = -1
        ComboBox6.Text = ""
        ReDim g_a_supps_alert(0)
        g_dimension_supp_alert = 0
        CheckedListBox3.Items.Clear()

        If TextBox1.Text <> "" Then

            ComboBox1.Text = db_access_general.GET_INSTITUTION(TextBox1.Text, conn)

            ComboBox2.Items.Clear()
            ComboBox2.Text = ""

            Dim dr As OracleDataReader

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

                'g_selected_category = ""

            End If

            dr.Dispose()
            dr.Close()

        End If

        ReDim g_a_loaded_categories_default(0)
        ReDim g_a_loaded_supplies_default(0)
        ReDim g_a_selected_supplies_default(0)
        g_index_selected_supplies_from_default = 0

        Cursor = Cursors.Arrow
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged

        If ComboBox8.SelectedIndex >= 0 Then

            If ComboBox7.Text = g_activity_desc Then

                g_selected_category = g_activity_flag

            ElseIf ComboBox7.Text = g_implants_desc Then

                g_selected_category = g_implants_flag

            ElseIf ComboBox7.Text = g_kits_desc Then

                g_selected_category = g_kits_flag

            ElseIf ComboBox7.Text = g_sets_desc Then

                g_selected_category = g_sets_flag

            ElseIf ComboBox7.Text = g_supplies_desc Then

                g_selected_category = g_supplies_flag

            ElseIf ComboBox7.Text = g_surgical_desc Then

                g_selected_category = g_surgical_flag

            Else

                g_selected_category = "ALL"

            End If

            Cursor = Cursors.WaitCursor

            'Limpar arrays
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

            CheckedListBox1.Items.Clear()
            CheckedListBox2.Items.Clear()
            CheckedListBox3.Items.Clear()

            ComboBox3.Items.Clear()
            ComboBox3.Text = ""
            ComboBox4.Items.Clear()
            ComboBox4.Text = ""
            ComboBox5.Items.Clear()
            ComboBox5.Text = ""

            g_selected_supplycategory = ""
            ReDim g_a_loaded_categories_default(0)
            ReDim g_a_loaded_supplies_default(0)
            ReDim g_a_selected_supplies_default(0)
            g_index_selected_supplies_from_default = 0

            '1 - Fill Version combobox

            Dim dr_def_versions As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not db_supplies.GET_DEFAULT_VERSIONS(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox8.SelectedIndex).id_supply_area, g_selected_category, conn, dr_def_versions) Then

#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                MsgBox("ERROR LOADING DEFAULT VERSIONS -  ComboBox2_SelectedIndexChanged", MsgBoxStyle.Critical)

            Else

                While dr_def_versions.Read()

                    ComboBox3.Items.Add(dr_def_versions.Item(0))

                End While

            End If

            dr_def_versions.Dispose()
            dr_def_versions.Close()

            '        'Box de categorias na instituição/software
            '        Dim dr_exam_cat As OracleDataReader

            '#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            '        If Not db_intervention.GET_INTERV_CATS_INST_SOFT(TextBox1.Text, g_selected_soft, g_procedure_type, conn, dr_exam_cat) Then
            '#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            '            MsgBox("ERROR LOADING INTERVENTION CATEGORIES FROM INSTITUTION!", vbCritical)

            '        Else

            '            ComboBox5.Items.Add("ALL")

            '            ReDim g_a_interv_cats_alert(0)
            '            g_a_interv_cats_alert(0) = 0

            '            Dim l_index As Int16 = 1

            '            While dr_exam_cat.Read()

            '                ComboBox5.Items.Add(dr_exam_cat.Item(1))
            '                ReDim Preserve g_a_interv_cats_alert(l_index)
            '                g_a_interv_cats_alert(l_index) = dr_exam_cat.Item(0)
            '                l_index = l_index + 1

            '            End While

            '        End If

            '        dr_exam_cat.Dispose()
            '        dr_exam_cat.Close()

            '        'Preencher os Clinical Services

            '        Dim dr_clin_serv As OracleDataReader

            '#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            '        If Not db_access_general.GET_CLIN_SERV(TextBox1.Text, g_selected_soft, conn, dr_clin_serv) Then
            '#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            '            MsgBox("ERROR GETTING CLINICAL SERVICES!")

            '        Else

            '            Dim i As Integer = 0

            '            Dim l_index_dep_clin_serv As Integer = 0
            '            ReDim g_a_dep_clin_serv_inst(l_index_dep_clin_serv)

            '            While dr_clin_serv.Read()

            '                ComboBox6.Items.Add(dr_clin_serv.Item(0))

            '                ReDim Preserve g_a_dep_clin_serv_inst(l_index_dep_clin_serv)
            '                g_a_dep_clin_serv_inst(l_index_dep_clin_serv) = dr_clin_serv.Item(1)
            '                l_index_dep_clin_serv = l_index_dep_clin_serv + 1

            '            End While

            '        End If

            '        dr_clin_serv.Dispose()
            '        dr_clin_serv.Close()

            Cursor = Cursors.Arrow

        End If

    End Sub


    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        If ComboBox2.Text <> "" And ComboBox8.Text <> "" And ComboBox7.Text <> "" Then

            Cursor = Cursors.WaitCursor

            'Limpar arrays
            'ReDim g_a_loaded_interventions_default(0)
            'ReDim g_a_selected_default_interventions(0)
            ' g_index_selected_intervention_from_default = 0

            ComboBox4.Items.Clear()
            ComboBox4.Text = ""

            CheckedListBox2.Items.Clear()

            'Determinar as categorias disponíveis para a versão escolhida
            'Array g_a_loaded_categories_default vai gaurdar os ids de todas as categorias

            ReDim g_a_loaded_categories_default(0)
            Dim l_index_loaded_categories As Int64 = 0

            Dim dr_lab_cat_def As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not db_supplies.GET_SUPP_CATS_DEFAULT(TextBox1.Text, g_selected_soft, ComboBox3.Text, g_a_SUP_AREAS(ComboBox8.SelectedIndex).id_supply_area, g_selected_category, conn, dr_lab_cat_def) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                MsgBox("ERROR LOADING DEFAULT SUPPLIES CATEGORIS -  ComboBox3_SelectedIndexChanged", MsgBoxStyle.Critical)

            Else

                ComboBox4.Items.Add("ALL")

                While dr_lab_cat_def.Read()

                    ComboBox4.Items.Add(dr_lab_cat_def.Item(1) & "   -  [" & dr_lab_cat_def.Item(0) & "]")
                    g_a_loaded_categories_default(l_index_loaded_categories) = dr_lab_cat_def.Item(0)
                    l_index_loaded_categories = l_index_loaded_categories + 1
                    ReDim Preserve g_a_loaded_categories_default(l_index_loaded_categories)

                End While

            End If

            dr_lab_cat_def.Dispose()
            dr_lab_cat_def.Close()

            CheckedListBox1.Items.Clear()

            Cursor = Cursors.Arrow

        End If

    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged

        ComboBox7.SelectedIndex = -1
        ComboBox7.Text = ""

        CheckedListBox1.Items.Clear()
        CheckedListBox2.Items.Clear()
        CheckedListBox3.Items.Clear()

        ComboBox3.Items.Clear()
        ComboBox3.Text = ""
        ComboBox4.Items.Clear()
        ComboBox4.Text = ""
        ComboBox5.Items.Clear()
        ComboBox5.Text = ""

        g_selected_category = ""
        g_selected_supplycategory = ""

        ReDim g_a_loaded_categories_default(0)
        ReDim g_a_loaded_supplies_default(0)
        ReDim g_a_selected_supplies_default(0)
        g_index_selected_supplies_from_default = 0

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

        '1 - Determinar o id_content da categoria selecionada
        If ComboBox4.SelectedIndex = 0 Then
            g_selected_supplycategory = 0
        Else
            g_selected_supplycategory = g_a_loaded_categories_default(ComboBox4.SelectedIndex - 1)
        End If

        Cursor = Cursors.WaitCursor
        CheckedListBox2.Items.Clear()
        ''2 - Carregar a grelha de análises por categoria
        ''e    
        ''3 - Criar estrutura com os elementos das análises carregados
        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_supplies.GET_SUPS_DEFAULT_BY_CAT(TextBox1.Text, g_selected_soft, ComboBox3.Text, g_a_SUP_AREAS(ComboBox8.SelectedIndex).id_supply_area, g_selected_category, g_selected_supplycategory, conn, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING SUPPLIES BY CATEGORY >> ComboBox4_SelectedIndexChanged")

        Else
            ReDim g_a_loaded_supplies_default(0) ''Limpar estrutura
            Dim l_dimension_array_loaded_supplies As Int64 = 0

            While dr.Read()

                'Colocar a categoria quando se seleciona o ALL

                If g_selected_supplycategory = "0" Then

                    CheckedListBox2.Items.Add(dr.Item(3) & "  -  [" & dr.Item(2) & "]  >> " & dr.Item(1))

                Else

                    CheckedListBox2.Items.Add(dr.Item(3) & "  -  [" & dr.Item(2) & "]")

                End If


                ReDim Preserve g_a_loaded_supplies_default(l_dimension_array_loaded_supplies)

                g_a_loaded_supplies_default(l_dimension_array_loaded_supplies).id_content_category = dr.Item(0)
                g_a_loaded_supplies_default(l_dimension_array_loaded_supplies).desc_category = dr.Item(1)
                g_a_loaded_supplies_default(l_dimension_array_loaded_supplies).id_content_supply = dr.Item(2)
                g_a_loaded_supplies_default(l_dimension_array_loaded_supplies).desc_supply = dr.Item(3)

                l_dimension_array_loaded_supplies = l_dimension_array_loaded_supplies + 1

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

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        'Limpar Versão
        ComboBox3.Text = ""
        ComboBox3.Items.Clear()

        'Limpar Seleção da box de Supply Area
        ComboBox8.SelectedIndex = -1

        'Limpar Array de categorias
        ReDim g_a_loaded_categories_default(0)

        ComboBox4.Items.Clear()
        ComboBox4.Text = ""

        'Limpar categoria selecionada
        g_selected_category = ""
        ComboBox7.SelectedIndex = -1

        g_selected_supplycategory = ""

        'Limpar Box de Supplies
        CheckedListBox2.Items.Clear()

        'Bloco do ALERT
        ComboBox5.Items.Clear()
        ComboBox5.Text = ""
        ReDim g_a_supp_cats_alert(0)
        ComboBox9.SelectedIndex = -1
        ComboBox9.Text = ""
        ComboBox6.SelectedIndex = -1
        ComboBox6.Text = ""
        ReDim g_a_supps_alert(0)
        g_dimension_supp_alert = 0

        g_selected_soft = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, TextBox1.Text, conn)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Cursor = Cursors.WaitCursor
        For Each indexChecked In CheckedListBox2.CheckedIndices
            'If para verificar se já está incluido na checkbox da direita

            Dim l_record_already_selected As Boolean = False

            Dim j As Integer = 0

            For j = 0 To CheckedListBox1.Items.Count() - 1

                If (g_a_loaded_supplies_default(indexChecked.ToString()).id_content_category = g_a_selected_supplies_default(j).id_content_category And g_a_loaded_supplies_default(indexChecked.ToString()).id_content_supply = g_a_selected_supplies_default(j).id_content_supply) Then

                    l_record_already_selected = True
                    Exit For

                End If

            Next

            If l_record_already_selected = False Then

                ReDim Preserve g_a_selected_supplies_default(g_index_selected_supplies_from_default)

                g_a_selected_supplies_default(g_index_selected_supplies_from_default).id_content_category = g_a_loaded_supplies_default(indexChecked.ToString()).id_content_category
                g_a_selected_supplies_default(g_index_selected_supplies_from_default).id_content_supply = g_a_loaded_supplies_default(indexChecked.ToString()).id_content_supply
                g_a_selected_supplies_default(g_index_selected_supplies_from_default).desc_category = g_a_loaded_supplies_default(indexChecked.ToString()).desc_category
                g_a_selected_supplies_default(g_index_selected_supplies_from_default).desc_supply = g_a_loaded_supplies_default(indexChecked.ToString()).desc_supply

                CheckedListBox1.Items.Add((g_a_selected_supplies_default(g_index_selected_supplies_from_default).desc_supply) & " - [" & g_a_selected_supplies_default(g_index_selected_supplies_from_default).desc_category & "]")

                CheckedListBox1.SetItemChecked((CheckedListBox1.Items.Count() - 1), True)

                g_index_selected_supplies_from_default = g_index_selected_supplies_from_default + 1

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

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        g_selected_soft = -1

        ComboBox8.SelectedIndex = -1
        ComboBox8.Text = ""

        ComboBox7.SelectedIndex = -1
        ComboBox7.Text = ""

        CheckedListBox1.Items.Clear()
        CheckedListBox2.Items.Clear()
        CheckedListBox3.Items.Clear()

        ComboBox3.Items.Clear()
        ComboBox3.Text = ""
        ComboBox4.Items.Clear()
        ComboBox4.Text = ""
        ComboBox5.Items.Clear()
        ComboBox5.Text = ""

        g_selected_category = ""
        g_selected_supplycategory = ""

        ReDim g_a_loaded_categories_default(0)
        ReDim g_a_loaded_supplies_default(0)
        ReDim g_a_selected_supplies_default(0)
        g_index_selected_supplies_from_default = 0

        TextBox1.Text = db_access_general.GET_INSTITUTION_ID(ComboBox1.SelectedIndex, conn)

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""

        'Bloco do ALERT
        ComboBox5.Items.Clear()
        ComboBox5.Text = ""
        ReDim g_a_supp_cats_alert(0)
        ComboBox9.SelectedIndex = -1
        ComboBox9.Text = ""
        ComboBox6.SelectedIndex = -1
        ComboBox6.Text = ""
        ReDim g_a_supps_alert(0)
        g_dimension_supp_alert = 0

        Dim dr As OracleDataReader

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_SOFT_INST(TextBox1.Text, conn, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING SOFTWARES!", vbCritical)

        Else

            While dr.Read()

                ComboBox2.Items.Add(dr.Item(1))

            End While

        End If

        dr.Dispose()
        dr.Close()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Cursor = Cursors.WaitCursor

        If CheckedListBox1.Items.Count() > 0 Then

            Dim l_a_checked_supplies() As SUPPLIES_API.supplies_default
            Dim l_index As Integer = 0

            For Each indexChecked In CheckedListBox1.CheckedIndices

                ReDim Preserve l_a_checked_supplies(l_index)

                l_a_checked_supplies(l_index).id_content_category = g_a_selected_supplies_default(indexChecked).id_content_category
                l_a_checked_supplies(l_index).id_content_supply = g_a_selected_supplies_default(indexChecked).id_content_supply
                l_a_checked_supplies(l_index).desc_category = g_a_selected_supplies_default(indexChecked).desc_category
                l_a_checked_supplies(l_index).desc_supply = g_a_selected_supplies_default(indexChecked).desc_supply

                l_index = l_index + 1

            Next

            If Not db_supplies.SET_SUPPLY_TYPE(TextBox1.Text, l_a_checked_supplies, conn) Then

                MsgBox("ERROR INSERTING SUPPLIES CATEGORIES!", vbCritical)

            ElseIf Not db_supplies.SET_SUPPLY(TextBox1.Text, l_a_checked_supplies, conn) Then

                MsgBox("ERROR INSERTING SUPPLIES!", vbCritical)

            ElseIf Not db_supplies.SET_SUPPLY_SOFT_INST(TextBox1.Text, g_selected_soft, l_a_checked_supplies, conn) Then

                MsgBox("ERROR INSERTING SUPPLY_SOFT_INST!", vbCritical)

            ElseIf Not db_supplies.SET_SUPPLY_SUP_AREA(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox8.SelectedIndex).id_supply_area, l_a_checked_supplies, conn) Then

                MsgBox("ERROR INSERTING SUPPLY_SOFT_INST!", vbCritical)

            ElseIf Not db_supplies.SET_SUPPLY_LOC_DEFAULT(TextBox1.Text, g_selected_soft, l_a_checked_supplies, conn) Then

                MsgBox("ERROR INSERTING SUPPLY_LOC_DEFAULT!", vbCritical)

            Else

                MsgBox("Record(s) successfully inserted.", vbInformation)

                '1 - Processo Limpeza
                '1.1 - Limpar a box de materiais a gravar no alert
                CheckedListBox1.Items.Clear()

                '1.2 - Remover o check dos materiais do default
                For i As Integer = 0 To CheckedListBox2.Items.Count - 1

                    CheckedListBox2.SetItemChecked(i, False)

                Next

                '1.3 - Limpar g_a_selected_supplies_default (Array materiais do default selecionadas pelo utilizador)
                ReDim g_a_selected_supplies_default(0)
                g_index_selected_supplies_from_default = 0

                '1.4 - Limpar a caixa de categorias de materiais do ALERT
                ComboBox5.Items.Clear()
                ComboBox5.SelectedItem = ""

                'Bloco do ALERT
                ComboBox5.Items.Clear()
                ComboBox5.Text = ""
                ReDim g_a_supp_cats_alert(0)
                ComboBox9.SelectedIndex = -1
                ComboBox9.Text = ""
                ComboBox6.SelectedIndex = -1
                ComboBox6.Text = ""
                ReDim g_a_supps_alert(0)
                g_dimension_supp_alert = 0
                CheckedListBox3.Items.Clear()

            End If

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        'Determinar o id_content da categoria selecionada
        If ComboBox5.SelectedIndex = 0 Then
            g_selected_supplycategory_alert = 0
        Else
            g_selected_supplycategory_alert = g_a_supp_cats_alert(ComboBox5.SelectedIndex)
        End If

        CheckedListBox3.Items.Clear()

        Dim dr_supplies As OracleDataReader

        g_dimension_supp_alert = 0
        ReDim g_a_supps_alert(g_dimension_supp_alert)


#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_supplies.GET_SUPS_ALERT_BY_CAT(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox6.SelectedIndex).id_supply_area, g_type_supply_alert, g_selected_supplycategory_alert, conn, dr_supplies) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING INTERVENTIONS FROM INSTITUTION!", MsgBoxStyle.Critical)

        Else

            While dr_supplies.Read()

                g_a_supps_alert(g_dimension_supp_alert).id_content_category = dr_supplies.Item(0)
                g_a_supps_alert(g_dimension_supp_alert).desc_category = dr_supplies.Item(1)

                g_a_supps_alert(g_dimension_supp_alert).id_content_supply = dr_supplies.Item(2)
                g_a_supps_alert(g_dimension_supp_alert).desc_supply = dr_supplies.Item(3)

                g_dimension_supp_alert = g_dimension_supp_alert + 1
                ReDim Preserve g_a_supps_alert(g_dimension_supp_alert)

                If g_selected_supplycategory_alert = "0" Then

                    CheckedListBox3.Items.Add(dr_supplies.Item(3) & "  >>  [" & dr_supplies.Item(1) & "]")

                Else

                    CheckedListBox3.Items.Add((dr_supplies.Item(3)))

                End If

            End While

        End If

        dr_supplies.Dispose()
        dr_supplies.Close()

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged

        'Bloco Para Limpar Variaveis

        ComboBox5.Items.Clear()
        ComboBox5.Text = ""
        ReDim g_a_supp_cats_alert(0)

        'Fim BLoco para Limpar Variaveis

        If ComboBox9.Text = g_activity_desc Then

            g_type_supply_alert = g_activity_flag

        ElseIf ComboBox9.Text = g_implants_desc Then

            g_type_supply_alert = g_implants_flag

        ElseIf ComboBox9.Text = g_kits_desc Then

            g_type_supply_alert = g_kits_flag

        ElseIf ComboBox9.Text = g_sets_desc Then

            g_type_supply_alert = g_sets_flag

        ElseIf ComboBox9.Text = g_supplies_desc Then

            g_type_supply_alert = g_supplies_flag

        ElseIf ComboBox9.Text = g_surgical_desc Then

            g_type_supply_alert = g_surgical_flag

        Else

            g_type_supply_alert = "ALL"

        End If

        'Box de categorias na instituição/software
        Dim dr_supp_cat As OracleDataReader

        ReDim g_a_supp_cats_alert(0)
        Dim l_index_loaded_categories As Int64 = 0

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_supplies.GET_SUPP_CATS_ALERT(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox6.SelectedIndex).id_supply_area, g_type_supply_alert, conn, dr_supp_cat) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR LOADING SUPPLY CATEGORIES FROM INSTITUTION!", vbCritical)

        Else

            ComboBox5.Items.Add("ALL")

            ReDim g_a_supp_cats_alert(0)
            g_a_supp_cats_alert(0) = 0

            Dim l_index As Int16 = 1

            While dr_supp_cat.Read()

                ComboBox5.Items.Add(dr_supp_cat.Item(1))
                ReDim Preserve g_a_supp_cats_alert(l_index)
                g_a_supp_cats_alert(l_index) = dr_supp_cat.Item(0)
                l_index = l_index + 1

            End While

        End If

        dr_supp_cat.Dispose()
        dr_supp_cat.Close()

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged

        'Bloco Para Limpar Variaveis

        ComboBox5.Items.Clear()
        ComboBox5.Text = ""
        ReDim g_a_supp_cats_alert(0)

        ComboBox9.SelectedIndex = -1
        ComboBox9.Text = ""

        'Fim BLoco para Limpar Variaveis

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

        If CheckedListBox3.CheckedIndices.Count > 0 Then

            Cursor = Cursors.WaitCursor

            Dim result As Integer = 0
            Dim l_sucess As Boolean = False

            'Perguntar se utilizador pretende mesmo apagar todos os supplies de uma categoria
            If (CheckedListBox3.CheckedIndices.Count = CheckedListBox3.Items.Count()) Then

                result = MsgBox("All records from the chosen category will be deleted! Confirm?", MessageBoxButtons.YesNo)

            End If

            If (result = DialogResult.Yes Or CheckedListBox3.CheckedIndices.Count < CheckedListBox3.Items.Count()) Then

                Dim indexChecked As Integer

                '1 - Apagar Registos
                'Ciclo para correr todos os registos do ALERT marcados com o check


                Dim l_selected_supplies_alert() As SUPPLIES_API.supplies_default
                Dim l_dimension_selected_supplies As Integer = 0

                For Each indexChecked In CheckedListBox3.CheckedIndices

                    ReDim Preserve l_selected_supplies_alert(l_dimension_selected_supplies)
                    l_selected_supplies_alert(l_dimension_selected_supplies).id_content_category = g_a_supps_alert(indexChecked).id_content_category
                    l_selected_supplies_alert(l_dimension_selected_supplies).id_content_supply = g_a_supps_alert(indexChecked).id_content_supply
                    l_dimension_selected_supplies = l_dimension_selected_supplies + 1

                Next

                If Not db_supplies.DELETE_SUPPLY_SOFT_INST(TextBox1.Text, g_selected_soft, l_selected_supplies_alert, g_a_SUP_AREAS(ComboBox6.SelectedIndex).id_supply_area, conn) Then

                    MsgBox("ERROR DELETING SUPPLIES FROM ALERT!", vbCritical)

                Else

                    l_sucess = True

                End If

                ''2 - Refresh à grelha de registos do Alert
                ''3.1 - Se estão a ser apagados todos os registos de uma categoria:
                If CheckedListBox3.CheckedIndices.Count = CheckedListBox3.Items.Count() Then

                    CheckedListBox3.Items.Clear()
                    ComboBox5.Items.Clear()
                    ComboBox5.Text = ""

                    '3.1.1 - Obter novamente as categorias dos supplies
                    Dim dr_supp_cat As OracleDataReader

                    ReDim g_a_supp_cats_alert(0)
                    Dim l_index_loaded_categories As Int64 = 0

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                    If Not db_supplies.GET_SUPP_CATS_ALERT(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox6.SelectedIndex).id_supply_area, g_type_supply_alert, conn, dr_supp_cat) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                        MsgBox("ERROR LOADING SUPPLY CATEGORIES FROM INSTITUTION!", vbCritical)

                    Else

                        ComboBox5.Items.Add("ALL")

                        ReDim g_a_supp_cats_alert(0)
                        g_a_supp_cats_alert(0) = 0

                        Dim l_index As Int16 = 1

                        While dr_supp_cat.Read()

                            ComboBox5.Items.Add(dr_supp_cat.Item(1))
                            ReDim Preserve g_a_supp_cats_alert(l_index)
                            g_a_supp_cats_alert(l_index) = dr_supp_cat.Item(0)
                            l_index = l_index + 1

                        End While

                    End If

                    dr_supp_cat.Dispose()
                    dr_supp_cat.Close()

                    'Limpar arrays
                    ReDim g_a_supps_alert(0)
                    g_dimension_supp_alert = 0

                Else '3.2 - Eliminar apenas os registos selecionados

                    CheckedListBox3.Items.Clear()
                    Dim dr_supplies As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                    If Not db_supplies.GET_SUPS_ALERT_BY_CAT(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox6.SelectedIndex).id_supply_area, g_type_supply_alert, g_selected_supplycategory_alert, conn, dr_supplies) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                        MsgBox("ERROR GETTING INTERVENTIONS FROM INSTITUTION!", MsgBoxStyle.Critical)

                    Else

                        While dr_supplies.Read()

                            g_a_supps_alert(g_dimension_supp_alert).id_content_category = dr_supplies.Item(0)
                            g_a_supps_alert(g_dimension_supp_alert).desc_category = dr_supplies.Item(1)

                            g_a_supps_alert(g_dimension_supp_alert).id_content_supply = dr_supplies.Item(2)
                            g_a_supps_alert(g_dimension_supp_alert).desc_supply = dr_supplies.Item(3)

                            g_dimension_supp_alert = g_dimension_supp_alert + 1
                            ReDim Preserve g_a_supps_alert(g_dimension_supp_alert)

                            If g_selected_supplycategory_alert = "0" Then

                                CheckedListBox3.Items.Add(dr_supplies.Item(3) & "  >>  [" & dr_supplies.Item(1) & "]")

                            Else

                                CheckedListBox3.Items.Add((dr_supplies.Item(3)))

                            End If

                        End While

                    End If

                    dr_supplies.Dispose()
                    dr_supplies.Close()

                End If

            End If


            If l_sucess = True Then

                MsgBox("Record(s) Successfuly deleted.", vbInformation)

            End If

            Cursor = Cursors.Arrow

        End If

    End Sub
End Class
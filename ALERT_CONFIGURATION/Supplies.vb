Imports Oracle.DataAccess.Client

Public Class Supplies

    Dim groupbox As myGroupBox

    Dim db_access_general As New General
    Dim db_supplies As New SUPPLIES_API
    Dim g_selected_soft As Int16 = -1

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
    Dim g_type_supply_alert_barcode As String = ""

    Dim g_a_loaded_categories_default() As String ' Array que vai guardar os id_contents das categorias carregadas do default
    Dim g_selected_supplycategory As String = ""

    'Array que vai guardar os id_contents dos supplies caregados do default
    Dim g_a_loaded_supplies_default() As SUPPLIES_API.supplies_default

    'Array que vai guardar os id_contents dos supplies escolhidos do default
    Dim g_a_selected_supplies_default() As SUPPLIES_API.supplies_default

    Dim g_index_selected_supplies_from_default As Integer = 0 ''Variavel utilizada no botão de adicionar à box da direita (CHECKBOX 1)

    'Array que vai guardar as categorias disponíveis no ALERT
    Dim g_a_supp_cats_alert() As String
    Dim g_a_supp_cats_alert_barcode() As String

    'Array que vai guardar os supplies carregadas do ALERT
    Dim g_a_supps_alert() As SUPPLIES_API.supplies_default
    Dim g_dimension_supp_alert As Int64 = 0

    Dim g_selected_supplycategory_alert As String = ""
    Dim g_selected_supplycategory_alert_barcode As String = ""

    Public Structure TABLE_BARCODE

        Public desc_supply_type As String
        Public desc_supply As String
        Public barcode As String
        Public lote As String
        Public serial As String
        Public id_supply As Int64
    End Structure

    Dim g_supply_barcode() As TABLE_BARCODE

    Private Sub Supplies_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "SUPPLIES  ::::  Connected to " & Connection.db

        Me.BackColor = Color.FromArgb(215, 215, 180)

        Me.Location = New Point(Form_location.x_position, Form_location.y_position)

        CheckedListBox2.BackColor = Color.FromArgb(195, 195, 165)
        CheckedListBox1.BackColor = Color.FromArgb(195, 195, 165)
        CheckedListBox3.BackColor = Color.FromArgb(195, 195, 165)

        DataGridView1.BackgroundColor = Color.FromArgb(195, 195, 165)

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

        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

        ''Popular SUP_AREAS
        Dim dr_sup_areas As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_supplies.GET_SUP_AREAS(dr_sup_areas) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING SUPPLY AREAS!", vbCritical)

        End If

        Dim l_index_sup_area As Integer = 0
        ReDim g_a_SUP_AREAS(0)

        While dr_sup_areas.Read()

            ComboBox8.Items.Add(dr_sup_areas.Item(1))
            ComboBox6.Items.Add(dr_sup_areas.Item(1))
            ComboBox10.Items.Add(dr_sup_areas.Item(1))

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

        ComboBox11.Items.Add("ALL")
        ComboBox11.Items.Add(g_activity_desc)
        ComboBox11.Items.Add(g_implants_desc)
        ComboBox11.Items.Add(g_kits_desc)
        ComboBox11.Items.Add(g_sets_desc)
        ComboBox11.Items.Add(g_supplies_desc)
        ComboBox11.Items.Add(g_surgical_desc)

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Me.Enabled = False
        Me.Dispose()
        Form1.Show()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Cursor = Cursors.WaitCursor

        'Limpar arrays
        g_selected_soft = -1

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
        ReDim g_a_supp_cats_alert_barcode(0)
        ComboBox9.SelectedIndex = -1
        ComboBox9.Text = ""
        ComboBox6.SelectedIndex = -1
        ComboBox6.Text = ""
        ReDim g_a_supps_alert(0)
        g_dimension_supp_alert = 0
        CheckedListBox3.Items.Clear()

        ComboBox10.SelectedIndex = -1
        ComboBox10.Text = ""

        ComboBox11.SelectedIndex = -1
        ComboBox11.Text = ""

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

                ComboBox3.Text = ""
                ComboBox3.Items.Clear()

                ComboBox4.Text = ""
                ComboBox4.Items.Clear()

                CheckedListBox2.Items.Clear()

                CheckedListBox1.Items.Clear()

                ComboBox5.Text = ""
                ComboBox5.Items.Clear()

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

        If CheckedListBox1.Items.Count() > 0 Then

            Dim result As Integer = 0

            result = MsgBox("There are unsaved records. Do you wish to save them?", vbYesNo)

            If (result = DialogResult.Yes) Then

                Cursor = Cursors.WaitCursor

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

                If Not db_supplies.SET_SUPPLY_TYPE(TextBox1.Text, l_a_checked_supplies) Then

                    MsgBox("ERROR INSERTING SUPPLIES CATEGORIES!", vbCritical)

                ElseIf Not db_supplies.SET_SUPPLY(TextBox1.Text, l_a_checked_supplies) Then

                    MsgBox("ERROR INSERTING SUPPLIES!", vbCritical)

                ElseIf Not db_supplies.SET_SUPPLY_SOFT_INST(TextBox1.Text, g_selected_soft, l_a_checked_supplies) Then

                    MsgBox("ERROR INSERTING SUPPLY_SOFT_INST!", vbCritical)

                ElseIf Not db_supplies.SET_SUPPLY_SUP_AREA(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox8.SelectedIndex).id_supply_area, l_a_checked_supplies) Then

                    MsgBox("ERROR INSERTING SUPPLY_SOFT_INST!", vbCritical)

                ElseIf Not db_supplies.SET_SUPPLY_LOC_DEFAULT(TextBox1.Text, g_selected_soft, l_a_checked_supplies) Then

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
                    ReDim g_a_supp_cats_alert_barcode(0)
                    ComboBox9.SelectedIndex = -1
                    ComboBox9.Text = ""
                    ComboBox6.SelectedIndex = -1
                    ComboBox6.Text = ""
                    ReDim g_a_supps_alert(0)
                    g_dimension_supp_alert = 0
                    CheckedListBox3.Items.Clear()

                    ComboBox10.SelectedIndex = -1
                    ComboBox10.Text = ""
                    ComboBox11.SelectedIndex = -1
                    ComboBox11.Text = ""

                End If

            End If

            Cursor = Cursors.Arrow

        End If

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

            CheckedListBox2.Items.Clear()
            CheckedListBox1.Items.Clear()

            ComboBox3.Items.Clear()
            ComboBox3.Text = ""
            ComboBox4.Items.Clear()
            ComboBox4.Text = ""

            g_selected_supplycategory = ""
            ReDim g_a_loaded_categories_default(0)
            ReDim g_a_loaded_supplies_default(0)
            ReDim g_a_selected_supplies_default(0)
            g_index_selected_supplies_from_default = 0

            '1 - Fill Version combobox

            Dim dr_def_versions As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not db_supplies.GET_DEFAULT_VERSIONS(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox8.SelectedIndex).id_supply_area, g_selected_category, dr_def_versions) Then

#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                MsgBox("ERROR LOADING DEFAULT VERSIONS -  ComboBox2_SelectedIndexChanged", MsgBoxStyle.Critical)

            Else

                While dr_def_versions.Read()

                    ComboBox3.Items.Add(dr_def_versions.Item(0))

                End While

            End If

            dr_def_versions.Dispose()
            dr_def_versions.Close()

            Cursor = Cursors.Arrow

        End If

    End Sub


    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        If CheckedListBox1.Items.Count() > 0 Then

            Dim result As Integer = 0

            result = MsgBox("There are unsaved records. Do you wish to save them?", vbYesNo)

            If (result = DialogResult.Yes) Then

                Cursor = Cursors.WaitCursor

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

                If Not db_supplies.SET_SUPPLY_TYPE(TextBox1.Text, l_a_checked_supplies) Then

                    MsgBox("ERROR INSERTING SUPPLIES CATEGORIES!", vbCritical)

                ElseIf Not db_supplies.SET_SUPPLY(TextBox1.Text, l_a_checked_supplies) Then

                    MsgBox("ERROR INSERTING SUPPLIES!", vbCritical)

                ElseIf Not db_supplies.SET_SUPPLY_SOFT_INST(TextBox1.Text, g_selected_soft, l_a_checked_supplies) Then

                    MsgBox("ERROR INSERTING SUPPLY_SOFT_INST!", vbCritical)

                ElseIf Not db_supplies.SET_SUPPLY_SUP_AREA(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox8.SelectedIndex).id_supply_area, l_a_checked_supplies) Then

                    MsgBox("ERROR INSERTING SUPPLY_SOFT_INST!", vbCritical)

                ElseIf Not db_supplies.SET_SUPPLY_LOC_DEFAULT(TextBox1.Text, g_selected_soft, l_a_checked_supplies) Then

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
                    ReDim g_a_supp_cats_alert_barcode(0)
                    ComboBox9.SelectedIndex = -1
                    ComboBox9.Text = ""
                    ComboBox6.SelectedIndex = -1
                    ComboBox6.Text = ""
                    ReDim g_a_supps_alert(0)
                    g_dimension_supp_alert = 0
                    CheckedListBox3.Items.Clear()

                    ComboBox10.SelectedIndex = -1
                    ComboBox10.Text = ""
                    ComboBox11.SelectedIndex = -1
                    ComboBox11.Text = ""

                End If

            End If

            Cursor = Cursors.Arrow

        End If

        If ComboBox2.Text <> "" And ComboBox8.Text <> "" And ComboBox7.Text <> "" Then

            Cursor = Cursors.WaitCursor

            CheckedListBox1.Items.Clear()

            ComboBox4.Items.Clear()
            ComboBox4.Text = ""

            CheckedListBox2.Items.Clear()

            ReDim g_a_loaded_categories_default(0)
            Dim l_index_loaded_categories As Int64 = 0

            Dim dr_lab_cat_def As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not db_supplies.GET_SUPP_CATS_DEFAULT(TextBox1.Text, g_selected_soft, ComboBox3.Text, g_a_SUP_AREAS(ComboBox8.SelectedIndex).id_supply_area, g_selected_category, dr_lab_cat_def) Then
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

            Cursor = Cursors.Arrow

        End If

    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged

        If CheckedListBox1.Items.Count() > 0 Then

            Dim result As Integer = 0

            result = MsgBox("There are unsaved records. Do you wish to save them?", vbYesNo)

            If (result = DialogResult.Yes) Then

                Cursor = Cursors.WaitCursor

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

                If Not db_supplies.SET_SUPPLY_TYPE(TextBox1.Text, l_a_checked_supplies) Then

                    MsgBox("ERROR INSERTING SUPPLIES CATEGORIES!", vbCritical)

                ElseIf Not db_supplies.SET_SUPPLY(TextBox1.Text, l_a_checked_supplies) Then

                    MsgBox("ERROR INSERTING SUPPLIES!", vbCritical)

                ElseIf Not db_supplies.SET_SUPPLY_SOFT_INST(TextBox1.Text, g_selected_soft, l_a_checked_supplies) Then

                    MsgBox("ERROR INSERTING SUPPLY_SOFT_INST!", vbCritical)

                ElseIf Not db_supplies.SET_SUPPLY_SUP_AREA(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox8.SelectedIndex).id_supply_area, l_a_checked_supplies) Then

                    MsgBox("ERROR INSERTING SUPPLY_SOFT_INST!", vbCritical)

                ElseIf Not db_supplies.SET_SUPPLY_LOC_DEFAULT(TextBox1.Text, g_selected_soft, l_a_checked_supplies) Then

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
                    ReDim g_a_supp_cats_alert_barcode(0)
                    ComboBox9.SelectedIndex = -1
                    ComboBox9.Text = ""
                    ComboBox6.SelectedIndex = -1
                    ComboBox6.Text = ""
                    ReDim g_a_supps_alert(0)
                    g_dimension_supp_alert = 0
                    CheckedListBox3.Items.Clear()

                    ComboBox10.SelectedIndex = -1
                    ComboBox10.Text = ""
                    ComboBox11.SelectedIndex = -1
                    ComboBox11.Text = ""

                End If

            End If

            Cursor = Cursors.Arrow

        End If

        CheckedListBox1.Items.Clear()

        ComboBox7.SelectedIndex = -1
        ComboBox7.Text = ""

        CheckedListBox2.Items.Clear()

        ComboBox3.Items.Clear()
        ComboBox3.Text = ""
        ComboBox4.Items.Clear()
        ComboBox4.Text = ""

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
        If Not db_supplies.GET_SUPS_DEFAULT_BY_CAT(TextBox1.Text, g_selected_soft, ComboBox3.Text, g_a_SUP_AREAS(ComboBox8.SelectedIndex).id_supply_area, g_selected_category, g_selected_supplycategory, dr) Then
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
        ReDim g_a_supp_cats_alert_barcode(0)
        ComboBox9.SelectedIndex = -1
        ComboBox9.Text = ""
        ComboBox6.SelectedIndex = -1
        ComboBox6.Text = ""
        ReDim g_a_supps_alert(0)
        g_dimension_supp_alert = 0

        ComboBox10.SelectedIndex = -1
        ComboBox10.Text = ""
        ComboBox11.SelectedIndex = -1
        ComboBox11.Text = ""

        g_selected_soft = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, TextBox1.Text)

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

        TextBox1.Text = db_access_general.GET_INSTITUTION_ID(ComboBox1.SelectedIndex)

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""

        'Bloco do ALERT
        ComboBox5.Items.Clear()
        ComboBox5.Text = ""
        ReDim g_a_supp_cats_alert(0)
        ReDim g_a_supp_cats_alert_barcode(0)
        ComboBox9.SelectedIndex = -1
        ComboBox9.Text = ""
        ComboBox6.SelectedIndex = -1
        ComboBox6.Text = ""
        ReDim g_a_supps_alert(0)
        g_dimension_supp_alert = 0

        ComboBox10.SelectedIndex = -1
        ComboBox10.Text = ""
        ComboBox11.SelectedIndex = -1
        ComboBox11.Text = ""

        Dim dr As OracleDataReader

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""

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

            If Not db_supplies.SET_SUPPLY_TYPE(TextBox1.Text, l_a_checked_supplies) Then

                MsgBox("ERROR INSERTING SUPPLIES CATEGORIES!", vbCritical)

            ElseIf Not db_supplies.SET_SUPPLY(TextBox1.Text, l_a_checked_supplies) Then

                MsgBox("ERROR INSERTING SUPPLIES!", vbCritical)

            ElseIf Not db_supplies.SET_SUPPLY_SOFT_INST(TextBox1.Text, g_selected_soft, l_a_checked_supplies) Then

                MsgBox("ERROR INSERTING SUPPLY_SOFT_INST!", vbCritical)

            ElseIf Not db_supplies.SET_SUPPLY_SUP_AREA(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox8.SelectedIndex).id_supply_area, l_a_checked_supplies) Then

                MsgBox("ERROR INSERTING SUPPLY_SOFT_INST!", vbCritical)

            ElseIf Not db_supplies.SET_SUPPLY_LOC_DEFAULT(TextBox1.Text, g_selected_soft, l_a_checked_supplies) Then

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
                ReDim g_a_supp_cats_alert_barcode(0)
                ComboBox9.SelectedIndex = -1
                ComboBox9.Text = ""
                ComboBox6.SelectedIndex = -1
                ComboBox6.Text = ""
                ReDim g_a_supps_alert(0)
                g_dimension_supp_alert = 0
                CheckedListBox3.Items.Clear()

                ComboBox10.SelectedIndex = -1
                ComboBox10.Text = ""
                ComboBox11.SelectedIndex = -1
                ComboBox11.Text = ""

                g_type_supply_alert_barcode = ""
                g_selected_supplycategory_alert_barcode = ""

                DataGridView1.DataBindings.Clear()
                DataGridView1.DataSource = Nothing
                DataGridView1.Rows.Clear()

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
        If Not db_supplies.GET_SUPS_ALERT_BY_CAT(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox6.SelectedIndex).id_supply_area, g_type_supply_alert, g_selected_supplycategory_alert, dr_supplies) Then
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
        If Not db_supplies.GET_SUPP_CATS_ALERT(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox6.SelectedIndex).id_supply_area, g_type_supply_alert, dr_supp_cat) Then
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

                If Not db_supplies.DELETE_SUPPLY_SOFT_INST(TextBox1.Text, g_selected_soft, l_selected_supplies_alert, g_a_SUP_AREAS(ComboBox6.SelectedIndex).id_supply_area) Then

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
                    If Not db_supplies.GET_SUPP_CATS_ALERT(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox6.SelectedIndex).id_supply_area, g_type_supply_alert, dr_supp_cat) Then
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
                    If Not db_supplies.GET_SUPS_ALERT_BY_CAT(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox6.SelectedIndex).id_supply_area, g_type_supply_alert, g_selected_supplycategory_alert, dr_supplies) Then
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

            'Bloco para limpar grelha de barcodes caso se tenha apagado um material da categoria selecionada na grelha de barcode
            If (ComboBox6.Text = ComboBox10.Text And ComboBox9.Text = ComboBox11.Text And ComboBox5.Text = ComboBox12.Text) Or (ComboBox5.Text = "") Then

                ReDim g_a_supp_cats_alert_barcode(0)

                ComboBox12.Items.Clear()
                ComboBox12.Text = ""
                ReDim g_a_supp_cats_alert_barcode(0)

                g_selected_supplycategory_alert_barcode = ""

                DataGridView1.DataBindings.Clear()
                DataGridView1.DataSource = Nothing
                DataGridView1.Rows.Clear()

                If ComboBox11.SelectedIndex > -1 Then

                    'Box de categorias na instituição/software
                    Dim dr_supp_cat As OracleDataReader

                    ReDim g_a_supp_cats_alert_barcode(0)
                    Dim l_index_loaded_categories As Int64 = 0

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                    If Not db_supplies.GET_SUPP_CATS_ALERT(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox10.SelectedIndex).id_supply_area, g_type_supply_alert_barcode, dr_supp_cat) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                        MsgBox("ERROR LOADING SUPPLY CATEGORIES FROM INSTITUTION!", vbCritical)

                    Else

                        ComboBox12.Items.Add("ALL")

                        ReDim g_a_supp_cats_alert_barcode(0)
                        g_a_supp_cats_alert_barcode(0) = 0

                        Dim l_index As Int16 = 1

                        While dr_supp_cat.Read()

                            ComboBox12.Items.Add(dr_supp_cat.Item(1))
                            ReDim Preserve g_a_supp_cats_alert_barcode(l_index)
                            g_a_supp_cats_alert_barcode(l_index) = dr_supp_cat.Item(0)
                            l_index = l_index + 1

                        End While

                    End If

                    dr_supp_cat.Dispose()
                    dr_supp_cat.Close()

                End If

            End If

            Cursor = Cursors.Arrow

            End If

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged

        ComboBox11.SelectedIndex = -1
        ComboBox11.Text = ""

    End Sub

    Private Sub ComboBox11_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox11.SelectedIndexChanged

        'Bloco Para Limpar Variaveis
        ComboBox12.Items.Clear()
        ComboBox12.Text = ""
        ReDim g_a_supp_cats_alert_barcode(0)

        'Fim BLoco para Limpar Variaveis

        If ComboBox11.Text = g_activity_desc Then

            g_type_supply_alert_barcode = g_activity_flag

        ElseIf ComboBox11.Text = g_implants_desc Then

            g_type_supply_alert_barcode = g_implants_flag

        ElseIf ComboBox11.Text = g_kits_desc Then

            g_type_supply_alert_barcode = g_kits_flag

        ElseIf ComboBox11.Text = g_sets_desc Then

            g_type_supply_alert_barcode = g_sets_flag

        ElseIf ComboBox11.Text = g_supplies_desc Then

            g_type_supply_alert_barcode = g_supplies_flag

        ElseIf ComboBox11.Text = g_surgical_desc Then

            g_type_supply_alert_barcode = g_surgical_flag

        Else

            g_type_supply_alert = "ALL"

        End If

        If ComboBox11.SelectedIndex > -1 Then

            'Box de categorias na instituição/software
            Dim dr_supp_cat As OracleDataReader

            ReDim g_a_supp_cats_alert_barcode(0)
            Dim l_index_loaded_categories As Int64 = 0

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not db_supplies.GET_SUPP_CATS_ALERT(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox10.SelectedIndex).id_supply_area, g_type_supply_alert_barcode, dr_supp_cat) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                MsgBox("ERROR LOADING SUPPLY CATEGORIES FROM INSTITUTION!", vbCritical)

            Else

                ComboBox12.Items.Add("ALL")

                ReDim g_a_supp_cats_alert_barcode(0)
                g_a_supp_cats_alert_barcode(0) = 0

                Dim l_index As Int16 = 1

                While dr_supp_cat.Read()

                    ComboBox12.Items.Add(dr_supp_cat.Item(1))
                    ReDim Preserve g_a_supp_cats_alert_barcode(l_index)
                    g_a_supp_cats_alert_barcode(l_index) = dr_supp_cat.Item(0)
                    l_index = l_index + 1

                End While

            End If

            dr_supp_cat.Dispose()
            dr_supp_cat.Close()

        End If

    End Sub

    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        'Determinar o id_content da categoria selecionada
        If ComboBox12.SelectedIndex = 0 Then
            g_selected_supplycategory_alert_barcode = 0
        Else
            g_selected_supplycategory_alert_barcode = g_a_supp_cats_alert_barcode(ComboBox12.SelectedIndex)
        End If

        'g_dimension_supp_alert_barcode = 0
        'ReDim g_a_supps_alert_barcode(g_dimension_supp_alert)

        Dim dr As OracleDataReader

        If Not db_supplies.GET_SUPS_ALERT_BARCODE(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox10.SelectedIndex).id_supply_area, g_type_supply_alert_barcode, g_selected_supplycategory_alert_barcode, dr) Then

            MsgBox("ERROR GETTING INTERVENTIONS FROM INSTITUTION!", MsgBoxStyle.Critical)

        Else

            ReDim g_supply_barcode(0)
            Dim index As Integer = 0

            While dr.Read()

                ReDim Preserve g_supply_barcode(index)
                g_supply_barcode(index).desc_supply_type = dr.Item(0)
                g_supply_barcode(index).desc_supply = dr.Item(1)
                Try
                    g_supply_barcode(index).barcode = dr.Item(2)
                Catch ex As Exception
                    g_supply_barcode(index).barcode = ""
                End Try

                Try
                    g_supply_barcode(index).lote = dr.Item(3)
                Catch ex As Exception
                    g_supply_barcode(index).lote = ""
                End Try

                Try
                    g_supply_barcode(index).serial = dr.Item(4)
                Catch ex As Exception
                    g_supply_barcode(index).serial = ""
                End Try

                g_supply_barcode(index).id_supply = dr.Item(5)

                index = index + 1

            End While

            DataGridView1.DataBindings.Clear()

            DataGridView1.DataSource = Nothing

            DataGridView1.Rows.Clear()

            DataGridView1.ColumnCount = 5
            DataGridView1.Columns(0).Name = "CATEGORY"
            DataGridView1.Columns(1).Name = "SUPPLY"
            DataGridView1.Columns(2).Name = "BARCODE"
            DataGridView1.Columns(3).Name = "LOT"
            DataGridView1.Columns(4).Name = "SERIAL_NUMBER"

            DataGridView1.Columns(0).Width = 170
            DataGridView1.Columns(1).Width = 265
            DataGridView1.Columns(2).Width = 105
            DataGridView1.Columns(3).Width = 105
            DataGridView1.Columns(4).Width = 123

            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable

            DataGridView1.Columns(0).ReadOnly = True
            DataGridView1.Columns(1).ReadOnly = True
            DataGridView1.Columns(2).ReadOnly = False
            DataGridView1.Columns(3).ReadOnly = False
            DataGridView1.Columns(4).ReadOnly = False


            For i As Integer = 0 To g_supply_barcode.Count() - 1

                DataGridView1.Rows.Add()
                DataGridView1.Rows(i).Cells(0).Value = g_supply_barcode(i).desc_supply_type
                DataGridView1.Rows(i).Cells(1).Value = g_supply_barcode(i).desc_supply
                DataGridView1.Rows(i).Cells(2).Value = g_supply_barcode(i).barcode
                DataGridView1.Rows(i).Cells(3).Value = g_supply_barcode(i).lote
                DataGridView1.Rows(i).Cells(4).Value = g_supply_barcode(i).serial

            Next

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Cursor = Cursors.WaitCursor

        Dim l_success As Boolean = True 'Variável de controlo para os erros

        Dim l_records_updated As Boolean = False 'Variável de controlo para mensagem de sucesso

        For i As Integer = 0 To g_supply_barcode.Count() - 1

            If DataGridView1.Rows(i).Cells(2).Value <> g_supply_barcode(i).barcode Then

                If Not db_supplies.SET_BARCODE(TextBox1.Text, g_supply_barcode(i).id_supply, DataGridView1.Rows(i).Cells(2).Value) Then

                    l_success = False

                Else

                    l_records_updated = True

                End If


            End If

            If DataGridView1.Rows(i).Cells(3).Value <> g_supply_barcode(i).lote Then

                If Not db_supplies.SET_LOT(TextBox1.Text, g_supply_barcode(i).id_supply, DataGridView1.Rows(i).Cells(3).Value) Then

                    l_success = False

                Else

                    l_records_updated = True

                End If

            End If

            If DataGridView1.Rows(i).Cells(4).Value <> g_supply_barcode(i).serial Then

                If Not db_supplies.SET_SERIAL_NUMBER(TextBox1.Text, g_supply_barcode(i).id_supply, DataGridView1.Rows(i).Cells(4).Value) Then

                    l_success = False

                Else

                    l_records_updated = True

                End If

            End If

        Next

        If l_success = False Then

            MsgBox("ERROR UPDATING SUPPLY INFORMATION!", vbCritical)

            'Bloco para Limpara a DATAGRIDVIEW e a estrutura
            Dim dr As OracleDataReader

            If Not db_supplies.GET_SUPS_ALERT_BARCODE(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox10.SelectedIndex).id_supply_area, g_type_supply_alert_barcode, g_selected_supplycategory_alert_barcode, dr) Then

                MsgBox("ERROR GETTING INTERVENTIONS FROM INSTITUTION!", MsgBoxStyle.Critical)

            Else

                ReDim g_supply_barcode(0)
                Dim index As Integer = 0

                While dr.Read()

                    ReDim Preserve g_supply_barcode(index)
                    g_supply_barcode(index).desc_supply_type = dr.Item(0)
                    g_supply_barcode(index).desc_supply = dr.Item(1)
                    Try
                        g_supply_barcode(index).barcode = dr.Item(2)
                    Catch ex As Exception
                        g_supply_barcode(index).barcode = ""
                    End Try

                    Try
                        g_supply_barcode(index).lote = dr.Item(3)
                    Catch ex As Exception
                        g_supply_barcode(index).lote = ""
                    End Try

                    Try
                        g_supply_barcode(index).serial = dr.Item(4)
                    Catch ex As Exception
                        g_supply_barcode(index).serial = ""
                    End Try

                    g_supply_barcode(index).id_supply = dr.Item(5)

                    index = index + 1

                End While

                DataGridView1.DataBindings.Clear()

                DataGridView1.DataSource = Nothing

                DataGridView1.Rows.Clear()

                DataGridView1.ColumnCount = 5
                DataGridView1.Columns(0).Name = "CATEGORY"
                DataGridView1.Columns(1).Name = "SUPPLY"
                DataGridView1.Columns(2).Name = "BARCODE"
                DataGridView1.Columns(3).Name = "LOT"
                DataGridView1.Columns(4).Name = "SERIAL_NUMBER"

                DataGridView1.Columns(0).Width = 170
                DataGridView1.Columns(1).Width = 265
                DataGridView1.Columns(2).Width = 105
                DataGridView1.Columns(3).Width = 105
                DataGridView1.Columns(4).Width = 123

                DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable

                DataGridView1.Columns(0).ReadOnly = True
                DataGridView1.Columns(1).ReadOnly = True
                DataGridView1.Columns(2).ReadOnly = False
                DataGridView1.Columns(3).ReadOnly = False
                DataGridView1.Columns(4).ReadOnly = False


                For i As Integer = 0 To g_supply_barcode.Count() - 1

                    DataGridView1.Rows.Add()
                    DataGridView1.Rows(i).Cells(0).Value = g_supply_barcode(i).desc_supply_type
                    DataGridView1.Rows(i).Cells(1).Value = g_supply_barcode(i).desc_supply
                    DataGridView1.Rows(i).Cells(2).Value = g_supply_barcode(i).barcode
                    DataGridView1.Rows(i).Cells(3).Value = g_supply_barcode(i).lote
                    DataGridView1.Rows(i).Cells(4).Value = g_supply_barcode(i).serial

                Next
            End If


        ElseIf l_success = True And l_records_updated = True Then

            MsgBox("Records successfully updated.", vbInformation)

            'Bloco para Limpara a DATAGRIDVIEW e a estrutura e para repopular grelha e estrutura
            Dim dr As OracleDataReader

            If Not db_supplies.GET_SUPS_ALERT_BARCODE(TextBox1.Text, g_selected_soft, g_a_SUP_AREAS(ComboBox10.SelectedIndex).id_supply_area, g_type_supply_alert_barcode, g_selected_supplycategory_alert_barcode, dr) Then

                MsgBox("ERROR GETTING INTERVENTIONS FROM INSTITUTION!", MsgBoxStyle.Critical)

            Else

                ReDim g_supply_barcode(0)
                Dim index As Integer = 0

                While dr.Read()

                    ReDim Preserve g_supply_barcode(index)
                    g_supply_barcode(index).desc_supply_type = dr.Item(0)
                    g_supply_barcode(index).desc_supply = dr.Item(1)
                    Try
                        g_supply_barcode(index).barcode = dr.Item(2)
                    Catch ex As Exception
                        g_supply_barcode(index).barcode = ""
                    End Try

                    Try
                        g_supply_barcode(index).lote = dr.Item(3)
                    Catch ex As Exception
                        g_supply_barcode(index).lote = ""
                    End Try

                    Try
                        g_supply_barcode(index).serial = dr.Item(4)
                    Catch ex As Exception
                        g_supply_barcode(index).serial = ""
                    End Try

                    g_supply_barcode(index).id_supply = dr.Item(5)

                    index = index + 1

                End While

                DataGridView1.DataBindings.Clear()

                DataGridView1.DataSource = Nothing

                DataGridView1.Rows.Clear()

                DataGridView1.ColumnCount = 5
                DataGridView1.Columns(0).Name = "CATEGORY"
                DataGridView1.Columns(1).Name = "SUPPLY"
                DataGridView1.Columns(2).Name = "BARCODE"
                DataGridView1.Columns(3).Name = "LOT"
                DataGridView1.Columns(4).Name = "SERIAL_NUMBER"

                DataGridView1.Columns(0).Width = 170
                DataGridView1.Columns(1).Width = 265
                DataGridView1.Columns(2).Width = 105
                DataGridView1.Columns(3).Width = 105
                DataGridView1.Columns(4).Width = 123

                DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable

                DataGridView1.Columns(0).ReadOnly = True
                DataGridView1.Columns(1).ReadOnly = True
                DataGridView1.Columns(2).ReadOnly = False
                DataGridView1.Columns(3).ReadOnly = False
                DataGridView1.Columns(4).ReadOnly = False


                For i As Integer = 0 To g_supply_barcode.Count() - 1

                    DataGridView1.Rows.Add()
                    DataGridView1.Rows(i).Cells(0).Value = g_supply_barcode(i).desc_supply_type
                    DataGridView1.Rows(i).Cells(1).Value = g_supply_barcode(i).desc_supply
                    DataGridView1.Rows(i).Cells(2).Value = g_supply_barcode(i).barcode
                    DataGridView1.Rows(i).Cells(3).Value = g_supply_barcode(i).lote
                    DataGridView1.Rows(i).Cells(4).Value = g_supply_barcode(i).serial

                Next

            End If

        End If

        Cursor = Cursors.Arrow

    End Sub
End Class
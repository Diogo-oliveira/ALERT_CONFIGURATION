Imports Oracle.DataAccess.Client
Public Class LAB_TESTS

    Dim db_access_general As New General
    Dim db_labs As New LABS_API
    Dim oradb As String = "Data Source=QC4V26522;User Id=alert_config;Password=qcteam"
    Dim l_selected_soft As Int16 = -1

    Dim l_loaded_categories_default() As String ' Array que vai guardar os id_contents das categorias carregadas do default
    Dim l_selected_category As String = ""

    Private Sub LAB_TESTS_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dr As OracleDataReader = db_access_general.GET_ALL_INSTITUTIONS(oradb)

        Dim i As Integer = 0

        While dr.Read()

            ComboBox1.Items.Add(dr.Item(0))

        End While

        dr.Dispose()

        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Dim form1 As New Form1

        Me.Enabled = False

        Me.Dispose()

        form1.Show()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox1.Text <> "" Then

            ComboBox1.Text = db_access_general.GET_INSTITUTION(TextBox1.Text, oradb)

            ComboBox2.Items.Clear()
            ComboBox2.Text = ""


            Dim dr As OracleDataReader = db_access_general.GET_SOFT_INST(TextBox1.Text, oradb)

            Dim i As Integer = 0

            While dr.Read()

                ComboBox2.Items.Add(dr.Item(1))

            End While

            ComboBox3.Items.Clear()
            ComboBox3.SelectedItem = ""

            ComboBox4.Items.Clear()
            ComboBox4.SelectedItem = ""

            CheckedListBox2.Items.Clear()

            l_selected_category = ""

        End If


    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        l_selected_soft = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, TextBox1.Text, oradb)

        '1 - Fill Version combobox

        ComboBox3.Items.Clear()
        ComboBox3.Text = ""

        Cursor = Cursors.WaitCursor

        Try

            Dim dr_def_versions As OracleDataReader = db_labs.GET_DEFAULT_VERSIONS(TextBox1.Text, l_selected_soft, oradb)

            While dr_def_versions.Read()

                ComboBox3.Items.Add(dr_def_versions.Item(0))

            End While

        Catch ex As Exception

            MsgBox("ERROR LOADING DEFAULT VERSIONS -  ComboBox2_SelectedIndexChanged", MsgBoxStyle.Critical)

        End Try

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        ComboBox4.Items.Clear()
        ComboBox4.Text = ""

        'Determinar as categorias disponíveis para a versão escolhida
        'Array l_loaded_categories_default vai gaurdar os ids de todas as categorias

        ReDim l_loaded_categories_default(0)
        Dim l_index_loaded_categories As Int16 = 0

        Try

            Dim dr_lab_cat_def As OracleDataReader = db_labs.GET_LAB_CATS_DEFAULT(ComboBox3.Text, TextBox1.Text, l_selected_soft, oradb)

            ComboBox4.Items.Add("ALL")

            While dr_lab_cat_def.Read()

                ComboBox4.Items.Add(dr_lab_cat_def.Item(1))
                l_loaded_categories_default(l_index_loaded_categories) = dr_lab_cat_def.Item(0)
                l_index_loaded_categories = l_index_loaded_categories + 1
                ReDim Preserve l_loaded_categories_default(l_index_loaded_categories)

            End While

        Catch ex As Exception

            MsgBox("ERROR LOADING DEFAULT LAB CATEGORIS -  ComboBox3_SelectedIndexChanged", MsgBoxStyle.Critical)

        End Try

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

        ''2 - Carregar a grelha de exames (fazer função - vai ser parecida à última que foi feita)
        ''3 Criar estrutura com os elementos dos exames carregados
        Dim dr As OracleDataReader = db_labs.GET_LABS_DEFAULT_BY_CAT(TextBox1.Text, l_selected_soft, ComboBox3.SelectedItem.ToString, l_selected_category, oradb)

        'ReDim loaded_exams(0) ''Limpar estrutura
        'Dim l_dimension_array_loaded_exams As Int64 = 0

        While dr.Read()

            CheckedListBox2.Items.Add(dr.Item(1) & " [" & dr.Item(3) & "]")

            ' ReDim Preserve loaded_exams(l_dimension_array_loaded_exams)

            'loaded_exams(l_dimension_array_loaded_exams).id_content_category = dr.Item(0)
            ' loaded_exams(l_dimension_array_loaded_exams).desc_category = dr.Item(1)
            'loaded_exams(l_dimension_array_loaded_exams).id_content_exam = dr.Item(2)
            'loaded_exams(l_dimension_array_loaded_exams).desc_exam = dr.Item(3)
            ' loaded_exams(l_dimension_array_loaded_exams).flg_first_result = dr.Item(4)
            ' loaded_exams(l_dimension_array_loaded_exams).flg_execute = dr.Item(5)
            ' loaded_exams(l_dimension_array_loaded_exams).flg_timeout = dr.Item(6)
            '  loaded_exams(l_dimension_array_loaded_exams).flg_result_notes = dr.Item(7)
            'loaded_exams(l_dimension_array_loaded_exams).flg_first_execute = dr.Item(8)
            '
            ''Determinar as idades e gender dos exames
            ''se não houver idades minimas/maximas, devolve -1
            ''se não houver gender, devolve vazio

            ' Try

            'loaded_exams(l_dimension_array_loaded_exams).age_min = dr.Item(9)

            ' Catch ex As Exception

            'loaded_exams(l_dimension_array_loaded_exams).age_min = -1

            ' End Try

            'Try

            'loaded_exams(l_dimension_array_loaded_exams).age_max = dr.Item(10)

            ' Catch ex As Exception

            'loaded_exams(l_dimension_array_loaded_exams).age_max = -1

            'End Try

            ' Try

            'loaded_exams(l_dimension_array_loaded_exams).gender = dr.Item(11)

            ' Catch ex As Exception

            'loaded_exams(l_dimension_array_loaded_exams).gender = ""

            ' End Try

            ' l_dimension_array_loaded_exams = l_dimension_array_loaded_exams + 1


        End While

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

        '  For Each indexChecked In CheckedListBox2.CheckedIndices

        'If para verificar se já está incluido na checkbox da direita

        Dim l_record_already_selected As Boolean = False

            Dim j As Integer = 0

        '    For j = 0 To CheckedListBox1.Items.Count() - 1

        '        If (loaded_exams(indexChecked.ToString()).id_content_exam = l_selected_default_exams(j).id_content_exam) Then

        '            l_record_already_selected = True
        '            Exit For

        '        End If

        '    Next

        '    If l_record_already_selected = False Then

        '        ReDim Preserve l_selected_default_exams(l_index_selected_exams_from_default)

        '        l_selected_default_exams(l_index_selected_exams_from_default).age_max = loaded_exams(indexChecked.ToString()).age_max
        '        l_selected_default_exams(l_index_selected_exams_from_default).age_min = loaded_exams(indexChecked.ToString()).age_min
        '        l_selected_default_exams(l_index_selected_exams_from_default).desc_category = loaded_exams(indexChecked.ToString()).desc_category
        '        l_selected_default_exams(l_index_selected_exams_from_default).flg_execute = loaded_exams(indexChecked.ToString()).flg_execute
        '        l_selected_default_exams(l_index_selected_exams_from_default).flg_first_execute = loaded_exams(indexChecked.ToString()).flg_first_execute
        '        l_selected_default_exams(l_index_selected_exams_from_default).flg_first_result = loaded_exams(indexChecked.ToString()).flg_first_result
        '        l_selected_default_exams(l_index_selected_exams_from_default).flg_result_notes = loaded_exams(indexChecked.ToString()).flg_result_notes
        '        l_selected_default_exams(l_index_selected_exams_from_default).flg_timeout = loaded_exams(indexChecked.ToString()).flg_timeout
        '        l_selected_default_exams(l_index_selected_exams_from_default).gender = loaded_exams(indexChecked.ToString()).gender
        '        l_selected_default_exams(l_index_selected_exams_from_default).id_content_category = loaded_exams(indexChecked.ToString()).id_content_category
        '        l_selected_default_exams(l_index_selected_exams_from_default).id_content_exam = loaded_exams(indexChecked.ToString()).id_content_exam
        '        l_selected_default_exams(l_index_selected_exams_from_default).desc_exam = loaded_exams(indexChecked.ToString()).desc_exam

        '        CheckedListBox1.Items.Add(l_selected_default_exams(l_index_selected_exams_from_default).desc_exam)
        '        CheckedListBox1.SetItemChecked((CheckedListBox1.Items.Count() - 1), True)

        '        l_index_selected_exams_from_default = l_index_selected_exams_from_default + 1

        '    End If

        'Next

        ''APAGAR
        MsgBox(db_access_general.GET_DEFAULT_TRANSLATION((db_access_general.GET_ID_LANG(2799, oradb)), "ANALYSIS.CODE_ANALYSIS.21557", oradb))

    End Sub
End Class
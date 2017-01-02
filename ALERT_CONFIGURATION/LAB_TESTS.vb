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
End Class
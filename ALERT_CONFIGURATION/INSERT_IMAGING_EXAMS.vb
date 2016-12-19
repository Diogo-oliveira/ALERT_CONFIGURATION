Imports Oracle.DataAccess.Client
Public Class INSERT_IMAGING_EXAMS

    Dim db_access As New EXAMS_API
    Dim oradb As String = "Data Source=QC4V26522;User Id=alert_config;Password=qcteam"
    Dim l_selected_soft As Int16 = -1
    Dim l_selected_category As String = ""


    ''Estrutura dos exames carregados do default
    Dim loaded_exams() As EXAMS_API.exams_default


    Private Sub INSERT_IMAGING_EXAMS_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dr As OracleDataReader = db_access.GET_ALL_INSTITUTIONS(oradb)

        Dim i As Integer = 0

        While dr.Read()

            ComboBox1.Items.Add(dr.Item(0))

        End While


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox1.Text <> "" Then

            ComboBox1.Text = db_access.GET_INSTITUTION(TextBox1.Text, oradb)

            ComboBox2.Items.Clear()
            ComboBox2.Text = ""


            Dim dr As OracleDataReader = db_access.GET_SOFT_INST(TextBox1.Text, oradb)

            Dim i As Integer = 0

            While dr.Read()

                ComboBox2.Items.Add(dr.Item(1))

            End While

            'l_selected_all_most_frequent = False

        End If

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        CheckedListBox1.Items.Clear()
        CheckedListBox2.Items.Clear()

        l_selected_soft = db_access.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, TextBox1.Text, oradb)

        ComboBox4.Items.Clear()
        ComboBox4.Text = ""

        ComboBox3.Items.Clear()
        ComboBox3.SelectedItem = ""


        Try

            Dim dr_def_versions As OracleDataReader = db_access.GET_DEFAULT_VERSIONS(TextBox1.Text, l_selected_soft, oradb)

            While dr_def_versions.Read()

                ComboBox3.Items.Add(dr_def_versions.Item(0))

            End While

        Catch ex As Exception

            MsgBox("ERROR LOADING DEFAULT VERSIONS -  ComboBox2_SelectedIndexChanged", MsgBoxStyle.Critical)

        End Try

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged


        ComboBox4.Items.Clear()
        ComboBox4.Text = ""

        CheckedListBox2.Items.Clear()

        Try

            Dim dr_exam_def As OracleDataReader = db_access.GET_EXAMS_CAT_DEFAULT(ComboBox3.Text, TextBox1.Text, l_selected_soft, oradb)

            ComboBox4.Items.Add("ALL")

            While dr_exam_def.Read()

                ComboBox4.Items.Add(dr_exam_def.Item(1))

            End While

        Catch ex As Exception

            MsgBox("ERROR LOADING DEFAULT EXAMS CATEGORY -  ComboBox3_SelectedIndexChanged", MsgBoxStyle.Critical)

        End Try

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        ''To DO
        ''1 - Determinar o id da categroia selecionada l_selected_category

        If ComboBox4.SelectedIndex = 0 Then

            l_selected_category = 0

        Else

            Try

                Dim dr_exam_def As OracleDataReader = db_access.GET_EXAMS_CAT_DEFAULT(ComboBox3.Text, TextBox1.Text, l_selected_soft, oradb)
                Dim l_index_aux As Int64 = 1


                While dr_exam_def.Read()



                    If l_index_aux = ComboBox4.SelectedIndex Then

                        l_selected_category = dr_exam_def.Item(0)
                        Exit While

                    End If

                    l_index_aux = l_index_aux + 1

                End While

            Catch ex As Exception

                MsgBox("ERROR DETERMINING ID_CONTENT OF CATEGORY -  ComboBox4_SelectedIndexChanged", MsgBoxStyle.Critical)

            End Try

        End If

        CheckedListBox2.Items.Clear()

        ''2 - Carregar a grelha de exames (fazer função - vai ser parecida à última que foi feita)
        ''3 Criar estrutura com os elementos dos exames carregados
        Dim dr As OracleDataReader = db_access.GET_EXAMS_DEFAULT_BY_CAT(TextBox1.Text, l_selected_soft, ComboBox3.SelectedItem.ToString, l_selected_category, oradb)

        ReDim loaded_exams(0) ''Limpar estrutura
        Dim l_dimension_array_loaded_exams As Int64 = 0

        While dr.Read()

            CheckedListBox2.Items.Add(dr.Item(3))

            ReDim Preserve loaded_exams(l_dimension_array_loaded_exams)

            loaded_exams(l_dimension_array_loaded_exams).id_content_category = dr.Item(0)
            loaded_exams(l_dimension_array_loaded_exams).id_content_exam = dr.Item(2)
            loaded_exams(l_dimension_array_loaded_exams).flg_first_result = dr.Item(4)
            loaded_exams(l_dimension_array_loaded_exams).flg_execute = dr.Item(5)
            loaded_exams(l_dimension_array_loaded_exams).flg_timeout = dr.Item(6)
            loaded_exams(l_dimension_array_loaded_exams).flg_result_notes = dr.Item(7)
            loaded_exams(l_dimension_array_loaded_exams).flg_first_execute = dr.Item(8)

            ''Determinar as idades e gender dos exames
            ''se não houver idades minimas/maximas, devolve -1
            ''se não houver gender, devolve vazio

            Try

                loaded_exams(l_dimension_array_loaded_exams).age_min = dr.Item(9)

            Catch ex As Exception

                loaded_exams(l_dimension_array_loaded_exams).age_min = -1

            End Try

            Try

                loaded_exams(l_dimension_array_loaded_exams).age_max = dr.Item(10)

            Catch ex As Exception

                loaded_exams(l_dimension_array_loaded_exams).age_max = -1

            End Try

            Try

                loaded_exams(l_dimension_array_loaded_exams).gender = dr.Item(11)

            Catch ex As Exception

                loaded_exams(l_dimension_array_loaded_exams).gender = ""

            End Try

            l_dimension_array_loaded_exams = l_dimension_array_loaded_exams + 1


        End While

        ''4 criar função que vai inserir os registos no alert. Função será chamada no botão >>

    End Sub

    Private Sub CheckedListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox2.SelectedIndexChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If CheckedListBox2.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox2.Items.Count - 1

                CheckedListBox2.SetItemChecked(i, True)

            Next

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ' For Each indexChecked In CheckedListBox2.CheckedIndices

        ' l_array_selected_indexes(i_index_checked_aux) = indexChecked.ToString()

        'i_index_checked_aux = i_index_checked_aux + 1



        '  Next

        If db_access.SET_EXAM_ALERT(470, 11, loaded_exams, oradb) Then

            MsgBox("SUCESS")

        Else

            MsgBox("ERROR")

        End If

    End Sub
End Class
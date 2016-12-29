Imports Oracle.DataAccess.Client
Public Class INSERT_OTHER_EXAM

    Dim db_access As New EXAMS_API
    Dim oradb As String = "Data Source=QC4V26522;User Id=alert_config;Password=qcteam"
    Dim l_selected_soft As Int16 = -1
    Dim l_selected_category As String = ""


    ''Estrutura dos exames carregados do default
    Dim loaded_exams() As EXAMS_API.exams_default

    'Estrutura que vai guardar os exames de default selecionados
    Dim l_selected_default_exams() As EXAMS_API.exams_default


    Dim l_index_selected_exams_from_default As Integer = 0 ''Variavel utilizada no botão de adicionar à box da direita (CHECKBOX 1)

    Dim l_total_cats As Int64 = 0

    Dim a_loaded_exams_by_cat_alert() As EXAMS_API.exams_alert ''Array que vai carregar todos os ids e descritivos de uma categoria do alert
    Dim a_selected_exams_alert() As EXAMS_API.exams_alert      ''Array que vai guardar os exames selecionados do alert
    Dim l_index_selected_exams_from_alert As Integer = 0 ''Variavel utilizada no botão de adicionar à box da direita (CHECKBOX 4 - do alert para o clinical service)

    Dim a_dep_clin_serv_inst() As Int64 ''Array que vai guardar os dep_clin_serv da instituição

    Private Sub INSERT_OTHER_EXAM_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dr As OracleDataReader = db_access.GET_ALL_INSTITUTIONS(oradb)

        Dim i As Integer = 0

        While dr.Read()

            ComboBox1.Items.Add(dr.Item(0))

        End While

        dr.Dispose()

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

        End If

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        TextBox1.Text = db_access.GET_INSTITUTION_ID(ComboBox1.SelectedIndex, oradb)

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""


        Dim dr As OracleDataReader = db_access.GET_SOFT_INST(TextBox1.Text, oradb)

        Dim i As Integer = 0

        While dr.Read()

            ComboBox2.Items.Add(dr.Item(1))

        End While

        'l_selected_all_most_frequent = False

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

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        CheckedListBox1.Items.Clear()
        CheckedListBox2.Items.Clear()

        l_selected_soft = db_access.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, TextBox1.Text, oradb)

        ComboBox4.Items.Clear()
        ComboBox4.Text = ""

        ComboBox3.Items.Clear()
        ComboBox3.Text = ""

        Try

            Dim dr_def_versions As OracleDataReader = db_access.GET_DEFAULT_VERSIONS(TextBox1.Text, l_selected_soft, "E", oradb)

            While dr_def_versions.Read()

                ComboBox3.Items.Add(dr_def_versions.Item(0))

            End While

        Catch ex As Exception

            MsgBox("ERROR LOADING DEFAULT VERSIONS -  ComboBox2_SelectedIndexChanged", MsgBoxStyle.Critical)

        End Try

        ComboBox5.Items.Clear()
        ComboBox5.Text = ""

        Try

            Dim dr_exam_cat As OracleDataReader = db_access.GET_EXAMS_CAT(TextBox1.Text, l_selected_soft, "E", oradb)

            ComboBox5.Items.Add("ALL")

            While dr_exam_cat.Read()

                ComboBox5.Items.Add(dr_exam_cat.Item(0))
                l_total_cats = l_total_cats + 1

            End While

        Catch ex As Exception

            MsgBox("Error Loading Exams Categories!", MsgBoxStyle.Critical)

        End Try

        ''''''''''''''''''''

        Dim dr As OracleDataReader = db_access.GET_CLIN_SERV(TextBox1.Text, l_selected_soft, oradb)

        Dim i As Integer = 0

        ComboBox6.Items.Clear()
        ComboBox6.Text = ""

        Dim l_index_dep_clin_serv As Integer = 0
        ReDim a_dep_clin_serv_inst(l_index_dep_clin_serv)

        While dr.Read()

            ComboBox6.Items.Add(dr.Item(0))

            ReDim Preserve a_dep_clin_serv_inst(l_index_dep_clin_serv)
            a_dep_clin_serv_inst(l_index_dep_clin_serv) = dr.Item(1)
            l_index_dep_clin_serv = l_index_dep_clin_serv + 1
        End While

        ''''''''''''''''''''''''''''

        CheckedListBox3.Items.Clear()
        CheckedListBox4.Items.Clear()

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        ComboBox4.Items.Clear()
        ComboBox4.Text = ""

        CheckedListBox2.Items.Clear()

        Try

            Dim dr_exam_def As OracleDataReader = db_access.GET_EXAMS_CAT_DEFAULT(ComboBox3.Text, TextBox1.Text, l_selected_soft, "E", oradb)

            ComboBox4.Items.Add("ALL")

            While dr_exam_def.Read()

                ComboBox4.Items.Add(dr_exam_def.Item(1))

            End While

        Catch ex As Exception

            MsgBox("ERROR LOADING DEFAULT EXAMS CATEGORY -  ComboBox3_SelectedIndexChanged", MsgBoxStyle.Critical)

        End Try

        CheckedListBox1.Items.Clear()

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        ''To DO
        ''1 - Determinar o id da categroia selecionada l_selected_category

        If ComboBox4.SelectedIndex = 0 Then

            l_selected_category = 0

        Else

            Try

                Dim dr_exam_def As OracleDataReader = db_access.GET_EXAMS_CAT_DEFAULT(ComboBox3.Text, TextBox1.Text, l_selected_soft, "E", oradb)
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
        Dim dr As OracleDataReader = db_access.GET_EXAMS_DEFAULT_BY_CAT(TextBox1.Text, l_selected_soft, ComboBox3.SelectedItem.ToString, l_selected_category, "E", oradb)

        ReDim loaded_exams(0) ''Limpar estrutura
        Dim l_dimension_array_loaded_exams As Int64 = 0

        While dr.Read()

            CheckedListBox2.Items.Add(dr.Item(3))

            ReDim Preserve loaded_exams(l_dimension_array_loaded_exams)

            loaded_exams(l_dimension_array_loaded_exams).id_content_category = dr.Item(0)
            loaded_exams(l_dimension_array_loaded_exams).desc_category = dr.Item(1)
            loaded_exams(l_dimension_array_loaded_exams).id_content_exam = dr.Item(2)
            loaded_exams(l_dimension_array_loaded_exams).desc_exam = dr.Item(3)
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

        For Each indexChecked In CheckedListBox2.CheckedIndices

            'If para verificar se já está incluido na checkbox da direita

            Dim l_record_already_selected As Boolean = False

            Dim j As Integer = 0

            For j = 0 To CheckedListBox1.Items.Count() - 1

                If (loaded_exams(indexChecked.ToString()).id_content_exam = l_selected_default_exams(j).id_content_exam) Then

                    l_record_already_selected = True
                    Exit For

                End If

            Next

            If l_record_already_selected = False Then

                ReDim Preserve l_selected_default_exams(l_index_selected_exams_from_default)

                l_selected_default_exams(l_index_selected_exams_from_default).age_max = loaded_exams(indexChecked.ToString()).age_max
                l_selected_default_exams(l_index_selected_exams_from_default).age_min = loaded_exams(indexChecked.ToString()).age_min
                l_selected_default_exams(l_index_selected_exams_from_default).desc_category = loaded_exams(indexChecked.ToString()).desc_category
                l_selected_default_exams(l_index_selected_exams_from_default).flg_execute = loaded_exams(indexChecked.ToString()).flg_execute
                l_selected_default_exams(l_index_selected_exams_from_default).flg_first_execute = loaded_exams(indexChecked.ToString()).flg_first_execute
                l_selected_default_exams(l_index_selected_exams_from_default).flg_first_result = loaded_exams(indexChecked.ToString()).flg_first_result
                l_selected_default_exams(l_index_selected_exams_from_default).flg_result_notes = loaded_exams(indexChecked.ToString()).flg_result_notes
                l_selected_default_exams(l_index_selected_exams_from_default).flg_timeout = loaded_exams(indexChecked.ToString()).flg_timeout
                l_selected_default_exams(l_index_selected_exams_from_default).gender = loaded_exams(indexChecked.ToString()).gender
                l_selected_default_exams(l_index_selected_exams_from_default).id_content_category = loaded_exams(indexChecked.ToString()).id_content_category
                l_selected_default_exams(l_index_selected_exams_from_default).id_content_exam = loaded_exams(indexChecked.ToString()).id_content_exam
                l_selected_default_exams(l_index_selected_exams_from_default).desc_exam = loaded_exams(indexChecked.ToString()).desc_exam

                CheckedListBox1.Items.Add(l_selected_default_exams(l_index_selected_exams_from_default).desc_exam)
                CheckedListBox1.SetItemChecked((CheckedListBox1.Items.Count() - 1), True)

                l_index_selected_exams_from_default = l_index_selected_exams_from_default + 1

            End If

        Next

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        If CheckedListBox2.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox2.Items.Count - 1

                CheckedListBox2.SetItemChecked(i, False)

            Next

        End If

    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim l_error As Boolean = False

        If CheckedListBox1.Items.Count() > 0 Then

            For Each indexChecked In CheckedListBox1.CheckedIndices

                'Estrutura auxiliar para ir gravando os dados dos exames marcados com o check
                Dim l_checked_default_exams_temp(0) As EXAMS_API.exams_default

                l_checked_default_exams_temp(0).age_max = l_selected_default_exams(indexChecked).age_max
                l_checked_default_exams_temp(0).age_min = l_selected_default_exams(indexChecked).age_min
                l_checked_default_exams_temp(0).desc_category = l_selected_default_exams(indexChecked).desc_category
                l_checked_default_exams_temp(0).desc_exam = l_selected_default_exams(indexChecked).desc_exam
                l_checked_default_exams_temp(0).flg_execute = l_selected_default_exams(indexChecked).flg_execute
                l_checked_default_exams_temp(0).flg_first_execute = l_selected_default_exams(indexChecked).flg_first_execute
                l_checked_default_exams_temp(0).flg_first_result = l_selected_default_exams(indexChecked).flg_first_result
                l_checked_default_exams_temp(0).flg_result_notes = l_selected_default_exams(indexChecked).flg_result_notes
                l_checked_default_exams_temp(0).flg_timeout = l_selected_default_exams(indexChecked).flg_timeout
                l_checked_default_exams_temp(0).gender = l_selected_default_exams(indexChecked).gender
                l_checked_default_exams_temp(0).id_content_category = l_selected_default_exams(indexChecked).id_content_category
                l_checked_default_exams_temp(0).id_content_exam = l_selected_default_exams(indexChecked).id_content_exam

                ''Função para inserir no ALERT os exames selecionados
                If Not db_access.SET_EXAM_ALERT(TextBox1.Text, l_selected_soft, l_checked_default_exams_temp, "E", oradb) Then

                    MsgBox("ERROR INSERTING EXAM(S)!", vbCritical)
                    l_error = True

                End If
            Next

            If l_error = False Then

                MsgBox("Record(s) inserted!")

            End If

            CheckedListBox1.Items.Clear()

            For i As Integer = 0 To CheckedListBox2.Items.Count - 1

                CheckedListBox2.SetItemChecked(i, False)

            Next

            ComboBox5.Items.Clear()
            ComboBox5.SelectedItem = ""

            Try

                Dim dr_exam_cat As OracleDataReader = db_access.GET_EXAMS_CAT(TextBox1.Text, l_selected_soft, "E", oradb)

                ComboBox5.Items.Add("ALL")

                While dr_exam_cat.Read()

                    ComboBox5.Items.Add(dr_exam_cat.Item(0))
                    l_total_cats = l_total_cats + 1

                End While

            Catch ex As Exception

                MsgBox("Error Loading Exams Categories!", MsgBoxStyle.Critical)

            End Try

            CheckedListBox3.Items.Clear()

        Else

            MsgBox("No records selected!", vbInformation)

        End If

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

        Try

            CheckedListBox3.Items.Clear()

            Dim l_exam_cat(l_total_cats)

            l_exam_cat(0) = 0 ''Referente ao all

            Dim dr_exam_cat As OracleDataReader = db_access.GET_EXAMS_CAT(TextBox1.Text, l_selected_soft, "E", oradb)

            Dim i_cats As Integer = 1

            While dr_exam_cat.Read()

                l_exam_cat(i_cats) = dr_exam_cat.Item(1)
                i_cats = i_cats + 1
            End While

            Dim dr As OracleDataReader = db_access.GET_EXAMS(TextBox1.Text, l_selected_soft, l_exam_cat(ComboBox5.SelectedIndex), "E", oradb)

            Dim i As Integer = 0

            ReDim a_loaded_exams_by_cat_alert(0)

            Dim l_index_aux_loaded_exams = 0

            While dr.Read()

                CheckedListBox3.Items.Add(dr.Item(0))
                ReDim Preserve a_loaded_exams_by_cat_alert(l_index_aux_loaded_exams)
                a_loaded_exams_by_cat_alert(l_index_aux_loaded_exams).id_exam = dr.Item(2)
                a_loaded_exams_by_cat_alert(l_index_aux_loaded_exams).desc_exam = dr.Item(0)

                l_index_aux_loaded_exams = l_index_aux_loaded_exams + 1

            End While

        Catch ex As Exception

            MsgBox("Error selecting exams - GET_EXAMS", MsgBoxStyle.Critical)

        End Try

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Dim form1 As New Form1

        form1.Show()

        Me.Close()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        For Each indexChecked In CheckedListBox3.CheckedIndices

            'If para verificar se já está incluido na checkbox da direita

            Dim l_record_already_selected As Boolean = False

            Dim j As Integer = 0

            For j = 0 To CheckedListBox4.Items.Count() - 1

                If (a_loaded_exams_by_cat_alert(indexChecked.ToString()).id_exam = a_selected_exams_alert(j).id_exam) Then

                    l_record_already_selected = True
                    Exit For

                End If

            Next

            If l_record_already_selected = False Then

                ReDim Preserve a_selected_exams_alert(l_index_selected_exams_from_alert)

                a_selected_exams_alert(l_index_selected_exams_from_alert).id_exam = a_loaded_exams_by_cat_alert(indexChecked.ToString()).id_exam
                a_selected_exams_alert(l_index_selected_exams_from_alert).desc_exam = a_loaded_exams_by_cat_alert(indexChecked.ToString()).desc_exam

                CheckedListBox4.Items.Add(a_selected_exams_alert(l_index_selected_exams_from_alert).desc_exam)
                CheckedListBox4.SetItemChecked((CheckedListBox4.Items.Count() - 1), True)

                l_index_selected_exams_from_alert = l_index_selected_exams_from_alert + 1

            End If

        Next

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        If ComboBox6.SelectedItem = "" Then

            MsgBox("No clincial Service selected", vbCritical)

        Else

            Dim l_id_dep_clin_serv As Int64 = a_dep_clin_serv_inst(ComboBox6.SelectedIndex)

            If CheckedListBox4.Items.Count() > 0 Then

                For Each indexChecked In CheckedListBox4.CheckedIndices

                    'Estrutura auxiliar para ir gravando os dados dos exames marcados com o check
                    Dim l_checked_alert_exams_temp(0) As EXAMS_API.exams_alert

                    l_checked_alert_exams_temp(0).desc_exam = a_selected_exams_alert(indexChecked).desc_exam
                    l_checked_alert_exams_temp(0).id_exam = a_selected_exams_alert(indexChecked).id_exam

                    If Not db_access.SET_EXAM_DEP_CLIN_SERV(l_checked_alert_exams_temp(0).id_exam, l_id_dep_clin_serv, "M", TextBox1.Text,
                                                                l_selected_soft, "DT", "Y", "N",
                                                                "N", "DT", oradb) Then

                        MsgBox("ERROR INSERTING EXAM AS FREQUENT", vbCritical)

                    End If

                Next

                MsgBox("Selected record(s) saved.", vbInformation)

                CheckedListBox4.Items.Clear()

                For ii As Integer = 0 To CheckedListBox3.Items.Count - 1

                    CheckedListBox3.SetItemChecked(ii, False)

                Next

            Else

                MsgBox("No records selected!", vbInformation)

            End If

        End If

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        If CheckedListBox3.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox3.Items.Count - 1

                CheckedListBox3.SetItemChecked(i, False)

            Next

        End If

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        If CheckedListBox3.Items.Count() > 0 Then

            For i As Integer = 0 To CheckedListBox3.Items.Count - 1

                CheckedListBox3.SetItemChecked(i, True)

            Next

        End If

    End Sub
End Class
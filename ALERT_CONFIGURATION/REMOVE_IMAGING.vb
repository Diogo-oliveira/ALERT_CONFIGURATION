Imports Oracle.DataAccess.Client
Public Class REMOVE_IMAGING
    Dim db_acces As New EXAMS_API

    Dim oradb As String = "Data Source=QC4V26522;User Id=alert_config;Password=qcteam"

    Dim l_selected_soft As Int16 = -1
    Dim l_selected_dep_clin_serv As Int64 = -1
    Dim l_selected_exam() As Integer

    Dim l_selected_all_most_frequent As Boolean = False
    Dim l_selected_all As Boolean = False

    Dim l_exam_cat() As Integer
    Dim l_total_cats As Int64 = 0

    Private Sub REMOVE_IMAGING_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        CheckedListBox1.CheckOnClick = True

        Dim dr As OracleDataReader = db_acces.GET_ALL_INSTITUTIONS(oradb)


        Dim i As Integer = 0

        While dr.Read()

            ComboBox1.Items.Add(dr.Item(0))

        End While

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        TextBox1.Text = db_acces.GET_INSTITUTION_ID(ComboBox1.SelectedIndex, oradb)

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""

        ComboBox3.Items.Clear()
        ComboBox3.Text = ""

        ComboBox4.Items.Clear()
        ComboBox4.Text = ""

        CheckedListBox1.Items.Clear()
        CheckedListBox2.Items.Clear()

        Dim dr As OracleDataReader = db_acces.GET_SOFT_INST(TextBox1.Text, oradb)

        Dim i As Integer = 0

        While dr.Read()

            ComboBox2.Items.Add(dr.Item(1))

        End While

        l_selected_all_most_frequent = False

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox1.Text <> "" Then

            ComboBox1.Text = db_acces.GET_INSTITUTION(TextBox1.Text, oradb)

            ComboBox2.Items.Clear()
            ComboBox2.Text = ""

            ComboBox3.Items.Clear()
            ComboBox3.Text = ""

            ComboBox4.Items.Clear()
            ComboBox4.Text = ""

            CheckedListBox1.Items.Clear()
            CheckedListBox2.Items.Clear()

            Dim dr As OracleDataReader = db_acces.GET_SOFT_INST(TextBox1.Text, oradb)

            Dim i As Integer = 0

            While dr.Read()

                ComboBox2.Items.Add(dr.Item(1))

            End While

            l_selected_all_most_frequent = False

        End If

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        CheckedListBox1.Items.Clear()
        CheckedListBox2.Items.Clear()

        l_selected_soft = db_acces.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, TextBox1.Text, oradb)

        Dim dr As OracleDataReader = db_acces.GET_CLIN_SERV(TextBox1.Text, l_selected_soft, oradb)

        Dim i As Integer = 0

        ComboBox3.Items.Clear()
        ComboBox3.Text = ""

        While dr.Read()

            ComboBox3.Items.Add(dr.Item(0))

        End While

        l_selected_all_most_frequent = False

        ComboBox4.Items.Clear()
        ComboBox4.Text = ""

        Try

            Dim dr_exam_cat As OracleDataReader = db_acces.GET_EXAMS_CAT(TextBox1.Text, l_selected_soft, oradb)

            ComboBox4.Items.Add("ALL")

            While dr_exam_cat.Read()

                ComboBox4.Items.Add(dr_exam_cat.Item(0))
                l_total_cats = l_total_cats + 1

            End While

        Catch ex As Exception

            MsgBox("Error Loading Exams Categories!", MsgBoxStyle.Critical)

        End Try

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        CheckedListBox1.Items.Clear()

        l_selected_dep_clin_serv = db_acces.GET_SELECTED_DEP_CLIN_SERV(TextBox1.Text, l_selected_soft, ComboBox3.SelectedIndex, oradb) 'Colocar IDs

        Dim dr As OracleDataReader = db_acces.GET_FREQ_EXAM(l_selected_soft, l_selected_dep_clin_serv, TextBox1.Text, oradb)

        Dim i As Integer = 0

        While dr.Read()

            CheckedListBox1.Items.Add(dr.Item(1))

        End While

        l_selected_all_most_frequent = False

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim i As Integer = 0

        Dim indexChecked As Integer

        Dim total_selected_exams As Integer = 0

        For Each indexChecked In CheckedListBox1.CheckedIndices

            total_selected_exams = total_selected_exams + 1

        Next

        ReDim l_selected_exam(total_selected_exams - 1)

        For Each indexChecked In CheckedListBox1.CheckedIndices

            Dim dr As OracleDataReader = db_acces.GET_FREQ_EXAM(l_selected_soft, l_selected_dep_clin_serv, TextBox1.Text, oradb)

            Dim i_index As Integer = 0

            While dr.Read()

                If i_index = indexChecked.ToString() Then

                    l_selected_exam(i) = dr.Item(0)

                End If

                i_index = i_index + 1

            End While

            i = i + 1
        Next


        If db_acces.DELETE_EXAMS_DEP_CLIN_SERV(l_selected_exam, l_selected_dep_clin_serv, oradb) Then

            MsgBox("Record(s) Deleted")

            CheckedListBox1.Items.Clear()

            l_selected_dep_clin_serv = db_acces.GET_SELECTED_DEP_CLIN_SERV(TextBox1.Text, l_selected_soft, ComboBox3.SelectedIndex, oradb)

            Dim dr_new As OracleDataReader = db_acces.GET_FREQ_EXAM(l_selected_soft, l_selected_dep_clin_serv, TextBox1.Text, oradb)

            Dim i_new As Integer = 0

            While dr_new.Read()

                CheckedListBox1.Items.Add(dr_new.Item(1))

            End While

        Else

            MsgBox("ERROR!")

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If CheckedListBox1.Items.Count() > 0 Then

            If l_selected_all_most_frequent = False Then

                For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                    CheckedListBox1.SetItemChecked(i, True)
                Next

                l_selected_all_most_frequent = True

            Else

                For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                    CheckedListBox1.SetItemChecked(i, False)
                Next

                l_selected_all_most_frequent = False

            End If

        End If

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

        Try

            CheckedListBox2.Items.Clear()

            Dim l_exam_cat(l_total_cats)

            l_exam_cat(0) = 0 ''Referente ao all

            Dim dr_exam_cat As OracleDataReader = db_acces.GET_EXAMS_CAT(TextBox1.Text, l_selected_soft, oradb)

            Dim i_cats As Integer = 1

            While dr_exam_cat.Read()

                l_exam_cat(i_cats) = dr_exam_cat.Item(1)
                i_cats = i_cats + 1
            End While

            Dim dr As OracleDataReader = db_acces.GET_EXAMS(TextBox1.Text, l_selected_soft, l_exam_cat(ComboBox4.SelectedIndex), oradb)

            Dim i As Integer = 0

            While dr.Read()

                CheckedListBox2.Items.Add(dr.Item(0))

            End While

        Catch ex As Exception

            MsgBox("Error selecting exams - GET_EXAMS", MsgBoxStyle.Critical)

        End Try


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim form1 As New Form1

        form1.Show()

        Me.Close()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        If CheckedListBox2.CheckedIndices.Count() > 0 Then

            Dim result As Integer = 0

            If (CheckedListBox2.CheckedIndices.Count = CheckedListBox2.Items.Count()) Then

                result = MsgBox("All records from the chosen category will be deleted! Confirm?", MessageBoxButtons.YesNo)

            End If


            If (result = DialogResult.Yes Or CheckedListBox2.CheckedIndices.Count < CheckedListBox2.Items.Count()) Then

                Dim indexChecked As Integer

                Dim total_selected_exams As Integer = 0

                For Each indexChecked In CheckedListBox2.CheckedIndices

                    total_selected_exams = total_selected_exams + 1

                Next

                ReDim l_selected_exam(total_selected_exams - 1)

                ''Determinar ID_EXAM
                '' 1 - Determinar a categoria selecionada
                '' 2 - Fazer um search a todos os exames da cat selecionada
                '' 3 - Ecolher os ids dos exames selecionados
                '' 4 - Apagar os exames selecionados
                '' 5 - Refresh à grid de exames

                '1

                Dim l_index_cat As Integer = ComboBox4.SelectedIndex
                Dim l_id_cat_exam As Int64 = 0

                Dim dr_exam_cat As OracleDataReader

                Dim i_index As Integer = 0

                Try

                    dr_exam_cat = db_acces.GET_EXAMS_CAT(TextBox1.Text, l_selected_soft, oradb)

                    While dr_exam_cat.Read()

                        If l_index_cat = 0 Then

                            l_id_cat_exam = 0
                            Exit While

                        ElseIf i_index = l_index_cat - 1 Then

                            l_id_cat_exam = dr_exam_cat.Item(1)
                            Exit While

                        End If

                        i_index = i_index + 1

                    End While

                Catch ex As Exception

                    MsgBox("ERROR GETTING EXAM CATEGORY - Button5_Click", vbCritical)

                End Try

                '2 e 3

                Dim l_array_exams(CheckedListBox2.Items.Count() - 1) As Int64

                Dim dr_exams As OracleDataReader

                Dim l_array_selected_exams(CheckedListBox2.CheckedIndices.Count() - 1) As Int64 ''Array que vai guardar o id dos exames selecionados


                Try

                    'Lista de exames de categoria selecionada
                    dr_exams = db_acces.GET_EXAMS(TextBox1.Text, l_selected_soft, l_id_cat_exam, oradb)

                    'Lista de indexes de exames selecionados
                    Dim l_array_selected_indexes(CheckedListBox2.CheckedIndices.Count()) As Integer
                    Dim i_index_checked_aux As Integer = 0

                    For Each indexChecked In CheckedListBox2.CheckedIndices

                        l_array_selected_indexes(i_index_checked_aux) = indexChecked.ToString()

                        i_index_checked_aux = i_index_checked_aux + 1

                    Next

                    ''Lista de exames selecionados - ERRO
                    i_index_checked_aux = 0
                    Dim i_selected_exams_aux As Int16 = 0

                    Dim l_index_selected_exams As Integer = 0

                    If (CheckedListBox2.CheckedIndices.Count() > 0) Then

                        While dr_exams.Read() ''Ler todos os exames da categoria selecionada

                            For ii As Integer = 0 To (CheckedListBox2.CheckedIndices.Count() - 1)

                                If (l_array_selected_indexes(ii) = i_selected_exams_aux) Then

                                    l_array_selected_exams(l_index_selected_exams) = dr_exams.Item(2)
                                    l_index_selected_exams = l_index_selected_exams + 1

                                End If

                            Next

                            i_selected_exams_aux = i_selected_exams_aux + 1

                        End While

                    End If


                Catch ex As Exception

                    MsgBox("ERROR GETTING SELECTED EXAMS - Button5_Click", vbCritical)

                End Try

                '4
                Try
                    If db_acces.DELETE_EXAMS(l_array_selected_exams, TextBox1.Text, l_selected_soft, oradb) Then

                        MsgBox("Record(s) deleted!")

                    Else

                        MsgBox("No records deleted.")

                    End If
                Catch ex As Exception

                    MsgBox("ERROR DELETING EXAMS - Button5_Click", vbCritical)

                End Try

                '5

                Try
                    CheckedListBox1.Items.Clear()
                    CheckedListBox2.Items.Clear()
                    ComboBox3.SelectedItem = ""


                    Dim dr_exams_cat As OracleDataReader = db_acces.GET_EXAMS(TextBox1.Text, l_selected_soft, l_id_cat_exam, oradb)

                    Dim i As Integer = 0

                    While dr_exams_cat.Read()

                        CheckedListBox2.Items.Add(dr_exams_cat.Item(0))

                    End While

                Catch ex As Exception

                    MsgBox("ERROR GETTING EXAMS BY CATEGORY - Button5_Click", vbCritical)

                End Try

                If ((result = DialogResult.Yes)) Then

                    ComboBox4.Items.Clear()
                    ComboBox4.Text = ""

                    Try

                        Dim dr_exam_cat_new As OracleDataReader = db_acces.GET_EXAMS_CAT(TextBox1.Text, l_selected_soft, oradb)

                        ComboBox4.Items.Add("ALL")

                        While dr_exam_cat_new.Read()

                            ComboBox4.Items.Add(dr_exam_cat_new.Item(0))
                            l_total_cats = l_total_cats + 1

                        End While

                    Catch ex As Exception

                        MsgBox("ERROR LOADING EXAMS CATEGORIES - Button5_Click", MsgBoxStyle.Critical)

                    End Try

                End If

            End If

            Else

            MsgBox("No selected records!")

        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        If CheckedListBox2.Items.Count() > 0 Then

            If l_selected_all = False Then

                For i As Integer = 0 To CheckedListBox2.Items.Count - 1
                    CheckedListBox2.SetItemChecked(i, True)
                Next

                l_selected_all = True

            Else

                For i As Integer = 0 To CheckedListBox2.Items.Count - 1
                    CheckedListBox2.SetItemChecked(i, False)
                Next

                l_selected_all = False

            End If

        End If

    End Sub
End Class
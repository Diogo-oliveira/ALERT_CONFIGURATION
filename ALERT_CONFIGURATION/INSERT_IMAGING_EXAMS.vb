Imports Oracle.DataAccess.Client
Public Class INSERT_IMAGING_EXAMS

    Dim db_access As New EXAMS_API
    Dim oradb As String = "Data Source=QC4V26522;User Id=alert_config;Password=qcteam"
    Dim l_selected_soft As Int16 = -1
    Dim l_selected_category As String = ""

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

        CheckedListBox2.Items.Add(l_selected_category) 'APAGAR



    End Sub

    Private Sub CheckedListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox2.SelectedIndexChanged

    End Sub
End Class
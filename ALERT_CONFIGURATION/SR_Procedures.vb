Imports Oracle.DataAccess.Client
Public Class SR_Procedures

    Dim db_access_general As New General
    Dim oradb As String
    Dim conn As New OracleConnection

    Dim db_sr_procedure As New SR_PROCEDURES_API

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

        If TextBox1.Text <> "" Then

            ComboBox1.Text = db_access_general.GET_INSTITUTION(TextBox1.Text, conn)

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

            ' g_selected_category = ""

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

        'Definir o tipo de registo
        ComboBox7.Items.Add("Surgical Procedures")
        ComboBox7.Items.Add("Kits")
        ComboBox7.Items.Add("Positionings")

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged

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

        Dim dr_def_versions As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_sr_procedure.GET_DEFAULT_VERSIONS(TextBox1.Text, ComboBox7.SelectedIndex, conn, dr_def_versions) Then


#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            MsgBox("ERROR LOADING DEFAULT VERSIONS -  ComboBox2_SelectedIndexChanged", MsgBoxStyle.Critical)
            dr_def_versions.Dispose()
            dr_def_versions.Close()

        Else
            While dr_def_versions.Read()

                ComboBox3.Items.Add(dr_def_versions.Item(0))

            End While

            dr_def_versions.Dispose()
            dr_def_versions.Close()

        End If


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

    End Sub
End Class
Imports Oracle.DataAccess.Client
Public Class SR_Procedures

    Dim db_access_general As New General
    Dim oradb As String
    Dim conn As New OracleConnection

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
                CheckedListBox3.Items.Clear()

                ComboBox6.Text = ""
                ComboBox6.Items.Clear()
                CheckedListBox4.Items.Clear()

                ' g_selected_category = ""

            End If

            dr.Dispose()
            dr.Close()

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

        CheckBox1.Checked = True
        CheckBox2.Checked = True
        ' g_procedure_type = 0

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

    End Sub
End Class
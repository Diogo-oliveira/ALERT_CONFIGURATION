Imports Oracle.DataAccess.Client

Public Class Supplies

    Dim db_access_general As New General
    Dim oradb As String
    Dim conn As New OracleConnection

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

    Public Sub New(ByVal i_oradb As String)

        InitializeComponent()
        oradb = i_oradb
        conn = New OracleConnection(oradb)

    End Sub

    Private Sub Supplies_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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

        ComboBox7.Items.Add(g_activity_desc)
        ComboBox7.Items.Add(g_implants_desc)
        ComboBox7.Items.Add(g_kits_desc)
        ComboBox7.Items.Add(g_sets_desc)
        ComboBox7.Items.Add(g_supplies_desc)
        ComboBox7.Items.Add(g_surgical_desc)

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

                'g_selected_category = ""

            End If

            dr.Dispose()
            dr.Close()

        End If

        Cursor = Cursors.Arrow
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged

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

        End If

    End Sub
End Class
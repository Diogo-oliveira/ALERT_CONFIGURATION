Imports Oracle.DataAccess.Client
Public Class MED_STD_NON_IV

    Dim db_access_general As New General
    Dim medication As New Medication_API

    Dim g_id_product As String = ""
    Dim g_id_product_supplier As String = ""
    Dim g_id_institution As Int64 = 0
    Dim g_id_software_index As Int16 = -1 ''index do software
    Dim g_selected_software As Int16 = -1
    Dim g_default_route As Int64 = -1

    Dim g_a_med_set_instructions() As Medication_API.MED_SET_INSTRUCTIONS
    Dim g_a_admin_methods() As Int64
    Dim g_a_admin_sites() As Int64

    Public Sub New(ByVal i_institution As Int64, ByVal i_software_index As Int16, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_default_route As Int64)

        InitializeComponent()
        g_id_product = i_id_product
        g_id_product_supplier = i_id_product_supplier
        g_id_institution = i_institution
        g_id_software_index = i_software_index
        g_default_route = i_default_route

    End Sub

    Private Sub MED_STD_NON_IV_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "MEDICATION - STANDARD INSTRUCTIONS ::::  Connected to " & Connection.db

        Me.BackColor = Color.FromArgb(215, 215, 180)

        TextBox2.Text = medication.GET_PRODUCT_DESC(g_id_institution, g_id_product, g_id_product_supplier)

        Dim dr_soft As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_SOFT_INST(g_id_institution, dr_soft) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING SOFTWARES!", vbCritical)
        Else
            While dr_soft.Read()
                ComboBox2.Items.Add(dr_soft.Item(1))
            End While
        End If

        dr_soft.Dispose()
        dr_soft.Close()

        ComboBox2.SelectedIndex = g_id_software_index
        g_selected_software = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, g_id_institution)


        Dim dr_freq As OracleDataReader

        If Not medication.GET_ALL_FREQS(g_id_institution, g_selected_software, dr_freq) Then
            MsgBox("Error getting all frequencies")
        Else
            While dr_freq.Read()
                ComboBox3.Items.Add(dr_freq.Item(1))
            End While
        End If

        ComboBox28.Items.Add("0 - ALL")
        ComboBox28.Items.Add("1 - External Prescription")
        ComboBox28.Items.Add("2 - Administer Here")
        ComboBox28.Items.Add("3 - Home Medication")

        ComboBox24.Items.Add("Y")
        ComboBox24.Items.Add("N")

    End Sub

    Private Sub ComboBox28_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox28.SelectedIndexChanged
        Dim dr_med_set_instruction As OracleDataReader
        ReDim g_a_med_set_instructions(0)
        ComboBox1.Items.Clear()
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not medication.GET_ALL_INSTRUCTIONS(g_id_institution, g_selected_software, g_id_product, g_id_product_supplier, ComboBox28.SelectedIndex, dr_med_set_instruction) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING LIST OF STANDARD INSTRUCTIONS!", vbCritical)
        Else
            Dim i As Integer = 0
            While dr_med_set_instruction.Read()
                ReDim Preserve g_a_med_set_instructions(i)
                g_a_med_set_instructions(i).id_product = dr_med_set_instruction.Item(0)
                g_a_med_set_instructions(i).id_std_presc_dir = dr_med_set_instruction.Item(1)
                g_a_med_set_instructions(i).rank = dr_med_set_instruction.Item(2)
                g_a_med_set_instructions(i).id_grant = dr_med_set_instruction.Item(3)
                g_a_med_set_instructions(i).market = dr_med_set_instruction.Item(4)
                g_a_med_set_instructions(i).market_desc = dr_med_set_instruction.Item(5)
                g_a_med_set_instructions(i).software = dr_med_set_instruction.Item(6)
                g_a_med_set_instructions(i).software_desc = dr_med_set_instruction.Item(7)
                g_a_med_set_instructions(i).id_pick_list = dr_med_set_instruction.Item(8)
                g_a_med_set_instructions(i).institution = dr_med_set_instruction.Item(9)

                ComboBox1.Items.Add(g_a_med_set_instructions(i).rank)

                i = i + 1

            End While
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox27.Text = g_a_med_set_instructions(ComboBox1.SelectedIndex).id_grant

        Dim dr_std_presc_dir As OracleDataReader
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not medication.GET_STD_PRESC_DIR(g_id_institution, g_a_med_set_instructions(ComboBox1.SelectedIndex).id_std_presc_dir, dr_std_presc_dir) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING STANDARD_PRESC_DIR!", vbCritical)
        Else
            While dr_std_presc_dir.Read()
                ComboBox24.Text = dr_std_presc_dir.Item(1)
                Try
                    TextBox24.Text = dr_std_presc_dir.Item(3)
                Catch ex As Exception
                    TextBox24.Text = ""
                End Try

                Try
                    ComboBox26 = dr_std_presc_dir.Item(4)
                Catch ex As Exception
                    ComboBox26.Text = ""
                End Try

                Try
                    ComboBox27 = dr_std_presc_dir.Item(5)
                Catch ex As Exception
                    ComboBox27.Text = ""
                End Try

            End While
        End If

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex > -1 And ComboBox2.Text <> "" Then
            g_selected_software = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, g_id_institution)

            Dim l_dr_sos As OracleDataReader
            If Not medication.GET_SOS_LIST(g_id_institution, g_selected_software, l_dr_sos) Then
                MsgBox("Error geting list of SOS reasons.", vbCritical)
            End If

            While l_dr_sos.Read()
                ComboBox25.Items.Add(l_dr_sos(1))
            End While

            Dim l_dr_admin_method As OracleDataReader
            If Not medication.GET_ADMIN_METHOD_LIST(g_id_institution, g_default_route, g_id_product_supplier, l_dr_admin_method) Then
                MsgBox("Error getting list of administration methods!", vbCritical)
            End If

            ReDim g_a_admin_methods(0)
            Dim i As Integer = 0
            While l_dr_admin_method.Read()
                ComboBox27.Items.Add(l_dr_admin_method.Item(1))
                ReDim Preserve g_a_admin_methods(i)
                g_a_admin_methods(i) = l_dr_admin_method(0)
                i = i + 1
            End While

            Dim l_dr_admin_sites As OracleDataReader
            If Not medication.GET_ADMIN_SITE_LIST(g_id_institution, g_default_route, g_id_product_supplier, l_dr_admin_sites) Then
                MsgBox("Error getting list of administration sotes!", vbCritical)
            End If

            ReDim g_a_admin_sites(0)
            Dim ii As Integer = 0
            While l_dr_admin_sites.Read()
                ComboBox26.Items.Add(l_dr_admin_sites.Item(1))
                ReDim Preserve g_a_admin_sites(ii)
                g_a_admin_sites(ii) = g_a_admin_sites(0)
                ii = ii + 1
            End While

        End If
    End Sub

    Private Sub ComboBox25_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox25.SelectedIndexChanged
        TextBox24.Text = ""
    End Sub

    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged
        ComboBox25.SelectedIndex = -1
    End Sub
End Class
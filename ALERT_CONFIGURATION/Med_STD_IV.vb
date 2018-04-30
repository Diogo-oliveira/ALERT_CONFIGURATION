Imports Oracle.DataAccess.Client
Public Class Med_STD_IV

    Dim db_access_general As New General
    Dim medication As New Medication_API

    Dim g_id_product As String = ""
    Dim g_id_product_supplier As String = ""
    Dim g_id_institution As Int64 = 0
    Dim g_id_software_index As Int16 = -1 ''index do software
    Dim g_selected_software As Int16 = -1
    Dim g_default_route As String = -1
    Dim g_id_market As Int16 = -1

    Dim g_a_frequencies() As Int64
    Dim g_a_product_um() As Int64
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

    Private Sub Med_STD_IV_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "MEDICATION - STANDARD INSTRUCTIONS ::::  Connected to " & Connection.db

        Me.BackColor = Color.FromArgb(215, 215, 180)

        TextBox2.Text = medication.GET_PRODUCT_DESC(g_id_institution, g_id_product, g_id_product_supplier)
        g_id_market = db_access_general.GET_INSTITUTION_MARKET(g_id_institution)

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
        ReDim g_a_frequencies(0)
        Dim l_index_freq As Int16 = 0
        If Not medication.GET_ALL_FREQS(g_id_institution, g_selected_software, dr_freq) Then
            MsgBox("Error getting all frequencies")
        Else
            ComboBox3.Items.Add("")
            While dr_freq.Read()
                ComboBox3.Items.Add(dr_freq.Item(1))

                ReDim Preserve g_a_frequencies(l_index_freq)
                g_a_frequencies(l_index_freq) = dr_freq.Item(0)
                l_index_freq = l_index_freq + 1
            End While
        End If

        ComboBox28.Items.Add("0 - ALL")
        ComboBox28.Items.Add("1 - External Prescription")
        ComboBox28.Items.Add("2 - Administer Here")
        ComboBox28.Items.Add("3 - Home Medication")

        ComboBox24.Items.Add("Y")
        ComboBox24.Items.Add("N")

        Dim l_dr_product_um As OracleDataReader
        Dim i As Integer = 0
        If Not medication.GET_PRODUCT_UM(g_id_institution, g_id_product, g_id_product_supplier, 1, l_dr_product_um) Then
            MsgBox("Error getting product unit measures!", vbCritical)
        Else

            ReDim g_a_product_um(0)

            While l_dr_product_um.Read()
                If i = 0 Then
                    ComboBox4.Items.Add("")
                    ComboBox6.Items.Add("")
                    ComboBox7.Items.Add("")
                    ComboBox8.Items.Add("")
                    ComboBox9.Items.Add("")
                    ComboBox10.Items.Add("")
                    ComboBox11.Items.Add("")
                End If
                ComboBox4.Items.Add(l_dr_product_um.Item(1))
                ComboBox6.Items.Add(l_dr_product_um.Item(1))
                ComboBox7.Items.Add(l_dr_product_um.Item(1))
                ComboBox8.Items.Add(l_dr_product_um.Item(1))
                ComboBox9.Items.Add(l_dr_product_um.Item(1))
                ComboBox10.Items.Add(l_dr_product_um.Item(1))
                ComboBox11.Items.Add(l_dr_product_um.Item(1))
                ReDim Preserve g_a_product_um(i)
                g_a_product_um(i) = l_dr_product_um(0)
                i = i + 1
            End While
            l_dr_product_um.Close()
        End If

        Dim l_dr_admin_method As OracleDataReader
        If Not medication.GET_ADMIN_METHOD_LIST(g_id_institution, g_default_route, g_id_product_supplier, l_dr_admin_method) Then
            MsgBox("Error getting list of administration methods!", vbCritical)
        End If

        ReDim g_a_admin_methods(0)
        i = 0
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
            g_a_admin_sites(ii) = l_dr_admin_sites.Item(0)
            ii = ii + 1
        End While

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox2.SelectedIndex > -1 Then
            Dim l_id_grant As Int64
            l_id_grant = medication.GET_ID_GRANT(g_id_institution, g_selected_software, "LNK_PRODUCT_STD_PRESC_DIR")

            'SE GRANT FOR = -1 ENTÃO É NECESSÁRIO CRIAR UM NOVO GRANT
            If l_id_grant = -1 Then
                If Not medication.SET_ID_GRANT(g_id_institution, g_selected_software, "LNK_PRODUCT_STD_PRESC_DIR") Then
                    MsgBox("Error creating ID_GRANT!", vbCritical)
                Else
                    l_id_grant = medication.GET_ID_GRANT(g_id_institution, g_selected_software, "LNK_PRODUCT_STD_PRESC_DIR")
                End If
            End If

            TextBox27.Text = l_id_grant

        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex > -1 And ComboBox2.Text <> "" Then

            ComboBox25.Items.Clear()

            g_selected_software = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, g_id_institution)

            Dim l_dr_sos As OracleDataReader
            If Not medication.GET_SOS_LIST(g_id_institution, g_selected_software, l_dr_sos) Then
                MsgBox("Error geting list of SOS reasons.", vbCritical)
            End If

            While l_dr_sos.Read()
                ComboBox25.Items.Add(l_dr_sos(1))
            End While

        End If
    End Sub
End Class
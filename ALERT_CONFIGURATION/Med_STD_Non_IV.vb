Imports Oracle.DataAccess.Client
Public Class MED_STD_NON_IV

    Dim db_access_general As New General
    Dim medication As New Medication_API

    Dim g_id_product As String = ""
    Dim g_id_product_supplier As String = ""
    Dim g_id_institution As Int64 = 0
    Dim g_id_software_index As Int16 = -1 ''index do software
    Dim g_selected_software As Int16 = -1

    Public Sub New(ByVal i_institution As Int64, ByVal i_software_index As Int16, ByVal i_id_product As String, ByVal i_id_product_supplier As String)

        InitializeComponent()
        g_id_product = i_id_product
        g_id_product_supplier = i_id_product_supplier
        g_id_institution = i_institution
        g_id_software_index = i_software_index

    End Sub

    Private Sub MED_STD_NON_IV_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

    End Sub
End Class
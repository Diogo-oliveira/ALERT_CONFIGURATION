Public Class YES_to_ALL
    'Yes - Yes
    'No - No
    'Yest to All - OK
    'No to All - Abort

    Dim g_desc_interv As String = ""

    Public Sub New(ByVal i_desc_interv As String)

        InitializeComponent()
        g_desc_interv = i_desc_interv

    End Sub


    Private Sub YES_to_ALL_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Label1.Text = g_desc_interv
        '"Record '" & g_desc_interv & "' exists for software 'ALL'. If you delete this record, it will also be deleted for all softwares. Confirm?"

    End Sub


End Class
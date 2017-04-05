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

        Me.Location = New Point(Form_location.x_position, Form_location.y_position)

        Me.BackColor = Color.FromArgb(215, 215, 180)

        Label1.Text = g_desc_interv

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Cursor = Cursors.WaitCursor
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Cursor = Cursors.WaitCursor
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Cursor = Cursors.WaitCursor
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Cursor = Cursors.WaitCursor
    End Sub

End Class
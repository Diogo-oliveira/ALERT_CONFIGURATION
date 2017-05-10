Public Class CONTENT_DISPLAY

    Dim g_instr_title As String = ""
    Dim g_instr As String = ""

    Public Sub New(ByVal i_instr_title As String, ByVal i_instr As String)

        InitializeComponent()
        g_instr_title = i_instr_title
        g_instr = i_instr

    End Sub

    Private Sub CONTENT_DISPLAY_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Location = New Point(Form_location.x_position, Form_location.y_position)

        Me.CenterToScreen()

        Me.BackColor = Color.FromArgb(215, 215, 180)

        TextBox1.BackColor = Color.FromArgb(215, 215, 180)

        Label1.Text = g_instr_title

        TextBox1.ReadOnly = True
        TextBox1.Multiline = True
        ' TextBox1.BorderStyle = BorderStyle.None
        TextBox1.ScrollBars = ScrollBars.Vertical
        TextBox1.Height = 280

        'Converter o chr(10) para parágrafo
        TextBox1.Text = g_instr.Replace(Chr(10), Environment.NewLine)

    End Sub

End Class
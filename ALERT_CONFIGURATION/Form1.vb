Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Application.Exit()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim FORM_ADD_IMAGING As New INSERT_IMAGING_EXAMS

        Me.Hide()

        FORM_ADD_IMAGING.ShowDialog()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim Form_SYS_CONFIG As New SYS_CONFIGS

        Me.Hide()

        Form_SYS_CONFIG.ShowDialog()

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim Form_Remove_Other As New REMOVE_OTHER_EXAMS

        Me.Hide()

        Form_Remove_Other.ShowDialog()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim form_insert_other As New INSERT_OTHER_EXAM

        Me.Hide()

        form_insert_other.ShowDialog()

    End Sub
End Class

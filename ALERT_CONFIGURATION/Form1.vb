Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim FORM_REMOVE_IMAGING As New REMOVE_IMAGING

        Me.Hide()

        FORM_REMOVE_IMAGING.ShowDialog()

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
End Class

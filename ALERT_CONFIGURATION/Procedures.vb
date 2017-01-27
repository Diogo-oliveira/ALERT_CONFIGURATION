Public Class Procedures
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Dim form1 As New Form1

        Me.Enabled = False

        Me.Dispose()

        form1.Show()

    End Sub
End Class
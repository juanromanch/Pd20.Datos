Public Class Acceder
    Dim enc As New Pd20.Encriptador("esta es la clave")


    Private Sub OK_Click(sender As Object, e As EventArgs) Handles OK.Click
        'PasswordTextBox.Text = enc.Encriptar(UsernameTextBox.Text)
        Dim Rejilla As New frmRejilla("sa", "A3erp4848")
        Rejilla.Show()





    End Sub

    Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles Cancel.Click
        'PasswordTextBox.Text = enc.Desencriptar(UsernameTextBox.Text)
    End Sub

End Class
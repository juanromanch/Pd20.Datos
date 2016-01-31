Public Class frmRegilla
    Public conecta As New Pd20.Datos.SQLServer("AJTIENDA\A3ERP", "JR")
    Public usu As String
    Public pass As String
    Public Sub New(u As String, c As String)

        ' Esta llamada es exigida por el diseñador.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        usu = u
        pass = c
        conecta.Ususario = usu
        conecta.Password = pass
        ' conecta.Base = "JR"


    End Sub

    Private Sub ButtonX1_Click(sender As Object, e As EventArgs) Handles ButtonX1.Click
        Dim dt As DataTable
        lCadena.Text = conecta.CadenaConexion
        dt = conecta.ObtenerDataTableConsulta("SELECT NOMCLI AS CLIENTE, DIRCLI1 AS DIRECCION FROM CLIENTES")
        Me.DataGridView1.DataSource = dt


    End Sub
End Class
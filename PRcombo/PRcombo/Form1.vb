Imports System.Data.OleDb
Public Class Form1
    Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\casa\Desktop\Base de datos\sistema.mdb")
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            con.Open()
            MessageBox.Show("Exito al conectar")
        Catch ex As Exception
            MessageBox.Show("error al iniciar", "Mensaje del sistema", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        loadingCombo()
        loadinglistbox()
    End Sub

    Public Sub loadingCombo()
        Dim table As New DataTable
        Dim sql As String = "SELECT IdDepartamento,nombre FROM departamentos"
        Dim adp As New OleDbDataAdapter(sql, con)

        adp.Fill(table)

        ComboBox1.DataSource = table
        ComboBox1.DisplayMember = "nombre"
        ComboBox1.ValueMember = "IdDepartamento"


    End Sub
    Public Sub loadinglistbox()
        Dim table As New DataTable
        Dim sql As String = "SELECT IdDepartamento, nombre FROM departamentos"
        Dim adp As New OleDbDataAdapter(sql, con)
        adp.Fill(table)

        ListBox1.DataSource = table
        ListBox1.DisplayMember = "nombre"
        ListBox1.ValueMember = "IdDepartamento"


    End Sub
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

        'Dim cod As Integer = Val(ListBox1.SelectedValue.ToString)


        'Dim table As New DataTable
        'Dim sql As String = "SELECT nombrep FROM Rprovincias WHERE Idregion = " & cod
        'Dim adp As New OleDbDataAdapter(sql, con)
        'adp.Fill(table)

        'ListBox2.DataSource = table
        'ListBox2.DisplayMember = "nombrep"


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
       

    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Dim cod As Integer = Val(ComboBox1.SelectedValue.ToString)



        Dim table As New DataTable
        Dim sql As String = "SELECT nombre FROM distritos WHERE IdDepartamento = " & cod
        Dim adp As New OleDbDataAdapter(sql, con)
        adp.Fill(table)

        ComboBox2.DataSource = table
        ComboBox2.DisplayMember = "nombre"
    End Sub

    
    Private Sub ListBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedValueChanged

        Dim cod As Integer = Val(ListBox1.SelectedValue.ToString)


        Dim table As New DataTable
        Dim sql As String = "SELECT nombre FROM distritos WHERE IdDepartamento = " & cod
        Dim adp As New OleDbDataAdapter(sql, con)
        adp.Fill(table)

        ListBox2.DataSource = table
        ListBox2.DisplayMember = "nombre"
    End Sub
End Class

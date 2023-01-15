Imports MySql.Data.MySqlClient

Module Module1

    Dim con As MySqlConnection
    Dim db As MySqlDataAdapter
    Dim ds As DataSet

    Public Sub koneksi()
        con = New MySqlConnection("server=localhost; user=root;" + "password=; database=prpenjualan;")
    End Sub

    Public Sub tampil(ByVal a As String, ByVal b As DataGridView)

        koneksi()
        db = New MySqlDataAdapter(a, con)
        ds = New DataSet
        db.Fill(ds)
        b.DataSource = ds
        b.DataMember = ds.Tables(0).ToString

    End Sub

    Public Sub teks(ByVal a As String)
        koneksi()
        db = New MySqlDataAdapter(a, con)
        ds = New DataSet
        db.Fill(ds)
    End Sub

    Public Sub kolom(ByVal a As TextBox, ByVal b As String)
        a.Text = ds.Tables(0).Rows(0)(b).ToString
    End Sub


End Module

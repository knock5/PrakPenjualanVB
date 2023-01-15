Public Class Form2

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.MdiParent = Form3
        ComboBox1.Items.Add("barang")
        ComboBox1.Items.Add("beli")
        ComboBox1.Items.Add("jual")

        noJual()
        noJualPlus()
        tampil("SELECT * FROM barang", DataGridView1)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "barang" Then
            tampil("SELECT * FROM barang", DataGridView1)
        ElseIf ComboBox1.Text = "beli" Then
            tampil("SELECT * FROM beli", DataGridView1)
        ElseIf ComboBox1.Text = "jual" Then
            tampil("SELECT * FROM jual", DataGridView1)
        End If
    End Sub

    Sub kosong()
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
    End Sub

    Sub noJual()
        tampil("SELECT MAX(kdjual)+1 as no from jual", DataGridView1)
        kolom(TextBox2, "no")
        If TextBox2.Text = "" Then
            TextBox2.Text = 1
        End If
    End Sub
    Sub noJualPlus()
        teks("SELECT MAX(noref)+1 as no from jual")
        kolom(TextBox1, "no")
        If TextBox1.Text = "" Then
            TextBox1.Text = 1
        End If
    End Sub
    Sub noJualReset()
        teks("SELECT MAX(noref) as no from jual")
        kolom(TextBox1, "no")
        If TextBox1.Text = "" Then
            TextBox1.Text = 1
        End If
    End Sub

    Sub vjual()
        tampil("SELECT jual.noref, jual.kodebarang, barang.namabarang, jual.jmlbrgjual, jual.tothrg from barang, jual where barang.kodebarang=jual.kodebarang and noref=" + TextBox1.Text + "", DataGridView1)
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        Try
            teks("SELECT * FROM barang where kodebarang=" + TextBox3.Text + "")
            kolom(TextBox4, "namabarang")
            kolom(TextBox5, "hargabarang")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        Try
            TextBox7.Text = TextBox5.Text * TextBox6.Text
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If TextBox3.Text = "" Then
                MsgBox("Silahkan lengkapi data!")
            ElseIf TextBox3.Text = "" Then
                MsgBox("Silahkan lengkapi data!")
            Else
                tampil("INSERT INTO jual (noref, kdjual, kodebarang, jmlbrgjual, tothrg) values('" + TextBox1.Text + "','" + TextBox2.Text + "','" + TextBox3.Text + "','" + TextBox6.Text + "','" + TextBox7.Text + "')", DataGridView1)
            End If
        Catch ex As Exception
        End Try

        Try
            teks("UPDATE barang set jumlahstok=(jumlahstok)-" + TextBox6.Text + " where kodebarang=" + TextBox3.Text + "")
            noJual()
        Catch ex As Exception
        End Try

        vjual()
        kosong()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        noJualPlus()
        vjual()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        noJualReset()
        vjual()
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            GroupBox1.Visible = False
            Button2.Visible = False
            Button3.Visible = False
            Button4.Visible = False
            Label8.Visible = False
            ComboBox1.Visible = False
            PrintForm1.PrintAction = Printing.PrintAction.PrintToPreview
            PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
        Catch ex As Exception
        End Try
        kosong()
        vjual()
        GroupBox1.Visible = True
        Button2.Visible = True
        Button3.Visible = True
        Button4.Visible = True
        Label8.Visible = True
        ComboBox1.Visible = True
    End Sub
End Class
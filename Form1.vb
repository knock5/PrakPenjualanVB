Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.MdiParent = Form3
        ComboBox1.Items.Add("barang")
        ComboBox1.Items.Add("beli")
        ComboBox1.Items.Add("jual")

        noBeli()
        noBeliPlus()
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

    Sub noBeli()
        tampil("SELECT MAX(kodebeli)+1 as no from beli", DataGridView1)
        kolom(TextBox2, "no")
        If TextBox2.Text = "" Then
            TextBox2.Text = 1
        End If
    End Sub
    Sub noBeliPlus()
        teks("SELECT MAX(noref)+1 as no from beli")
        kolom(TextBox1, "no")
        If TextBox1.Text = "" Then
            TextBox1.Text = 1
        End If
    End Sub
    Sub noBeliReset()
        teks("SELECT MAX(noref) as no from beli")
        kolom(TextBox1, "no")
        If TextBox1.Text = "" Then
            TextBox1.Text = 1
        End If
    End Sub

    Sub vbeli()
        tampil("SELECT beli.noref, beli.kodebarang, barang.namabarang, beli.jmlbrgbeli, beli.tothrgbeli from barang, beli where barang.kodebarang=beli.kodebarang and noref=" + TextBox1.Text + "", DataGridView1)
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
            ElseIf TextBox6.Text = "" Then
                MsgBox("Silahkan lengkapi data!")
            Else
                teks("INSERT INTO beli (noref, kodebeli, kodebarang, jmlbrgbeli, tothrgbeli) values('" + TextBox1.Text + "','" + TextBox2.Text + "','" + TextBox3.Text + "','" + TextBox6.Text + "','" + TextBox7.Text + "')")
            End If
        Catch ex As Exception
        End Try

        Try
            teks("UPDATE barang set jumlahstok=(jumlahstok)+" + TextBox6.Text + " where kodebarang=" + TextBox3.Text + "")
            noBeli()
        Catch ex As Exception
        End Try

        vbeli()
        kosong()
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        noBeliPlus()
        vbeli()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        noBeliReset()
        vbeli()
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
        vbeli()
        GroupBox1.Visible = True
        Button2.Visible = True
        Button3.Visible = True
        Button4.Visible = True
        Label8.Visible = True
        ComboBox1.Visible = True
    End Sub


End Class

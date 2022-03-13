Imports System.Data.Odbc
Public Class Form1
    Dim CONN As OdbcConnection
    Dim CMD As OdbcCommand
    Dim DS As New DataSet
    Dim DA As OdbcDataAdapter
    Dim RD As OdbcDataReader
    Dim LokasiData As String

    Sub Koneksi()
        LokasiData = "Driver={MySQL ODBC 3.51 Driver};Database=siswa;server=localhost;uid=root"
        CONN = New OdbcConnection(LokasiData)
        If CONN.State = ConnectionState.Closed Then
            CONN.Open()
        End If
    End Sub

    Sub TampilGrid()
        Call Koneksi()
        DA = New OdbcDataAdapter("select * From datasiswa ", CONN)
        DS = New DataSet
        DA.Fill(DS, "datasiswa")
        DataGridView1.DataSource = DS.Tables("datasiswa")
        DataGridView1.ReadOnly = True
    End Sub

    Sub KosongkanData()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox1.Text = ""
    End Sub

    Sub MunculCombo()
        ComboBox1.Items.Add("RPL")
        ComboBox1.Items.Add("MM")
        ComboBox1.Items.Add("TBG")
        ComboBox1.Items.Add("TBU")
        ComboBox1.Items.Add("TBSM")
        ComboBox1.Items.Add("TKR")
        ComboBox1.Items.Add("TPM")
        ComboBox1.Items.Add("TPTL")
        ComboBox1.Items.Add("TITL")
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call TampilGrid()
        Call MunculCombo()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Silahkan Isi Semua Form")
        Else
            Call Koneksi()
            Dim simpan As String = "insert into datasiswa values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & ComboBox1.Text & "')"
            CMD = New OdbcCommand(simpan, CONN)
            CMD.ExecuteNonQuery()
            MsgBox("Input data berhasil")
            Call TampilGrid()
            Call KosongkanData()
        End If
    End Sub

    Private Sub TextBox1_KeyPress1(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 6
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            CMD = New OdbcCommand("Select * From datasiswa  where no_absen='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("No Absen Tidak Ada, Silahkan coba lagi!")
                TextBox1.Focus()
            Else
                TextBox2.Text = RD.Item("nama")
                TextBox3.Text = RD.Item("alamat")
                TextBox4.Text = RD.Item("no_hp")
                ComboBox1.Text = RD.Item("jurusan")
                TextBox2.Focus()
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call Koneksi()
        Dim edit As String = "update datasiswa set nama='" & TextBox2.Text & "',alamat='" & TextBox3.Text & "',no_hp='" & TextBox4.Text & "',jurusan='" & ComboBox1.Text & "' where no_absen='" & TextBox1.Text & "'"
        CMD = New OdbcCommand(edit, CONN)
        CMD.ExecuteNonQuery()
        MsgBox("Data Berhasil diUpdate")
        Call TampilGrid()
        Call KosongkanData()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan Pilih Data yang akan di hapus dengan Masukan no_absen dan ENTER")
        Else
            If MessageBox.Show("Yakin akan dihapus..?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call Koneksi()
                Dim hapus As String = "delete From datasiswa  where no_absen='" & TextBox1.Text & "'"
                CMD = New OdbcCommand(hapus, CONN)
                CMD.ExecuteNonQuery()
                Call TampilGrid()
                Call KosongkanData()
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        End
    End Sub
End Class
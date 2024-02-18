Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1

    ' Deklarasi variabel Excel
    Private excelApp As Excel.Application
    Private excelWorkbook As Excel.Workbook
    Private excelWorksheet As Excel.Worksheet

    'Fungsi untuk mengekspor data ke Excel
    Private Sub ExportToExcel()
        ' Inisialisasi Excel
        excelApp = New Excel.Application
        excelWorkbook = excelApp.Workbooks.Add()
        excelWorksheet = excelWorkbook.Sheets("Sheet1")

        ' Menulis data dari TextBox ke Excel
        excelWorksheet.Cells(1, 1).Value = "Nama"
        excelWorksheet.Cells(1, 2).Value = "Pekerjaan"
        excelWorksheet.Cells(1, 3).Value = "Gaji Pokok"
        excelWorksheet.Cells(1, 4).Value = "Tunjangan"
        excelWorksheet.Cells(1, 5).Value = "Bonus"
        excelWorksheet.Cells(1, 6).Value = "Lembur"
        excelWorksheet.Cells(1, 7).Value = "Potongan"
        excelWorksheet.Cells(1, 8).Value = "Total Gaji"

        ' Menulis data dari TextBox ke Excel
        excelWorksheet.Cells(2, 1).Value = TextBox1.Text
        excelWorksheet.Cells(2, 2).Value = TextBox2.Text
        excelWorksheet.Cells(2, 3).Value = TextBox3.Text
        excelWorksheet.Cells(2, 4).Value = TextBox4.Text
        excelWorksheet.Cells(2, 5).Value = TextBox5.Text
        excelWorksheet.Cells(2, 6).Value = TextBox6.Text
        excelWorksheet.Cells(2, 7).Value = TextBox7.Text
        excelWorksheet.Cells(2, 8).Value = TextBox8.Text

        ' Menyimpan Excel
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "Excel Files|*.xlsx"
        saveFileDialog1.Title = "Save an Excel File"
        saveFileDialog1.ShowDialog()

        If saveFileDialog1.FileName <> "" Then
            excelWorkbook.SaveAs(saveFileDialog1.FileName)
            excelWorkbook.Close()
            excelApp.Quit()
        End If
    End Sub

    ' Store Data
    Private dataMapping As New Dictionary(Of String, String) From {
        {"880001", "Raihan, IT, Rp.6.000.000"},
        {"880002", "Rahma, Akuntan, Rp.6.000.000"},
        {"880003", "Bagas, General Affair, Rp.6.200.000"},
        {"880004", "Yudi, Administrasi, Rp.5.400.000"},
        {"880005", "Ilham, Human Resources, Rp.6.200.000"}
    }

    ' Menampilkan dataMapping pada ComboBox1
    ' Membuat TextBox1,2,3 menjadi read only
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.AddRange(dataMapping.Keys.ToArray())

        TextBox1.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True
    End Sub

    ' Deklarasi TextBox
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        UpdateGajiTotal()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        UpdateGajiTotal()
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        UpdateGajiTotal()
    End Sub

    Private Sub TextBox7_TextChanged(Sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        UpdateGajiTotal()
    End Sub

    ' Fungsi perhitungan gaji
    Private Sub UpdateGajiTotal()
        ' Pastikan TextBox3 berisi angka sebelum melakukan pertambahan
        Dim gajiPokok As Integer
        If Integer.TryParse(TextBox3.Text.Replace("Rp.", "").Replace(".", ""), gajiPokok) Then
            ' Pastikan TextBox4 berisi angka
            Dim tambahan As Integer = If(Integer.TryParse(TextBox4.Text, tambahan), tambahan, 0)
            Dim tambahan1 As Integer = If(Integer.TryParse(TextBox5.Text, tambahan1), tambahan1, 0)
            Dim tambahan2 As Integer = If(Integer.TryParse(TextBox6.Text, tambahan2), tambahan2, 0)
            Dim tambahan3 As Integer = If(Integer.TryParse(TextBox7.Text, tambahan3), tambahan3, 0)

            ' Hitung total gaji dengan tambahan dari TextBox5, TextBox6, dan TextBox4
            Dim totalGaji As Integer = gajiPokok + tambahan + tambahan1 + tambahan2 - tambahan3
            TextBox8.Text = "Rp." & totalGaji.ToString("N0")

        Else
            TextBox8.Text = "Invalid Input"
        End If
    End Sub

    ' Menampilkan dataMapping pada ComboBox1 jika diklik untuk mengisi otomatis TextBox1,2,3
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim selectedValue As String = ComboBox1.SelectedItem.ToString()

        If dataMapping.ContainsKey(selectedValue) Then
            Dim values() As String = dataMapping(selectedValue).Split(", ")
            TextBox1.Text = values(0) ' Nama
            TextBox2.Text = values(1) ' Pekerjaan
            TextBox3.Text = values(2) ' Gaji Pokok
        End If

        TextBox1.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True
    End Sub

    ' Tombol reset TextBox untuk menghitung kembali
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
    End Sub

    ' Deklarasi MenuBar
    Private Sub MenuToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MenuToolStripMenuItem.Click

    End Sub

    ' Deklarasi pemanggilan fungi ExportToExcel pada sub menu Export to excel
    Private Sub ExportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToolStripMenuItem.Click
        ExportToExcel()
    End Sub

    ' Deklarasi fungsi penambahan data staff baru
    Private Sub TambahStaffToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TambahStaffToolStripMenuItem.Click
        ' Meminta pengguna untuk memasukan NIK staff baru
        Dim id As String = InputBox("Masukkan NIK staf baru:", "Tambah Staf")

        ' Periksa apakah pengguna memasukkan NIK atau membatalkan
        If Not String.IsNullOrEmpty(id) Then
            ' Periksa apakah NIK sudah ada dalam dataMapping
            If Not dataMapping.ContainsKey(id) Then
                ' Jika NIK belum ada, minta pengguna untuk memasukkan detail staf
                Dim detail As String = InputBox("Masukkan detail staf (Nama, Pekerjaan, Gaji):", "Tambah Staf")
                ' Tambahkan data baru ke dataMapping
                dataMapping.Add(id, detail)
                ' Perbarui ComboBox1
                ComboBox1.Items.Add(id)
                MessageBox.Show("Staf berhasil ditambahkan.", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("NIK staf sudah ada.", "Peringatan", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If
    End Sub

    ' Deklarasi fungsi keluar aplikasi
    Private Sub KeluarToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles KeluarToolStripMenuItem1.Click
        Me.Close()
    End Sub

End Class
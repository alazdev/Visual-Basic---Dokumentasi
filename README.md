---
typora-copy-images-to: Image
---

# Visual Basic Documentation 

By Azim



[TOC]



## Yang Diperlukan

1. Microsoft **SQL Server**
2. Microsoft **SQL Server Express**
3. Microsoft **SQL Server Management Studio**
4. Microsoft **Visual Studio Community** (Visual Basic)

## Yang Perlu Diingat

1. Rancangan **Database**
2. Rancangan **Tampilan Form**
3. **Coding**

## 

## Coding

------

### Koneksi

1. Buka **Visual Studio** lalu buatlah project baru (Visual Basic)

2. Buat **Class** baru, dengan cara buka **Solution Explorer** klik kanan pada nama project klik **Add > New Item...**![Add New Item...](Image/Add%20New%20Item....png)

3. Lalu pilih **Class** dan masukkan nama class di sudut kiri bawah![Add New Item...](Image/Class%20-%20KoneksiDB.png)

4. **Imports** **Sql** dan **SqlClient**

   ```vb
   Imports System.Data.Sql
   Imports System.Data.SqlClient
   ```

5. Buatlah variabel **New SqlConnection** dan **String**

   ```vb
   Public Connection As New SqlConnection
   Public SQL As String
   ```

6. Buatlah **Function** untuk membuka Koneksi

   ```vb
   Public Sub BukaKoneksiDB()
   	SQL = "Data Source=DESKTOP-V7HVKL2\SQLEXPRESS;Initial Catalog=sistem_informasi;Integrated Security=True"
   	Connection = New SqlConnection(SQL)
   	Try
   		If Connection.State = ConnectionState.Closed Then
   			Connection.Open()
            Else
            	MsgBox("Koneksi bermasalah")
            End If
   	Catch ex As Exception
   		MsgBox("Koneksi bermasalah")
   	End Try
   End Sub
   ```

7. Buat juga **Function** untuk menutup Koneksi

   ```vb
   Public Sub TutupKoneksiDB()
       Connection.Close()
   End Sub
   ```

8. Akan terlihat kurang lebih seperti ini![1 KoneksiDB](Image/1%20KoneksiDB.png)

------

### (Kodingan) Login

1. **Imports** **SqlClient** terlebih dahulu

   ```vb
   Imports System.Data.SqlClient
   ```

2. Panggil **KoneksiDB** yang dibuat tadi (diatas) dengan kodingan seperti berikut

   ```vb
   Dim Koneksi As New KoneksiDB
   ```

3. Buka koneksi

   ```vb
   Koneksi.BukaKoneksiDB()
   ```

4. Buatlah Proses **Login**

   ```vb
   Try
       Koneksi.BukaKoneksiDB()
       If Not String.IsNullOrEmpty(TextBox1.Text) And Not String.IsNullOrEmpty(TextBox2.Text) Then
           Dim cmd As SqlCommand
           Dim Bacakan As SqlDataReader
           Dim SQL As String
           
           SQL = "SELECT * FROM staf WHERE nip='" & TextBox1.Text & "' AND password='" & TextBox2.Text & "'"
           cmd = New SqlCommand(SQL, Koneksi.Connection)
           Bacakan = cmd.ExecuteReader()
           If Bacakan.HasRows = True Then
               While Bacakan.Read()
                   MenuKaryawan.NamaStaf = Bacakan.Item("nama").ToString()
               End While
               'Disini Anda Bisa Memasukkan Form apa yang ditampilkan setelah login
           Else
               TextBox1.Clear()
               TextBox2.Clear()
               MsgBox("NIP Tidak Ditemukan!")
           End If
       Else
           TextBox1.Clear()
           TextBox2.Clear()
           MsgBox("Masukkan Username dan Password Terlebih Dahulu")
       End If
   Catch ex As Exception
       MsgBox("Koneksi Bermasalah")
   Finally
       Koneksi.TutupKoneksiDB()
   End Try
   ```

5. Jangan lupa untuk menutup koneksi

   ```vb
   Koneksi.TutupKoneksiDB()
   ```

6. Akan terlihat  kurang lebih seperti ini![2 Login](Image/2%20Login.png)

------

### (Kodingan) Create, Update, dan Delete

1. **Imports** **SqlClient** terlebih dahulu

   ```vb
   Imports System.Data.SqlClient
   ```

2. Panggil **KoneksiDB** yang dibuat tadi (diatas) dengan kodingan seperti berikut

   ```vb
   Dim Koneksi As New KoneksiDB
   ```

3. Buka koneksi setiap function yang membutuhkan koneksi

   ```vb
   Koneksi.BukaKoneksiDB()
   ```

4. Kodingan **Create** kurang lebih akan seperti berikut ini

   ```vb
   Dim cmd As SqlCommand
   Dim Insert As String = "INSERT INTO nama_tabel (nama_kolom1,nama_kolom2,nama_kolom3) VALUES ('" & TextBox1.Text() & "','" & TextBox2.Text() & "','" & TextBox3.Text() & "')"
   
   cmd = New SqlCommand(Insert, Koneksi.Connection)
   cmd.ExecuteNonQuery()
   ```

5. Kodingan untuk memasukkan value dari **Tabel** ke **Form** dan **String**

   ```vb
   Dim Data As Integer = DataGridView1.CurrentRow.Index()
   
   nama_kolom1 = DataGridView1.Item(0, Data).Value
   TextBox2.Text() = DataGridView1.Item(1, Data).Value
   TextBox3.Text() = DataGridView1.Item(2, Data).Value
   ```

6. Kodingan **Update** kurang lebih akan seperti berikut ini

   ```vb
   Dim cmd As SqlCommand
   Dim Update As String = "UPDATE nama_tabel SET nama_kolom2='" & TextBox2.Text() & "',nama_kolom3='" & TextBox3.Text() & "' WHERE nama_kolom1='" & nama_kolom1 & "'"
   
   cmd = New SqlCommand(Update, Koneksi.Connection)
   cmd.ExecuteNonQuery()
   ```

7. Kodingan untuk memasukkan value dari **Tabel** ke **String**, kemudian dihapus kurang lebih akan seperti berikut

   ```vb
   Dim Data As Integer = DataGridView1.CurrentRow.Index()
   Dim cmd As SqlCommand
   
   nama_kolom1 = DataGridView1.Item(0, Data).Value
   Dim Delete As String = "DELETE nama_tabel WHERE nama_kolom1='" & id_jurusan & "'"
   
   cmd = New SqlCommand(Delete, Koneksi.Connection)
   cmd.ExecuteNonQuery()
   ```

   

8. Jangan lupa untuk menutup koneksi setiap function yang membutuhkan koneksi

   ```vb
   Koneksi.TutupKoneksiDB()
   ```

------

### (Kodingan) Search

1. **Imports** **SqlClient** terlebih dahulu

   ```vb
   Imports System.Data.SqlClient
   ```

2. Panggil **KoneksiDB** yang dibuat tadi (diatas) dengan kodingan seperti berikut

   ```vb
   Dim Koneksi As New KoneksiDB
   ```

3. Buka koneksi

   ```vb
   Koneksi.BukaKoneksiDB()
   ```

4. Buatlah proses **Search** dan masukkan hasilnya ke **Tabel**

   ```vb
   Dim ds As New DataSet
   
   Dim Search As String = "SELECT * FROM nama_tabel WHERE nama_kolom2 LIKE '%" & TextBox0.Text() & "%' OR nama_kolom3 LIKE '%" & TextBox0.Text() & "%'
   
   Dim da As New SqlDataAdapter(Search, Koneksi.Connection)
   da.Fill(ds, "nama_tabel")
   DataGridView1.DataSource = ds.Tables("nama_tabel")
   ```

5. Jangan lupa untuk menutup koneksi

   ```vb
   Koneksi.TutupKoneksiDB()
   ```

------

### 

**<u>Maaf ya :) Dokumentasi ini hanya diperuntukkan pribadi, maaf jika Anda kurang paham atau tidak paham..</u>**

**<u>Jika Anda paham, kodingan bisa disesuaikan dengan kebutuhan...</u>**# Visual-Basic---Dokumentasi
Option Strict On

Imports Microsoft.Data.Sqlite

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '▼ファイルのダイアログからｃｓｖファイルを取得
        Dim dialog As New OpenFileDialog
        dialog.Filter = "CSVファイル|*.csv|すべてのファイル|*.*"
        dialog.InitialDirectory = Application.StartupPath
        dialog.FileName = "19YAMANA.CSV"

        If dialog.ShowDialog = DialogResult.Cancel Then
            Return
        End If

        Dim csvFileName As String = dialog.FileName

        '▼ｄｂにインサート

        'データ格納用のテーブルがなければ作成
        Using database As New SqliteConnection("Data Source=mydatabase.db")
            Using sql = database.CreateCommand
                sql.CommandText = "CREATE TABLE IF NOT EXISTS postal(" &
                        "PostalCode TEXT, " &
                        "Address TEXT)"
                database.Open()
                sql.ExecuteNonQuery()
                database.Close()
            End Using
        End Using


        'CSVを読み込んでデータベースに格納
        Const maxCount As Integer = 100 'サンプルなので最大100件まで処理することにします。
        Dim count As Integer

        Using database As New SqliteConnection("Data Source=mydatabase.db")

            '.NET Frameworkで実行する場合、この行は不要でエラーになるのでコメント化すること
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance)

            Dim reader As New IO.StreamReader(csvFileName, System.Text.Encoding.GetEncoding("shift_jis"))

            database.Open()
            Do Until reader.EndOfStream
                Dim line As String = reader.ReadLine
                Dim items As String() = line.Split(",")
                Dim postalCode As String = items(2).Trim(""""c)
                Dim address As String = (items(6) & items(7) & items(8)).Replace("""", "")

                InsertToDb(database, postalCode, address)

                count += 1
                If count = maxCount Then
                    Exit Do
                End If

            Loop
            database.Close()

        End Using

        '▼ちゃんとデータベースに格納されたか確認
        'データベースから読み込んで画面上のDataGridView1に表示します。

        Using database As New SqliteConnection("Data Source=mydatabase.db")
            database.Open()
            Dim table As DataTable = ReadDb(database)
            database.Close()
            DataGridView1.DataSource = table
        End Using

    End Sub

    Private Sub InsertToDb(database As SqliteConnection, postalCode As String, address As String)

        Using sql = database.CreateCommand
            sql.CommandText = "INSERT INTO postal VALUES (@postalCode, @address)"
            sql.Parameters.AddWithValue("@postalCode", postalCode)
            sql.Parameters.AddWithValue("@address", address)

            sql.ExecuteNonQuery()
        End Using
    End Sub

    Private Function ReadDb(database As SqliteConnection) As DataTable

        Dim table As New DataTable
        table.Columns.Add("postalCode", GetType(String))
        table.Columns.Add("address", GetType(String))

        Using sql = database.CreateCommand
            sql.CommandText = "SELECT * FROM postal"
            Using reader As SqliteDataReader = sql.ExecuteReader
                Do While reader.Read
                    table.Rows.Add(reader("postalCode"), reader("address"))
                Loop
            End Using
        End Using

        Return table

    End Function
End Class

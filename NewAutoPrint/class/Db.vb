Option Strict On
Option Explicit On
Option Infer On

' import module
Imports MySql.Data.MySqlClient
Imports Mysqlx.Expect.Open.Types.Condition.Types

' ◇ MySQLモジュール
Public Class Db
    ' モジュール変数
    Private Property mysqlCon As New MySqlConnection

    '--------------------------------------------------------
    '*番  号：SQL1
    '*関数名：Sql_st
    '*機  能：データベース接続
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub Sql_st()
        ' 変数定義
        Dim mysqldb As String ' SQL接続文

        Try
            ' SQL接続文
            mysqldb = "Server=localhost" _
            & ";Port=3306" _
            & ";Database=*****" _
            & ";UserID=*****" _
            & ";Password='*****'"

            ' SQL設定
            mysqlCon.ConnectionString = mysqldb
            ' SQL接続
            mysqlCon.Open()

            Debug.Print("db connected.")

        Catch ex As Exception
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：SQL2
    '*関数名：Sql_cl
    '*機  能：データベース切断
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub Sql_cl()
        Try
            ' データベースの切断
            mysqlCon.Close()
            Debug.Print("db closed.")

        Catch ex As Exception
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：SQL3
    '*関数名：Sql_select
    '*出　力：抽出結果（DataTable）
    '*機  能：SELECT処理
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Function Sql_select(query As String) As DataTable
        ' 変数定義
        Dim mySqlDataAdapter As MySqlDataAdapter ' MySQL接続アダプタ
        Dim dt As DataTable ' データテーブル

        ' DataTable初期化
        dt = New DataTable()

        Try
            Debug.Print("db select started.")
            ' DBアダプタ
            mySqlDataAdapter = New MySqlDataAdapter(query, mysqlCon)
            ' データを取得し、アダプタにセットする
            mySqlDataAdapter.Fill(dt)

        Catch ex As Exception
            Console.WriteLine(ex)
        End Try
        ' データテーブル返し
        Return dt
    End Function

    '--------------------------------------------------------
    '*番  号：SQL4
    '*関数名：Sql_update
    '*出　力：処理結果
    '*機  能：UPDATE処理
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Function Sql_update(table As String, valueList As List(Of Hashtable), Optional whereHash As Hashtable = Nothing) As String
        ' 変数定義
        Dim valIndex As Integer ' 列カウンタ
        Dim whIndex As Integer ' whereカウンタ
        Dim query As String ' クエリ
        Dim cmd As MySqlCommand ' DB実行体

        ' SQL読込
        cmd = New MySqlCommand()

        Try
            ' クエリ初期化
            query = "Update " + table + " SET "

            ' 値
            For Each hash As Hashtable In valueList
                ' プレースホルダー
                For Each key1 As String In hash.Keys
                    ' ‘@‘ を連結
                    query += CStr(key1)
                    ' = を連結
                    query += " = "
                    ' ‘@‘ を連結
                    query += "@" + CStr(key1)
                    ' where句あり
                    If whereHash IsNot Nothing Then
                        ' 最後以外
                        If (valIndex < whereHash.Keys.Count - 1) Then
                            ' 最後以外はカンマ付
                            query += ", "
                        End If
                    End If
                    ' インクリメント
                    valIndex += 1
                Next
            Next

            ' where句
            If whereHash IsNot Nothing Then
                ' クエリ
                query += "WHERE "

                ' プレースホルダー
                For Each key2 As String In whereHash.Keys
                    ' ‘@‘ を連結
                    query += CStr(whereHash(key2))
                    ' = を連結
                    query += " = "
                    ' ‘@‘ を連結
                    query += "@" + CStr(key2)
                    ' 最後以外
                    If (whIndex < whereHash.Keys.Count - 1) Then
                        ' 最後以外はカンマ付
                        query += " AND "
                    End If
                    ' インクリメント
                    whIndex += 1
                Next
            End If

            ' 値
            For Each hash As Hashtable In valueList
                ' SQL設定
                With cmd
                    ' 事前設定
                    .CommandText = query
                    .Connection = mysqlCon
                    .CommandType = CommandType.Text
                    ' HashTableループ
                    For Each key As String In hash.Keys
                        ' 値セット
                        .Parameters.AddWithValue("@" + key, hash(key))
                        Console.WriteLine("column: " + key)
                        Console.WriteLine("value: " + CStr(hash(key)))
                    Next
                    ' Insert実行
                    .Prepare()
                    .ExecuteNonQuery()
                    ' 値クリア
                    .Parameters.Clear()
                End With
            Next

            ' 最終insertId
            Return "success"

        Catch ex As MySqlException
            Console.WriteLine(ex)
            ' エラー時は0
            Return "error"
        End Try
    End Function

    '--------------------------------------------------------
    '*番  号：SQL5
    '*関数名：Sql_insert
    '*出　力：最終処理ID（Long）、0はエラー
    '*機  能：INSERT処理
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Function Sql_insert(table As String, valueList As List(Of Hashtable)) As Long
        ' 変数定義
        Dim valIndex As Integer ' 列カウンタ
        Dim query As String ' クエリ
        Dim columns As String = "" ' カラム名
        Dim placeholder As String = "" ' プレースホルダー
        Dim cmd As MySqlCommand ' DB実行体

        Try
            ' クエリ
            query = ""
            ' SQL読込
            cmd = New MySqlCommand()

            ' プレースホルダー
            For Each key As String In valueList(0).Keys
                columns += key
                placeholder +=  "@" + CStr(key)
                ' 最後以外
                If (valIndex < valueList(0).Keys.Count - 1) Then
                    ' 最後以外はカンマ付
                    columns += ", "
                    placeholder += ", "
                End If
                ' インクリメント
                valIndex += 1
            Next

            ' クエリ
            query = "insert into " + table + " (" + columns + ") values (" + placeholder + ");"

            ' 値
            For Each hash As Hashtable In valueList
                ' SQL設定
                With cmd
                    ' 事前設定
                    .CommandText = query
                    .Connection = mysqlCon
                    .CommandType = CommandType.Text
                    ' HashTableループ
                    For Each key As String In hash.Keys
                        ' 値セット
                        .Parameters.AddWithValue("@" + key, hash(key))
                        Console.WriteLine("column: " + key)
                        Console.WriteLine("value: " + CStr(hash(key)))
                    Next
                    ' Insert実行
                    .Prepare()
                    .ExecuteNonQuery()
                    ' 値クリア
                    .Parameters.Clear()
                End With
            Next

            ' 最終insertId
            Return cmd.LastInsertedId

        Catch ex As MySqlException
            Console.WriteLine("Error: ")
            Console.WriteLine(ex)
            Return 0
        End Try
    End Function
End Class
Option Strict On
Option Explicit On
Option Infer On

' □ MainFormモジュール
Public Class MainForm
    ' ■ ButtonClick
    ' DataGridView1
    '--------------------------------------------------------
    '*番  号：B3
    '*関数名：SelectWorkButton
    '*機  能：作業フォルダ選択ボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub SelectWorkButton(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ' 変数定義
        Dim ymd As String ' 年月日
        Dim selectFullPathStr As String ' ファイル名フルパス
        Dim workPathStr As String ' Workフォルダフルパス
        Dim fbd As New FolderBrowserDialog ' ダイアログ
        ' オブジェクト定義
        Dim itemDataTable As DataTable' itemテーブル
            
        Try
            ' 作業進捗ヘッダ初期化
            ComboBox1.SelectedItem = ""
            ComboBox2.SelectedItem = ""
            TextBox10.Text = ""
            TextBox11.Text = ""
            CheckBox1.Checked = True
            ' 作業進捗ヘッダ初期化
            Button7.Text = "Check"
            TextBox1.Text = "0"
            TextBox2.Text = "0"
            TextBox4.Text = "0"
            TextBox5.Text = "0"
            TextBox6.Text = "0"
            TextBox7.Text = "0"
            TextBox8.Text = "0"
            TextBox9.Text = "0"
            Label22.Text = "0"
            Label27.Text = ""
            
            ' 上部に表示する説明テキストを指定する
            fbd.Description = "フォルダを指定してください。"
            ' ルートフォルダを指定するデフォルトでDesktop
            fbd.RootFolder = Environment.SpecialFolder.Desktop
            ' 最初に選択するフォルダを指定するRootFolder以下にあるフォルダである必要がある
            fbd.SelectedPath = GetAppPath() + "\Work\"
            ' ユーザーが新しいフォルダを作成できるようにするデフォルトでTrue
            fbd.ShowNewFolderButton = True

            ' ダイアログを表示する
            If fbd.ShowDialog(Me) = DialogResult.OK Then
                ' ファイル名（拡張子なし）の取得
                ymd = IO.Path.GetFileNameWithoutExtension(fbd.SelectedPath)
                ' グローバルに日付を格納
                globalNowDate = ymd
                ' ファイル名のフルパス
                selectFullPathStr = IO.Path.GetFullPath(fbd.SelectedPath).ToLower
                ' Workフォルダのフルパス
                workPathStr = (GetAppPath() + "\Work").ToLower
                ' Workフォルダあり
                If selectFullPathStr.IndexOf(workPathStr) >= 0 AndAlso IO.Path.GetFullPath(fbd.SelectedPath).Length > (GetAppPath() + "\Work").Length Then
                    ' YMDテキストラベル更新
                    Label24.Text = ymd.Substring(0, 2) + "年"
                    Label25.Text = ymd.Substring(2, 2) + "月"
                    Label26.Text = ymd.Substring(4, 2) + "日"
                    ' itemTable読込直し
                    itemDataTable = InstructionsFilter()
                    MsgBox("Workフォルダ内の発送日フォルダを選択してください。")
                End If
            End If
            ' DataGridView初期化
            itemDataTable = InstructionsFilter()
            ' DataGridView1流し込み
            DataGridView1.DataSource = itemDataTable

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：B4
    '*関数名：BundleFolderButton
    '*機  能：一括フォルダ選択ボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub BundleFolderButton(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ' 変数定義
        Dim dirCnt As Integer ' カウンタ1
        Dim selCnt As Integer ' カウンタ2
        Dim outPath As String ' ディレクトリパス
        Dim updateItemResult As String ' item更新結果
        ' オブジェクト定義
        Dim resultItemTable As DataTable ' itemテーブル

        Try
            ComboBox1.SelectedItem = ""
            ComboBox2.SelectedItem = ""
            TextBox10.Text = ""
            CheckBox1.Checked = True
            resultItemTable = InstructionsFilter()

            ' itemDataが空
            If resultItemTable Is Nothing OrElse resultItemTable.Rows.Count = 0 Then
                ' データなし
                MsgBox("登録がありません")
                Throw New System.Exception("An exception has occurred.")
            End If

            ' DataTableループ
            For Each itemRow As DataRow In resultItemTable.Rows
                ' ディレクトリパス
                outPath = GetAppPath() + "\Work\" + globalNowDate + "\" + itemRow("指示書").ToString + "\"
                ' フォルダ (ディレクトリ) が存在しているかどうか確認する
                Call MakeEmptyDir(outPath)
                ' 存在しなければフォルダ作成
                If Not IO.Directory.Exists(outPath) Then
                    ' 配置進捗クリア
                    itemRow("配置進捗") = "　"
                    ' カウントアップ
                    dirCnt = dirCnt + 1
                End If
                ' 配置進捗空欄
                If itemRow("配置進捗") Is "　" Then
                    ' チェック付ける
                    itemRow("配置選択") = True
                    ' 選択数カウント表示更新
                    Label22.Text = CStr(CInt(Label22.Text) + 1)
                    ' カウントアップ
                    selCnt = selCnt + 1
                End If
            Next
            ' itemDB更新
            updateItemResult = UpdateItemDb(resultItemTable)
            ' 更新成功
            If updateItemResult <> "error" Then
                ' DataGridView1流し込み
                DataGridView1.DataSource = resultItemTable
            End If
            ' メッセージ表示
            MsgBox("ﾌｫﾙﾀﾞ一括作成[" + dirCnt.ToString + "件]" + vbCrLf + "　　と" + vbCrLf + "未配置ﾌｫﾙﾀﾞ選択[" + selCnt.ToString + "件]" + vbCrLf + vbCrLf + "完了しました")

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：B5
    '*関数名：LayoutLabelsButton
    '*機  能：検索・配置・出力ボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub LayoutLabelsButton(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        ' 変数定義
        Dim x As Integer ' 幅
        Dim y As Integer ' 高さ
        Dim cnt As Integer ' 高さ
        Dim wCount As Integer ' 幅カウンタ
        Dim hCount As Integer ' 高さカウンタ
        Dim counter As Integer ' カウンタ1
        Dim txtCount As Integer ' カウンタ2
        Dim endHaichi As Integer ' カウンタ3
        Dim checkJuchu As Integer ' カウンタ4
        Dim shipYmdDir As String ' 発送日
        Dim itemFullPath As String ' フルパス
        Dim tmpItemFolderPath As String ' フルパス
        Dim sameIdFileKey As String ' ai検索キー
        Dim sameIdCopyName As String ' 重複ファイル
        Dim updateItemResult As String ' item更新結果
        Dim updatePrintResult As String ' print更新結果
        Dim newPageFlag As Boolean ' 初回フラグ
        ' 配列定義
        Dim sameIdfiles As String() ' カウンタ行
        ' ジェネリック定義
        Dim txtRows As New List(Of Integer) ' カウンタ行
        ' オブジェクト定義
        Dim resultItemTable As DataTable ' itemテーブル
        Dim resultPrintTable As DataTable ' printテーブル
        Dim finalPrintTable As DataTable ' 最終printテーブル
        Dim resultTables As DataSet ' データセット

        Try
            ' DBインスタンス作成
            Dim dbMaker = New Db()
            ' illustratorインスタンス作成
            Dim adobeMaker = New adobeAI()
            ' 初期化
            wCount = 0
            hCount = 0
            x = 70
            y = 1700
            Label22.Text = ""
            Label27.Text = ""
            ' DB取得
            resultTables = GetAllDbData(globalNowDate)
            ' itemDB取得
            resultItemTable = resultTables.Tables(0)
            ' printDB取得
            resultPrintTable = resultTables.Tables(1)
            ' プログレスバーを初期化する
            ProgressBar2.Minimum = 0
            ProgressBar2.Maximum = Integer.Parse(Label22.Text)
            ProgressBar2.Value = 0

            ' itemDataが空
            If resultItemTable Is Nothing OrElse resultItemTable.Rows.Count = 0 Then
                ' データなし
                MsgBox("登録がありません")
                Throw New System.Exception("An exception has occurred.")
            End If

            ' itemDataTableループ
            For Each tmpRow As DataRow In resultItemTable.Rows
                ' EPS処理済みをカウント
                If tmpRow("配置進捗") Is "EPS処理済" Then
                    txtCount = txtCount + 1
                    txtRows(txtCount) = (counter)
                End If
                ' インクリメント
                counter = counter + 1
            Next

            ' itemDataTableループ
            For Each itemRow As DataRow In resultItemTable.Rows
                ' 変数初期化
                checkJuchu = 0
                ' aiﾌｧｲﾙ検出
                If Not CBool(itemRow("配置選択")) OrElse itemRow("配置進捗").ToString <> "　" Then
                    Console.WriteLine("配置無し")
                    Continue For
                End If
                ' カレントセル指定
                itemRow("配置進捗") = "検出中"
                ' プログレスバー更新
                ProgressBar1.Minimum = 0
                ProgressBar1.Maximum = checkJuchu
                ProgressBar1.Value = 0
                ' aiファイル検索
                Call SearchAiFiles(itemRow, resultPrintTable)

                If CInt(itemRow("検出OK")) > 0 Then
                    itemRow("検出ﾌｧｲﾙﾊﾟｽ") = "検出有"
                Else
                    itemRow("検出ﾌｧｲﾙﾊﾟｽ") = "検出無"
                End If
                ' printDataが空
                If resultPrintTable Is Nothing OrElse resultPrintTable.Rows.Count = 0 Then
                    ' データなし
                    Console.WriteLine("データなし")
                    Continue For
                End If
                ' printDataTableループ
                For Each printRow As DataRow In resultPrintTable.Rows
                    ' 初回フラグ
                    newPageFlag = True
                    ' 指示書内受注書確認
                    If printRow("出力").ToString = "" Then
                        Console.WriteLine("出力無し")
                        Continue For
                    End If
                    ' aiファイル配置
                    ' 発送日フォルダ
                    shipYmdDir = GetAppPath() + "\Work\" + globalNowDate + "\"
                    ' フルパス
                    itemFullPath = shipYmdDir + printRow("顧客番号").ToString
                    ' ai検索キー
                    sameIdFileKey = itemFullPath + " *.ai"
                    ' ai検索対象フォルダ
                    sameIdfiles = IO.Directory.GetFiles(shipYmdDir, sameIdFileKey, IO.SearchOption.AllDirectories)
                    ' 検索ヒット
                    If sameIdfiles.Length > 0 Then
                        ' 重複ファイル
                        sameIdCopyName = IO.Path.GetDirectoryName(itemRow("指示書").ToString + "\" + IO.Path.GetFileName(sameIdfiles(0)))
                        ' ファイルを移動する
                        IO.File.Move(sameIdfiles(0), sameIdCopyName)
                        ' 検出状況更新
                        printRow("検出状況") = "検出成功"
                        ' 作業ﾌｧｲﾙﾊﾟｽ更新
                        printRow("作業ﾌｧｲﾙﾊﾟｽ") = sameIdCopyName
                    End If
                    ' 最初以外
                    If counter > 0 Then
                        ' ヘッダ配置
                        adobeMaker.TextLayout("▼", 36, x + wCount * 470 - 30, (y - 20) - hCount * 470)
                    End If
                    ' 検出成功 | 手動検出 | 変更変換済み
                    If printRow("検出状況").ToString = "検出成功" Or printRow("検出状況").ToString = "手動検出" Or printRow("検出状況").ToString = "変更変換済み" Then
                        printRow("出力") = "NG"
                    Else
                        printRow("出力") = "OK"
                    End If
                    ' ai配置
                    If CheckBox2.Checked = False And CheckBox12.Checked = True Then
                        Dim tmp As String = resultItemTable.Rows(txtRows(txtCount - 1)).Item("フォルダ").ToString
                        ' 対象フォルダパス
                        tmpItemFolderPath = System.IO.Path.GetDirectoryName(tmp)
                        ' printDataTableループ
                        cnt = cnt + 1
                        ' ai配置関数配置
                        makeLayoutPrint(itemRow, printRow, cnt, True)
                    End If
                Next
                ' チェックはずす
                itemRow("配置選択") = False
                ' 選択数カウント表示更新
                Label22.Text = CStr(CInt(Label22.Text) - 1)
                ' 配置進捗更新
                itemRow("配置進捗") = "完了"
                ' インクリメント
                endHaichi = endHaichi + 1
                ' プログレスバー更新
                ProgressBar2.Value = endHaichi
                ' 背景色
                itemRow("背景色") = Color.LightGray
                ' 進捗ﾌｧｲﾙ更新
                itemRow("進捗") = "検索済"
            Next
            
            ' 印刷作業
            If CheckBox13.Checked = True Then
                ' printDataTableループ
                For Each printRow As DataRow In resultPrintTable.Rows
                    ' ラベル印刷
                    makeLabelPrint(printRow, False, newPageFlag)
                    newPageFlag = False
                Next
            End If

            ' itemDB更新
            updateItemResult = UpdateItemDb(resultItemTable)
            ' printDB更新
            updatePrintResult = UpdatePrintDb(resultPrintTable)

            ' item更新成功
            If updateItemResult <> "error" Then
                ' DataGridView1流し込み
                DataGridView1.DataSource = resultItemTable
            End If
            ' print更新成功
            If updatePrintResult <> "error" Then
                ' printDB更新
                finalPrintTable = PrintDataTableModify(resultPrintTable)
                ' DataGridView2流し込み
                DataGridView2.DataSource = finalPrintTable
            End If
            ' ソート無効
            For Each c As DataGridViewColumn In DataGridView2.Columns
                c.SortMode = DataGridViewColumnSortMode.NotSortable
            Next c
            
            ' チェックあり
            If Not CheckBox1.Checked Then
                MsgBox("配置処理完了")
            Else
                MsgBox("配置/面付け処理完了")
            End If

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：B6
    '*関数名：OutputLayoutButton
    '*機  能：面付出力ボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub OutputLayoutButton(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ' 変数定義
        Dim cnt As Integer
        Dim updateItemResult As String ' item更新結果
        Dim updatePrintResult As String ' print更新結果
        ' オブジェクト定義
        Dim resultDataSet As DataSet ' データセット
        Dim resultItemTable As DataTable ' itemテーブル
        Dim resultPrintTable As DataTable ' printテーブル
        Dim finalPrintTable As DataTable ' 最終printテーブル

        Try
            ' DB取得
            resultDataSet = GetAllDbData(globalNowDate)
            ' itemDB取得
            resultItemTable = resultDataSet.Tables(0)
            ' printDB取得
            resultPrintTable = resultDataSet.Tables(1)

            ' itemDataが空
            If resultItemTable Is Nothing OrElse resultItemTable.Rows.Count = 0 Then
                ' データなし
                MsgBox("登録がありません")
                Throw New System.Exception("An exception has occurred.")
            End If

            ' 面付選択数確認
            If Integer.Parse(Label23.Text) > 0 Then
                ' 未選択メッセージ
                MsgBox("選択されていません")
                Throw New System.Exception("An exception has occurred.")
            End If

            ' itemDataTableループ
            For Each itemRow As DataRow In resultItemTable.Rows
                ' 面付選択チェック中
                If Not CBool(itemRow("面付選択")) Then
                    Console.WriteLine("面付選択無し")
                    Continue For
                End If
                ' チェックはずす
                itemRow("面付選択") = False
                ' 進捗更新
                Label27.Text = itemRow("指示書").ToString
                ' 選択行進捗とラジオボタンの表示同期
                Select Case itemRow("進捗").ToString
                    Case ""
                        RadioButton1.Checked = True
                    Case "未処理"
                        RadioButton1.Checked = True
                    Case "検索済"
                        RadioButton2.Checked = True
                    Case "配置出力済"
                        RadioButton3.Checked = True
                    Case "配置確認済"
                        RadioButton4.Checked = True
                    Case "面付出力済"
                        RadioButton5.Checked = True
                    Case "確認待ち"
                        RadioButton6.Checked = True
                End Select
                ' 進捗更新
                RadioButton5.Checked = True
                ' printDataTableループ
                For Each printRow As DataRow In resultPrintTable.Rows
                    ' 配置図出力処理へ
                    makeLayoutPrint(itemRow, printRow, cnt, False)
                    cnt = cnt + 1
                Next
                ' 選択数カウント表示更新
                Label23.Text = CStr(CInt(Label23.Text) - 1)
                ' 表示更新
                Refresh()
            Next
            
            ' 指示書パスクリア
            Label27.Text = ""
            ' itemDB更新
            updateItemResult = UpdateItemDb(resultItemTable)
            ' printDB更新
            updatePrintResult = UpdatePrintDb(resultPrintTable)

            ' item更新成功
            If updateItemResult <> "error" Then
                ' DataGridView1流し込み
                DataGridView1.DataSource = resultItemTable
            End If
            ' print更新成功
            If updatePrintResult <> "error" Then
                ' printDB更新
                finalPrintTable = PrintDataTableModify(resultPrintTable)
                ' DataGridView2流し込み
                DataGridView2.DataSource = finalPrintTable
            End If
            ' 完了メッセージ
            MsgBox("面付出力処理完了")

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：B7
    '*関数名：LabelCheckButton
    '*機  能：Checkボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub LabelCheckButton(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        ' 変数定義
        Dim ymd As String ' 年月日
        Dim shipDir As String ' Sドライブ発送日フォルダ
        Dim baseDir As String ' Lドライブ発送日フォルダ
        Dim fileCheck As String ' チェックOK
        Dim sakuseisyoTxtPath As String ' 作成書ファイルフルパス名
        Dim sakuseisyoDir As String ' 作成書ファイルフォルダ
        Dim searchFilePathAi As String ' aiチェック
        Dim searchFileKeyAi As String ' aiチェック
        Dim searchFilePathEPS As String ' epsファイルパス
        Dim searchFileKeyEPS As String ' epsファイルパス
        Dim updateItemResult As String ' item更新結果
        Dim updatePrintResult As String ' print更新結果 
        Dim tmpYmd As String ' 一時年月日

        ' 配列定義
        Dim filesEPS() As String ' epsファイル一式
        Dim filesAi() As String ' ファイル一式

        ' オブジェクト定義
        Dim resultDataSet As DataSet ' データセット
        Dim resultItemTable As DataTable ' itemテーブル
        Dim resultPrintTable As DataTable ' printテーブル
        Dim finalPrintTable As DataTable ' 最終printテーブル

        Try
            ' DB取得
            resultDataSet = GetAllDbData(globalNowDate)
            ' itemDB取得
            resultItemTable = resultDataSet.Tables(0)
            ' printDB取得
            resultPrintTable = resultDataSet.Tables(1)

            ' itemDataが空
            If resultItemTable Is Nothing OrElse resultItemTable.Rows.Count = 0 Then
                ' データなし
                MsgBox("登録がありません")
                ' エラー
                Throw New System.Exception("An exception has occurred.")
            End If

            ' 一時年月日
            tmpYmd = resultItemTable.Rows(0).Item("前回発送日").ToString
            ' 年月日取得
            ymd = tmpYmd.Substring(0, 6)
            ' Sドライブ発送日フォルダ取得
            shipDir = GetAppPath() + "\work\" + ymd + "\"
            ' Lドライブaiファイルコピー先日付フォルダベースディレクトリ
            baseDir = testLabelHDPath + ymd.Substring(0, 2) + "年度ラベル作成済み分\20" + ymd.Substring(0, 4) + "\" + ymd.Substring(2, 4)
            ' Lドライブaiファイルコピー
            If Not IO.Directory.Exists(baseDir) Then
                ' Lドライブaiファイルコピー
                IO.Directory.CreateDirectory(baseDir)
            End If
            ' 日付フォルダ＞★仕分け済み内「★jpg」フォルダ確認
            If Not IO.Directory.Exists(shipDir + "★仕分け済み\★jpg\") Then
                ' 日付フォルダ＞★仕分け済み内「★jpg」フォルダ生成
                IO.Directory.CreateDirectory(shipDir + "★仕分け済み\★jpg\")
            End If
            ' チェックOK
            fileCheck = "OK"

            ' itemDataTableループ
            For Each itemRow As DataRow In resultItemTable.Rows
                ' 作成書ファイルフルパス名取得
                sakuseisyoTxtPath = itemRow("指示書").ToString
                ' 作成書ファイルフォルダ取得
                sakuseisyoDir = itemRow("フルパス").ToString
                ' 作成書ファイルフォルダ取得
                Label27.Text = sakuseisyoTxtPath
                ' printDataが空
                If resultPrintTable Is Nothing OrElse resultPrintTable.Rows.Count = 0 Then
                    ' データなし
                    Console.WriteLine("データなし")
                    Continue For
                End If

                ' printDataTableループ
                For Each printRow As DataRow In resultPrintTable.Rows
                    ' 変数初期化
                    searchFilePathAi = ""
                    searchFileKeyAi = ""
                    searchFilePathEPS = ""
                    searchFileKeyEPS = ""

                    ' 枚数確認
                    If CInt(printRow("枚数")) = 0 OrElse printRow("作業ﾌｧｲﾙﾊﾟｽ").ToString.Length = 0 Then
                        Console.WriteLine("枚数・作業ﾌｧｲﾙﾊﾟｽなし")
                        Continue For
                    End If
                    ' aiチェック
                    searchFilePathAi = IO.Path.GetDirectoryName(printRow("作業ﾌｧｲﾙﾊﾟｽ").ToString)
                    ' ai検索キー
                    searchFileKeyAi = IO.Path.GetFileNameWithoutExtension(printRow("作業ﾌｧｲﾙﾊﾟｽ").ToString) + ".ai"
                    ' ファイル一式
                    filesAi = IO.Directory.GetFiles(searchFilePathAi, searchFileKeyAi, IO.SearchOption.AllDirectories)
                    ' ファイル数が1以外でOK
                    If filesAi.Length <> 1 And fileCheck = "OK" Then
                        ' ファイルチェック
                        fileCheck = searchFilePathAi + "フォルダ内" + vbCrLf + vbCrLf + "作成書：" + Label27.Text + vbCrLf + vbCrLf + searchFileKeyAi + vbCrLf + vbCrLf
                    End If

                    ' ファイルなし
                    If filesAi.Length = 0 Then
                        ' ファイルチェック
                        fileCheck = fileCheck + "・Aiが見つかりません" + vbCrLf
                    ' ファイル数が2以上
                    ElseIf filesAi.Length > 1 Then
                        ' ファイルチェック
                        fileCheck = fileCheck + "・Aiが複数あります" + vbCrLf
                    End If
                    ' epsファイルパス
                    searchFilePathEPS = IO.Path.GetDirectoryName(printRow("作業ﾌｧｲﾙﾊﾟｽ").ToString)
                    ' epsファイル検索キー
                    searchFileKeyEPS = IO.Path.GetFileNameWithoutExtension(printRow("作業ﾌｧｲﾙﾊﾟｽ").ToString) + ".eps"
                    ' epsファイル一式
                    filesEPS = IO.Directory.GetFiles(searchFilePathEPS, searchFileKeyEPS, IO.SearchOption.AllDirectories)

                    ' epsファイル一式が一つ以外
                    If filesEPS.Length <> 1 And fileCheck = "OK" Then
                        ' ファイルチェック
                        fileCheck = searchFilePathEPS + "フォルダ内" + vbCrLf + vbCrLf + "作成書：" + Label27.Text + vbCrLf + vbCrLf + searchFileKeyAi + vbCrLf + vbCrLf
                    End If

                    ' epsファイルなし
                    If filesEPS.Length = 0 Then
                        ' ファイルチェック
                        fileCheck = fileCheck + "・EPSが見つかりません" + vbCrLf
                    ElseIf filesEPS.Length > 1 Then
                        ' ファイルチェック
                        fileCheck = fileCheck + "・EPSが複数あります" + vbCrLf
                    End If

                    ' ファイルOK
                    If fileCheck <> "OK" Then
                        Exit For
                    End If
                Next
                ' ファイルOK
                If fileCheck <> "OK" Then
                    Exit For
                End If
            Next

            ' itemDB更新
            updateItemResult = UpdateItemDb(resultItemTable)
            ' printDB更新
            updatePrintResult = UpdatePrintDb(resultPrintTable)

            ' item更新成功
            If updateItemResult <> "error" Then
                ' DataGridView1流し込み
                DataGridView1.DataSource = resultItemTable
            End If
            ' print更新成功
            If updatePrintResult <> "error" Then
                ' printDB更新
                finalPrintTable = PrintDataTableModify(resultPrintTable)
                ' DataGridView2流し込み
                DataGridView2.DataSource = finalPrintTable
            End If

            ' ファイルOK
            If fileCheck <> "OK" Then
                ' ファイルチェック
                MsgBox(fileCheck)
            Else
                ' ファイルOK
                Button14.Text = "CheckOK"
            End If

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：B8
    '*関数名：ExchangeJpgButton
    '*機  能：Jpg変換ボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub ExchangeJpgButton(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        ' 変数定義
        Dim ymd As String ' 年月日
        Dim fontInfo As String ' 画像ファイル名
        Dim sake As String ' 商品名
        Dim shipDir As String ' Sドライブ発送日フォルダ
        Dim baseDir As String ' Lドライブaiファイル
        Dim outputPath As String ' 出力ファイルパス
        Dim jpegFileName As String ' 一時ファイル名
        Dim sakuseisyoTxtPath As String ' 作成書ファイルフルパス名
        Dim sakuseisyoDir As String ' 作成書ファイルフォルダ
        Dim checkAiDirFlag As String ' チェックフラグ
        Dim searchFilePathEPS As String ' epsファイルパス
        Dim searchFileKeyEPS As String ' epsファイルパス
        Dim workingDir As String ' 枚数確認
        Dim updateItemResult As String ' item更新結果
        Dim updatePrintResult As String ' print更新結果
        ' 配列定義
        Dim filesEPS() As String ' epsファイル一覧
        ' オブジェクト定義
        Dim resultDataSet As DataSet ' データセット
        Dim resultItemTable As DataTable ' itemテーブル
        Dim resultPrintTable As DataTable ' printテーブル
        Dim finalPrintTable As DataTable ' 最終printテーブル

        Try
            ' DBインスタンス作成
            Dim dbMaker = New Db()
            ' DB取得
            resultDataSet = GetAllDbData(globalNowDate)
            ' itemDB取得
            resultItemTable = resultDataSet.Tables(0)
            ' printDB取得
            resultPrintTable = resultDataSet.Tables(1)

            ' 管理アカウントのみ実行
            If globalUserName <> adminUser OrElse Button14.Text <> "CheckOK" Then
                ' データなし
                MsgBox("管理ユーザ専用機能です")
                Throw New System.Exception("An exception has occurred.")
            End If
            ' 年月日取得
            ymd = resultItemTable.Rows(0).ToString.Substring(0, 6)
            ' Sドライブ発送日フォルダ取得
            shipDir = GetAppPath() + "\work\" + ymd + "\"
            ' Lドライブaiファイルコピー先日付フォルダベースディレクトリ
            baseDir = testLabelHDPath + ymd.Substring(0, 2) + "年度ラベル作成済み分\20" + ymd.Substring(0, 4) + "\" + ymd.Substring(2, 4)

            ' Lドライブaiファイルコピー先日付フォルダ確認
            If Not IO.Directory.Exists(baseDir) Then
                ' Lドライブaiファイルコピー先日付フォルダ確認
                IO.Directory.CreateDirectory(baseDir)
            End If
            ' 日付フォルダ＞★仕分け済み内「★jpg」フォルダ確認
            If Not IO.Directory.Exists(shipDir + "★仕分け済み\★jpg\") Then
                ' 日付フォルダ＞★仕分け済み内「★jpg」フォルダ生成
                IO.Directory.CreateDirectory(shipDir + "★仕分け済み\★jpg\")
            End If

            ' itemDataTableループ
            For Each itemRow As DataRow In resultItemTable.Rows
                ' 作成書ファイルフルパス名取得
                sakuseisyoTxtPath = itemRow("colFullPath").ToString
                ' 作成書ファイルフォルダ取得
                sakuseisyoDir = itemRow("colDir").ToString
                ' チェックフラグ初期化
                checkAiDirFlag = ""
                ' printDB取得
                resultPrintTable = dbMaker.Sql_select("Select * from printtable where ID = " + itemRow("itemID").ToString)
                ' printDataが空
                If resultPrintTable Is Nothing OrElse resultPrintTable.Rows.Count = 0 Then
                    ' データなし
                    Console.WriteLine("データなし")
                    Continue For
                End If
                ' printDataTableループ
                For Each printRow As DataRow In resultPrintTable.Rows
                    ' 初期化
                    sake = ""
                    jpegFileName = ""
                    outputPath = ""
                    fontInfo = ""
                    workingDir = printRow("作業ﾌｧｲﾙﾊﾟｽ").ToString
                    ' illustratorインスタンス作成
                    Dim adobeMaker = New adobeAI()
                    ' 枚数確認
                    If CInt(printRow("枚数")) = 0 OrElse workingDir.Length = 0 Then
                        Console.WriteLine("枚数無し")
                        Continue For
                    End If
                    ' epsファイルパス
                    searchFilePathEPS = IO.Path.GetDirectoryName(workingDir)
                    ' eps検索キー
                    searchFileKeyEPS = IO.Path.GetFileNameWithoutExtension(workingDir) + ".eps"
                    ' epsファイル一覧
                    filesEPS = IO.Directory.GetFiles(searchFilePathEPS, searchFileKeyEPS, IO.SearchOption.AllDirectories)
                    ' 検索に1件ヒット」
                    If filesEPS.Length <> 1 Then
                        Console.WriteLine("EPSなし")
                        Continue For
                    End If
                    ' フラグ空欄
                    If checkAiDirFlag = "" Then
                        ' コピー先作成書フォルダ確認
                        If Not IO.Directory.Exists(baseDir + "\" + sakuseisyoDir + "\") Then
                            ' コピー先作成書フォルダ生成
                            IO.Directory.CreateDirectory(baseDir + "\" + sakuseisyoDir + "\")
                        End If
                        ' フラグOK
                        checkAiDirFlag = "OK"
                    End If
                    ' コピー
                    IO.File.Copy(workingDir, baseDir + "\" + sakuseisyoDir + "\" + IO.Path.GetFileName(workingDir), True)
                    ' 一時ファイル名
                    jpegFileName = IO.Path.GetFileName(filesEPS(0)).Replace("eps", "jpg")
                    ' 商品名
                    sake = printRow("商品名").ToString
                    sake = sake.Replace(" ", "")
                    ' 10バイト未満ならハイフン
                    While System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(sake) < 10
                        sake = sake + "-"
                    End While
                    ' アンダーバー移行除去
                    If jpegFileName.IndexOf("_") > 0 Then
                        fontInfo = jpegFileName.Substring(jpegFileName.IndexOf("_"), jpegFileName.Length - jpegFileName.IndexOf("_"))
                    Else
                        fontInfo = ".jpg"
                    End If
                    ' 最終ファイル名
                    jpegFileName = printRow("新ID").ToString + printRow("顧客番号").ToString + "-" + sake + fontInfo
                    ' 出力ファイルパス
                    outputPath = shipDir + "★仕分け済み\★jpg\" + jpegFileName
                    ' epsからjpgへ変換
                    adobeMaker.ExchangeJPG(filesEPS(0), outputPath, 600, 600)
                    ' 対象ファイル削除
                    IO.File.Delete(outputPath.Replace("-", " "))
                    ' ハイフン除去
                    IO.File.Move(outputPath, outputPath.Replace("-", " "))
                Next
            Next

            ' itemDB更新
            updateItemResult = UpdateItemDb(resultItemTable)
            ' printDB更新
            updatePrintResult = UpdatePrintDb(resultPrintTable)

            ' item更新成功
            If updateItemResult <> "error" Then
                ' DataGridView1流し込み
                DataGridView1.DataSource = resultItemTable
            End If 
            ' print更新成功
            If updatePrintResult <> "error" Then
                ' printDB更新
                finalPrintTable = PrintDataTableModify(resultPrintTable)
                ' DataGridView2流し込み
                DataGridView2.DataSource = finalPrintTable
            End If 
            ' 終了メッセージ
            MsgBox("変換終了")

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：B9
    '*関数名：InitUiButton
    '*機  能：UI初期化（×）ボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub InitUiButton(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        ComboBox1.SelectedItem = ""
        ComboBox2.SelectedItem = ""
        TextBox10.Text = ""
    End Sub

    '--------------------------------------------------------
    '*番  号：B10
    '*関数名：UpdateUiButton
    '*機  能：UI更新ボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub UpdateUiButton(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click

        Dim updateItemResult As String ' item更新結果
        ' オブジェクト定義
        Dim resultItemTable As DataTable ' itemテーブル

        Try
            ' DataTableセット
            resultItemTable = InstructionsFilter()
            ' itemDB更新
            updateItemResult = UpdateItemDb(resultItemTable)
            ' item更新成功
            If updateItemResult <> "error" Then
                ' DataGridView1流し込み
                DataGridView1.DataSource = resultItemTable
            End If 

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：B16
    '*関数名：OpenWorkButton
    '*機  能：Work開くボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub OpenWorkButton(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        ' 変数定義
        Dim psi As New System.Diagnostics.ProcessStartInfo() ' psi

        Try
            psi.FileName = GetAppPath() + "/Work"
            psi.Verb = "explore"
            System.Diagnostics.Process.Start(psi)

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：B16
    '*関数名：LabelOkButton
    '*機  能：ラベルOKボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    ' Private Sub LabelOkButton(sender As System.Object, e As System.EventArgs) Handles Button16.Click
    '  Try
    '   If ComboBox3.Items.Count = 0 Then
    '     Exit Sub
    '   End If
    '   ' 配置PB：OK
    '   Dim currentJPG As String = ComboBox3.SelectedItem.ToString
    '   Dim currentIndex As Integer = ComboBox3.SelectedIndex
    '   Dim currentPath As String = IO.Path.GetDirectoryName(currentJPG)
    '   Dim jpgMaxCount As Integer = ComboBox3.Items.Count
    '   Dim jpegFiles As String() = IO.Directory.GetFiles(currentPath, "*.jpg")

    '   If currentIndex < jpgMaxCount - 1 Then
    '     ComboBox3.SelectedIndex = currentIndex + 1
    '     PictureBox1.Image = Image.FromFile(jpegFiles(ComboBox3.SelectedIndex))
    '   Else
    '     MessageBox.Show("最後です")
    '   End If

    '   Catch ex As System.IO.IOException
    ' 	  Console.WriteLine(ex)
    ' 	End Try
    ' End Sub

    '--------------------------------------------------------
    '*番  号：B17
    '*関数名：LabelNgButton
    '*機  能：ラベルNGボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    ' Private Sub LabelNgButton(sender As System.Object, e As System.EventArgs) Handles Button16.Click
    '  Try
    '   If ComboBox3.Items.Count = 0 Then
    '     Exit Sub
    '   End If
    '   ' 配置PB：OK
    '   Dim currentJPG As String = ComboBox3.SelectedItem.ToString
    '   Dim currentIndex As Integer = ComboBox3.SelectedIndex
    '   Dim currentPath As String = IO.Path.GetDirectoryName(currentJPG)
    '   Dim jpgMaxCount As Integer = ComboBox3.Items.Count
    '   Dim jpegFiles As String() = IO.Directory.GetFiles(currentPath, "*.jpg")

    '   If currentIndex < jpgMaxCount - 1 Then
    '     ComboBox3.SelectedIndex = currentIndex + 1
    '     PictureBox1.Image = Image.FromFile(jpegFiles(ComboBox3.SelectedIndex))
    '   Else
    '     MessageBox.Show("最後です")
    '   End If

    '   Catch ex As System.IO.IOException
    ' 	  Console.WriteLine(ex)
    ' 	End Try
    ' End Sub

    ' DataGridView2
    '--------------------------------------------------------
    '*番  号：B11
    '*関数名：SelectAllButton
    '*機  能：全選択ボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub SelectAllButton(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        ' オブジェクト定義
        Dim resultTables As DataSet ' データセット
        Dim finalPrintTable As DataTable ' finalprintテーブル
        Dim resultPrintTable As DataTable ' printテーブル

        Try
            ' DB取得
            resultTables = GetAllDbData(globalNowDate)
            ' printDB取得
            resultPrintTable = resultTables.Tables(1)

            ' printDataが空
            If resultPrintTable Is Nothing OrElse resultPrintTable.Rows.Count = 0 Then
                ' データなし
                Console.WriteLine("データなし")
            End If

            ' printDataTableループ
            For Each printRow As DataRow In resultPrintTable.Rows
                ' 全選択
                If Button10.Text = "全選択" Then
                    Button10.Text = "全解除"
                Else
                    Button10.Text = "全選択"
                End If
                ' 出力がOK以外なら次ループ
                If printRow("出力").ToString <> "OK" Then
                    Continue For
                End If
                If Button10.Text = "全選択" AndAlso CInt(printRow("枚数")) > 0 Then
                    printRow("部分配置") = True
                    printRow("部分面付") = True
                ElseIf Button10.Text <> "全選択" Then
                    printRow("部分配置") = False
                    printRow("部分面付") = False
                End If
            Next

            ' printDB更新
            UpdatePrintDb(resultPrintTable)
            ' printDB更新
            finalPrintTable = PrintDataTableModify(resultPrintTable)
            ' DataGridView2流し込み
            DataGridView2.DataSource = finalPrintTable

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：B12
    '*関数名：OutputPartLayout
    '*機  能：部分配置出力ボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub OutputPartLayout(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        ' 変数定義
        Dim cnt As Integer ' 部分配置カウンタ
        Dim selCount As Integer ' 部分配置カウンタ
        Dim counter As Integer ' 最終インデックス
        Dim selRowsLast As Integer ' インデックス保存
        Dim tmpDepartureDate As String ' 変更チェック
        Dim tmpInstruction As String ' 指示書をコピー
        Dim checkSheetName As String ' 配置ﾌｧｲﾙ名称
        Dim searchFilePathSelCheck As String ' 部分配置検索
        Dim searchFileKeySelCheck As String ' 部分配置検索パス
        Dim tmpStrSelCheckSheetLast As String ' 配置ﾌｧｲﾙ
        Dim tmpStrSelCheckSheetCnt As String ' 配置ﾌｧｲﾙ名称
        Dim dtNow As DateTime ' 時刻
        Dim tsNow As TimeSpan ' 時刻の部分
        ' 配列定義
        Dim filesSelCheck As String() ' 取得ファイル一式
        Dim tmpStrSelCheckSheetCntArr As String() ' カウンタ一式
        Dim selRows As List(Of Integer) ' インデックス保存
        ' オブジェクト定義
        Dim resultItemTable As New DataTable ' itemテーブル
        Dim resultPrintTable As New DataTable ' printテーブル
        Dim resultDataSet As New DataSet ' データセット
        Try
            selCount = 0
            selRowsLast = 0
            tmpDepartureDate = ""
            tmpInstruction = ""
            dtNow = DateTime.Now
            selRows = New List(Of Integer)
            ' 時刻の部分だけを取得する
            tsNow = dtNow.TimeOfDay
            ' DB取得
            resultDataSet = GetAllDbData(TextBox11.Text)
            ' itemDB取得
            resultItemTable = resultDataSet.Tables(0)
            ' printDB取得
            resultPrintTable = resultDataSet.Tables(1)

            ' printDataTableループ
            For Each tmpPrintRow As DataRow In resultPrintTable.Rows
                ' 初期化
                searchFilePathSelCheck = ""
                searchFileKeySelCheck = ""
                tmpStrSelCheckSheetLast = ""
                tmpStrSelCheckSheetCnt = ""

                ' 部分配置true
                If CBool(tmpPrintRow("部分配置")) Then
                    ' 部分配置カウンタ
                    selCount = selCount + 1
                    ' 最終インデックス
                    selRowsLast = counter
                    ' インデックス保存
                    selRows.Add(counter)
                    ' 変更チェック
                    If tmpDepartureDate = "" AndAlso tmpPrintRow("前回発送日").ToString.Substring(0, 1) = "*" Then
                        ' 前回発送日
                        tmpDepartureDate = tmpPrintRow("前回発送日").ToString.Substring(tmpPrintRow("前回発送日").ToString.LastIndexOf("_") + 1, 5)
                        ' itemTableループ
                        For Each itemRow As DataRow In resultItemTable.Rows
                            ' 日付が一致
                            If tmpDepartureDate = itemRow("日付").ToString Then
                                ' 指示書をコピー
                                tmpInstruction = itemRow("指示書").ToString
                                Exit For
                            End If
                        Next
                    End If
                End If
                ' インクリメント
                counter = counter + 1
            Next
            
            ' 空欄ならTextBox1を取得
            If tmpInstruction = "" Then
                tmpInstruction = TextBox1.Text
            End If

            If selCount > 0 Then
                ' 配置ﾌｧｲﾙ名称宣言
                checkSheetName = "配_部分_" + Label3.Text + " " + tmpInstruction + " "
                ' 部分配置検索
                searchFilePathSelCheck = IO.Path.GetDirectoryName(resultPrintTable.Rows(selRowsLast).Item("作業ﾌｧｲﾙﾊﾟｽ").ToString)
                ' 部分配置検索パス
                searchFileKeySelCheck = checkSheetName + "*.ai"
                ' 取得ファイル一式
                filesSelCheck = IO.Directory.GetFiles(searchFilePathSelCheck, searchFileKeySelCheck, IO.SearchOption.AllDirectories)
                ' 配置ﾌｧｲﾙ名称　連番付加
                If filesSelCheck.Length = 0 Then
                    ' ファイル無しなら_1_を追加
                    checkSheetName = checkSheetName + "_1_"
                Else
                    ' ファイル有りなら_で追加
                    tmpStrSelCheckSheetLast = filesSelCheck(filesSelCheck.Length - 1).ToString
                    tmpStrSelCheckSheetCntArr = tmpStrSelCheckSheetLast.Split("_")
                    tmpStrSelCheckSheetCnt = tmpStrSelCheckSheetCntArr(tmpStrSelCheckSheetCntArr.Length - 2)
                    checkSheetName = checkSheetName + "_" + (Integer.Parse(tmpStrSelCheckSheetCnt) + 1).ToString + "_"
                End If
                ' printDataTableループ
                For Each printRow As DataRow In resultPrintTable.Rows
                    ' ai配置関数配置
                    cnt = cnt + 1
                    makeLayoutPrint(resultItemTable.Rows(Integer.Parse(TextBox11.Text)), printRow, cnt, False)
                Next
            End If

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try   
    End Sub

    '--------------------------------------------------------
    '*番  号：B13
    '*関数名：OutputPartImposition
    '*機  能：部分面付出力ボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub OutputPartImposition(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        ' 変数定義
        Dim newPageFlag As Boolean ' 初回フラグ
        ' オブジェクト定義
        Dim resultPrintTable As DataTable ' printテーブル

        Try
            ' DB取得
            resultPrintTable = GetPrintDbData(globalSelectedItem.toString())
            ' printDataTableループ
            For Each printRow As DataRow In resultPrintTable.Rows
                ' 面付け出力処理へ
                Call makeLabelPrint(printRow, True, newPageFlag)
                ' 初回フラグ
                newPageFlag = True
            Next

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try    
    End Sub

    '--------------------------------------------------------
    '*番  号：B14
    '*関数名：SearchFolder
    '*機  能：フォルダ内再検索ボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub SearchFolder(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        ' 変数定義
        Dim searchFilePath As String ' 検索パス
        Dim searchFileKey As String' サイズ取得
        ' 配列定義
        Dim files As String()' ファイル一覧
        Dim epsSizes As Double() ' 検索パス
        ' オブジェクト定義
        Dim resultTables As DataSet ' データセット
        Dim resultPrintTable As DataTable ' printテーブル

        Try
            ' DB取得
            resultTables = GetAllDbData(globalNowDate)
            ' printDB取得
            resultPrintTable = resultTables.Tables(1)
            ' printDataが空
            If resultPrintTable Is Nothing OrElse resultPrintTable.Rows.Count = 0 Then
                ' データなし
                Console.WriteLine("データなし")
            End If

            ' printDataTableループ
            For Each printRow1 As DataRow In resultPrintTable.Rows
                ' 変数初期化
                searchFilePath = ""
                searchFileKey = ""
                files = {}
                ' 検出状況が「検出成功」「手動検出」「変更変換済み」以外
                If printRow1("検出状況").ToString = "検出成功" OrElse printRow1("検出状況").ToString = "手動検出" OrElse printRow1("検出状況").ToString = "変更変換済み" Then
                    Console.WriteLine("未検出")
                    Continue For
                End If
                ' 前回発送日の最初が"*"
                If printRow1("前回発送日").ToString.Substring(0, 1) <> "*" OrElse printRow1("作業ﾌｧｲﾙﾊﾟｽ").ToString = "" Then
                    Console.WriteLine("前回発送日が*")
                    Continue For
                End If
                ' 検索パス
                searchFilePath = IO.Path.GetDirectoryName(Label27.Text).ToString
                ' 検索キー
                searchFileKey = printRow1("新ID").ToString + " " + printRow1("前回発送日").ToString + " *.ai"
                ' ファイル一覧
                files = IO.Directory.GetFiles(searchFilePath, searchFileKey, IO.SearchOption.AllDirectories)
                '検出数１個なら
                If files.Length = 1 AndAlso files(0).ToString.IndexOf("_F.ai") < 0 Then
                    ' 作業ﾌｧｲﾙﾊﾟｽ保存
                    printRow1("作業ﾌｧｲﾙﾊﾟｽ") = files(0).ToString
                    ' EPS変換用インスタンス作成
                    Dim adobeMaker1 = New adobeAI()
                    ' EPS変換用
                    adobeMaker1.ExchangeEPS(files(0).ToString, printRow1("作業ﾌｧｲﾙﾊﾟｽ").ToString)
                    ' サイズ取得用インスタンス作成
                    Dim adobeMaker2 = New adobeAI()
                    ' サイズ取得
                    epsSizes = adobeMaker2.getEpsSize(files(0).ToString, 600, 600)
                    ' 幅と高さを保存
                    printRow1("幅") = epsSizes(0)
                    printRow1("高さ") = epsSizes(1)
                    printRow1("検出状況") = "検出成功"
                ElseIf files.Length = 0 Then
                    ' 検出状況を更新
                    printRow1("検出状況") = "検出失敗"
                    ' 列のセルの背景色を水色にする
                    printRow1("背景色") = Color.Aqua
                ElseIf files.Length > 1 Then
                    ' 検出状況を更新
                    printRow1("検出状況") = "複数検出"
                    ' 列のセルの背景色を水色にする
                    printRow1("背景色") = Color.Aqua
                End If
            Next

            ' printDataTableループ
            For Each printRow2 As DataRow In resultPrintTable.Rows
                ' 出力有
                If printRow2("出力").ToString = "" Then
                    Console.WriteLine("出力有")
                    Continue For
                End If
                ' 検出状況
                If printRow2("検出状況") Is "検出成功" Or printRow2("検出状況") Is "手動検出" Then
                    ' 検出NG
                    printRow2("出力") = "NG"
                Else
                    ' 検出OK
                    printRow2("出力") = "OK"
                End If
            Next
            ' printDB更新
            UpdatePrintDb(resultPrintTable)
            ' DataGridView2流し込み
            'Call DataGridView2_Render(resultPrintTable)
            ' 終了メッセージ
            MsgBox("再検索終了")

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：B15
    '*関数名：SendAllConfirmation
    '*機  能：確認全送信ボタン
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub SendAllConfirmation(sender As Object, e As EventArgs) Handles Button15.Click
        ' 変数定義
        Dim instructionName As String ' 指示書名
        ' オブジェクト定義
        Dim resultDataSet As DataSet ' データセット
        Dim resultPrintTable As DataTable ' printテーブル

        Try
            ' PB確認希望
            If globalPBCheck = False Then
                MessageBox.Show("確認希望無効モードです")
                Exit Sub
            End If
            ' DB取得
            resultDataSet = GetAllDbData(globalNowDate)
            ' printDB取得
            resultPrintTable = resultDataSet.Tables(1)
            ' printDataが空
            If resultPrintTable Is Nothing OrElse resultPrintTable.Rows.Count = 0 Then
                ' データなし
                Console.WriteLine("データなし")
            End If
            ' 指示書名
            instructionName = Label27.Text
            ' 指示書名に「確認」が含まれる
            If instructionName.IndexOf("確認") = 0 Then
                ' データなし
                MsgBox("登録がありません")
                Throw New System.Exception("An exception has occurred.")
            End If
            ' printDataTableループ
            For Each printRow As DataRow In resultPrintTable.Rows
                ' 前回発送日の頭が
                SendConfirmation(printRow)
            Next
            ' printDB更新
            UpdatePrintDb(resultPrintTable)
            ' DataGridView2流し込み
            'Call DataGridView2_Render(resultPrintTable)
            ' 終了メッセージ
            MsgBox("確認終了")

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    ' ■ PanelDrag
    '--------------------------------------------------------
    '*番  号：D1
    '*関数名：Panel2_DragEnter
    '*機  能：txtファイルドラッグ開始
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub Panel2_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Panel2.DragEnter
        Try
            ' ファイル形式の場合のみ、ドラッグを受け付けます。
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                e.Effect = DragDropEffects.Copy
            Else
                e.Effect = DragDropEffects.None
            End If

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：D2
    '*関数名：Panel2_DragDrop
    '*機  能：txtファイルドラッグ終了
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub Panel2_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Panel2.DragDrop
        ' 変数定義
        Dim itemRs As Integer ' itemDB結果
        Dim printRs As Integer ' printDB結果
        Dim fileName As String ' 最初のファイル名
        Dim strOutPut As String ' 本日の日付
        Dim strOutDate As String ' 4桁日付
        Dim updateItemResult As String ' item更新結果
        Dim tmpFileNameWithoutExtension As String ' 拡張子無しtxtファイル名
        ' 配列定義
        Dim strFileNames As String() ' ファイルパス一覧
        ' オブジェクト定義
        Dim itemDataTable As DataTable ' itemDataTable
        Dim printDataTable As DataTable ' printDataTable
        Dim tmpPrintDataTable As DataTable ' 一時printDataTable

        Try
            ' DBインスタンス作成
            Dim dbMaker = New Db()
            ' データベースと接続
            Call dbMaker.Sql_st()
            ' ユーザ名
            globalUserName = adminUser
            ' 新規行追加OFF
            DataGridView1.AllowUserToAddRows = False
            DataGridView2.AllowUserToAddRows = False
            DataGridView1.RowHeadersVisible = False
            DataGridView2.RowHeadersVisible = False
            ' 全ファイル
            strFileNames = CType(e.Data.GetData(DataFormats.FileDrop, False), String())
            ' 最初のファイル名
            fileName = IO.Path.GetFileName(strFileNames(0).ToString)
            ' itemdataTable初期化
            itemDataTable = New DataTable("itemTable")
            ' printdataTable初期化
            printDataTable = New DataTable("printTable")

            ' ◇ 共通
            ' 本日の日付
            strOutPut = Strings.Left(fileName, 6)
            ' グローバルに格納
            globalNowDate = strOutPut
            ' 4桁日付
            strOutDate = Strings.Right(strOutPut, 4)
            ' 日付表示
            Label5.Text = Strings.Left(strOutDate, 2) + "\" + Strings.Right(strOutDate, 2)
            ' 管理アカウントのみ実行
            If globalUserName = adminUser Then
                ' NAS日付フォルダ作成
                Call MakeEmptyDir(globalRootPath + globalNowDate + "\")
            End If
            ' ローカル日付フォルダ作成
            Call MakeEmptyDir(globalLocalRootPath + globalNowDate + "\")

            ' 日付空ファイル作成
            For i As Integer = 0 To strFileNames.Length - 1
                ' 作成ディレクトリ名
                tmpFileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(strFileNames(i).ToString)
                ' ファイルの存在確認(管理アカウントのみ実行()
                If IO.File.Exists(tmpFileNameWithoutExtension) And globalUserName = adminUser Then
                    ' 存在しない場合フォルダ作成
                    Call MakeEmptyDir(globalRootPath + globalNowDate + "\" + tmpFileNameWithoutExtension + "\")
                End If
            Next
            ' itemDB作成
            itemRs = MakeItemDb(strFileNames)
            
            ' 結果あり
            If itemRs > 0 Then
                ' printDB作成
                printRs = MakePrintDb(itemRs, strFileNames)
                ' 初期itemDatatable作成
                itemDataTable = MakeItemInitialTable()
            End If
            ' 結果なし
            If printRs = 0 Then
                ' データなし
                MsgBox("結果がありません")
                Throw New System.Exception("An exception has occurred.")
            End If
            ' 対象セルを更新  
            For Each itemRow As DataRow In itemDataTable.Rows
                ' 初期printDatatable作成
                tmpPrintDataTable = MakePrintInitialTable(itemRow("itemID").ToString)
                ' 逐次追加
                printDataTable.Merge(tmpPrintDataTable)
            Next

            ' マスタDT初期化
            ' itemDB更新
           updateItemResult = UpdateItemDb(itemDataTable)
           ' item更新成功
            If updateItemResult <> "error" Then
                ' printDB更新
                UpdatePrintDb(printDataTable)
                ' データセット追加
                DataGridView1.DataSource = itemDataTable
            End If 
            ' DataGridView2描画後処理
            ' データセット追加
            DataGridView2.DataSource = printDataTable
            ' DBクローズ
            dbMaker.Sql_cl()

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    ' ■ DataGridView
    '--------------------------------------------------------
    '*番  号：G1
    '*関数名：DataGridView1_CellContentClick
    '*機  能：ボタンクリック処理
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub DataGridView1_CellContentClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        ' 変数定義
        Dim idx As Integer ' クリック行数
        Dim fileNameWithoutExtension As String ' 拡張子無しファイル名
        Dim targetRootPath As String ' 対象のルートフォルダパス
        Dim targetLocalPath As String ' 対象のローカルフォルダパス
        Dim tmpLayoutCheck As Boolean ' 配置選択
        Dim tmpImpositionCheck As Boolean ' 面付選択
        Dim tmpInstruction As String ' 指示書
        Dim tmpLayoutProgress As String ' 配置進捗
        Dim tmpEditProgress As String '指示書編集
        Dim tmpProgress As String ' 進捗
        Dim searchNullAiFileKey As String ' ai判定用
        Dim updateItemResult As String ' item更新結果
        Dim updatePrintResult As String ' print更新結果
        Dim tmpColor As Color ' カラー
        Dim tmpPbColor As Color ' PBカラー
        ' ジェネリック定義
        Dim filesNullAiCheck As String() ' nullAIチェック用
        Dim itemColumnList As List(Of String) ' itemヘッダ
        Dim itemButtonLabelList As List(Of String) ' ボタンラベルリスト
        Dim itemCheckLabelList As List(Of String) ' チェックラベルリスト
        Dim itemButtonColumnList As New List(Of String) ' ボタンカラム名
        Dim itemCheckColumnList As New List(Of String) ' チェックカラム名
        ' オブジェクト定義
        Dim tmpRows As DataRow() ' data行
        Dim psi As New System.Diagnostics.ProcessStartInfo() ' プロセス
        Dim resultDs As DataSet ' result
        Dim emptyDt As New DataTable ' データセット
        Dim newItemDt As DataTable ' itemDataTable
        Dim newPrintDt As DataTable ' printDataTable
        Dim finalPrintTable As DataTable ' 最終printテーブル

        Try
            ' DBインスタンス作成
            Dim dbMaker = New Db()
            ' DB取得
            resultDs = GetAllDbData(globalNowDate)
            ' itemDB取得
            newItemDt = resultDs.Tables(0)
            ' クリックした行
            idx = e.RowIndex
            ' itemDB取得
            newPrintDt = resultDs.Tables(1)
            ' 初期化
            filesNullAiCheck = {}
            ' ヘッダリスト
            itemColumnList = GetFixedData("itemheader")
            ' 対象ID
            tmpInstruction = newItemDt.Rows(idx).Item("指示書").ToString
            ' 対象ファイル名（拡張子無し）
            fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(tmpInstruction)
            ' 対象のルートフォルダパス
            targetRootPath = globalRootPath + globalNowDate + "\" + fileNameWithoutExtension + "\"
            ' 対象のローカルフォルダパス
            targetLocalPath = globalLocalRootPath + globalNowDate + "\" + fileNameWithoutExtension + "\"
            ' ボタンラベルリスト
            itemButtonLabelList = GetFixedData("itembutton")
            ' チェックラベルリスト
            itemCheckLabelList = GetFixedData("itemcheck")
            ' 表示ボタン列追加
            For i As Integer = 0 To itemButtonLabelList.Count - 1
                ' 表示対象のカラムを格納
                If itemButtonLabelList(i) <> "" Then
                    itemButtonColumnList.Add(itemColumnList(i))
                End If
            Next
            ' 表示チェック列追加
            For j As Integer = 0 To itemCheckLabelList.Count - 1
                ' 表示対象のカラムを格納
                If itemCheckLabelList(j) <> "" Then
                    itemCheckColumnList.Add(itemColumnList(j))
                End If
            Next
            ' 行あり
            If newItemDt.Rows.Count = 0 OrElse idx < 0 Then
                MsgBox("結果がありません")
                Throw New System.Exception("An exception has occurred.")
            End If
            ' 配置進捗
            tmpLayoutProgress = newItemDt.Rows(idx).Item("配置進捗").ToString
            ' 指示書編集
            tmpEditProgress = newItemDt.Rows(idx).Item("指示書編集").ToString
            ' 進捗
            tmpProgress = newItemDt.Rows(idx).Item("進捗").ToString

            ' 列によって処理
            Select Case DataGridView1.Columns(0).Name
                ' フォルダ作成ボタン押下
                Case itemButtonColumnList(0)
                    ' ラベルの編集モード表示を削除更新
                    Label11.Text = ""
                    Label27.Text = ""
                    ' フォルダ作成
                    Call MakeEmptyDir(targetRootPath)
                    ' セル更新
                    newItemDt.Rows(idx).Item("配置進捗") = "　"
                    MsgBox("フォルダ作成完了")

                ' フォルダ開くボタン押下
                Case itemButtonColumnList(1)
                    ' フォルダ開く
                    If System.IO.Directory.Exists(targetRootPath) Then
                        ' エクスプローラで開く
                        psi.FileName = IO.Path.GetDirectoryName(targetRootPath)
                        psi.Verb = "explore"
                        System.Diagnostics.Process.Start(psi)
                    End If

                ' 指示書編集ボタン押下
                Case itemButtonColumnList(2)
                    ' 配置進捗が完了でない
                    If tmpLayoutProgress <> "完了" Then
                        MsgBox("配置を作成してください")
                        Exit Sub
                    End If
                    ' 編集モードでなければ編集モードに入る
                    If tmpEditProgress <> "編集中" Then
                        MsgBox("編集モードでないです")
                        Exit Sub
                    End If
                    ' 選択中ID更新
                    TextBox11.Text = CStr(idx + 1)
                    ' 編集モード
                    tmpProgress = newItemDt.Rows(idx).Item("進捗").ToString
                    ' 配置進捗
                    tmpLayoutProgress = newItemDt.Rows(idx).Item("配置進捗").ToString
                    ' 指示書編集
                    tmpEditProgress = newItemDt.Rows(idx).Item("指示書編集").ToString
                    ' ラベルを全選択へ更新
                    Button11.Text = "全選択"

                    ' 現在モード　ラベルチェック
                    If Button11.Text = "[編集]" Then
                        ' ラベルの編集モード表示を削除更新
                        Button11.Text = ""
                        ' 修正作業
                        For Each printRow As DataRow In newPrintDt.Rows
                            ' 変更aiファイル判定初期化
                            searchNullAiFileKey = ""
                            ' 枚数ゼロならキャンセル
                            If printRow("枚数").ToString = "0" Then
                                ' エラー表示
                                printRow("printID") = "ｷｬﾝｾﾙ"
                                ' ファイル更新
                                MyVariable.printButtonIndexes.fileUpdateIdx = 1
                                ' 確認希望
                                MyVariable.printButtonIndexes.needCheckIdx = 1
                                ' 変更aiファイル判定ループ
                                If printRow("検出状況").ToString = "変更前Ai" Then
                                    ' 検索キー
                                    searchNullAiFileKey = printRow("新ID").ToString + " " + printRow("顧客番号").ToString  + " *.ai"
                                    ' チェック結果8
                                    filesNullAiCheck = IO.Directory.GetFiles(globalRootPath + "\Work\" + globalNowDate + "\" + newItemDt.Rows(idx).Item("指示書").ToString, searchNullAiFileKey)
                                    ' aiチェック
                                    newItemDt.Rows(idx).Item("変更前") = filesNullAiCheck.length.toString()
                                    ' ヒットなし
                                    If filesNullAiCheck(0).IndexOf("F.ai") = 0 Then
                                        ' 変更済みAiに
                                        printRow("検出状況") = "変更済みAi"
                                        ' 対象ファイルのパス
                                        printRow("作業ﾌｧｲﾙﾊﾟｽ") = filesNullAiCheck(0)
                                        ' OKを増やす
                                        newItemDt.Rows(idx).Item("検出OK") = Integer.Parse(newItemDt.Rows(idx).Item("検出OK").ToString) + 1
                                        ' NGを減らす
                                        newItemDt.Rows(idx).Item("検出NG") = Integer.Parse(newItemDt.Rows(idx).Item("検出NG").ToString) - 1
                                    End If
                                End If
                            End If
                        Next

                        ' 選択行進捗とラジオボタンの表示同期
                        Select tmpProgress
                            Case ""
                                RadioButton1.Checked = True
                            Case "未処理"
                                RadioButton1.Checked = True
                            Case "検索済"
                                RadioButton2.Checked = True
                            Case "配置出力済"
                                RadioButton3.Checked = True
                            Case "配置確認済"
                                RadioButton4.Checked = True
                            Case "面付出力済"
                                RadioButton5.Checked = True
                            Case "確認待ち"
                                RadioButton6.Checked = True
                        End Select

                        ' ラベルを編集モードへ更新
                        Label11.Text = "[編集]"

                        ' 指示書名が一致ならチェック
                        If newItemDt.Rows(idx).Item("指示書") Is tmpInstruction Then
                            ' 配置選択ボタンにチェック
                            newItemDt.Rows(idx).Item("配置選択") = True
                            ' 指示書編集ボタンを「編集中」に
                            newItemDt.Rows(idx).Item("指示書編集") = "編集中"
                            ' 選択行をイエローでハイライト
                            newItemDt.Rows(idx).Item("背景色") = Color.Yellow
                            ' ラベルに指示書名を表示
                            Label27.Text = tmpInstruction
                        End If
                        ' ソート禁止
                        For Each c As DataGridViewColumn In DataGridView2.Columns
                            c.SortMode = DataGridViewColumnSortMode.NotSortable
                        Next

                    Else
                        ' ラベルの編集モード表示を削除更新
                        Label11.Text = ""
                        Label27.Text = ""
                        ' 指示書編集を編集に
                        newItemDt.Rows(idx).Item("指示書編集") = "編集"

                        ' 選択行ハイライトを戻す
                        For Each itemRow As DataRow In newItemDt.Rows
                            ' 指示書名が一致ならチェック
                            If itemRow("指示書") Is tmpInstruction Then
                                ' 配置選択チェック外し
                                itemRow("配置選択") = False
                                ' 編集中ボタンを編集へ更新
                                itemRow("指示書編集") = "編集"
                            End If
                        Next
                    End If

                    ' 編集中ボタンを編集へ更新
                    newItemDt.Rows(idx).Item("指示書編集") = "編集"
                    ' 進捗に応じて色を変える
                    Select Case tmpProgress
                        Case ""
                            tmpColor = Color.White
                        Case "未処理"
                            tmpColor = Color.White
                        Case "検索済"
                            tmpColor = Color.LightGray
                        Case "配置出力済"
                            tmpColor = Color.LightCyan
                        Case "配置確認済"
                            tmpColor = Color.LightBlue
                        Case "面付出力済"
                            tmpColor = Color.LightGreen
                        Case "確認待ち"
                            tmpColor = Color.LightCoral
                    End Select

                    ' PB確認希望 進捗に応じて色を変える
                    Select Case newItemDt.Rows(idx).Item("PB確認希望").ToString
                        Case ""
                            tmpPbColor = Color.White
                        Case "待ち"
                            tmpPbColor = Color.Yellow
                        Case "済×"
                            tmpPbColor = Color.LightCoral
                        Case "済〇"
                            tmpPbColor = Color.LightBlue
                        Case "面付出力済"
                            tmpPbColor = Color.LightGreen
                    End Select
                    ' 背景色変更
                    newItemDt.Rows(idx).Item("背景色") = tmpPbColor
                    ' グローバルに選択IDを保存
                    globalSelectedItem = idx + 1

                ' 作業用DLボタン押下
                Case itemButtonColumnList(3)
                    ' 存在しない場合フォルダ作成
                    Call MakeEmptyDir(targetLocalPath)

                ' 配置選択
                Case itemCheckColumnList(0)
                    ' ラベルの編集モード表示を削除更新
                    Label11.Text = ""
                    Label27.Text = ""
                    ' 配置選択
                    tmpLayoutCheck = CBool(newItemDt.Rows(idx).Item("配置選択"))
                    ' 重複数
                    tmpRows = newItemDt.Select("指示書" + " = " + tmpInstruction)
                    ' 重複有
                    If tmpRows.Length > 0 Then
                        ' 対象チェックを外す
                        For Each row As DataRow In tmpRows
                            row("配置選択") = Not tmpLayoutCheck
                        Next
                    End If
                    ' 配置選択がON
                    If tmpLayoutCheck Then
                        ' 選択数カウント表示更新
                        Label22.Text = (Integer.Parse(Label22.Text) - 1).ToString
                    Else
                        ' 配置進捗が「完了」
                        If tmpLayoutProgress = "完了" Then
                            MsgBox("配置済みです")
                        Else
                            ' 選択数カウント表示更新
                            Label23.Text = (Integer.Parse(Label23.Text) - 1).ToString
                        End If
                    End If

                ' 面付選択
                Case itemCheckColumnList(1)
                    ' ラベルの編集モード表示を削除更新
                    Label11.Text = ""
                    Label27.Text = ""
                    ' 面付選択
                    tmpImpositionCheck = CBool(newItemDt.Rows(idx).Item("配置選択"))
                    ' 重複数
                    tmpRows = newItemDt.Select("指示書" + " = " + tmpImpositionCheck.ToString)
                    ' 重複有
                    If tmpRows.Length > 0 Then
                        ' 対象チェックを外す
                        For Each row As DataRow In tmpRows
                            row("面付選択") = Not tmpImpositionCheck
                        Next
                    End If
                    ' 配置選択がON
                    If tmpImpositionCheck Then
                        ' 選択数カウント表示更新
                        Label23.Text = (Integer.Parse(Label23.Text) - 1).ToString
                    Else
                        ' 配置進捗が「完了」
                        If tmpLayoutProgress = "完了" Then
                            MsgBox("配置を作成してください")
                        Else
                            ' 選択数カウント表示更新
                            Label23.Text = (Integer.Parse(Label23.Text) + 1).ToString
                        End If
                    End If

                Case Else
                    Console.WriteLine("Error: エラーです")
            End Select

            ' itemDB更新
            updateItemResult = UpdateItemDb(newItemDt)
            ' printDB更新
            updatePrintResult = UpdatePrintDb(newPrintDt)

            ' item更新成功
            If updateItemResult <> "error" Then
                ' DataGridView1流し込み
                DataGridView1.DataSource = newItemDt
            End If 
            ' print更新成功
            If updatePrintResult <> "error" Then
                ' printDB更新
                finalPrintTable = PrintDataTableModify(newPrintDt)
                ' DataGridView2流し込み
                DataGridView2.DataSource = finalPrintTable
            End If

             ' ラベルの編集モード表示を削除更新
            Label11.Text = ""
            Label27.Text = ""

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：G2
    '*関数名：DataGridView2_CellContentClick
    '*機  能：ボタンクリック処理
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        ' 変数定義
        Dim idx As Integer ' クリックした行
        Dim tmpId As String ' 該当ID
        Dim expression As String ' 選択式
        Dim updateItemResult As String ' item更新結果
        Dim updatePrintResult As String ' print更新結果
        ' ジェネリック定義
        Dim printButtonLabelList As List(Of String) ' ボタン列リスト
        Dim printCheckLabelList As List(Of String) ' チェック列リスト
        Dim printButtonColumnList As New List(Of String) ' ボタン列リスト
        Dim printCheckColumnList As New List(Of String) ' チェック列リスト
        ' オブジェクト定義
        Dim resultTables As DataSet ' データセット
        Dim newItemDt As DataTable ' itemテーブル
        Dim newPrintDt As DataTable ' printテーブル
        Dim tmpNewPrintDt As DataTable ' 一時printテーブル
        Dim finalPrintTable As DataTable ' 最終printテーブル
        Dim printDr As DataRow ' print行
        Dim OpenFileDialog1 As New OpenFileDialog() ' ダイアログ

        Try
            ' illustratorインスタンス作成
            Dim adobeMaker = New adobeAI()
            ' DB取得
            resultTables = GetAllDbData(globalNowDate)
            ' itemDB取得
            newItemDt = resultTables.Tables(0)
            ' printDB取得
            tmpNewPrintDt = resultTables.Tables(1)
            ' クリックした行
            idx = e.RowIndex
            ' 対象printRowデータ
            printDr = tmpNewPrintDt.Rows(idx + 1)
            ' ボタン列リスト
            printButtonLabelList = GetFixedData("printbutton")
            ' チェック列リスト
            printCheckLabelList = GetFixedData("printcheck")

            ' ボタン列リスト
            For i As Integer = 0 To printButtonLabelList.Count - 1
                ' 表示対象あり
                If printButtonLabelList(i) <> "" Then
                    ' カラム格納
                    printButtonColumnList.Add(printButtonLabelList(i))
                End If
            Next

            ' チェック列リスト
            For j As Integer = 0 To printCheckLabelList.Count - 1
                ' 表示対象あり
                If printCheckLabelList(j) <> "" Then
                    ' カラム格納
                    printCheckColumnList.Add(printCheckLabelList(j))
                End If
            Next

            ' データあり
            If DataGridView2.Rows.Count > 0 AndAlso idx >= 0 Then
                Console.WriteLine("clicked: " + DataGridView2.Columns(idx).Name)
                ' 列により分岐
                Select Case DataGridView2.Columns(0).Name
                    ' ファイル更新押下
                    Case printButtonColumnList(0)
                        ' 該当ID
                        tmpId = printDr("printID").ToString
                        ' 選択式
                        expression = "printID = " + CStr(Integer.Parse(tmpId) + 1)
                        ' 抽出データ
                        newPrintDt = tmpNewPrintDt.Select(expression).CopyToDataTable()
                        ' ファイル更新
                        If newItemDt.Rows(0).Item("ファイル更新").ToString = "更新" Then
                            ' 背景色変更
                            newItemDt.Rows(0).Item("背景色") = Color.Yellow
                            ' 前回発送日の最初が*
                            If newItemDt.Rows(0).Item("前回発送日").ToString.Substring(0, 1) <> "*" Then
                                ' チェック無し
                                If Not CheckBox8.Checked Then
                                    'タイトル 
                                    OpenFileDialog1.Title = "顧客番号：[" + newItemDt.Rows(0).Item("顧客番号").ToString + "]  酒類：[" + newItemDt.Rows(0).Item("商品名").ToString + "] の更新するファイルを選択してください。"
                                    OpenFileDialog1.Filter = "aiファイル|*.ai|すべて|*.*"
                                    OpenFileDialog1.FilterIndex = 1
                                End If
                            End If
                            MsgBox("ファイル更新完了")
                        End If

                    ' AiOpen押下
                    Case printButtonColumnList(1)
                        If printDr("検出ﾌｧｲﾙﾊﾟｽ").ToString.Length > GetAppPath().ToString.Length Then
                            adobeMaker.Open(printDr("検出ﾌｧｲﾙﾊﾟｽ").ToString)
                        End If
                        
                    ' 確認希望押下
                    Case printButtonColumnList(2)
                        ' PB確認希望
                        SendConfirmation(printDr)
                        MsgBox("確認希望完了")

                    ' 部分配置
                    Case printCheckColumnList(0)
                        ' 部分配置
                        If printDr("出力").ToString = "OK" Then
                            printDr("部分配置") = CBool(Not CBool(printDr("部分配置"))).ToString
                        End If

                    ' 部分面付
                    Case printCheckColumnList(1)
                        ' 部分面付
                        If printDr("出力").ToString = "OK" Then
                            printDr("部分面付") = CBool(Not CBool(printDr("部分面付"))).ToString
                        End If

                    Case Else
                        Console.WriteLine("")
                End Select
            End If
            ' itemDB更新
            updateItemResult = UpdateItemDb(newItemDt)
            ' printDB更新
            updatePrintResult = UpdatePrintDb(tmpNewPrintDt)
            
            ' item更新成功
            If updateItemResult <> "error" Then
                ' DataGridView1流し込み
                DataGridView1.DataSource = newItemDt
            End If 
            ' print更新成功
            If updatePrintResult <> "error" Then
                ' printDB更新
                finalPrintTable = PrintDataTableModify(tmpNewPrintDt)
                ' DataGridView2流し込み
                DataGridView2.DataSource = finalPrintTable
            End If

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    ' ■ 汎用関数
    '--------------------------------------------------------
    '*番  号：M1
    '*関数名：PrintDataTableModify
    '*機  能：DataGridView2描画
    '*分　類：プライベート
    '--------------------------------------------------------
    Private Function PrintDataTableModify(ByRef printDt As DataTable) As DataTable
        ' 変数定義
        Dim selectedId As Integer ' 選択中のitemID
        Dim searchNullAiFileKey As String ' AIファイル検索キー
        Dim tmpColor As Color ' カラー
        Dim filesNullAiCheck() As String ' nullチェック
        ' オブジェクト定義
        Dim newItemDt As New DataTable ' itemDataTable
        Dim emptyDt As New DataTable ' データセット

        Try
            ' 初期化
            filesNullAiCheck = {}
            ' 選択中のitemID
            selectedId = Integer.Parse(TextBox11.Text)

            ' 修正作業
            For Each printRow1 As DataRow In printDt.Rows
                ' 枚数ゼロならキャンセル
                If printRow1("枚数").ToString = "0" Then
                    ' エラー表示
                    printRow1("printID") = "ｷｬﾝｾﾙ"
                    ' ファイル更新
                    MyVariable.printButtonIndexes.fileUpdateIdx = 1
                End If
                
                ' PB確認希望でIDの開始文字が"*""
                If Label27.Text.IndexOf("確認") > 0 AndAlso globalPBCheck AndAlso printRow1("printID").ToString.Substring(0, 1) = "*" Then
                    ' 確認希望
                    MyVariable.printButtonIndexes.needCheckIdx = 1
                    ' PB確認を更新
                    MyVariable.printTextIndexes.pbCheckResult = 1
                End If

                ' 変更aiファイル判定
                If printRow1("検出状況").ToString = "変更前Ai" Then
                    ' aiファイル検索キー
                    searchNullAiFileKey = printRow1("新ID").ToString + " " + printRow1("顧客番号").ToString + " *.ai"
                    ' aiファイル取得
                    filesNullAiCheck = IO.Directory.GetFiles(globalRootPath + "\Work\" + printRow1("指示書").ToString.Substring(0, 6) + "\" + printRow1("指示書").ToString, searchNullAiFileKey)
                    ' aiチェック
                    If filesNullAiCheck.Length > 0 AndAlso filesNullAiCheck(0).IndexOf("F.ai") = 0 Then
                        ' 検出状況
                        printRow1("検出状況") = "変更済みAi"
                        ' 更新
                        newItemDt.Rows(selectedId).Item("検出OK") = Integer.Parse(newItemDt.Rows(selectedId).Item("検出OK").ToString) - 1
                        newItemDt.Rows(selectedId).Item("検出NG") = Integer.Parse(newItemDt.Rows(selectedId).Item("検出NG").ToString) + 1
                    End If
                End If
            Next

            ' PB確認希望処理
            For Each printRow2 As DataRow In printDt.Rows
                ' PB確認希望で前回発送日が*起点
                If Label27.Text.IndexOf("確認") > 0 AndAlso globalPBCheck AndAlso printRow2("前回発送日").ToString.Substring(0, 1) = "*" Then
                    ' 送信
                    printRow2("確認希望") = "送信"
                    ' 列によって処理
                    Select Case printRow2("PB確認").ToString
                        Case "待ち"
                            tmpColor = Color.Yellow
                        Case "済×"
                            tmpColor = Color.LightCoral
                        Case "済〇"
                            tmpColor = Color.LightBlue
                    End Select
                    ' 作業ファイルパスが空欄
                    If printRow2("検出ﾌｧｲﾙﾊﾟｽ") Is "" Then
                        tmpColor = Color.LightCoral
                    End If
                    ' 背景色変更
                    printRow2("背景色") = tmpColor
                End If
            Next
            Return printDt

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            Return emptyDt
        End Try
    End Function
End Class
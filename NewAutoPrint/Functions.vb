' ■ Functions
Public Module Functions
    '--------------------------------------------------------
    '*番  号：M1
    '*関数名：MakeEmptyDir
    '*機  能：空フォルダ作成
    '*分  類：パブリック
    '--------------------------------------------------------
    Public Sub MakeEmptyDir(outPath As String)
        Try
            ' 存在しなければフォルダ作成
            If Not IO.Directory.Exists(outPath) Then
                ' フォルダ作成
                IO.Directory.CreateDirectory(outPath)
            End If

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：M2
    '*メソッド名：makeLabelPrint
    '*機  能：ラベル印刷
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub makeLabelPrint(printRow As DataRow, selFlg As Boolean, firstFlg As Boolean)
        ' 変数定義
        Dim x As Integer
        Dim y As Integer
        Dim dx As Integer
        Dim dy As Integer
        Dim w As Integer
        Dim h As Integer
        Dim wCntTmp As Integer
        Dim hCntTmp As Integer
        Dim pageCount As Integer
        Dim mai As Integer
        Dim cnt As Integer
        Dim wCountMax As Integer
        Dim hCountMax As Integer
        Dim workingPath As String
        Dim patternId As String
        Dim printSheetName As String
        Dim finalSavePath As String
        Dim aitemplateFilePath As String
        Dim searchFilePathEPS As String
        Dim searchFileKeyEPS As String
        Dim templatePath As String
        
        ' 配列定義
        Dim copyPathTemp1 As String()
        Dim copyPathTemp2 As String()
        Dim copyPathTemp3 As String()
        Dim filesEPS As String()
        ' オブジェクト定義
        Dim patternRow As DataRow
        Dim patternDt As DataTable
        Dim templateDt As DataTable
        Dim masterDataSet As DataSet

        Try 
            ' データセット初期化
            masterDataSet = New DataSet
            ' illustratorインスタンス作成
            Dim adobeMaker = New adobeAI()
            ' 面付けループ
            patternId = PatternRecognition(printRow("幅"), printRow("高さ"), printRow("商品名"))
            ' 該当パターンなし
            If patternId = "" Then
                ' データなし
                MsgBox("該当するサイズのラベルがありません")
                Throw New System.Exception("An exception has occurred.")
            End If
            ' マスタ取得
            masterDataSet = GetMasterData(patternId)
            ' patternMasterDB取得
            patternDt = masterDataSet.Tables(0)
            ' patternRow取得
            patternRow = patternDt.Rows(0)
            ' templateMasterDB取得
            templateDt = masterDataSet.Tables(1)
            ' 初期値設定
            x = 0
            y = 0
            wCountMax = patternRow("横枚数") ' 横に表示する最大枚数
            hCountMax = patternRow("縦枚数") ' 縦に表示する最大枚数
            w = patternRow("ｼｰﾄ幅") / wCountMax ' 一枚あたりの幅
            h = patternRow("ｼｰﾄ高") / hCountMax ' 一枚あたりの高さ
            dx = w ' 横のオフセット
            dy = h ' 縦のオフセット
            wCntTmp = 0 ' カウンタ（テンポラリ）
            hCntTmp = 0 ' カウンタ（テンポラリ）
            pageCount = 1
            ' テンプレートパス
            templatePath = templateDt.Rows(0).Item("ﾃﾝﾌﾟﾚｰﾄﾌｧｲﾙ名").ToString
            ' テンプレート名
            aitemplateFilePath = GetAppPath() + "\work\" + templatePath + ".ai"
            
            If firstFlg Then
                ' イラレドキュメントOPEN
                adobeMaker.Open(aitemplateFilePath)
                cnt = 0
                ' 開始X座標
                x = CInt(patternRow("幅開始"))
                ' 開始y座標
                y = CInt(patternRow("高さ開始"))
                ' xカウンタ（テンポラリ）
                wCntTmp = 0 
                ' yカウンタ（テンポラリ）
                hCntTmp = 0 
            End If

            ' 枚数取得
            mai = printRow("枚数")
            ' 作業ﾌｧｲﾙﾊﾟｽ
            workingPath = printRow("作業ﾌｧｲﾙﾊﾟｽ").ToString

            ' epsチェック
            If MainForm.CheckBox3.Checked = False Then
                ' 検索ファイルパス
                searchFilePathEPS = IO.Path.GetDirectoryName(workingPath)
                ' 検索ファイルパスキー
                searchFileKeyEPS = IO.Path.GetFileNameWithoutExtension(workingPath) + ".eps"
            Else
                ' 検索ファイルパス
                copyPathTemp1 = workingPath.ToString.Split("\")
                searchFilePathEPS = "C:\ebisu\TempOut\" + copyPathTemp1(copyPathTemp1.Length - 2) + "\"
                ' 検索ファイルパスキー
                copyPathTemp2 = workingPath.ToString.Split("\")
                searchFileKeyEPS = copyPathTemp2(copyPathTemp2.Length - 1).Replace(".ai", ".eps")
            End If
            ' epsチェック
            filesEPS = IO.Directory.GetFiles(searchFilePathEPS, searchFileKeyEPS, IO.SearchOption.AllDirectories)

            ' 一枚の面付け枚数ループ
            For i As Integer = 0 To mai - 1
                ' 初期化
                printSheetName = ""
                ' epsファイル配置
                adobeMaker.EPSLayout(cnt, filesEPS(0), workingPath, x - dx * wCntTmp, y - dy * hCntTmp)
                ' 上限超えで終了
                If cnt > 11 Then
                    Exit for
                End If
                ' インクリメント
                wCntTmp = wCntTmp + 1
                ' 幅上限を超えるとカウントアップ
                If wCntTmp >= wCountMax Then
                    wCntTmp = 0
                    hCntTmp = hCntTmp + 1
                End If
                ' 高さ上限を超えるとカウントアップ
                If hCntTmp >= hCountMax Then
                    hCntTmp = 0
                    ' 部分印刷
                    If selFlg Then
                        ' ラベル名
                        printSheetName = "帳_部分_" + IO.Path.GetFileName(IO.Path.GetDirectoryName(workingPath))
                    Else
                        ' ラベル名
                        printSheetName = "帳" + IO.Path.GetFileName(IO.Path.GetDirectoryName(workingPath)) + " 720 "
                    End If
                    ' テキスト配置
                    adobeMaker.TextLayout(printSheetName + " 720 " + " Page." + pageCount.ToString, 18, 77.5 + 20, y + 3)

                    ' 部分印刷
                    If selFlg Then
                        ' 保存パス
                        finalSavePath = IO.Path.GetDirectoryName(workingPath).ToString + "\" + printSheetName + "page" + pageCount.ToString + ".ai"
                    Else
                        ' 面付けaiファイル保存
                        If MainForm.CheckBox3.Checked Then
                            ' 作業ﾌｧｲﾙﾊﾟｽを分割
                            copyPathTemp3 = workingPath.Split("\")
                            ' 保存パス
                            finalSavePath = "C:\ebisu\TempOut\" + copyPathTemp3(copyPathTemp3.Length - 2) + "\" + printSheetName + " 720 " + "page" + pageCount.ToString + ".ai"
                        Else
                            ' 保存パス
                            finalSavePath = IO.Path.GetDirectoryName(workingPath).ToString + "\" + printSheetName + " 720 " + "page" + pageCount.ToString + ".ai"
                        End If
                    End If
                    ' ファイル保存
                    adobeMaker.Save(finalSavePath)
                    '面付けaiファイル出力
                    adobeMaker.PrintDocument(patternDt.Rows(0).item("宇宿ﾌﾟﾘﾝﾀID"), patternDt.Rows(0).item("山梨ﾌﾟﾘﾝﾀID"))
                    '面付けaiファイル閉じる
                    adobeMaker.Close()
                    ' カウントアップ
                    pageCount = pageCount + 1
                End If
            Next

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：M3
    '*関数名：makeLayoutPrint
    '*機  能：配置図プリント関数
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub makeLayoutPrint(itemRow As DataRow, printRow As DataRow, cnt As Integer, selFlg As Boolean)
        ' 変数定義
        Dim counter As Integer ' 配置epsカウンタ
        Dim wCount As Integer ' 幅カウンタ
        Dim hCount As Integer ' 高さカウンタ
        Dim x As Integer ' 幅設定用
        Dim y As Integer ' 高さ設定用
        Dim selCount As Integer ' カウンタ
        Dim selRowsLast As Integer ' 最後のインデックス
        Dim newPageFlag As Integer ' 新ページフラグ
        Dim tmpWidth As Integer ' 一時ラベル幅
        Dim pageCount As Integer ' 番号表示用カウンタ
        Dim fileEditor As String ' 汎用キャッシュ
        Dim tmpDetect As String ' 一時検出状況
        Dim patternId As String ' パターン番号
        Dim checkSheetName As String ' シート名
        Dim tmpLocalPath As String ' 一時作業パス
        Dim tmpLabelText As String ' 配置テキスト
        Dim tmpResultText As String ' 検出成功
        Dim tmpTopText As String ' 一時テキスト
        Dim tmpHeadText As String ' 一時テキスト
        Dim tmpSubText As String ' 一時テキスト
        Dim tmpSavePath As String ' 保存先パス
        Dim searchFilePathEPS As String ' eps検索パス
        Dim searchFileKeyEPS As String ' eps検索キー
        Dim sizeCheck As Boolean ' サイズチェック
        ' 配列定義
        Dim selRows As New List(Of Integer)
        Dim filesEPS As String() ' ファイルEPS
        ' オブジェクト定義
        Dim patternDr As DataRow ' パターン対象行
        Dim templateDt As New DataTable ' テンプレート対象行
        Dim resultItemTable As DataTable ' itemテーブル
        Dim resultDataset As New DataSet ' データ
        Dim masterDataSet As New DataSet ' マスタデータ

        Try
            ' 初期値代入
            counter = 0
            wCount = 0
            hCount = 0
            newPageFlag = 0
            pageCount = 1
            x = 70
            y = 1700
            fileEditor = ""
            ' illustratorインスタンス作成
            Dim adobeMaker = New adobeAI()
            ' DB取得
            resultDataset = GetAllDbData(MainForm.TextBox11.Text)
            ' itemDB取得
            resultItemTable = resultDataset.Tables(0)
            ' パターンID
            patternId = PatternRecognition(printRow("幅"), printRow("高さ"), printRow("商品名"))
            ' 該当パターンなし
            If patternId = "" Then
                ' データなし
                MsgBox("該当するサイズのラベルがありません")
                Throw New System.Exception("An exception has occurred.")
            End If
            ' マスタ取得
            masterDataSet = GetMasterData(patternId)
            ' patternMasterDB取得
            patternDr = masterDataSet.Tables(0).Rows(0)
            ' templateMasterDB取得
            templateDt = masterDataSet.Tables(1)
            '配置ﾌｧｲﾙ名称
            checkSheetName = "配" + itemRow("枚数")

            ' フラグあり
            If selFlg Then
                tmpWidth = 1500
                tmpDetect = printRow("検出状況")
                tmpResultText = "検出成功"
            Else
                tmpWidth = 130
                tmpDetect = printRow("出力")
                tmpResultText = "OK"
            End If
            
            ' 検出成功
            If tmpDetect = tmpResultText Then
                ' インクリメント
                selCount = selCount + 1
                selRowsLast = cnt
                selRows.Add(cnt)
                ' 作業ﾌｧｲﾙﾊﾟｽ
                tmpLocalPath = printRow("作業ﾌｧｲﾙﾊﾟｽ")

                ' 変更チェック
                If fileEditor = "" AndAlso printRow("前回発送日").Substring(0, 1) = "*" Then
                    fileEditor = tmpLocalPath.Substring(tmpLocalPath.LastIndexOf("_") + 1, 5)
                        If fileEditor = itemRow("日付") Then
                            fileEditor = itemRow("指示書")
                        End If
                End If
                ' eps検索パス
                searchFilePathEPS = IO.Path.GetDirectoryName(tmpLocalPath)
                ' 配置テキスト
                tmpLabelText = IO.Path.GetFileNameWithoutExtension(tmpLocalPath)
                ' eps検索キー
                searchFileKeyEPS = tmpLabelText + ".eps"
                ' 対象epsファイル
                filesEPS = IO.Directory.GetFiles(searchFilePathEPS, searchFileKeyEPS, IO.SearchOption.AllDirectories)
                ' epsファイル配置
                adobeMaker.EPSLayout(counter, filesEPS(0), tmpLocalPath, x + wCount * 470, y - hCount * 470)
                ' テキスト配置
                adobeMaker.TextLayout(tmpLabelText, 12, x + wCount * 470, y - hCount * 470 - 430)

                ' 通常ラベルサイズチェック
                sizeCheck = True
                If printRow("商品名").IndexOf("300NG") < 0 Then
                    If printRow("幅") < 400 AndAlso printRow("高さ") < 400 Then
                        If printRow("幅") < 314.6457 Then
                            sizeCheck = True
                        End If
                        If printRow("幅") > 320.315 Then
                            sizeCheck = True
                        End If
                        If printRow("高さ") < 265.6063 Then
                            sizeCheck = True
                        End If
                        If printRow("高さ") > 272.126 Then
                            sizeCheck = True
                        End If
                    End If
                End If

                ' 通常ラベルサイズチェックによりテキスト変更
                If selFlg And sizeCheck Then
                    tmpLabelText = "【ｻｲｽﾞ確認】" + tmpLabelText
                Else
                    tmpLabelText = IO.Path.GetFileNameWithoutExtension(tmpLocalPath)
                End If
            Else
                ' 一時テキスト
                tmpTopText = printRow("新ID") + "-" + printRow("前回発送日") + " - " + printRow("顧客番号") + " - " + printRow("商品名") + " - " + printRow("検出状況")
                ' テキスト配置
                adobeMaker.TextLayout(tmpTopText, 12, x + wCount * 470, y - hCount * 470 - 430)
            End If
            ' インクリメント
            counter = counter + 1
            wCount = wCount + 1
            ' 幅がリミット超え
            If wCount > 4 Then
                ' インクリメント
                hCount = hCount + 1
                ' カウンタリセット
                wCount = 0
            End If
            ' 高さが3以上
            If hCount > 2 Then
                ' 一時テキスト
                tmpHeadText = checkSheetName + " Page." + pageCount.ToString + "【出:" + MainForm.TextBox1.Text + "】"
                ' テキスト配置
                adobeMaker.TextLayout(tmpHeadText, 48, tmpWidth, 130)
                '配置aiファイル保存
                adobeMaker.Save(IO.Path.GetDirectoryName(itemRow("フォルダ") + "\" + checkSheetName + "page" + pageCount.ToString + ".ai"))
                '配置aiファイル出力
                adobeMaker.PrintDocument(0)
                '配置aiファイル閉じる
                adobeMaker.Close()
                ' カウンタリセット
                newPageFlag = 0
                hCount = 0
                ' インクリメント
                pageCount = pageCount + 1
            End If

            ' カウンタ3未満
            If hCount < 3 AndAlso newPageFlag = 1 Then
                ' 一時テキスト
                tmpSubText = checkSheetName + " Page." + pageCount.ToString + "【出:" + MainForm.TextBox1.Text + "】"
                ' テキスト配置
                adobeMaker.TextLayout(tmpSubText, 48, tmpWidth, 130)
                ' 保存先パス
                tmpSavePath = IO.Path.GetDirectoryName(itemRow("フォルダ")).ToString + "\" + checkSheetName + "page" + pageCount.ToString + ".ai"
                ' 配置aiファイル保存
                adobeMaker.Save(tmpSavePath)
                ' 配置aiファイルプリントアウト
                adobeMaker.PrintDocument(0)
                ' ドキュメントを閉じる
                adobeMaker.Close()
                ' 変数初期化
                newPageFlag = 0
                hCount = 0
                ' インクリメント
                pageCount = pageCount + 1
            End If

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：M4
    '*関数名：SearchAiFiles
    '*機  能：aiファイル検索関数
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub SearchAiFiles(ByRef itemRow As DataRow, ByRef printDt As DataTable)
        ' 変数定義
        Dim counter As Integer ' プログレスバー用カウンタ
        Dim tmpFileIndex As Integer ' 一時ファイルインデックス
        Dim code8StartYmdCheck As Integer ' 日付チェック用
        Dim systemStartYmdCheck As Integer ' チェック用前回発送日
        Dim tateText As String ' 文字列連結
        Dim newCode As String ' 新コード
        Dim oldCode As String ' 旧コード
        Dim tmpCustomerCode As String ' 一時顧客番号
        Dim tmpFileName As String ' 一時ファイル名
        Dim tmpProductName As String ' 一時商品名
        Dim tmpDepartureDate As String ' 一時発送日時
        Dim searchFilePath As String ' 検索ファイルパス
        Dim searchFileKey As String ' 検索ファイルキー
        Dim copyName As String ' コピー名
        Dim tmpLabelName As String ' 一時ラベル名
        Dim tmpDetectedPath As String ' 一時検索パス
        Dim tmpItemFolderPath As String ' 一時アイテムフォルダパス
        ' 配列定義
        Dim files As String() ' 新コード

        Try
            ' LabelHD検索対象
            searchFilePath = "H:\TEST\" + Strings.Right(globalNowDate, 4) + "年度ラベル作成済み分\" + globalNowDate + "\" + globalNowDate.Substring(4, 4)
            ' printDataTableループ
            For Each printRow As DataRow In printDt.Rows
                ' 初期化
                systemStartYmdCheck = 0
                code8StartYmdCheck = 0
                searchFileKey = ""
                oldCode = ""
                newCode = ""
                tateText = ""
                copyName = ""
                tmpFileName = ""
                tmpDetectedPath = ""
                tmpProductName = ""
                tmpLabelName = ""
                files = {}
                ' 顧客番号
                tmpCustomerCode = printRow("顧客番号")
                ' 前回発送日
                tmpDepartureDate = printRow("前回発送日")
                ' 空欄の場合0に
                If printRow("旧ID") = "" Then
                    printRow("旧ID") = "0"
                End If
                ' 新ID
                newCode = printRow("新ID")
                ' 商品名
                tmpProductName = printRow("商品名").Replace(" ", "")
                ' ラベル名
                tmpLabelName = printRow("ラベル名").Replace(" ", "")
                ' 対象フォルダパス
                tmpItemFolderPath = IO.Path.GetDirectoryName(itemRow.Item("フォルダ")).ToString

                ' 前回発送フォルダの有無確認
                If IO.Directory.Exists(searchFilePath) Then
                    ' 日付チェック用
                    code8StartYmdCheck = Integer.Parse(tmpDepartureDate.Substring(0, 8))
                    ' 日付が20160618より前で顧客番号が8桁
                    If code8StartYmdCheck < 20160618 And tmpCustomerCode.Length = 8 Then
                        ' 1-2文字目を削除
                        tmpCustomerCode = tmpCustomerCode.Substring(2, 6)
                    End If
                    ' ai検出
                    If Integer.Parse(printRow("旧ID")) = 0 Then
                        ' 検索キー設定
                        searchFileKey = "* " + tmpCustomerCode + " * " + tmpProductName.Replace(" ", "") + "_*.ai"
                    Else
                        ' 旧コード
                        oldCode = printRow("旧ID")
                        ' チェック用前回発送日
                        systemStartYmdCheck = Integer.Parse(tmpDepartureDate.Substring(0, 8))
                        ' 旧IDが0から始まり日付が20130115より前
                        If oldCode.Substring(0, 1) = "0" AndAlso systemStartYmdCheck < 20130115 Then
                            ' 末尾削除
                            oldCode = oldCode.Substring(1, oldCode.Length - 1)
                        End If
                        ' 検索キー設定
                        searchFileKey = oldCode + " " + tmpCustomerCode + " *.ai"
                    End If
                    ' ステータス変更
                    printRow("検出状況") = "検出中..."
                    ' ファイル検索
                    files = IO.Directory.GetFiles(searchFilePath, searchFileKey, IO.SearchOption.AllDirectories)
                    ' 検出数0で旧ID4ｹﾀ時、頭0削除し3桁で再検索
                    If files.Length = 0 AndAlso oldCode.Length = 4 AndAlso oldCode.Substring(0, 1) = 0 Then
                        ' 先頭の0を削除
                        oldCode = oldCode.Substring(1, oldCode.Length - 1)
                        ' 検索キー再設定
                        searchFileKey = oldCode + " " + tmpCustomerCode + " *.ai"
                        ' 3桁で再検索
                        files = IO.Directory.GetFiles(searchFilePath, searchFileKey, IO.SearchOption.AllDirectories)
                    End If
                    ' 検出数0で旧ID3ｹﾀ時、頭に0付加し4桁で再検索
                    If files.Length = 0 AndAlso oldCode.Length = 3 Then
                        ' 頭に0を付加
                        oldCode = "0" + oldCode
                        ' 検索キー再設定
                        searchFileKey = oldCode + " " + tmpCustomerCode + " *.ai"
                        ' 4桁で再検索
                        files = IO.Directory.GetFiles(searchFilePath, searchFileKey, IO.SearchOption.AllDirectories)
                    End If

                    ' 検出数１個なら
                    If files.Length = 1 Then
                        ' 検出状況更新
                        printRow("検出状況") = "検出成功"
                        ' 検出ファイルパス
                        printRow("検出ﾌｧｲﾙﾊﾟｽ") = files(0).ToString
                        ' 一時保存
                        tmpDetectedPath = printRow("検出ﾌｧｲﾙﾊﾟｽ")
                        ' ..ai対策
                        If tmpDetectedPath.IndexOf("..ai") > 0 Then
                            ' リネーム
                            IO.File.Move(tmpDetectedPath, tmpDetectedPath.Replace("..ai", ".ai"))
                            ' 検出ファイルパス更新
                            printRow("検出ﾌｧｲﾙﾊﾟｽ") = tmpDetectedPath.Replace("..ai", ".ai")
                        End If
                        ' 空白対策
                        If tmpDetectedPath.IndexOf("      ") > 0 Then
                            ' リネーム
                            IO.File.Move(tmpDetectedPath, tmpDetectedPath.Replace("      ", ""))
                            ' 検出ファイルパス更新
                            printRow("検出ﾌｧｲﾙﾊﾟｽ") = tmpDetectedPath.Replace("      ", "")
                        End If
                        ' 一時ファイル名
                        tmpFileName = IO.Path.GetFileName(printRow("検出ﾌｧｲﾙﾊﾟｽ"))
                        ' 一時ファイルインデックス
                        tmpFileIndex = tmpFileName.IndexOf(" ", 5)
                        ' コピー
                        copyName = tmpItemFolderPath + "\" + newCode + " " + tmpCustomerCode + tmpFileName.Substring(tmpFileName, tmpFileName.Length - tmpFileIndex)
                        ' 既に同名のファイルが存在していても上書
                        IO.File.Copy(printRow("検出ﾌｧｲﾙﾊﾟｽ"), copyName, True)
                        ' 再設定
                        printRow("検出ﾌｧｲﾙﾊﾟｽ") = copyName
                        ' 検出OK
                        itemRow("検出OK") = itemRow("検出OK") + 1

                    ElseIf files.Length = 0 Then
                        ' ステータス更新
                        printRow("検出状況") = "検出失敗"
                        ' 検出NG
                        itemRow("検出NG") = itemRow("検出NG") + 1
                        '列のセルの背景色を薄桃色に
                        itemRow("背景色") = Color.LightCoral

                    ElseIf files.Length > 1 Then
                        ' ステータス更新
                        printRow("検出状況") = "複数検出"
                        ' 検出NG
                        itemRow("検出NG") = itemRow("検出NG") + 1
                        ' 列のセルの背景色を薄桃色に
                        itemRow("背景色") = Color.LightCoral
                    End If

                Else
                    ' 前回発送日１文字目が*→変更データだった場合 ここから
                    If tmpDepartureDate.Substring(0, 1) = "*" Then
                        ' 旧IDが0より大きい
                        If Integer.Parse(printRow("旧ID")) > 0 Then
                            ' 前回発送フォルダよりaiコピーここから
                            searchFilePath = "H:\TEST\" + tmpDepartureDate.Substring(1, 4) + "年度ラベル作成済み分\" + tmpDepartureDate.Substring(1, 6) + "\" + tmpDepartureDate.Substring(5, 4)
                            ' 前回発送フォルダの有無確認
                            If IO.Directory.Exists(searchFilePath) Then
                                ' 前回発送日
                                code8StartYmdCheck = Integer.Parse(tmpDepartureDate.Substring(1, 8))
                                ' 前回発送日が20160618以前
                                If code8StartYmdCheck < 20160618 And tmpCustomerCode.Length = 8 Then
                                    ' 1・2文字目を除去
                                    tmpCustomerCode = tmpCustomerCode.Substring(2, 6)
                                End If
                                ' 検索キー設定
                                If Integer.Parse(printRow("旧ID")) = 0 Then
                                    ' 旧IDなし
                                    searchFileKey = "* " + tmpCustomerCode + " * " + tmpProductName + "_*.ai"
                                Else
                                    ' 旧IDあり
                                    oldCode = printRow("旧ID")
                                    ' 日付チェック
                                    systemStartYmdCheck = Integer.Parse(tmpDepartureDate.Substring(1, 8))
                                    ' 1文字目が0で20130115以前
                                    If oldCode.Substring(0, 1) = "0" AndAlso systemStartYmdCheck < 20130115 Then
                                        ' 旧コードの頭0除去
                                        oldCode = oldCode.Substring(1, oldCode.Length - 1)
                                    End If
                                    ' 検索キー
                                    searchFileKey = oldCode + " " + tmpCustomerCode + " *.ai"
                                End If
                                ' 検出状況更新
                                printRow("検出状況") = "検出中..."
                                ' ファイル一式
                                files = IO.Directory.GetFiles(searchFilePath, searchFileKey, IO.SearchOption.AllDirectories)
                                ' 検出数0で旧ID3ｹﾀ時、頭に0付加し4桁で再検索
                                If files.Length = 0 AndAlso oldCode.Length = 4 AndAlso oldCode.Substring(0, 1) = 0 Then
                                    ' 旧ID保存
                                    oldCode = oldCode.Substring(1, oldCode.Length - 1)
                                    ' 検索キー
                                    searchFileKey = oldCode + " " + tmpCustomerCode + " *.ai"
                                    ' ファイル一式
                                    files = IO.Directory.GetFiles(searchFilePath, searchFileKey, IO.SearchOption.AllDirectories)
                                End If
                                ' 検出数１個なら
                                If files.Length = 1 Then
                                    ' 検出状況更新
                                    printRow("検出状況") = "検出成功"
                                    ' 検出ファイルパス更新
                                    printRow("検出ﾌｧｲﾙﾊﾟｽ") = files(0).ToString
                                    ' 検出ファイルパス一時保存
                                    tmpDetectedPath = printRow("検出ﾌｧｲﾙﾊﾟｽ")
                                    ' ..ai対策
                                    If tmpDetectedPath.IndexOf("..ai") > 0 Then
                                        ' ..aiは.aiにリネーム
                                        IO.File.Move(tmpDetectedPath, tmpDetectedPath.Replace("..ai", ".ai"))
                                        ' リネームしたファイルパスを保存
                                        printRow("検出ﾌｧｲﾙﾊﾟｽ") = tmpDetectedPath.Replace("..ai", ".ai")
                                    End If
                                    ' 一時ファイル名
                                    tmpFileName = IO.Path.GetFileName(printRow("検出ﾌｧｲﾙﾊﾟｽ"))
                                    ' 検出ﾌｧｲﾙﾊﾟｽ
                                    copyName = tmpItemFolderPath + "\" + newCode + " " + tmpCustomerCode + tmpFileName.Substring(tmpFileName.IndexOf(" ", 5), tmpFileName.Length - tmpFileName.IndexOf(" ", 5))
                                    ' 検出ﾌｧｲﾙﾊ名に_Fを付与
                                    copyName = copyName.Replace(".ai", "_F.ai")
                                    ' 既に同名のファイルが存在していても上書
                                    IO.File.Copy(printRow("検出ﾌｧｲﾙﾊﾟｽ"), copyName, True)
                                    ' 作業ﾌｧｲﾙﾊﾟｽ格納
                                    printRow("作業ﾌｧｲﾙﾊﾟｽ") = copyName
                                    ' 検出OKを増やす
                                    itemRow("検出OK") = itemRow("検出OK") + 1

                                ElseIf files.Length = 0 Then
                                    ' 検出状況更新
                                    printRow("検出状況") = "検出失敗"
                                    ' 検出NGを増やす
                                    itemRow("検出NG") = itemRow("検出NG") + 1

                                ElseIf files.Length > 1 Then
                                    ' 検出状況更新
                                    printRow("検出状況") = "複数検出"
                                    ' 検出NGを増やす
                                    itemRow("検出NG") = itemRow("検出NG") + 1
                                    ' 文字列連結
                                    For tate As Integer = 0 To tmpLabelName.Length - 1
                                        tateText = String.Concat(tateText, tmpLabelName.AsSpan(tate, 1), vbCrLf)
                                    Next
                                    ' 検出ﾌｧｲﾙﾊﾟｽ
                                    copyName = tmpItemFolderPath + "\" + newCode + " " + tmpCustomerCode + " " + tmpLabelName + " " + tmpProductName + "_F.ai"
                                    ' 空aiﾌｧｲﾙ生成
                                    MakeEmptyAi(tmpLabelName, copyName)
                                    ' 列のセルの背景色を薄桃色に
                                    printRow("背景色") = Color.LightCoral
                                End If
                            End If

                            ' 変更で旧コード無しここから
                        ElseIf Integer.Parse(printRow("旧ID")) = 0 Then
                            ' 前回発送フォルダよりaiコピーここから
                            searchFilePath = "H:\TEST\" + tmpDepartureDate.Substring(1, 4) + "年度ラベル作成済み分\" + tmpDepartureDate.Substring(1, 6) + "\" + tmpDepartureDate.Substring(5, 4)
                            ' 前回発送フォルダの有無確認
                            If IO.Directory.Exists(searchFilePath) Then
                                ' 前回発送日の頭1文字がアスタリスク
                                If tmpDepartureDate.Substring(0, 1) = "*" Then
                                    ' アスタリスク削除
                                    code8StartYmdCheck = Integer.Parse(tmpDepartureDate.Substring(1, 8))
                                Else
                                    code8StartYmdCheck = Integer.Parse(tmpDepartureDate.Substring(0, 8))
                                End If
                                ' 前回発送日が20160618以前
                                If code8StartYmdCheck < 20160618 And tmpCustomerCode.Length = 8 Then
                                    ' 顧客番号の前2文字を削除
                                    tmpCustomerCode = tmpCustomerCode.Substring(2, 6)
                                End If
                                ' 旧IDが0
                                If Integer.Parse(printRow("旧ID")) = 0 Then
                                    ' 検索キー
                                    searchFileKey = "* " + tmpCustomerCode + " * " + tmpProductName + "_*.ai"
                                Else
                                    ' 旧IDが0
                                    oldCode = printRow("旧ID")
                                    ' 前回発送日
                                    systemStartYmdCheck = Integer.Parse(tmpCustomerCode.Substring(1, 8))
                                    ' 旧IDが0で前回発送日が20130115以前
                                    If oldCode.Substring(0, 1) = "0" AndAlso systemStartYmdCheck < 20130115 Then
                                        ' 旧IDの頭1文字を削除
                                        oldCode = oldCode.Substring(1, oldCode.Length - 1)
                                    End If
                                    ' 検索キー
                                    searchFileKey = oldCode + " " + tmpCustomerCode + " *.ai"
                                End If
                                ' 検出状況変更
                                printRow("検出状況") = "検出中..."
                                ' 取得ファイル一覧
                                files = IO.Directory.GetFiles(searchFilePath, searchFileKey, IO.SearchOption.AllDirectories)
                                ' 検出数0で旧ID3ｹﾀ時、頭に0付加し4桁で再検索
                                If files.Length = 0 AndAlso oldCode.Length = 4 AndAlso oldCode.Substring(0, 1) = 0 Then
                                    ' 1文字目以降
                                    oldCode = oldCode.Substring(1, oldCode.Length - 1)
                                    ' 検索キー
                                    searchFileKey = oldCode + " " + tmpCustomerCode + " *.ai"
                                    ' 取得ファイル一覧
                                    files = IO.Directory.GetFiles(searchFilePath, searchFileKey, IO.SearchOption.AllDirectories)
                                End If
                                ' 検出数１個なら
                                If files.Length = 1 Then
                                    ' 検出状況変更
                                    printRow("検出状況") = "検出成功"
                                    ' 検出ﾌｧｲﾙﾊﾟｽ変更
                                    printRow("検出ﾌｧｲﾙﾊﾟｽ") = files(0).ToString
                                    ' ..ai対策
                                    If printRow("検出ﾌｧｲﾙﾊﾟｽ").IndexOf("..ai") > 0 Then
                                        ' ..aiは.aiにリネーム
                                        IO.File.Move(printRow("検出ﾌｧｲﾙﾊﾟｽ"), printRow("検出ﾌｧｲﾙﾊﾟｽ").Replace("..ai", ".ai"))
                                        ' リネームしたファイルパスを保存
                                        printRow("検出ﾌｧｲﾙﾊﾟｽ") = printRow("検出ﾌｧｲﾙﾊﾟｽ").ToString.Replace("..ai", ".ai")
                                    End If
                                    ' 検出ﾌｧｲﾙﾊﾟｽ
                                    copyName = tmpItemFolderPath + "\" + newCode + " " + tmpCustomerCode + IO.Path.GetFileName(printRow("検出ﾌｧｲﾙﾊﾟｽ")).Substring(IO.Path.GetFileName(printRow("検出ﾌｧｲﾙﾊﾟｽ")).IndexOf(" ", 5), IO.Path.GetFileName(printRow("検出ﾌｧｲﾙﾊﾟｽ")).Length - IO.Path.GetFileName(printRow("検出ﾌｧｲﾙﾊﾟｽ")).IndexOf(" ", 5))
                                    ' 末尾を「_F.ai」に
                                    copyName = copyName.Replace(".ai", "_F.ai")
                                    ' 既に同名のファイルが存在していても上書
                                    IO.File.Copy(printRow("検出ﾌｧｲﾙﾊﾟｽ"), copyName, True)
                                    ' 作業ﾌｧｲﾙﾊﾟｽ更新
                                    printRow("作業ﾌｧｲﾙﾊﾟｽ") = copyName
                                    ' 検出OK更新
                                    itemRow("検出OK") = itemRow("検出OK") + 1

                                ElseIf files.Length = 0 Then
                                    ' 検出状況更新
                                    printRow("検出状況") = "検出失敗"
                                    ' 検出NG更新
                                    itemRow("検出NG") = itemRow("検出NG") + 1
                                    ' 保存先パス
                                    copyName = tmpItemFolderPath + "\" + newCode + " " + tmpCustomerCode + " " + tmpLabelName + " " + tmpProductName + "_F.ai"
                                    ' 空aiﾌｧｲﾙ生成
                                    MakeEmptyAi(tmpLabelName, copyName)
                                    ' 列のセルの背景色を薄桃色に
                                    printRow("背景色") = Color.LightCoral

                                ElseIf files.Length > 1 Then
                                    ' 検出状況更新
                                    printRow("検出状況") = "複数検出"
                                    ' 検出NG更新
                                    itemRow("検出NG") = itemRow("検出NG") + 1
                                    ' 保存先パス
                                    copyName = tmpItemFolderPath + "\" + newCode + " " + tmpCustomerCode + " " + tmpLabelName + " " + tmpProductName + "_F.ai"
                                    ' 空aiﾌｧｲﾙ生成
                                    MakeEmptyAi(tmpLabelName, copyName)
                                    ' 列のセルの背景色を薄桃色に
                                    printRow("背景色") = Color.LightCoral
                                End If
                            End If
                        Else
                            ' 保存先パス
                            copyName = tmpItemFolderPath + "\" + newCode + " " + tmpCustomerCode + " " + tmpLabelName + " " + tmpProductName + "_F.ai"
                            ' 空aiﾌｧｲﾙ生成
                            MakeEmptyAi(tmpLabelName, copyName)
                        End If
                        ' 検出状況変更
                        printRow("検出状況") = "変更前Ai"
                        ' 検出OK更新
                        itemRow("変更前Ai") = itemRow("変更前Ai") + 1
                    Else
                        ' 検出状況変更
                        printRow("検出状況") = "ﾌｫﾙﾀﾞ無し"
                        ' 検出NG更新
                        itemRow("検出NG") = itemRow("検出NG") + 1
                    End If
                    ' 列のセルの背景色を薄桃色に
                    printRow("背景色") = Color.LightCoral
                End If
                ' インクリメント
                counter = counter + 1
                ' ProgressBar1の値を変更する
                MainForm.ProgressBar1.Value = counter
            Next

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：M5
    '*関数名：SendConfirmation
    '*機  能：確認希望送信関数
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub SendConfirmation(ByRef printRow As DataRow)
        ' 変数定義
        Dim pbid As Integer ' 新ID
        Dim instructionName As String ' PBファイル名
        Dim pbFileName As String ' PBファイル名
        Dim productItem As String ' 商品名
        Dim pbDir As String ' 検索結果
        Dim searchFileKey As String ' ファイル検索キー
        ' 配列定義
        Dim filesCheck As String() ' 検出ファイル

        Try
            ' 前回発送日の頭が*
            If printRow("前回発送日").ToString.Substring(0, 1) <> "*" Then
                Console.WriteLine("前回発送日が*")
                Exit Sub
            End If
            ' 初期化
            pbFileName = ""
            instructionName = printRow("指示書").ToString
            ' 新ID
            pbid = CInt(printRow("新ID"))
            ' 商品名
            productItem = printRow("商品名").ToString
            ' 検索結果
            pbDir = GetAppPath() + "\work\" + instructionName.Substring(0, 6) + "\" + instructionName
            ' ファイル検索キー
            searchFileKey = pbid.ToString.PadLeft(4, "0"c) + "-*.jpg"
            ' 検出ファイル
            filesCheck = IO.Directory.GetFiles(pbDir, searchFileKey)
            ' 検出ファイル数
            If filesCheck.Length > 0 Then
                ' PBファイル名
                pbFileName = IO.Path.GetFileName(filesCheck(0))
                ' 販売管理DB更新
                ' progressPbCheckSendDgv2(instructionName, pbid, productItem, pbFileName)
                ' PB確認更新
                ' printRow("PB確認") = progressPbCheckStatusDgv2(instructionName, pbid)
            Else
                MessageBox.Show("送信画像ファイルが見つかりません")
            End If
            ' PB確認希望
            If printRow("PB確認").ToString = "待ち" Then
                printRow("背景色") = Color.Yellow
            End If
            ' PB確認希望
            If printRow("PB確認").ToString = "済×" Then
                printRow("背景色") = Color.LightCoral
            End If
            ' PB確認希望
            If printRow("PB確認").ToString = "済〇" Then
                printRow("背景色") = Color.LightBlue
            End If
            ' PB確認希望
            If printRow("面付出力済").ToString = "済〇" Then
                printRow("背景色") = Color.LightGreen
            End If

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：F1
    '*関数名：InstructionsFilter
    '*機  能：指示書フィルタ関数
    '*戻り値：DataTable(item)
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Function InstructionsFilter() As DataTable
        ' 変数定義
        Dim intOk As Integer ' 検出OKファイル数
        Dim intNg As Integer ' 検出NGファイル数
        Dim ymdFlag As String ' 日付取得
        Dim searchNullAiFilePath As String ' 変更前ai検索対象パス
        Dim searchNullAiFileKey As String ' 変更前ai検索キー
        Dim ymd As String ' 日付
        Dim viewOK As Boolean ' 作成書表示
        ' 配列定義
        Dim filesNullAiCheck As String() ' 変更前AIカウント
        ' オブジェクト定義
        Dim resultItemDt As New DataTable ' itemDB取得
        Dim resultPrintDt As New DataTable ' printDB取得
        Dim tmpResultPrintDt As New DataTable ' 一時printDB取得
        Dim resultTables As New DataSet ' printテーブル

        Try
            ' ボタンラベル初期化
            MainForm.Button7.Text = "Check"
            ' YMDテキストラベル更新
            If MainForm.Label24.Text = "　　年" Or MainForm.Label25.Text = "　　月" Or MainForm.Label26.Text = "　　日" Then
                Throw New System.Exception("An exception has occurred.")
            End If
            ' 日付取得
            ymdFlag = MainForm.Label24.Text.Replace("年", "") + MainForm.Label25.Text.Replace("月", "") + MainForm.Label26.Text.Replace("日", "")
            ymdFlag = ymdFlag.Replace("　", "")
            ' 表示更新
            MainForm.TextBox1.Text = "0"
            MainForm.TextBox2.Text = "0"
            MainForm.TextBox3.Text = "0"
            MainForm.TextBox4.Text = "0"
            MainForm.TextBox5.Text = "0"
            MainForm.TextBox6.Text = "0"
            MainForm.TextBox7.Text = "0"
            MainForm.TextBox8.Text = "0"
            MainForm.TextBox9.Text = "0"
            MainForm.Label22.Text = "0"
            MainForm.Label27.Text = ""

            ' ソート無効
            For Each c As DataGridViewColumn In MainForm.DataGridView2.Columns
                c.SortMode = DataGridViewColumnSortMode.NotSortable
            Next c
            ' ファイル名（拡張子なし）の取得
            ymd = ymdFlag

            If ymd <> "" Then
                ' 前回発送フォルダの有無確認
                If Not IO.Directory.Exists(GetAppPath() + "\work\" + ymd.Substring(0, 6)) Then
                    MsgBox("Workフォルダ内の発送日フォルダを選択してください。")
                End If
                ' YMDテキストラベル更新
                MainForm.Label24.Text = ymd.Substring(0, 2) + "年"
                MainForm.Label25.Text = ymd.Substring(2, 2) + "月"
                MainForm.Label26.Text = ymd.Substring(4, 2) + "日"
            End If
            ' DB取得
            resultTables = GetAllDbData(globalNowDate)
            ' itemDB取得
            resultItemDt = resultTables.Tables(0)
            ' printDB取得
            resultPrintDt = resultTables.Tables(1)
            ' 作成書表示フィルタに従う
            viewOK = MainForm.CheckBox1.Checked

            ' itemDataTableループ
            For Each itemRow As DataRow In resultItemDt.Rows
                If Not viewOK Then
                    ' チェックボックス判定
                    Select Case MainForm.CheckBox2.Checked
                        Case True
                            If itemRow("進捗").ToString = "未処理" Then
                                viewOK = True
                            End If
                            If itemRow("進捗").ToString = "" Then
                                viewOK = True
                            End If
                    End Select
                    Select Case MainForm.CheckBox3.Checked
                        Case True
                            If itemRow("進捗").ToString = ("検索済") Then
                                viewOK = True
                            End If
                    End Select
                    Select Case MainForm.CheckBox6.Checked
                        Case True
                            If itemRow("進捗").ToString = ("配置出力済") Then
                                viewOK = True
                            End If
                    End Select
                    Select Case MainForm.CheckBox5.Checked
                        Case True
                            If itemRow("進捗").ToString = ("配置確認済") Then
                                viewOK = True
                            End If
                    End Select
                    Select Case MainForm.CheckBox4.Checked
                        Case True
                            If itemRow("進捗").ToString = ("面付出力済") Then
                                viewOK = True
                            End If
                    End Select
                    Select Case MainForm.CheckBox11.Checked
                        Case True
                            If itemRow("進捗").ToString = ("確認待ち") Then
                                viewOK = True
                            End If
                    End Select
                End If
                ' 作成書表示フィルタに従
                If MainForm.CheckBox1.Checked = False AndAlso
                  MainForm.CheckBox2.Checked = False AndAlso
                  MainForm.CheckBox3.Checked = False AndAlso
                  MainForm.CheckBox4.Checked = False AndAlso
                  MainForm.CheckBox5.Checked = False AndAlso
                  MainForm.CheckBox6.Checked = False AndAlso
                  MainForm.CheckBox7.Checked = False Then
                    viewOK = True
                End If

                ' コンボボックス判定
                ' 新・Rが含まれる場合
                If MainForm.ComboBox1.SelectedItem <> "" Then
                    If itemRow("指示書").IndexOf(MainForm.ComboBox1.SelectedItem) < 0 Then
                        viewOK = False
                    End If
                End If
                If MainForm.ComboBox2.SelectedItem <> "" Then
                    If itemRow("指示書").IndexOf(MainForm.ComboBox2.SelectedItem).IndexOf(MainForm.ComboBox2.SelectedItem) < 0 Then
                        viewOK = False
                    End If
                End If
                ' 配置判定
                If itemRow("配置進捗") <> "" Then
                    viewOK = False
                End If
                ' 日付判定
                If MainForm.TextBox10.Text <> "" Then
                    If itemRow("指示書").IndexOf(MainForm.TextBox10.Text, 6) < 0 Then
                        viewOK = False
                    End If
                End If

                ' 発送日が対象外
                If Not viewOK Then
                    MsgBox("Workフォルダ内の発送日フォルダを選択してください。")
                    Throw New System.Exception("Workフォルダ内の発送日フォルダを選択してください。")
                End If

                ' 変更前ai検索対象パス
                searchNullAiFilePath = GetAppPath() + "\Work\" + ymd + "\" + itemRow("指示書")
                ' 変更前ai検索キー
                searchNullAiFileKey = "*F.ai"
                ' 変更前AIカウント
                filesNullAiCheck = IO.Directory.GetFiles(searchNullAiFilePath, searchNullAiFileKey)
                ' AIファイル数更新
                itemRow("変更前") = filesNullAiCheck.Length
                ' 配置完了
                itemRow("配置進捗") = "完了"

                ' printTableループ
                For Each printRow As DataRow In resultPrintDt.Rows
                    ' 初期化
                    intOk = 0
                    intNg = 0
                    ' 検出成功/手動検出のとき
                    If printRow("検出状況") = "検出成功" Or printRow("検出状況") = "手動検出" Then
                        If Integer.Parse(printRow("枚数")) > 0 Then
                            ' OKインクリメント
                            intOk = intOk + 1
                        End If
                    Else
                        If Integer.Parse(printRow("枚数")) > 0 Then
                            ' NGインクリメント
                            intNg = intNg + 1
                        End If
                    End If
                Next
                ' カウンタ更新
                itemRow("検出OK") = intOk
                itemRow("検出NG") = intNg

                ' 進捗よりステータスを更新
                Select Case itemRow("進捗")
                    Case "", "未処理"
                        ' 未処理数+1
                        MainForm.TextBox2.Text = (Integer.Parse(MainForm.TextBox2.Text.Replace(","c, "")) + 1).ToString("#,##0")
                    Case "配置確認済"
                        ' 配置出力済数+1
                        MainForm.TextBox8.Text = (Integer.Parse(MainForm.TextBox8.Text.Replace(","c, "")) + 1).ToString("#,##0")
                    Case "面付出力済"
                        ' 面付出力済数+1
                        MainForm.TextBox7.Text = Integer.Parse(MainForm.TextBox7.Text.Replace(","c, "")) + itemRow("枚数")
                End Select

                ' 検索済数ﾌｧｲﾙ+
                If Integer.Parse(itemRow("検出OK")) > 0 Then
                    MainForm.TextBox6.Text = Integer.Parse(MainForm.TextBox6.Text.Replace(","c, "")) + Integer.Parse(itemRow("検出OK"))
                End If
                If Integer.Parse(itemRow("検出NG")) > 0 Then
                    MainForm.TextBox4.Text = Integer.Parse(MainForm.TextBox4.Text.Replace(","c, "")) + Integer.Parse(itemRow("検出NG"))
                End If
                
                ' 処理済数
                MainForm.TextBox9.Text = Integer.Parse(MainForm.TextBox1.Text.Replace(","c, "") - Integer.Parse(MainForm.TextBox2.Text.Replace(","c, ""))).ToString("#,##0")
                MainForm.TextBox6.Text = Integer.Parse(MainForm.TextBox6.Text).ToString("#,##0")
                MainForm.TextBox4.Text = Integer.Parse(MainForm.TextBox4.Text).ToString("#,##0")
                MainForm.TextBox7.Text = Integer.Parse(MainForm.TextBox7.Text).ToString("#,##0")
                ' 作業進捗ヘッダ　受注表件数更新
                MainForm.TextBox1.Text = resultItemDt.Rows.Count.ToString
                ' 読み込み中非表示
                If globalPBCheck Then
                    'tantoPB確認()
                End If
            Next
            ' データテーブル返還
            Return resultItemDt

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            Return resultItemDt
        End Try
    End Function

    ' ■ AI関係処理
    '--------------------------------------------------------
    '*番  号：A1
    '*関数名：MakeEmptyAi
    '*機  能：空Aiファイル作成
    '*分  類：プライベート
    '--------------------------------------------------------
    Private Sub MakeEmptyAi(labelName As String, savePath As String)
        ' 変数定義
        Dim heightText As String = ""

        Try
            ' インスタンス作成
            Dim adobeMaker = New adobeAI()
            ' ai配置テキスト作成
            For height As Integer = 0 To labelName.Length - 1
                heightText = heightText + labelName.Substring(height, 1) + vbCrLf
            Next
            ' 空ai作成
            adobeMaker.TextLayout(heightText, 30, 300, 400)
            adobeMaker.TextLayout(labelName, 30, 200, 500)
            ' 空ai保存
            adobeMaker.Save(savePath)
            ' 空ai閉じる
            adobeMaker.Close()

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*番  号：A3
    '*関数名：PatternRecognition
    '*機  能：パターン判定用
    '*戻り値：パターンID
    '*分  類：プライベート
    '--------------------------------------------------------
    Private Function PatternRecognition(width As Double, height As Double, pdName As String) As String
        ' 変数定義
        Dim miniFlg As boolean ' 300フラグ
        ' オブジェクト定義
        Dim patternDt As New DataTable ' パターンマスタ
        Dim masterDataSet As New DataSet  ' マスタ取得

        Try
            ' マスタ取得
            masterDataSet = GetMasterData()
            ' patternMasterTable取得
            patternDt = masterDataSet.Tables(0)
            
            ' patternMasterTableループ
            For Each patternRow As DataRow In patternDt.Rows
                ' 300フラグ
                miniFlg = False
                ' 通常か面付1800
                If patternRow("special") = "300NG" And pdName.IndexOf(patternRow("special")) Then
                    Continue For
                End If
                ' 通常か面付1800
                If patternRow("special") = "300" And pdName.IndexOf(patternRow("special")) Then
                    miniFlg = True
                End If
                ' 判定ループ
                If width > patternRow("wlower") AndAlso width < patternRow("wupper") AndAlso height > patternRow("hlower") AndAlso height < patternRow("hupper") Then
                    ' 略語判定あり
                    If miniFlg Then
                        ' 略語なし
                        If patternRow("patternname").IndexOf("tate") > 0 Then
                            ' パターンID
                            Return "4"
                        Else
                            ' パターンID
                            Return "3"
                        End If
                    Else
                        ' パターンID
                        Return patternRow("patternID")
                        Exit For
                    End If
                End If
            Next
            ' ヒットしなければ空欄
            Return ""

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            Return ""
        End Try
    End Function

    ' ■ 戻り処理
    '--------------------------------------------------------
    '*番  号：F1
    '*関数名：GetAppPath
    '*機  能：アプリパス取得
    '*出  力：アプリパス(String)
    '*分  類：パブリック
    '--------------------------------------------------------
    Public Function GetAppPath() As String
        Return IO.Path.GetDirectoryName(
        System.Reflection.Assembly.GetExecutingAssembly().Location)
    End Function
End Module
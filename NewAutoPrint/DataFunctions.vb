' ■ DataFunctions
Public Module DataFunctions
    '--------------------------------------------------------
    '*番  号：F1
    '*関数名：MakeItemInitialTable
    '*機  能：itemDataTable作成
    '*戻り値：DataTable(item)
    '*分  類：パブリック
    '--------------------------------------------------------
    Public Function MakeItemInitialTable(strFileNames As String()) As DataTable
        ' ジェネリック定義
        Dim itemColumnList As List(Of String) ' itemヘッダ
        Dim itemTypeList As List(Of String) ' item型
        Dim itemDbFlgList As List(Of String) ' itemdbフラグ
        Dim itemDataTableFlgList As List(Of String) ' itemdatatableフラグ
        ' オブジェクト定義
        Dim emptyDt As New DataTable ' 空のDataTable
        Dim emptyRow As DataRow ' 空のDataRow
        Dim dtView As DataView ' 重複削除用
        Dim resultItemDt As DataTable ' 取得itemDataTable
        Dim finalItemDt As DataTable ' 最終itemDataTable

        Try
            ' DBインスタンス作成
            Dim dbMaker = New Db()
            ' 空の行
            emptyRow = emptyDt.NewRow
            ' itemDB取得
            resultItemDt = dbMaker.Sql_select("Select * from itemtable where date = " + globalNowDate)
            ' item型
            itemTypeList = GetFixedData("itemtype")
            ' itemdatatableフラグ
            itemDbFlgList = GetFixedData("itemdbflg")
            ' itemdatatableフラグ
            itemDataTableFlgList = GetFixedData("itemdatatableflg")
            ' itemヘッダ
            itemColumnList = GetFixedData("itemheader")

            ' itemdatatable
            For i As Integer = 0 To itemDataTableFlgList.Count
                ' DBなしDataTableあり
                If itemDbFlgList(i) = "0" And itemDataTableFlgList(i) = "1" Then
                    ' カラムを格納
                    resultItemDt.Columns.Add(itemColumnList(i), System.Type.GetType("System." + itemTypeList(i)))
                    ' 値を格納
                    For Each itemRow As DataRow In resultItemDt.Rows
                        ' 配置選択
                        itemRow("配置選択") = False
                        ' 検出OK
                        itemRow("検出OK") = "0"
                        ' 検出NG
                        itemRow("検出NG") = "0"
                        ' 面付選択
                        itemRow("面付選択") = False
                        ' 指示書状態
                        itemRow("指示書状態") = ""
                        ' 変更前
                        itemRow("変更前") = "0"
                        ' 背景色
                        itemRow("背景色") = ""
                        ' 指示書重複
                        itemRow("指示書重複") = False
                    Next
                End If
            Next
            ' 重複削除処理
            dtView = New DataView(resultItemDt)
            ' 変換後データテーブル
            finalItemDt = dtView.ToTable(True, "itemID")
            ' 結果返し
            Return finalItemDt

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            ' 結果返し
            Return emptyDt
        End Try
    End Function

    '--------------------------------------------------------
    '*番  号：F2
    '*関数名：MakePrintInitialTable
    '*機  能：printDataTable作成
    '*戻り値：DataTable(print)
    '*分  類：パブリック
    '--------------------------------------------------------
    Public Function MakePrintInitialTable(id As String) As DataTable
        ' ジェネリック定義
        Dim printDbFlgList As List(Of String) ' printdbフラグ
        Dim printColumnList As List(Of String) ' printヘッダ
        Dim printTypeList As List(Of String) ' print型
        Dim printDataTableFlgList As List(Of String) ' printdatatableフラグ
        ' オブジェクト定義
        Dim emptyDt As New DataTable ' 空のDataTable
        Dim resultPrintDt As DataTable ' printDataTable

        Try
            ' 初期化
            ' DBインスタンス作成
            Dim dbMaker = New Db()
            ' printdb
            printDbFlgList = GetFixedData("printdb")
            ' printヘッダ
            printColumnList = GetFixedData("printheader")
            ' print型
            printTypeList = GetFixedData("printtype")
            ' printdatatableフラグ
            printDataTableFlgList = GetFixedData("printdatatableflg")
            ' printDB取得
            resultPrintDt = dbMaker.Sql_select("Select * from printtable where ID = " + id)

            ' printdatatable
            For i As Integer = 0 To printDataTableFlgList.Count - 1
                ' DBなしDataTableあり
                If printDbFlgList(i) = "0" Then
                    ' カラムを格納
                    resultPrintDt.Columns.Add(printColumnList(i), System.Type.GetType("System." + printTypeList(i)))
                    ' 値を格納
                    For Each printRow As DataRow In resultPrintDt.Rows
                        ' 部分配置
                        printRow("部分配置") = False
                        ' 部分面付
                        printRow("部分面付") = False
                        ' PB確認
                        printRow("背景色") = ""
                    Next
                End If
            Next
            ' 結果返し
            Return resultPrintDt

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            ' 空白返し
            Return emptyDt
        End Try
    End Function

    '--------------------------------------------------------
    '*番  号：F3
    '*関数名：MakePatternInitialTable
    '*機  能：patternDataTable作成
    '*戻り値：DataTable(pattern)
    '*分  類：パブリック
    '--------------------------------------------------------
    Public Function MakePatternInitialTable() As DataTable
        ' オブジェクト定義
        Dim emptyDt As New DataTable ' 空のDataTable
        Dim patternColumnList As List(Of String) ' patternヘッダ
        Dim patternTypeList As List(Of String) ' pattern型
        Dim resultPatternDt As DataTable ' patternDataTable

        Try
            ' DataTable初期化
            resultPatternDt = New DataTable
            ' patternヘッダ
            patternColumnList = GetFixedData("patternheader")
            ' pattern型
            patternTypeList = GetFixedData("patterntype")

            ' printdatatable
            For i As Integer = 0 To patternColumnList.Count - 1
                ' カラムを格納
                resultPatternDt.Columns.Add(patternColumnList(i), System.Type.GetType("System." + patternTypeList(i)))
            Next
            ' 結果返し
            Return resultPatternDt

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            ' 空白返し
            Return emptyDt
        End Try
    End Function

    '--------------------------------------------------------
    '*番  号：F4
    '*関数名：MakeTemplateInitialTable
    '*機  能：templateDataTable作成
    '*戻り値：DataTable(template)
    '*分  類：パブリック
    '--------------------------------------------------------
    Public Function makeTemplateInitialTable() As DataTable
        ' オブジェクト定義
        Dim emptyDt As New DataTable ' 空のDataTable
        Dim templateColumnList As List(Of String) ' templateヘッダ
        Dim templateTypeList As List(Of String) ' template型
        Dim resultTemplateDt As DataTable ' templateDataTable

        Try
            ' DataTable初期化
            resultTemplateDt = New DataTable
            ' templateヘッダ
            templateColumnList = GetFixedData("templateheader")
            ' template型
            templateTypeList = GetFixedData("templatetype")

            ' printdatatable
            For i As Integer = 0 To templateColumnList.Count - 1
                ' カラムを格納
                resultTemplateDt.Columns.Add(templateColumnList(i), System.Type.GetType("System." + templateTypeList(i)))
            Next
            ' 結果返し
            Return resultTemplateDt

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            ' 空白返し
            Return emptyDt
        End Try
    End Function

    '--------------------------------------------------------
    '*番  号：F5
    '*関数名：MakeItemDb
    '*機  能：itemDB作成
    '*戻り値：Integer(DB結果)
    '*分  類：パブリック
    '--------------------------------------------------------
    Public Function MakeItemDb(strFileNames As String()) As Integer
        ' 変数定義
        Dim itemRs As Integer ' itemDB結果
        Dim itemCnt As Integer ' 件数
        Dim pageCnt As Integer ' ラベル枚数
        Dim tmpStr As String ' 行読み込み用
        Dim fileName As String ' 最初のファイル名
        Dim rootParentPath As String ' ドラッグ元フォルダパス
        Dim tmpFileName As String ' txtファイル名
        Dim pbCheckName As String ' PB確認希望
        Dim tmpFileNameWithoutExtension As String ' 拡張子抜きファイル名
        Dim zendoFlg As Boolean ' 前回同様フラグ
        ' 配列定義
        Dim tmpStrArray As String() ' 行格納用
        ' ジェネリック定義
        Dim itemDbValuesArrayList As List(Of Hashtable) ' itemDB値リスト
        Dim templateItemColumnList As List(Of String) = New List(Of String)() ' templateヘッダ設定
        '
        Dim itemDbHashList As Hashtable ' itemDB値
        ' オブジェクト定義
        Dim tmpObjFile As IO.StreamReader ' txt読込用

        Try
            ' DBインスタンス作成
            Dim dbMaker = New Db()
            ' itemDB値リスト
            itemDbValuesArrayList = New List(Of Hashtable)
            ' 最初のファイル名
            fileName = IO.Path.GetFileName(strFileNames(0).ToString)
            ' ドラッグ元フォルダパス
            rootParentPath = strFileNames(0).Replace(fileName, "")

            ' itemtableクエリ用データ作成
            For i As Integer = 0 To strFileNames.Length - 1
                ' 前同フラグ初期化
                zendoFlg = False
                ' ファイルの存在確認
                If IO.File.Exists(strFileNames(i).ToString) Then
                    ' 初期化
                    itemCnt = 0
                    pageCnt = 0
                    ' txtファイル名
                    tmpFileName = IO.Path.GetFileName(strFileNames(i).ToString)
                    ' 拡張子抜きファイル名
                    tmpFileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(tmpFileName)
                    ' txt読込
                    tmpObjFile = New IO.StreamReader(rootParentPath + tmpFileName, System.Text.Encoding.Default)
                    ' 最初の行読込
                    tmpStr = tmpObjFile.ReadLine()
                    ' 最後の行まで読込
                    While (tmpStr <> "")
                        ' 件数
                        itemCnt += 1
                        ' 配列格納
                        tmpStrArray = tmpStr.Split(CChar(","))
                        ' ラベル枚数
                        pageCnt += Integer.Parse(tmpStrArray(3))
                        ' 次の行読込
                        tmpStr = tmpObjFile.ReadLine()
                    End While

                    ' 前同確認
                    If tmpFileNameWithoutExtension.IndexOf("前同") > 0 Then
                        ' 前同フラグオン
                        zendoFlg = True
                    Else
                        ' 前同フラグオフ
                        zendoFlg = False
                    End If

                    ' PB確認希望チェック
                    If tmpFileNameWithoutExtension.IndexOf("確認") > 0 AndAlso globalPBCheck Then
                        ' PB確認希望
                        pbCheckName = "検証用"
                    Else
                        ' PB確認希望は空欄
                        pbCheckName = ""
                    End If

                    ' db用ハッシュテーブル
                    itemDbHashList = New Hashtable From {
                      {"id", CStr(i + 1)}, ' ID
                      {"date", globalNowDate}, ' 日付
                      {"itemname", tmpFileNameWithoutExtension}, ' 指示書
                      {"amount", itemCnt}, ' 件数
                      {"labelnumber", pageCnt}, ' 枚数
                      {"folderpath", globalRootPath + globalNowDate + "\" + tmpFileNameWithoutExtension}, ' フォルダ
                      {"layoutprogress", "　"}, ' 配置進捗
                      {"progress", "未処理"}, ' 進捗
                      {"printtime", ""}, ' 配置出力時間
                      {"pbcheck", pbCheckName}, ' PB確認希望
                      {"dl", ""}, ' DL済
                      {"zendoflg", zendoFlg}, ' 前同
                      {"printstaffname", globalUserName}, ' 担当者
                      {"usable", 1} ' usable
                    }
                    ' 全体リストに追加
                    itemDbValuesArrayList.Add(itemDbHashList)
                End If
            Next
            ' 結果
            itemRs = dbMaker.Sql_insert("itemtable", itemDbValuesArrayList)

            ' データ登録完了
            If itemRs > 0 Then
                Console.WriteLine("Message: itemDB登録成功")
                ' データセット返し
                Return itemRs
            Else
                ' 0返し
                Return 0
            End If

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            ' 0返し
            Return 0
        End Try
    End Function

    '--------------------------------------------------------
    '*番  号：F6
    '*関数名：MakePrintDb
    '*機  能：printDB作成
    '*戻り値：Long(DB結果)
    '*分  類：パブリック
    '--------------------------------------------------------
    Public Function MakePrintDb(itemRs As Integer, strFileNames As String()) As Integer
        ' 変数定義
        Dim printRs As Integer ' printDB結果
        Dim tmpStr As String ' txt読込用
        Dim tmpNewId As String ' txt読込用
        Dim fileName As String ' 最初のファイル名
        Dim tmpFileName As String ' txtファイル名
        Dim rootParentPath As String ' ドラッグ元フォルダパス
        ' 配列定義
        Dim tmpStrArray As String() ' 行格納用
        ' ジェネリック定義
        Dim printDbHashList As Hashtable ' printdb値格納用
        Dim printDbValuesArrayList As List(Of Hashtable) ' printdb値リスト
        ' オブジェクト定義
        Dim tmpObjFile As IO.StreamReader ' txt読込

        Try
            ' DBインスタンス作成
            Dim dbMaker = New Db()
            ' printdb値リスト
            printDbValuesArrayList = New List(Of Hashtable)
            ' 最初のファイル名
            fileName = IO.Path.GetFileName(strFileNames(0).ToString)
            ' ドラッグ元フォルダパス
            rootParentPath = strFileNames(0).Replace(fileName, "")

            ' itemtableクエリ用データ作成
            For Each strFile As String In strFileNames
                ' ファイルの存在確認
                If IO.File.Exists(strFile) Then
                    ' txtファイル名
                    tmpFileName = IO.Path.GetFileName(strFile)
                    ' txt読込
                    tmpObjFile = New IO.StreamReader(rootParentPath + tmpFileName, System.Text.Encoding.Default)
                    ' 最初の行読込
                    tmpStr = tmpObjFile.ReadLine()
                    ' 最後の行まで読込
                    While (tmpStr <> "")
                        ' カンマ無し
                        If Not tmpStr.Contains(",") Then
                            ' ループを抜ける
                            Exit While
                        End If
                        ' 配列格納
                        tmpStrArray = tmpStr.Split(CChar(","))
                        ' ""除去
                        tmpNewId = tmpStrArray(5).Replace("""", "")
                        ' 3桁なら0を付与
                        If tmpNewId.Length = 3 Then
                            ' IDをゼロパディング
                            tmpNewId = "0" + tmpNewId
                        End If
                        ' ハッシュ用
                        printDbHashList = New Hashtable From {
                          {"itemtable_id", itemRs}, ' itemID
                          {"shippingdate", tmpStrArray(0).Replace("""", "")}, ' 前回発送日
                          {"customerno", Integer.Parse(tmpStrArray(1))}, ' 顧客番号
                          {"productname", tmpStrArray(2).Replace(" ", "").Replace("""", "")}, ' 商品名
                          {"pages", Integer.Parse(tmpStrArray(3))}, ' 枚数
                          {"oldid", tmpStrArray(4).Replace("""", "")}, ' 旧ID
                          {"newid", tmpNewId}, ' 新ID
                          {"searchprogress", ""}, ' 検出状況
                          {"findfilepath", ""}, ' 検出ﾌｧｲﾙﾊﾟｽ
                          {"localpath", ""}, ' 作業ﾌｧｲﾙﾊﾟｽ
                          {"width", 0}, ' 幅
                          {"height", 0}, ' 高さ
                          {"outputstatus", 0}, ' 出力
                          {"labelname", tmpStrArray(6).Replace("""", "")}, ' ラベル名
                          {"confirm", ""},' 確認希望
                          {"pbcheck", 0},' PB確認
                          {"usable", 1} ' usable
                        }
                        ' 全体リストに追加
                        printDbValuesArrayList.Add(printDbHashList)
                        ' 最初の行読込
                        tmpStr = tmpObjFile.ReadLine()
                    End While
                End If
            Next
            ' Mysql登録
            printRs = dbMaker.Sql_insert("printtable", printDbValuesArrayList)

            ' データ登録完了
            If printRs > 0 Then
                Console.WriteLine("Message: printDB登録成功")
                ' 登録結果返し
                Return printRs
            Else
                ' 0返し
                Return 0
            End If

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            ' 0返し
            Return 0
        End Try
    End Function

    '--------------------------------------------------------
    '*番  号：F7
    '*関数名：UpdateItemDb
    '*機  能：itemDB更新
    '*戻り値：String(DB結果)
    '*分  類：パブリック
    '--------------------------------------------------------
    Public Function UpdateItemDb(itemTable As dataTable) As String
        ' 変数定義
        Dim itemRs As String
        ' ジェネリック定義
        Dim itemDbHashList As Hashtable ' itemdb値格納用
        Dim itemDbValuesArrayList As List(Of Hashtable) ' itemdb値リスト

        Try
            ' DBインスタンス作成
            Dim dbMaker = New Db()
            ' itemdb値リスト
            itemDbValuesArrayList = New List(Of Hashtable)

            ' itemtableクエリ用データ作成
            For Each itemRow As dataRow In itemTable.Rows
                ' db用ハッシュテーブル
                itemDbHashList = New Hashtable From {
                    {"id", itemRow("itemID")}, ' ID
                    {"date", itemRow("日付")}, ' 日付
                    {"itemname", itemRow("指示書")}, ' 指示書
                    {"amount", itemRow("件数")}, ' 件数
                    {"labelnumber", itemRow("枚数")}, ' 枚数
                    {"folderpath", itemRow("フォルダ")}, ' フォルダ
                    {"fullpath", itemRow("フルパス")}, ' フォルダ
                    {"layoutprogress", itemRow("配置進捗")}, ' 配置進捗
                    {"progress", itemRow("進捗")}, ' 進捗
                    {"printtime", itemRow("配置出力時間")}, ' 配置出力時間
                    {"pbcheck", itemRow("PB確認希望")}, ' PB確認希望
                    {"dl", itemRow("DL済")}, ' DL済
                    {"zendoflg", itemRow("前同")}, ' 前同
                    {"printstaffname", itemRow("担当者")}, ' 担当者
                    {"usable", 1} ' usable
                }
                ' 全体リストに追加
                itemDbValuesArrayList.Add(itemDbHashList)
            Next
            ' Mysql登録
            itemRs = dbMaker.Sql_update("itemtable", itemDbValuesArrayList)
            Return itemRs

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            ' 0返し
            Return 0
        End Try
    End Function

    '--------------------------------------------------------
    '*番  号：F8
    '*関数名：UpdatePrintDb
    '*機  能：printDB更新
    '*戻り値：String(DB結果)
    '*分  類：パブリック
    '--------------------------------------------------------
    Public Function UpdatePrintDb(printTable As dataTable) As String
        ' 変数定義
        Dim printRs As String ' Mysql登録結果
        ' ジェネリック定義
        Dim printDbHashList As Hashtable ' printdb値格納用
        Dim printDbValuesArrayList As List(Of Hashtable) ' printdb値リスト

        Try
            ' DBインスタンス作成
            Dim dbMaker = New Db()
            ' printdb値リスト
            printDbValuesArrayList = New List(Of Hashtable)

            ' printtableクエリ用データ作成
            For Each printRow As dataRow In printTable.Rows
                ' ハッシュ用
                printDbHashList = New Hashtable From {
                    {"id", printRow("printID")}, ' itemID
                    {"itemtable_id", printRow("itemID")}, ' itemID
                    {"shippingdate", printRow("前回発送日")}, ' 前回発送日
                    {"customerno", printRow("顧客番号")}, ' 顧客番号
                    {"productname", printRow("商品名")}, ' 商品名
                    {"pages", printRow("枚数")}, ' 枚数
                    {"oldid", printRow("旧ID")}, ' 旧ID
                    {"newid", printRow("新ID")}, ' 新ID
                    {"searchprogress", printRow("検出状況")}, ' 検出状況
                    {"findfilepath", printRow("検出ﾌｧｲﾙﾊﾟｽ")}, ' 検出ﾌｧｲﾙﾊﾟｽ
                    {"localpath", printRow("作業ﾌｧｲﾙﾊﾟｽ")}, ' 作業ﾌｧｲﾙﾊﾟｽ
                    {"width", printRow("幅")}, ' 幅
                    {"height", printRow("高さ")}, ' 高さ
                    {"outputstatus", printRow("出力")}, ' 出力
                    {"labelname", printRow("ラベル名")}, ' ラベル名
                    {"confirm", printRow("確認希望")},' 確認希望
                    {"pbcheck", printRow("PB確認")},' PB確認
                    {"usable", 1} ' usable
                }
                ' 全体リストに追加
                printDbValuesArrayList.Add(printDbHashList)
            Next
            ' Mysql登録
            printRs = dbMaker.Sql_update("printtable", printDbValuesArrayList)

            ' データ登録完了
            If printRs <> "error" Then
                Console.WriteLine("Message: printDB更新成功")
            End If
            ' 登録結果返し
            Return printRs

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            ' 返し
            Return "error"
        End Try
    End Function

    '--------------------------------------------------------
    '*番  号：F9
    '*関数名：GetItemDbData
    '*機  能：itemDBデータ取得
    '*戻り値：DataTable(itemDB)
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Function GetItemDbData(itemId As String) As DataTable
        ' オブジェクト定義
        Dim emptyDt As New DataTable ' データセット
        Dim tmpResultItemTable As DataTable ' printテーブル

        Try
            ' DBインスタンス作成
            Dim dbMaker = New Db()
            ' DataTable初期化
            tmpResultItemTable = New DataTable
            ' itemDataTable抽出
            tmpResultItemTable = dbMaker.Sql_select("Select * from itemtable where id = " + itemId)
            ' DataTable返し
            Return tmpResultItemTable

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            Return emptyDt
        End Try
    End Function

     '--------------------------------------------------------
    '*番  号：F10
    '*関数名：GetPrintDbData
    '*機  能：DBデータ取得
    '*戻り値：DataTable(printDB)
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Function GetPrintDbData(itemid As String) As DataTable
        ' オブジェクト定義
        Dim emptyDt As New DataTable ' データセット
        Dim tmpResultPrintTable As DataTable ' printテーブル

        Try
            ' DBインスタンス作成
            Dim dbMaker = New Db()
            ' DataTable初期化
            tmpResultPrintTable = New DataTable
            ' itemDB取得
            tmpResultPrintTable = dbMaker.Sql_select("Select * from printtable where id = " + itemid)
            ' DataTable返し
            Return tmpResultPrintTable

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            ' DataTable返し
            Return emptyDt
        End Try
    End Function

     '--------------------------------------------------------
    '*番  号：F11
    '*関数名：GetAllDbData
    '*機  能：DBデータ取得
    '*戻り値：DataSet(取得データセット)
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Function GetAllDbData(tmpDate As String) As DataSet
        ' オブジェクト定義
        Dim resultItemTable As DataTable ' itemテーブル
        Dim tmpResultPrintTable As DataTable ' 一時printテーブル
        Dim resultPrintTable As DataTable ' printテーブル
        Dim ds As New DataSet ' データセット
        Dim emptyDs As New DataSet ' データセット

        Try
            ' DBインスタンス作成
            Dim dbMaker = New Db()
            ' DataTable初期化
            resultPrintTable = New DataTable
            ' itemDB取得
            resultItemTable = dbMaker.Sql_select("Select * from itemtable where date = " + tmpDate)
            ' printDB取得
            For Each itemRow As DataRow In resultItemTable.Rows
                ' printDataTable抽出
                tmpResultPrintTable = dbMaker.Sql_select("Select * from printtable where id = " + itemRow("itemID"))
                ' 過去履歴合体
                resultPrintTable.Merge(tmpResultPrintTable)
            Next
            ' itemDB返還
            ds.Tables.Add(resultItemTable)
            ' printDB返還
            ds.Tables.Add(resultPrintTable)
            ' DataSet返し
            Return ds

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            ' 空DataSet返し
            Return emptyDs
        End Try
    End Function

    '--------------------------------------------------------
    '*番  号：F12
    '*関数名：GetMasterData
    '*機  能：DBマスタデータ取得
    '*戻り値：DataSet(取得データセット)
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Function GetMasterData(Optional patternid As String = "") As DataSet
        ' 変数定義
        Dim tmpTemplateId As String
        ' オブジェクト定義
        Dim patternMasterHeadDt As DataTable ' patternheadmaster
        Dim templateMasterHeadDt As DataTable ' templateheadmaster
        Dim patternMasterDt As DataTable ' patternmaster
        Dim templateMasterDt As DataTable ' templatemaster
        Dim printerMasterDt As DataTable ' printermaster
        Dim ds As New DataSet ' データセット
        Dim emptyDs As New DataSet ' データセット

        Try
            ' DBインスタンス作成
            Dim dbMaker = New Db()
            ' パターン判定
            If patternid <> "" Then
                ' masterDB取得
                patternMasterDt = dbMaker.Sql_select("Select * from patternmaster where id = " + patternid)
                ' template_id取得
                tmpTemplateId = patternMasterDt.Rows(0).item(1)
                ' templateDB取得
                templateMasterDt = dbMaker.Sql_select("Select * from templatemaster where template_id = " + tmpTemplateId)
                ' printerDB取得
                printerMasterDt = dbMaker.Sql_select("Select * from printermaster")
            Else
                 ' masterDB取得
                patternMasterDt = dbMaker.Sql_select("Select * from patternmaster")
                ' templateDB取得
                templateMasterDt = dbMaker.Sql_select("Select * from templatemaster")
                ' printerDB取得
                printerMasterDt = dbMaker.Sql_select("Select * from printermaster")
            End If
            ' ヘッダ取得
            patternMasterHeadDt = MakePatternInitialTable()
            templateMasterHeadDt = makeTemplateInitialTable()
            ' ヘッダ合体
            patternMasterHeadDt.Merge(patternMasterDt)
            templateMasterHeadDt.Merge(templateMasterDt)
            ' masterDB返還
            ds.Tables.Add(patternMasterDt)
            ' templateDB返還
            ds.Tables.Add(templateMasterDt)
            ' printerDB返還
            ds.Tables.Add(printerMasterDt)
            ' DataSet返し
            Return ds

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            ' 空DataSet返し
            Return emptyDs
        End Try
    End Function

    '--------------------------------------------------------
    '*番  号：F13
    '*関数名：GetFixedData
    '*機  能：項目名取得
    '*戻り値：項目名一覧(String)
    '*分  類：パブリック
    '--------------------------------------------------------
    Public Function GetFixedData(str As String) As List(Of String)
        ' 変数定義
        Dim itemLength As Integer ' item列数
        Dim printLength As Integer ' print列数
        Dim patternLength As Integer ' pattern列数
        Dim templateLength As Integer ' template列数
        ' ジェネリック定義
        Dim itemColumnList As List(Of String) = New List(Of String)() ' itemヘッダ設定
        Dim itemTypeList As List(Of String) = New List(Of String)() ' item型設定
        Dim itemWidthList As List(Of String) = New List(Of String)() ' item幅設定
        Dim itemDataViewFlgList As List(Of String) = New List(Of String)() ' itemDataGridView設定
        Dim itemDataTableFlgList As List(Of String) = New List(Of String)() ' itemDataTable設定
        Dim itemDbFlgList As List(Of String) = New List(Of String)() ' itemDb設定
        Dim itemButtonList As List(Of String) = New List(Of String)() ' itemButton設定
        Dim itemCheckList As List(Of String) = New List(Of String)() ' itemCheck設定 
        Dim printColumnList As List(Of String) = New List(Of String)() ' printヘッダ設定
        Dim printTypeList As List(Of String) = New List(Of String)() ' print型設定
        Dim printWidthList As List(Of String) = New List(Of String)() ' print幅設定
        Dim printDataViewFlgList As List(Of String) = New List(Of String)() ' printDataGridView設定
        Dim printDataTableFlgList As List(Of String) = New List(Of String)() ' printDataTableFlg設定
        Dim printDbFlgList As List(Of String) = New List(Of String)() ' printDbFlg設定
        Dim printDbColumnList As List(Of String) = New List(Of String)() ' printColumn設定
        Dim printButtonList As List(Of String) = New List(Of String)() ' printButton設定
        Dim printCheckList As List(Of String) = New List(Of String)() ' printCheck設定
        Dim patternColumnList As List(Of String) = New List(Of String)() ' patternヘッダ設定
        Dim patternTypeList As List(Of String) = New List(Of String)() ' pattern型設定
        Dim templateColumnList As List(Of String) = New List(Of String)() ' templateヘッダ設定
        Dim templateTypeList As List(Of String) = New List(Of String)() ' template型設定

        Try
            ' itemDataTableの列数
            itemLength = MyConst.Variables.ItemColumn.GetLength(0) - 1
            ' printDataTableの列数
            printLength = MyConst.Variables.PrintColumn.GetLength(0) - 1
            ' patternDataTableの列数
            patternLength = MyConst.Variables.PatternColumn.GetLength(0) - 1
            ' templateDataTableの列数
            templateLength = MyConst.Variables.TemplateColumn.GetLength(0) - 1

            ' 分岐
            Select Case str
                ' itemヘッダ設定
                Case "itemheader"
                    ' Forステートメントを使って列挙
                    For i As Integer = 0 To itemLength
                        itemColumnList.Add(MyConst.Variables.ItemColumn(i, 0))
                    Next
                    ' itemヘッダ返し
                    Return itemColumnList

                ' item型設定
                Case "itemtype"
                    ' Forステートメントを使って列挙
                    For j As Integer = 0 To itemLength
                        itemTypeList.Add(MyConst.Variables.ItemColumn(j, 1))
                    Next
                    ' item型返し
                    Return itemTypeList

                ' item幅設定
                Case "itemwidth"
                    ' Forステートメントを使って列挙
                    For k As Integer = 0 To itemLength
                        itemWidthList.Add(sizeDictionary(MyConst.Variables.ItemColumn(k, 2)))
                    Next
                    ' item幅返し
                    Return itemWidthList

                ' itemDataGridView設定
                Case "itemdataviewflg"
                    ' Forステートメントを使って列挙
                    For m As Integer = 0 To itemLength
                        itemDataViewFlgList.Add(MyConst.Variables.ItemColumn(m, 3))
                    Next
                    ' itemDataGridViewフラグ返し
                    Return itemDataViewFlgList

                ' itemDataTable設定
                Case "itemdatatableflg"
                    ' Forステートメントを使って列挙
                    For n As Integer = 0 To itemLength
                        itemDataTableFlgList.Add(MyConst.Variables.ItemColumn(n, 4))
                    Next
                    ' itemDataTableフラグ返し
                    Return itemDataTableFlgList

                ' itemDb設定
                Case "itemdbflg"
                    ' Forステートメントを使って列挙
                    For o As Integer = 0 To itemLength
                        itemDbFlgList.Add(MyConst.Variables.ItemColumn(o, 5))
                    Next
                    ' itemDbフラグ返し
                    Return itemDbFlgList

                ' itemColumn設定
                Case "itemcolumnflg"
                    ' Forステートメントを使って列挙
                    For p As Integer = 0 To itemLength
                        itemButtonList.Add(MyConst.Variables.ItemColumn(p, 6))
                    Next
                    ' dbカラム返し
                    Return itemButtonList

                ' itemButton設定
                Case "itembutton"
                    ' Forステートメントを使って列挙
                    For q As Integer = 0 To itemLength
                        itemButtonList.Add(MyConst.Variables.ItemColumn(q, 7))
                    Next
                    ' ボタン返し
                    Return itemButtonList

                ' check設定
                Case "itemcheck"
                    ' Forステートメントを使って列挙
                    For r As Integer = 0 To itemLength
                        itemCheckList.Add(MyConst.Variables.ItemColumn(r, 8))
                    Next
                    ' ボタン返し
                    Return itemCheckList

                ' printヘッダ設定
                Case "printheader"
                    ' Forステートメントを使って列挙
                    For s As Integer = 0 To printLength
                        printColumnList.Add(MyConst.Variables.PrintColumn(s, 0))
                    Next
                    ' printヘッダ返し
                    Return printColumnList

                ' print型設定
                Case "printtype"
                    ' Forステートメントを使って列挙
                    For t As Integer = 0 To printLength
                        printTypeList.Add(MyConst.Variables.PrintColumn(t, 1))
                    Next
                    ' print型返し
                    Return printTypeList

                ' printDataGridView幅設定
                Case "printwidth"
                    ' Forステートメントを使って列挙
                    For u As Integer = 0 To printLength
                        printWidthList.Add(sizeDictionary(MyConst.Variables.PrintColumn(u, 2)))
                    Next
                    ' print幅返し
                    Return printWidthList

                ' printDataGridViewFlg設定
                Case "printdataviewflg"
                    ' Forステートメントを使って列挙
                    For v As Integer = 0 To printLength
                        printDataViewFlgList.Add(MyConst.Variables.PrintColumn(v, 3))
                    Next
                    ' printDataGridViewフラグ返し
                    Return printDataViewFlgList

                ' printDataTableFlg設定
                Case "printdatatableflg"
                    ' Forステートメントを使って列挙
                    For w As Integer = 0 To printLength
                        printDataTableFlgList.Add(MyConst.Variables.PrintColumn(w, 4))
                    Next
                    ' printDataTableフラグ返し
                    Return printDataTableFlgList

                ' printDbFlg設定
                Case "printdbflg"
                    ' Forステートメントを使って列挙
                    For x As Integer = 0 To printLength
                        printDbFlgList.Add(MyConst.Variables.PrintColumn(x, 5))
                    Next
                    ' printDbフラグ返し
                    Return printDbFlgList

                ' printColumn設定
                Case "printcolumnflg"
                    ' Forステートメントを使って列挙
                    For y As Integer = 0 To printLength
                        printDbColumnList.Add(MyConst.Variables.PrintColumn(y, 6))
                    Next
                    ' printColumn返し
                    Return printDbColumnList

                ' printボタン設定
                Case "printbutton"
                    ' Forステートメントを使って列挙
                    For z As Integer = 0 To printLength
                        printButtonList.Add(MyConst.Variables.PrintColumn(z, 7))
                    Next
                    ' printボタン返し
                    Return printButtonList

                ' check設定
                Case "printcheck"
                    ' Forステートメントを使って列挙
                    For l As Integer = 0 To itemLength
                        printCheckList.Add(MyConst.Variables.ItemColumn(l, 8))
                    Next
                    ' ボタン返し
                    Return printCheckList
                
                ' patternヘッダ設定
                Case "patternheader"
                    ' Forステートメントを使って列挙
                    For a As Integer = 0 To patternLength
                        patternColumnList.Add(MyConst.Variables.PatternColumn(a, 0))
                    Next
                    ' printヘッダ返し
                    Return patternColumnList

                ' pattern型設定
                Case "patterntype"
                    ' Forステートメントを使って列挙
                    For b As Integer = 0 To patternLength
                        patternTypeList.Add(MyConst.Variables.PatternColumn(b, 1))
                    Next
                    ' pattern型返し
                    Return patternTypeList

                ' templateヘッダ設定
                Case "templateheader"
                    ' Forステートメントを使って列挙
                    For c As Integer = 0 To templateLength
                        templateColumnList.Add(MyConst.Variables.TemplateColumn(c, 0))
                    Next
                    ' templateヘッダ返し
                    Return templateColumnList

                ' template型設定
                Case "templatetype"
                    ' Forステートメントを使って列挙
                    For d As Integer = 0 To templateLength
                        templateTypeList.Add(MyConst.Variables.TemplateColumn(d, 1))
                    Next
                    ' template型返し
                    Return templateTypeList

                Case Else
                    ' 上記以外はエラー
                    Return New List(Of String)({})
                    Console.WriteLine("Error: item取得エラー")
            End Select

        Catch ex As System.IO.IOException
            ' 空ジェネリック返し
            Return New List(Of String)({})
            Console.WriteLine(ex)
        End Try
    End Function
End Module
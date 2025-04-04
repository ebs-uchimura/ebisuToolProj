Option Strict On
Option Explicit On
Option Infer On

' import module
' 項目名, 型, 幅, GridView表示, DataTable表示, MySQL格納, MySQLカラム名, ボタン文字列, チェックボックス
Namespace MyConst
    Public Module Variables
        ' itemカラム
        Public ItemColumn As String(,) = {
          {"itemID", "String", "supershort", "1", "1", "1", "id", "", "0"},
          {"日付", "String", "supershort", "0", "1", "1", "date", "", "0"},
          {"指示書", "String", "superwide", "1", "1", "1", "itemname", "", "0"},
          {"件数", "String", "short", "1", "1", "1", "amount", "", "0"},
          {"枚数", "String", "short", "1", "1", "1", "labelnumber", "", "0"},
          {"フォルダ", "String", "short", "0", "1", "1", "folderpath", "", "0"},
          {"フォルダ作成", "String", "short", "0", "0", "0", "", "作成", "0"},
          {"配置進捗", "String", "short", "1", "1", "1", "layoutprogress", "", "0"},
          {"配置選択", "Boolean", "short", "1", "1", "0", "", "", "1"},
          {"フォルダ開く", "String", "short", "0", "0", "0", "", "開く", "0"},
          {"検出OK", "String", "short", "1", "1", "0", "", "", "0"},
          {"検出NG", "String", "short", "1", "1", "0", "", "", "0"},
          {"面付選択", "Boolean", "short", "1", "1", "0", "", "", "1"},
          {"指示書状態", "String", "short", "0", "1", "0", "", "", "0"},
          {"指示書編集", "String", "short", "0", "0", "0", "", "編集", "0"},
          {"進捗", "String", "short", "1", "1", "1", "progress", "", "0"},
          {"変更前", "String", "supershort", "1", "1", "0", "", "", "0"},
          {"配置出力時間", "String", "short", "1", "1", "1", "printtime", "", "0"},
          {"PB確認希望", "String", "short", "1", "1", "1", "pbcheck", "", "0"},
          {"DL済", "String", "short", "1", "1", "1", "dl", "", "0"},
          {"作業用DL", "String", "short", "0", "0", "0", "", "DL", "0"},
          {"前同", "String", "short", "0", "1", "1", "zendoflg", "", "0"},
          {"担当者", "String", "short", "0", "1", "1", "printstaffname", "", "0"},
          {"背景色", "String", "short", "0", "1", "0", "", "", "0"},
          {"指示書重複", "String", "short", "0", "1", "0", "", "", "0"}
        }

        ' printカラム
        Public PrintColumn As String(,) = {
          {"printID", "String", "supershort", "1", "1", "1", "id", "", "0"},
          {"itemID", "String", "supershort", "1", "1", "1", "itemtable_id", "", "0"},
          {"前回発送日", "String", "standard", "1", "1", "1", "shippingdate", "", "0"},
          {"顧客番号", "String", "standard", "1", "1", "1", "customerno", "", "0"},
          {"商品名", "String", "middle", "1", "1", "1", "productname", "", "0"},
          {"枚数", "String", "short", "1", "1", "1", "pages", "", "0"},
          {"旧ID", "String", "short", "1", "1", "1", "oldid", "", "0"},
          {"新ID", "String", "short", "1", "1", "1", "newid", "", "0"},
          {"検出状況", "String", "short", "1", "1", "1", "searchprogress", "", "0"},
          {"検出ﾌｧｲﾙﾊﾟｽ", "String", "standard", "1", "1", "1", "findfilepath", "", "0"},
          {"作業ﾌｧｲﾙﾊﾟｽ", "String", "standard", "1", "1", "1", "localpath", "", "0"},
          {"幅", "String", "short", "1", "1", "1", "width", "", "0"},
          {"高さ", "String", "short", "1", "1", "1", "height", "", "0"},
          {"出力", "String", "short", "1", "1", "1", "outputstatus", "", "0"},
          {"ファイル更新", "String", "short", "0", "0", "0", "", "更新", "0"},
          {"部分配置", "Boolean", "short", "1", "1", "0", "", "", "1"},
          {"部分面付", "Boolean", "short", "1", "1", "0", "", "", "1"},
          {"ラベル名", "String", "wide", "1", "1", "1", "labelname", "", "0"},
          {"AiOpen", "String", "short", "0", "0", "0", "", "開く", "0"},
          {"確認希望", "String", "short", "1", "1", "1", "confirm", " ", "0"},
          {"PB確認", "String", "short", "1", "1", "1", "pbcheck", "", "0"},
          {"背景色", "String", "short", "0", "1", "0", "", "", "0"}
        }

        ' patternカラム
        Public PatternColumn As String(,) = {
          {"patternID", "String", "id"},
          {"ﾊﾟﾀｰﾝ名", "String", "patternname"},
          {"略称", "String", "abbreviation"},
          {"ﾃﾝﾌﾟﾚｰﾄID", "String", "template_id"},
          {"宇宿ﾌﾟﾘﾝﾀID", "String", "uprinter_id"},
          {"山梨ﾌﾟﾘﾝﾀID", "String", "yprinter_id"},
          {"ｼｰﾄ幅", "String", "sheetwidth"},
          {"ｼｰﾄ高", "String", "sheetheight"},
          {"横枚数", "String", "wcount"},
          {"縦枚数", "String", "hcount"},
          {"幅ｵﾌｾｯﾄ", "String", "woffset"},
          {"高さｵﾌｾｯﾄ", "String", "hoffset"},
          {"幅上限", "String", "wupper"},
          {"幅下限", "String", "wlower"},
          {"高さ上限", "String", "hupper"},
          {"高さ下限", "String", "hlower"},
          {"300", "String", "special"},
          {"幅開始", "String", "wstart"},
          {"高さ開始", "String", "hstart"}
        }

        ' templateカラム
        Public TemplateColumn As String(,) = {
          {"templateID", "String", "id"},
          {"ﾀｲﾌﾟID", "String", "type_id"},
          {"ﾃﾝﾌﾟﾚｰﾄ名", "String", "templatename"},
          {"ﾃﾝﾌﾟﾚｰﾄﾌｧｲﾙ名", "String", "templatefilename"},
          {"ﾃﾝﾌﾟﾚｰﾄ幅", "String", "templatewidth"},
          {"ﾃﾝﾌﾟﾚｰﾄ高さ", "String", "templateheight"}
        }
        
    End Module

    Public Module ItemTextNames
        ' 配置進捗
        Public layoutProgress As String() = {" ", "検出中", "検出有", "検出無", "EPS処理済", "完了", "編集", "編集中"}
        ' 進捗
        Public progress As String() = {" ", "未処理", "検索済", "配置出力済", "配置確認済", "面付出力済", "確認待ち"}
    End Module

    Public Module PrintTextNames
        ' 検出状況
        Public detectStatus As String() = {"検出中", "検出成功", "手動検出", "変更変換済み", "検出中…", "検出有", "検出無", "EPS処理済", "完了", "変更済みAi"}
        ' 出力
        Public outResult As String() = {"OK", "NG"}
        ' 出力
        Public pbCheckResult As String() = {"待ち", "済×", "済〇"}
    End Module

    Public Module ItemButtonNames
        ' 作成書編集
        Public instructionEditButton As String() = {"編集", "編集中"}
    End Module

    Public Module PrintButtonNames
        ' ファイル更新
        Public fileUpdate As String() = {"更新", "×"}
        ' 確認希望
        Public needCheck As String() = {" ", "送信"}
    End Module
End Namespace
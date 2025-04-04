Option Strict On
Option Explicit On
Option Infer On

Module conf
    ' 1point=2.835mm
    Friend globalPtm As Double = 2.835
    ' 111.5mm
    Friend globalPbw As Double = 111.5 * globalPtm
    ' 84mm
    Friend globalPbh As Double = 94 * globalPtm
    ' 選択中のアイテム
    Friend globalSelectedItem As Integer
    ' 現在のユーザ
    Friend globalUserName As String
    ' 現在の日付
    Friend globalNowDate As String
    ' printer設定記述ファイル
    Friend sPrinterFilePath As String = "C:\Users\ebisudo\Documents\■EbisuAutoPrintSet\printConfig.txt"
    ' サーバ・ルートパス
    Friend globalRootPath As String = "H:\TEST\■EbisuAutoPrintSet\work\"
    ' ローカル・ルートパス
    Friend globalLocalRootPath As String = "C:\Users\ebisudo\Documents\01_ebisuDEV\10_AutoPrint\Work\"
    ' AdminUser
    Friend adminUser As String = "ADMIN_USER"
    ' testUser
    Friend testUser As String = "TEST_USER"
    ' PB確認機能
    Friend globalPBCheck As Boolean
    ' 検証用LabelHDパス
    Friend testLabelHDPath As String = "H:\TEST\20"

    ' プリンタ一覧取得
    Public Function getPrinter() As String()
        Dim linesPrinter As String() = System.IO.File.ReadAllLines(sPrinterFilePath, System.Text.Encoding.Default)
        Return linesPrinter
    End Function

    ' プリンタ設定
    Public Function setPrinter(ByVal linesPrinter As String()) As String()
        IO.File.WriteAllText(sPrinterFilePath, "", System.Text.Encoding.GetEncoding("shift_jis"))
        For i As Integer = 0 To linesPrinter.Length - 1
            IO.File.AppendAllText(sPrinterFilePath, linesPrinter(i) & vbCrLf, System.Text.Encoding.GetEncoding("shift_jis"))
        Next
        Return linesPrinter
    End Function

    ' 変換用dictionary
    Friend sizeDictionary As Dictionary(Of String, String) = New Dictionary(Of String, String) From {
            {"supershort", "35"},
            {"short", "55"},
            {"standard", "75"},
            {"middle", "120"},
            {"wide", "120"},
            {"superwide", "350"},
            {"none", "0"}
    }
End Module

Option Explicit On
Option Infer On

' ◇ Adobe Illustrator クラス
Public Class adobeAI
    ' 共通プロパティ
    Private Property appRefCnv As Illustrator.Application
    Private Property docRefCnv As Illustrator.Document

    '--------------------------------------------------------
    '*メソッド名：New
    '*機  能：初期化
    '*分　類：パブリック
    '--------------------------------------------------------
    Sub New()
        appRefCnv = CType(CreateObject("Illustrator.Application"), Illustrator.Application)
    End Sub

    '--------------------------------------------------------
    '*メソッド名：Open
    '*機  能：ドキュメント開く
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub Open(path As String)
        docRefCnv = appRefCnv.Open(path)
    End Sub

    '--------------------------------------------------------
    '*メソッド名：Add
    '*機  能：ドキュメント追加
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub Add(Optional width As Integer = 600, Optional height As Integer = 600)
        docRefCnv = appRefCnv.Documents.Add(Illustrator.AiDocumentColorSpace.aiDocumentCMYKColor, width, height)
    End Sub

    '--------------------------------------------------------
    '*メソッド名：Save
    '*機  能：ドキュメント閉じる
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub Save(path As String)
        docRefCnv.SaveAs(path)
    End Sub

    '--------------------------------------------------------
    '*メソッド名：Close
    '*機  能：ドキュメント閉じる
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub Close()
        docRefCnv.Close(2)
    End Sub

    '--------------------------------------------------------
    '*メソッド名：TextLayout
    '*機  能：テキスト配置
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub TextLayout(str As String, fontSize As Double, Optional transX As Double = 0, Optional transY As Double = 0)
        Dim txtObj = docRefCnv.TextFrames.Add()
        Try
            txtObj.Contents = str
            txtObj.TextRange.CharacterAttributes.Size = fontSize
            txtObj.Translate(transX, transY)

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*メソッド名：EPSLayout
    '*機  能：ESPファイル配置
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub EPSLayout(idx As Integer, filePath As String, fileName As String, Optional transX As Double = 600, Optional transY As Double = 600)
        ' 変数定義
        Dim tmpX
        Dim tmpY
        Dim placedImags(100) As Illustrator.PlacedItem
        Dim txtObj = docRefCnv.TextFrames.Add()

        Try
            tmpX = placedImags(idx).Position(0) * -1
            tmpY = placedImags(idx).Position(1) * -1
            placedImags(idx) = docRefCnv.PlacedItems.Add()
            placedImags(idx).File = filePath
            placedImags(idx).Name = fileName
            placedImags(idx).Translate(tmpX, tmpY)
            placedImags(idx).Translate(transX, transY)

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*メソッド名：ExchangeEPS
    '*機  能：EPS変換
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub ExchangeEPS(path As String, outputPath As String)
        ' 変数定義
        Dim newSaveOptionsCnv As Illustrator.EPSSaveOptions = CType(CreateObject("Illustrator.EPSSaveOptions"), Illustrator.EPSSaveOptions)
        Try
            newSaveOptionsCnv.CMYKPostScript = True
            newSaveOptionsCnv.EmbedAllFonts = True
            docRefCnv.SaveAs(outputPath, newSaveOptionsCnv)

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*メソッド名：ExchangeJPG
    '*機  能：JPG変換
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub ExchangeJPG(epspath As String, outputPath As String, width As Double, height As Double)
        ' 変数定義
        Dim placedEpsSize(0) As Illustrator.PlacedItem
        Try
            docRefCnv = appRefCnv.Documents.Add(Illustrator.AiDocumentColorSpace.aiDocumentCMYKColor, width, height)
            placedEpsSize(0) = docRefCnv.PlacedItems.Add()
            placedEpsSize(0).File = epspath
            docRefCnv.Export(outputPath, Illustrator.AiExportType.aiJPEG)

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*メソッド名：MakeEmptyESP
    '*機  能：空EPS作成
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub MakeEmptyESP()
        ' 変数定義
        Dim newSaveOptionsCnv = CreateObject("Illustrator.EPSSaveOptions")
        Try
            newSaveOptionsCnv.CMYKPostScript = True
            newSaveOptionsCnv.EmbedAllFonts = True
            Me.Close()

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*メソッド名：LayoutLabels
    '*機  能：ラベル配置
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub LayoutLabels(idx As Integer, filePath As String, fileName As String, Optional transX As Double = 600, Optional transY As Double = 600, Optional initFlg As Boolean = False)
        ' 変数定義
        Dim tmpX
        Dim tmpY
        Dim placedImags(100) As Illustrator.PlacedItem
        Dim txtObj = docRefCnv.TextFrames.Add()

        Try
            tmpX = placedImags(idx).Position(0) * -1
            tmpY = placedImags(idx).Position(1) * -1
            placedImags(idx) = docRefCnv.PlacedItems.Add()
            placedImags(idx).File = filePath
            placedImags(idx).Name = fileName
            placedImags(idx).Translate(tmpX, tmpY)
            placedImags(idx).Translate(transX, transY)

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*メソッド名：PrintDocument
    '*機  能：ラベル・配置図印刷
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Sub PrintDocument(usukiId As Integer, Optional yamanashiId As Integer = 0)
        ' 変数定義
        Dim coordinateOptions As New Illustrator.PrintCoordinateOptions
        Dim printOptions As New Illustrator.PrintOptions
        ' オブジェクト定義
        Dim masterDataSet As DataSet
        Dim printerDataTable As DataTable
        Try
            ' オプション設定
            printOptions.CoordinateOptions = coordinateOptions
            coordinateOptions.FitToPage = True
            coordinateOptions.Orientation = 1
            ' マスタ取得
            masterDataSet = GetMasterData()
            printerDataTable = masterDataSet.Tables(2)

            If MainForm.RadioButton8.Checked Then
                ' 宇宿
                printOptions.PrinterName = printerDataTable.Rows(usukiId).item(1)
            Else
                If yamanashiId > 0 Then
                    ' 山梨
                    printOptions.PrinterName = printerDataTable.Rows(yamanashiId).item(1)
                Else
                    ' データなし
                    MsgBox("該当するプリンタがありません")
                    Throw New System.Exception("An exception has occurred.")
                End If
            End If
            ' プリントアウト
            docRefCnv.PrintOut(printOptions)

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
        End Try
    End Sub

    '--------------------------------------------------------
    '*メソッド名：getEpsSize
    '*機  能：EPSサイズ取得
    '*分　類：パブリック
    '--------------------------------------------------------
    Public Function getEpsSize(path As String, width As Double, height As Double) As Double()
        ' 変数定義
        Dim placedEpsSize(0) As Illustrator.PlacedItem
        Try
            docRefCnv = appRefCnv.Documents.Add(Illustrator.AiDocumentColorSpace.aiDocumentCMYKColor, width, height)
            placedEpsSize(0) = docRefCnv.PlacedItems.Add()
            placedEpsSize(0).File = path
            Return {placedEpsSize(0).Width, placedEpsSize(0).Height}

        Catch ex As System.IO.IOException
            Console.WriteLine(ex)
            Return {0, 0}
        End Try
    End Function
End Class

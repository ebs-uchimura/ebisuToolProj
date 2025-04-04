Imports Microsoft.VisualBasic.FileIO
Imports System.Text
Imports System.IO.Path
Imports System.Runtime.InteropServices.Marshal
Imports System.GC

Public Class Form1
    ' 変数定義
    Dim printoutFlg As Boolean = True ' 検証用プリントフラグ
    Dim configCsv As New List(Of String())
    Dim configSetCsv As New List(Of String())

    ' フォーム読み込み時処理
    ' ラジオボタンリセット
    Public Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ' 変数定義
        ' ラジオボタンリセット
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        RadioButton5.Checked = False
        RadioButton6.Checked = False
        RadioButton7.Checked = False
        RadioButton8.Checked = False
        RadioButton9.Checked = False
        RadioButton10.Checked = False
        RadioButton11.Checked = False
        RadioButton12.Checked = False
        RadioButton13.Checked = False
        RadioButton14.Checked = False
        RadioButton15.Checked = False
        RadioButton16.Checked = False
        RadioButton17.Checked = False
        RadioButton18.Checked = False
        RadioButton19.Checked = False
        RadioButton20.Checked = False
        RadioButton21.Checked = False
        RadioButton22.Checked = False

        configCsv = loadCsv(GetAppPath() + "\Work\config.csv")
        setAutoComplete(configCsv)
        configSetCsv = loadCsv(GetAppPath() + "\Work\configSet.csv")
        
        For Each p As String In Printing.PrinterSettings.InstalledPrinters
            If p.Contains("ポスター") Then
                If p.Contains("2号機") Then
                    RadioButton25.Text = p
                    RadioButton25.AutoCheck = True
                End If
            End If
            If p.Contains("名刺") Then
                If p.Contains("2号機") Then
                    RadioButton26.Text = p
                    RadioButton26.AutoCheck = True
                End If
            End If
            If p.Contains("7171") Then
                If p.Contains("new_poster") Then
                    RadioButton27.Text = p
                    RadioButton27.AutoCheck = True
                Else If p.Contains("名刺") Then
                    RadioButton28.Text = p
                    RadioButton28.AutoCheck = True
                End If
            End If
        Next

        ' poster変換exeフォルダ存在確認
        If System.IO.Directory.Exists("c:\ebisuPoster\exe\") Then
        Else
            ' フォルダ作成する
            System.IO.Directory.CreateDirectory("\ebisuPoster\exe\")
        End If

        ' 変換exeファイル確認＆copy
        Dim searchFilePathA As String = "c:\ebisuPoster\exe\"
        Dim searchFilePathB As String = GetAppPath() + "\Work\"
        Dim searchFileKey As String = "*.exe"
        Dim filesA As String()
        filesA = IO.Directory.GetFiles(searchFilePathA, searchFileKey)

        If filesA.Length <> 9 Then
            Dim filesB As String()
            filesB = IO.Directory.GetFiles(searchFilePathB, searchFileKey)
            ' 既に同名のファイルが存在していても上書
            For i As Integer = 0 To filesB.Length - 1
                IO.File.Copy(filesB(i), searchFilePathA + GetFileName(filesB(i)), True)
            Next
        End If
    End Sub

    ' テキストボックスA　ドラッグでファイルのみ受け付け
    Private Sub txtTemp_DragEnterA(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtTempA.DragEnter
        ' ファイル形式の場合のみ、ドラッグを受け付けます。
        If e.Data.GetDataPresent(DataFormats.FileDrop) = True Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    ' テキストボックスAへaiファイルドラッグドロップされた時の処理
    Private Sub txtTemp_DragDropA(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtTempA.DragDrop
        ' ドラッグされたファイル・フォルダのパスを格納
        Dim strFileName As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())
        ' ファイルの存在確認を行い、ある場合にのみ、テキストボックスにパスを表示します。（この処理でフォルダを対象外に）
        If System.IO.File.Exists(strFileName(0).ToString) = True Then
            Me.txtTempA.Text = strFileName(0).ToString
        End If
        aiA()
    End Sub

    'aiファイルA処理
    Private Sub aiA()
        'ラジオボタンリセット
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        RadioButton5.Checked = False
        RadioButton6.Checked = False
        RadioButton7.Checked = False
        RadioButton8.Checked = False
        RadioButton17.Checked = False
        RadioButton19.Checked = False
        RadioButton21.Checked = False

        'posterフォルダ存在確認
        If System.IO.Directory.Exists(GetDirectoryName(Me.txtTempA.Text) + "\poster\") Then
        Else
            ' フォルダ作成する
            System.IO.Directory.CreateDirectory(GetDirectoryName(Me.txtTempA.Text) + "\poster\")
        End If

        'EPS変換
        txtTempProgressA.BackColor = Color.PaleGreen
        txtTempProgressA.Text = "EPS変換中"
        cnvEps(txtTempA.Text) '.Replace(" "c, "_"c)

        'PBサイズ取得
        'PSDエクスポート
        'PBサイズ確定
        txtTempProgressA.BackColor = Color.PaleTurquoise
        txtTempProgressA.Text = "サイズ確定中"
        getPbSize(txtTempA.Text)

        'PSD変形
        txtTempProgressA.BackColor = Color.CornflowerBlue
        txtTempProgressA.Text = "PSD変形中"
        bendPsdA(txtTempA.Text)

        '屋号ai作成
        If CheckBox4.Checked = True Then
            createYagoAi(txtTempA.Text)
        End If

        txtTempProgressA.BackColor = Color.WhiteSmoke
        txtTempProgressA.Text = "PSD変形終了"

        Me.Activate()

    End Sub

    'テキストボックスB　ドラッグでファイルのみ受け付け
    Private Sub txtTemp_DragEnterB(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtTempB.DragEnter

        'ファイル形式の場合のみ、ドラッグを受け付けます。
        If e.Data.GetDataPresent(DataFormats.FileDrop) = True Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If

    End Sub

    'テキストボックスBへaiファイルドラッグドロップされた時の処理
    Private Sub txtTemp_DragDropB(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtTempB.DragDrop

        'ドラッグされたファイル・フォルダのパスを格納
        Dim strFileName As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())

        'ファイルの存在確認を行い、ある場合にのみ、テキストボックスにパスを表示します。
        '（この処理でフォルダを対象外に）
        If System.IO.File.Exists(strFileName(0).ToString) = True Then
            Me.txtTempB.Text = strFileName(0).ToString
        End If

        aiB()
    End Sub

    'aiファイルB処理
    Private Sub aiB()
        'ラジオボタンリセット
        RadioButton9.Checked = False
        RadioButton10.Checked = False
        RadioButton11.Checked = False
        RadioButton12.Checked = False
        RadioButton13.Checked = False
        RadioButton14.Checked = False
        RadioButton15.Checked = False
        RadioButton16.Checked = False
        RadioButton18.Checked = False
        RadioButton20.Checked = False
        RadioButton22.Checked = False

        'posterフォルダ存在確認
        If System.IO.Directory.Exists(GetDirectoryName(Me.txtTempB.Text) + "\poster\") Then
        Else
            ' フォルダ作成する
            System.IO.Directory.CreateDirectory(GetDirectoryName(Me.txtTempB.Text) + "\poster\")
        End If

        'EPS変換
        txtTempProgressB.BackColor = Color.PaleGreen
        txtTempProgressB.Text = "EPS変換中"
        cnvEps(txtTempB.Text)

        'PBサイズ取得
        'PSDエクスポート
        'PBサイズ確定
        txtTempProgressB.BackColor = Color.PaleTurquoise
        txtTempProgressB.Text = "サイズ確定中"
        getPbSizeB(txtTempB.Text)

        'PSD変形
        txtTempProgressB.BackColor = Color.CornflowerBlue
        txtTempProgressB.Text = "PSD変形中"
        bendPsdB(txtTempB.Text)

        'PSD変形終了
        txtTempProgressB.BackColor = Color.WhiteSmoke
        txtTempProgressB.Text = "PSD変形終了"

        Me.Activate()
    End Sub

    'テキストボックスC　ドラッグでファイルのみ受け付け
    Private Sub txtTemp_DragEnterC(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtTempC.DragEnter

        'ファイル形式の場合のみ、ドラッグを受け付けます。
        If e.Data.GetDataPresent(DataFormats.FileDrop) = True Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    'テキストボックスCへepsファイルドラッグドロップされた時の処理
    Private Sub txtTemp_DragDropC(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtTempC.DragDrop

        'ドラッグされたファイル・フォルダのパスを格納
        Dim strFileName As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())

        'ファイルの存在確認を行い、ある場合にのみ、テキストボックスにパスを表示します。
        '（この処理でフォルダを対象外に）
        If System.IO.File.Exists(strFileName(0).ToString) = True Then
            Me.txtTempC.Text = strFileName(0).ToString
        End If

        Me.Activate()
    End Sub

    'テキストボックスD　ドラッグでファイルのみ受け付け
    Private Sub txtTemp_DragEnterD(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtTempD.DragEnter
        'ファイル形式の場合のみ、ドラッグを受け付けます。
        If e.Data.GetDataPresent(DataFormats.FileDrop) = True Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    'テキストボックスDへepsファイルドラッグドロップされた時の処理
    Private Sub txtTemp_DragDropD(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtTempD.DragDrop
        'ドラッグされたファイル・フォルダのパスを格納
        Dim strFileName As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())

        'ファイルの存在確認を行い、ある場合にのみ、テキストボックスにパスを表示します。
        '（この処理でフォルダを対象外に）
        If System.IO.File.Exists(strFileName(0).ToString) = True Then
            Me.txtTempD.Text = strFileName(0).ToString
        End If

        Me.Activate()
    End Sub

    'アプリ実行パス取得
    Public Shared Function GetAppPath() As String
        Return System.IO.Path.GetDirectoryName(
            System.Reflection.Assembly.GetExecutingAssembly().Location)
    End Function

    'aiファイルをposterディレクトリへEPS保存
    Public Sub cnvEps(ByVal aiFilePath As String)
        'aiファイルオープン
        ''EPS変換
        Dim appRefCnv As Illustrator.Application = Nothing
        Dim docRefCnv As Illustrator.Document = Nothing
        Dim newSaveOptionsCnv As Illustrator.EPSSaveOptions

        Try
            appRefCnv = CreateObject("Illustrator.Application.29")
            docRefCnv = appRefCnv.Open(aiFilePath)
            newSaveOptionsCnv = CreateObject("Illustrator.EPSSaveOptions")
            newSaveOptionsCnv.CMYKPostScript = True
            newSaveOptionsCnv.EmbedAllFonts = True
            docRefCnv.SaveAs(
                GetDirectoryName(aiFilePath) + "\poster\" + GetFileName(aiFilePath), newSaveOptionsCnv)
            docRefCnv.Close(2)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            If Not docRefCnv Is Nothing Then ReleaseComObject(docRefCnv) 'COMオブジェクトの開放
            If Not appRefCnv Is Nothing Then ReleaseComObject(appRefCnv) 'COMオブジェクトの開放
            GC.Collect()
        End Try
    End Sub

    'epsファイルからサイズ取得しラベル確定
    Public Sub getPbSize(ByVal aiFilePath As String)
        
        Dim appRefCnv As Illustrator.Application = Nothing
        Dim docRefCnv As Illustrator.Document = Nothing
        Dim placedEpsSize(0) As Illustrator.PlacedItem
        Dim psdExportOptions = CreateObject("Illustrator.ExportOptionsPhotoshop")
        Dim epsWidth As Integer = 0
        Dim epsHeight As Integer = 0
        Dim pbPsdPath As String = GetDirectoryName(aiFilePath) + "\poster\" + GetFileNameWithoutExtension(aiFilePath) + ".psd"
Try
        appRefCnv = CreateObject("Illustrator.Application.29")
        'イラレドキュメント生成
        docRefCnv = appRefCnv.Documents.Add(Illustrator.AiDocumentColorSpace.aiDocumentCMYKColor, 600, 600)
        placedEpsSize(0) = docRefCnv.PlacedItems.Add()
        placedEpsSize(0).File = GetDirectoryName(aiFilePath) + "\poster\" + GetFileNameWithoutExtension(aiFilePath) + ".eps"

        '幅と高さを取得
        epsWidth = placedEpsSize(0).Width
        epsHeight = placedEpsSize(0).Height
        'MessageBox.Show(epsWidth.ToString + vbCrLf + epsHeight.ToString)

        'PSDエクスポート
        psdExportOptions.Resolution = 200
        docRefCnv.Export(pbPsdPath, 2, psdExportOptions) ' 2 = aiPhotoshop
        docRefCnv.Close(2)

        'PBサイズ確定
        'ラベルをサイズでチェック
        '720横 / 呑神
        If epsWidth >= 313 And epsWidth <= 317 Then
            If epsHeight >= 265 And epsHeight <= 267 Then
                '720mlyoko
                RadioButton1.Checked = True
            ElseIf epsHeight >= 400 Then
                '呑神
                RadioButton8.Checked = True
            Else
                RadioButton5.Checked = True
            End If
        End If

        'Wine / DRY
        If epsWidth >= 240 And epsWidth <= 246 Then
            If epsHeight >= 325 And epsHeight <= 330 Then
                RadioButton3.Checked = True
            Else
                RadioButton4.Checked = True
            End If
        End If

        '720縦
        If epsWidth >= 263 And epsWidth <= 269 Then
            RadioButton2.Checked = True
        End If

        '1800
        If epsWidth >= 405 And epsWidth <= 410 Then
            RadioButton6.Checked = True
        End If

        'LIGHT
        If epsWidth >= 212 And epsWidth <= 217 Then
            RadioButton7.Checked = True
        End If

        'SUIJIN WINE HELF
        If epsWidth >= 206 AndAlso epsWidth <= 208 Then
            If epsHeight >= 278 And epsHeight <= 281 Then
                RadioButton17.Checked = True
            End If
        End If

        'SUIJIN WINE Rouge
        If epsWidth >= 242 AndAlso epsWidth <= 246 Then
            If epsHeight >= 256 And epsHeight <= 260 Then
                RadioButton19.Checked = True
            End If
        End If

        'SUIJIN WINE Soleil
        If epsWidth >= 298 AndAlso epsWidth <= 302 Then
            If epsHeight >= 236 And epsHeight <= 240 Then
                RadioButton21.Checked = True
            End If
        End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            If Not docRefCnv Is Nothing Then ReleaseComObject(docRefCnv) 'COMオブジェクトの開放
            If Not appRefCnv Is Nothing Then ReleaseComObject(appRefCnv) 'COMオブジェクトの開放
            GC.Collect()
        End Try
    End Sub

    'epsファイルからサイズ取得しラベル確定
    Public Sub getPbSizeB(ByVal aiFilePath As String)
        
        Dim appRefCnv As Illustrator.Application = Nothing
        Dim docRefCnv As Illustrator.Document = Nothing
        Dim placedEpsSize(0) As Illustrator.PlacedItem
        Dim psdExportOptions = CreateObject("Illustrator.ExportOptionsPhotoshop")
        Dim epsWidth As Integer = 0
        Dim epsHeight As Integer = 0
        Dim pbPsdPath As String = GetDirectoryName(aiFilePath) + "\poster\" + GetFileNameWithoutExtension(aiFilePath) + ".psd"

        Try
        appRefCnv = CreateObject("Illustrator.Application.29")
        'イラレドキュメント生成
        docRefCnv = appRefCnv.Documents.Add(Illustrator.AiDocumentColorSpace.aiDocumentCMYKColor, 600, 600)

        placedEpsSize(0) = docRefCnv.PlacedItems.Add()
        placedEpsSize(0).File = GetDirectoryName(aiFilePath) + "\poster\" + GetFileNameWithoutExtension(aiFilePath) + ".eps"


        '幅と高さを取得
        epsWidth = placedEpsSize(0).Width
        epsHeight = placedEpsSize(0).Height
        'MessageBox.Show(epsWidth.ToString + vbCrLf + epsHeight.ToString)

        'PSDエクスポート
        psdExportOptions.Resolution = 200
        docRefCnv.Export(pbPsdPath, 2, psdExportOptions) ' 2 = aiPhotoshop
        docRefCnv.Close(2)

        'PBサイズ確定
        'ラベルをサイズでチェック
        '720横 / 呑神
        If epsWidth >= 313 And epsWidth <= 317 Then
            If epsHeight >= 265 And epsHeight <= 267 Then
                '720mlyoko
                RadioButton16.Checked = True
            ElseIf epsHeight >= 400 Then
                '呑神
                RadioButton9.Checked = True
            Else
                RadioButton12.Checked = True
            End If
        End If

        'Wine / DRY / 900
        If epsWidth >= 240 And epsWidth <= 246 Then
            If epsHeight >= 325 And epsHeight <= 330 Then
                RadioButton14.Checked = True
            Else
                RadioButton13.Checked = True
            End If
        End If

        '720縦
        If epsWidth >= 263 And epsWidth <= 269 Then
            RadioButton15.Checked = True
        End If

        '1800
        If epsWidth >= 405 And epsWidth <= 410 Then
            RadioButton11.Checked = True
        End If

        'LIGHT
        If epsWidth >= 212 And epsWidth <= 217 Then
            RadioButton10.Checked = True
        End If

        'SUIJIN WINE HELF
        If epsWidth >= 206 AndAlso epsWidth <= 208 Then
            If epsHeight >= 278 And epsHeight <= 281 Then
                RadioButton18.Checked = True
            End If
        End If

        'SUIJIN WINE Rouge
        If epsWidth >= 242 AndAlso epsWidth <= 246 Then
            If epsHeight >= 256 And epsHeight <= 260 Then
                RadioButton20.Checked = True
            End If
        End If

        'SUIJIN WINE Soleil
        If epsWidth >= 298 AndAlso epsWidth <= 302 Then
            If epsHeight >= 236 And epsHeight <= 240 Then
                RadioButton22.Checked = True
            End If
        End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            If Not docRefCnv Is Nothing Then ReleaseComObject(docRefCnv) 'COMオブジェクトの開放
            If Not appRefCnv Is Nothing Then ReleaseComObject(appRefCnv) 'COMオブジェクトの開放
            GC.Collect()
        End Try
    End Sub

    '屋号ai
    Public Sub createYagoAi(ByVal aiFilePath As String)
        Dim yago As String = GetFileNameWithoutExtension(aiFilePath)
        Dim stArrayData As String() = yago.Split(" "c)
        MessageBox.Show(stArrayData(2))
    End Sub

    'Aファイルの曲げ処理実行
    Public Sub bendPsdA(ByVal aiFilePath As String)
        Dim pbPsdPath As String = Chr(34) + GetDirectoryName(aiFilePath) + "\poster\" + GetFileNameWithoutExtension(aiFilePath) + ".psd" + Chr(34)
        Dim exeDir As String = "c:\ebisuPoster\exe\"
        If RadioButton1.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "720mlyoko.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton2.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "720mltate.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton3.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "SUIJINWINE.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton4.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "SUIJINDRY.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton5.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "900ml.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton6.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "1800ml.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton7.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "SUIJINLIGHT.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton8.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "donjin.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton17.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "SUIJINWINE_HALF.exe", pbPsdPath)
            p.WaitForExit()
        End If

        'Rouge
        If RadioButton19.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "SUIJINEWINE_Rouge.exe", pbPsdPath)
            p.WaitForExit()
        End If

        'Soleil
        If RadioButton21.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "SUIJINEWINE_Soleil.exe", pbPsdPath)
            p.WaitForExit()
        End If

        Me.Activate()
    End Sub

    'Bファイルの曲げ処理実行
    Public Sub bendPsdB(ByVal aiFilePath As String)
        Dim pbPsdPath As String = Chr(34) + GetDirectoryName(aiFilePath) + "\poster\" + GetFileNameWithoutExtension(aiFilePath) + ".psd" + Chr(34)
        Dim exeDir As String = "c:\ebisuPoster\exe\"
        If RadioButton16.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "720mlyoko.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton15.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "720mltate.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton14.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "SUIJINWINE.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton13.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "SUIJINDRY.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton12.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "900ml.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton11.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "1800ml.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton10.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "SUIJINLIGHT.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton9.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "donjin.exe", pbPsdPath)
            p.WaitForExit()
        End If

        If RadioButton18.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "SUIJINWINE_HALF.exe", pbPsdPath)
            p.WaitForExit()
        End If

        'Rouge
        If RadioButton20.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "SUIJINEWINE_Rouge.exe", pbPsdPath)
            p.WaitForExit()
        End If

        'Soleil
        If RadioButton22.Checked = True Then
            'ファイルを開いて終了まで待機する
            Dim p As System.Diagnostics.Process =
                System.Diagnostics.Process.Start(exeDir + "SUIJINEWINE_Soleil.exe", pbPsdPath)
            p.WaitForExit()
        End If
        Me.Activate()
    End Sub

    'List(Of string())へcsv読込
    Private Function loadCsv(csvFileName As String)
        'List(Of string())でCSVファイル読み込み
        Dim csvRecords As New List(Of String())
        Dim tfp As New FileIO.TextFieldParser(csvFileName, System.Text.Encoding.GetEncoding(932))
        tfp.TextFieldType = FileIO.FieldType.Delimited
        tfp.Delimiters = New String() {","}
        tfp.HasFieldsEnclosedInQuotes = True
        tfp.TrimWhiteSpace = True

        While Not tfp.EndOfData
            Dim fields As String() = tfp.ReadFields()
            csvRecords.Add(fields)
        End While

        tfp.Close()

        Return csvRecords
    End Function

    '資材番号オートコンプリート読込
    Private Sub setAutoComplete(dat As List(Of String()))
        Dim aeNoAutoList As New AutoCompleteStringCollection()
        TextBox1.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox1.AutoCompleteCustomSource = aeNoAutoList

        TextBox3.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox3.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox3.AutoCompleteCustomSource = aeNoAutoList

        TextBox5.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox5.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox5.AutoCompleteCustomSource = aeNoAutoList

        For loop1 = 0 To dat.Count - 1
            aeNoAutoList.Add(dat(loop1)(1))
        Next loop1
    End Sub

    '資材番号１　入力バリデーション
    Private Sub TextBox1_Validating(ByVal sender As Object,
                                    ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        For i = 0 To configCsv.Count - 1
            If configCsv(i)(1) = TextBox1.Text Then
                TextBox2.Text = configCsv(i)(2)

                For ii = 0 To configSetCsv.Count - 1
                    If configSetCsv(ii)(1) = TextBox1.Text Then

                        If configSetCsv(ii)(2).Length > 0 Then
                            TextBox3.Text = configSetCsv(ii)(2)
                            For iii = 0 To configCsv.Count - 1
                                If configCsv(iii)(1) = TextBox3.Text Then
                                    TextBox4.Text = configCsv(iii)(2)
                                    Exit For
                                Else
                                    TextBox4.Text = "入力・検索中"
                                End If
                            Next

                        End If

                        If configSetCsv(ii)(3).Length > 0 Then
                            TextBox5.Text = configSetCsv(ii)(3)
                            For iii = 0 To configCsv.Count - 1
                                If configCsv(iii)(1) = TextBox5.Text Then
                                    TextBox6.Text = configCsv(iii)(2)
                                    Exit For
                                Else
                                    TextBox6.Text = "入力・検索中"
                                End If
                            Next

                        End If

                        Exit For
                    End If

                Next

                Exit For
            Else
                TextBox2.Text = "入力・検索中"
            End If
        Next

        If TextBox1.Text = "" Then
            TextBox2.Text = ""
        End If

        If TextBox2.Text = "入力・検索中" And TextBox1.Text <> "" Then
            e.Cancel = True
            MessageBox.Show("資材番号が登録されていません。", "エラー",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    '資材番号２　入力バリデーション
    Private Sub TextBox3_Validating(ByVal sender As Object,
                                ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox3.Validating
        For i = 0 To configCsv.Count - 1
            If configCsv(i)(1) = TextBox3.Text Then
                TextBox4.Text = configCsv(i)(2)
                Exit For
            Else
                TextBox4.Text = "入力・検索中"
            End If
        Next

        If TextBox3.Text = "" Then
            TextBox4.Text = ""
        End If

        If TextBox4.Text = "入力・検索中" And TextBox3.Text <> "" Then
            e.Cancel = True
            MessageBox.Show("資材番号が登録されていません。", "エラー",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    '資材番号３　入力バリデーション
    Private Sub TextBox5_Validating(ByVal sender As Object,
                                ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox5.Validating
        For i = 0 To configCsv.Count - 1
            If configCsv(i)(1) = TextBox5.Text Then
                TextBox6.Text = configCsv(i)(2)
                Exit For
            Else
                TextBox6.Text = "入力・検索中"
            End If
        Next

        If TextBox5.Text = "" Then
            TextBox6.Text = ""
        End If

        If TextBox6.Text = "入力・検索中" And TextBox5.Text <> "" Then
            e.Cancel = True
            MessageBox.Show("資材番号が登録されていません。", "エラー",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    '資材番号１　枚数入力バリデーション
    Private Sub TextBox7_Validating(ByVal sender As Object,
                                    ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox7.Validating
        If IsNumeric(TextBox7.Text) = False And TextBox7.Text <> "" Then
            e.Cancel = True
            MessageBox.Show("数値以外が入力されています。", "エラー",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    '資材番号２　枚数入力バリデーション
    Private Sub TextBox8_Validating(ByVal sender As Object,
                                    ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox8.Validating
        If IsNumeric(TextBox8.Text) = False And TextBox8.Text <> "" Then
            e.Cancel = True
            MessageBox.Show("数値以外が入力されています。", "エラー",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    '資材番号３　枚数入力バリデーション
    Private Sub TextBox9_Validating(ByVal sender As Object,
                                ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox9.Validating
        If IsNumeric(TextBox9.Text) = False And TextBox9.Text <> "" Then
            e.Cancel = True
            MessageBox.Show("数値以外が入力されています。", "エラー",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    '2号機出力
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        printPoster("2号機")
    End Sub

    '山梨出力
    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        printPoster("山梨")
    End Sub

    '出力処理
    'printerName As String:出力先
    Private Sub printPoster(ByVal printerName As String)
        Dim checkSum As Integer = 0
        '資材チェックボックス1
        If CheckBox1.Checked = True Then
            If TextBox1.Text.Length > 0 Then

                If TextBox7.Text.Length = 0 Then
                    TextBox7.Text = 0
                End If

                If Integer.Parse(TextBox7.Text) > 0 Then
                    BackColor = Color.PaleGreen
                    'ポスター作製        
                    makePoster(TextBox1.Text, printerName, Integer.Parse(TextBox7.Text))
                    BackColor = SystemColors.Control
                Else
                    MessageBox.Show("資材番号1出力枚数が入力されていません")
                    TextBox7.Focus()
                End If

            Else
                MessageBox.Show("資材番号1が入力されていません")
                TextBox1.Focus()
            End If


        Else
            checkSum = checkSum + 1
        End If

        '資材チェックボックス2
        If CheckBox2.Checked = True Then
            If TextBox3.Text.Length > 0 Then

                If TextBox8.Text.Length = 0 Then
                    TextBox8.Text = 0
                End If
                If Integer.Parse(TextBox8.Text) > 0 Then
                    BackColor = Color.PaleGreen

                    'ポスター作製        
                    makePoster(TextBox3.Text, printerName, Integer.Parse(TextBox8.Text))

                    BackColor = SystemColors.Control
                Else
                    MessageBox.Show("資材番号2出力枚数が入力されていません")
                    TextBox8.Focus()
                End If

            Else
                MessageBox.Show("資材番号2が入力されていません")
                TextBox2.Focus()
            End If

        Else
            checkSum = checkSum + 1
        End If

        '資材チェックボックス3
        If CheckBox3.Checked = True Then
            If TextBox5.Text.Length > 0 Then

                If TextBox9.Text.Length = 0 Then
                    TextBox9.Text = 0
                End If

                If Integer.Parse(TextBox9.Text) > 0 Then
                    BackColor = Color.PaleGreen
                    'ポスター作製        
                    makePoster(TextBox5.Text, printerName, Integer.Parse(TextBox9.Text))
                    BackColor = SystemColors.Control
                Else
                    MessageBox.Show("資材番号3出力枚数が入力されていません")
                    TextBox9.Focus()
                End If

            Else
                MessageBox.Show("資材番号3が入力されていません")
                TextBox3.Focus()
            End If

        Else
            checkSum = checkSum + 1
        End If

        Me.Activate()

        If checkSum = 3 Then
            MessageBox.Show("チェックボックスが選択されていません")
            Exit Sub
        Else

            GC.Collect()

        End If

    End Sub

    '資材作成
    Private Sub makePoster(ByVal posterNo As String, ByVal printerName As String, ByVal printCount As Integer)
        Dim psdPath As String = ""
        Dim aiPath As String = ""
        Dim psdX As Double = 0
        Dim psdY As Double = 0
        Dim psdW As Double = 0
        Dim psdH As Double = 0
        Dim psdRotate As Double = 0

        'CSV file format
        '
        '  00:id 連番
        '  01:posterNo AE-nnnn
        '  02:資材名称
        '  03:path aiファイル
        '  04:psdXA1		 05:psdYA1		 06:psdWA1		 07:psdHA1		 08:psdRotateA1
        '  09:psdXA2		 10:psdYA2		 11:psdWA2		 12:psdHA2		 13:psdRotateA2
        '  14:psdXA3		 15:psdYA3		 16:psdWA3		 17:psdHA3		 18:psdRotateA3
        '  19:psdXA4		 20:psdYA4		 21:psdWA4		 22:psdHA4		 23:psdRotateA4
        '  24:psdXA5		 25:psdYA5		 26:psdWA5		 27:psdHA5		 28:psdRotateA5
        '  29:psdXA6		 30:psdYA6		 31:psdWA6		 32:psdHA6		 33:psdRotateA6
        '  34:psdXA7		 35:psdYA7		 36:psdWA7		 37:psdHA7		 38:psdRotateA7
        '  39:psdXA8		 40:psdYA8		 41:psdWA8		 42:psdHA8		 43:psdRotateA8

        '  44:psdXB1		 45:psdYB1		 46:psdWB1		 47:psdHB1		 48:psdRotateB1
        '  49:psdXB2		 50:psdYB2		 51:psdWB2		 52:psdHB2		 53:psdRotateB2
        '  54:psdXB3		 55:psdYB3		 56:psdWB3		 57:psdHB3		 58:psdRotateB3
        '  59:psdXB4		 60:psdYB4		 61:psdWB4		 62:psdHB4		 63:psdRotateB4
        '  64:psdXB5		 65:psdYB5		 66:psdWB5		 67:psdHB5		 68:psdRotateB5
        '  69:psdXB6		 70:psdYB6		 71:psdWB6		 72:psdHB6		 73:psdRotateB6
        '  74:psdXB7		 75:psdYB7		 76:psdWB7		 77:psdHB7		 78:psdRotateB7
        '  79:psdXB8		 80:psdYB8		 81:psdWB8		 82:psdHB8		 83:psdRotateB8

        '  84:epsXA1		 85:epsYA1		 86:epsWA1		 87:epsHA1		 88:epsRotateA1
        '  89:epsXA2		 90:epsYA2		 91:epsWA2		 92:epsHA2		 93:epsRotateA2
        '  94:epsXA3		 95:epsYA3		 96:epsWA3		 97:epsHA3		 98:epsRotateA3
        '  99:epsXA4		100:epsYA4		101:epsWA4		102:epsHA4		103:epsRotateA4
        ' 104:epsXA5		105:epsYA5		106:epsWA5		107:epsHA5		108:epsRotateA5
        ' 109:epsXA6		110:epsYA6		111:epsWA6		112:epsHA6		113:epsRotateA6
        ' 114:epsXA7		115:epsYA7		116:epsWA7		117:epsHA7		118:epsRotateA7
        ' 119:epsXA8		120:epsYA8		121:epsWA8		122:epsHA8		123:epsRotateA8


        '資材ナンバーから必要パラメータグループ取得
        For i = 0 To configCsv.Count - 1

            If configCsv(i)(1) = posterNo Then
                aiPath = GetAppPath() + "\" + configCsv(i)(3)
                'aifile open
                aiOpen(aiPath)

                'A1-A8、B1-B8、epsの計23セットを読込
                For ii = 0 To 23
                    '1セット読込
                    psdX = Double.Parse(configCsv(i)(4 + ii * 5))
                    psdY = Double.Parse(configCsv(i)(5 + ii * 5))
                    psdW = Double.Parse(configCsv(i)(6 + ii * 5))
                    psdH = Double.Parse(configCsv(i)(7 + ii * 5))
                    psdRotate = Double.Parse(configCsv(i)(8 + ii * 5))

                    'データ存在チェック
                    If psdX = 0 And psdY = 0 And psdW = 0 And psdH = 0 Then
                        '無ければ何もしない
                    Else
                        'PSDパス取得
                        If ii < 8 Then
                            psdPath = GetDirectoryName(txtTempA.Text) + "\poster\" + GetFileNameWithoutExtension(txtTempA.Text) + ".psd"
                            'PSD配置
                            addPlacedPsd(psdPath, psdX, psdY, psdW, psdH, psdRotate)
                        ElseIf ii < 16 Then
                            If txtTempB.Text <> "" Then
                                 psdPath = GetDirectoryName(txtTempB.Text) + "\poster\" + GetFileNameWithoutExtension(txtTempB.Text) + ".psd"
                                'PSD配置
                                addPlacedPsd(psdPath, psdX, psdY, psdW, psdH, psdRotate)
                            End If
                        Else
                            'psdPath = GetDirectoryName(txtTempB.Text) + "\poster\" + GetFileNameWithoutExtension(txtTempC.Text) + ".eps"
                            psdPath = txtTempC.Text
                            'eps配置
                            addPlacedPsd(psdPath, psdX, psdY, psdW, psdH, psdRotate)
                        End If
                    End If
                Next
                Exit For
            End If
        Next

        Dim printOK As Integer

        Select Case printerName
            Case "2号機"
                printOK += 1
                Dim appRef As Illustrator.Application = Nothing
                Dim docRef As Illustrator.Document = Nothing
                Try

                appRef = CreateObject("Illustrator.Application.29")
                docRef = appRef.ActiveDocument

                'aiファイル出力
                Dim jobOptionsRef As New Illustrator.PrintJobOptions
                Dim coordinateOptions As New Illustrator.PrintCoordinateOptions
                Dim printOptions As New Illustrator.PrintOptions
                printOptions.CoordinateOptions = coordinateOptions
                printOptions.CoordinateOptions.horizontalScale  = 106
                printOptions.CoordinateOptions.verticalScale = 106
                printOptions.JobOptions = jobOptionsRef
                coordinateOptions.Orientation = 4   'auto
                jobOptionsRef.Copies = printCount
                ' 選択プリンタ
                If RadioButton25.Checked Then
                    printOptions.PrinterName = RadioButton25.Text
                End If
                If RadioButton26.Checked Then
                    printOptions.PrinterName = RadioButton26.Text
                End If
                ' 印刷無し
                If printoutFlg Then
                    If CheckBox6.Checked = False Then
                        docRef.PrintOut(printOptions)
                    End If
                End If
                ' 保存
                If CheckBox5.Checked = True Then
                    Dim SavePath As String = GetDirectoryName(txtTempA.Text) + "\poster\" + posterNo + "_" + GetFileNameWithoutExtension(txtTempA.Text)
                    docRef.SaveAs(SavePath)
                End If
                docRef.Close(2)

                Catch ex As Exception
                            MessageBox.Show(ex.Message)
                Finally
                    If appRef IsNot Nothing Then ReleaseComObject(appRef) 'COMオブジェクトの開放
                    If docRef IsNot Nothing Then ReleaseComObject(docRef) 'COMオブジェクトの開放
                    GC.Collect()
                End Try

            Case "山梨"
                printOK += 1

                Dim appRef As Illustrator.Application = Nothing
                Dim docRef As Illustrator.Document = Nothing

                Try
                appRef = CreateObject("Illustrator.Application.29")
                docRef = appRef.ActiveDocument

                'aiファイル出力
                Dim jobOptionsRef As New Illustrator.PrintJobOptions
                Dim coordinateOptions As New Illustrator.PrintCoordinateOptions
                Dim printOptions As New Illustrator.PrintOptions
                printOptions.CoordinateOptions = coordinateOptions
                printOptions.JobOptions = jobOptionsRef
                coordinateOptions.Orientation = 4   'auto
                jobOptionsRef.Copies = printCount
                ' 選択プリンタ
                If RadioButton27.Checked Then
                    printOptions.PrinterName = RadioButton27.Text
                End If
                If RadioButton28.Checked Then
                    printOptions.PrinterName = RadioButton28.Text
                End If
                ' 印刷無し
                If printoutFlg Then
                    If CheckBox6.Checked = False Then
                        docRef.PrintOut(printOptions)
                    End If
                End If
                ' 保存
                If CheckBox5.Checked = True Then
                    Dim SavePath As String = GetDirectoryName(txtTempA.Text) + "\poster\" + posterNo + "_" + GetFileNameWithoutExtension(txtTempA.Text)
                    docRef.SaveAs(SavePath)
                End If
                docRef.Close(2)

                Catch ex As Exception
                                MessageBox.Show(ex.Message)
                    Finally
                        If appRef IsNot Nothing Then ReleaseComObject(appRef) 'COMオブジェクトの開放
                        If docRef IsNot Nothing Then ReleaseComObject(docRef) 'COMオブジェクトの開放
                        GC.Collect()
                    End Try
        End Select
    End Sub

    'aiファイルオープン
    Private Sub aiOpen(ByVal aiPath As String)
        Dim appRef As Illustrator.Application = Nothing
        Dim docRef As Illustrator.Document = Nothing

        Try
        appRef = CreateObject("Illustrator.Application.29")
        docRef = appRef.Open(aiPath)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            If appRef IsNot Nothing Then ReleaseComObject(appRef) 'COMオブジェクトの開放
            If docRef IsNot Nothing Then ReleaseComObject(docRef) 'COMオブジェクトの開放
            GC.Collect()
        End Try
    End Sub

    'aiファイルへPSDファイル配置
    Private Sub addPlacedPsd(ByVal psdPath As String,
                             ByVal psdX As Double, ByVal psdY As Double,
                             ByVal psdW As Double, ByVal psdH As Double, ByVal psdRotate As Double)
        
        Dim appRef As Illustrator.Application = Nothing
        Dim docRef As Illustrator.Document = Nothing
        Dim placedImags As Illustrator.PlacedItem

        Try
        appRef = CreateObject("Illustrator.Application.29")
        docRef = appRef.ActiveDocument

        placedImags = docRef.PlacedItems.Add()
        placedImags.File = psdPath
        placedImags.Width = psdW
        placedImags.Height = psdH
        placedImags.Translate(
            placedImags.Position(0) * -1,
            placedImags.Position(1) * -1)
        placedImags.Translate(psdX, psdY)
        placedImags.Rotate(psdRotate)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            If appRef IsNot Nothing Then ReleaseComObject(appRef) 'COMオブジェクトの開放
            If docRef IsNot Nothing Then ReleaseComObject(docRef) 'COMオブジェクトの開放
            GC.Collect()
        End Try
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
    End Sub
End Class


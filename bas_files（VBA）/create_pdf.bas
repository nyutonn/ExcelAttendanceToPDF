' エクセルファイルを別ファイルに書き込みPDF化して保存
Sub main()
    MsgBox Application.OperatingSystem
    メンバーリスト
    Dim wsAllMembers As Worksheet
    Dim ws2pdf As Worksheet

    ' メインの実装
    Call makeWorksheet(wsAllMembers, ws2pdf) ' Worksheetの作成
    Call extractTable(wsAllMembers, ws2pdf) '表を抽出
    Call writeOverview(wsAllMembers, ws2pdf) 'テキストを書き込む
    Call savePDF(ws2pdf)  'PDFに保存

End Sub

' エクセルファイルを別ファイルに書き込むだけ，PDF化しない
Sub makeTable()
    ' MsgBox Application.OperatingSystem
    ' メンバーリスト
    Dim wsAllMembers As Worksheet
    Dim ws2pdf As Worksheet

    ' メインの実装
    Call makeWorksheet(wsAllMembers, ws2pdf) ' Worksheetの作成
    Call extractTable(wsAllMembers, ws2pdf) '表を抽出
    Call writeOverview(wsAllMembers, ws2pdf) 'テキストを書き込む

End Sub

' Worksheetを作成する
Sub makeWorksheet(ByRef wsAllMembers As Worksheet, ByRef ws2pdf As Worksheet)
    ' wsAllMembersを作成
    Set wsAllMembers = ThisWorkbook.Sheets("メンバーリスト")
    ' 今日の日付 -> レッスン日
    ' 日付も外部からの入力を受け付けたい
    Dim today As Date
    Dim lessonDay As Date
    Dim sheetName As String
    today = Date '今日の日付
    lessonDay = wsAllMembers.Cells(4, 12).Value 'レッスン日
    sheetName = Format(lessonDay, "yyyy年mm月dd日")
    
    ' すでに同じ名前のシートが存在するか確認し、存在する場合は削除します
    On Error Resume Next 'エラーハンドリングスタート，エラーが起こっても停止しない
    Application.DisplayAlerts = False 'エラーメッセージの表示をOFFにする
    Set ws2pdf = ThisWorkbook.Sheets(sheetName)
    If Not ws2pdf Is Nothing Then
        ws2pdf.Delete
    End If
    Application.DisplayAlerts = True 'エラーメッセージの表示をONにする
    On Error GoTo 0 'エラーハンドリングストップ，以降は普通にエラーで停止する
    
    ' 新しいシートを作成し、名前を設定します
    Set ws2pdf = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))
    ws2pdf.Name = sheetName
    
    ' 結果をメッセージボックスで表示します
    MsgBox "シート '" & sheetName & "' が作成されました。"

End Sub

' メインの表を抽出して別ページに書き込み
Sub extractTable(ByVal wsAllMembers As Worksheet, ByVal ws2pdf As Worksheet)
    ' 今日の日付 -> レッスン日
    ' 日付も外部からの入力を受け付けたい
    Dim today As Date
    Dim lessonDay As Date
    Dim sheetName As String
    today = Date '今日の日付
    lessonDay = wsAllMembers.Cells(4, 12).Value 'レッスン日
    sheetName = Format(lessonDay, "yyyy年mm月dd日")

    ' 最終行を取得
    Dim lastRow As Long, i As Long
    lastRow = wsAllMembers.Cells(wsAllMembers.Rows.Count, 1).End(xlUp).row - 8
    Dim colIndex As Long
    colIndex = 2
    ' ヘッダーをコピー
    For j = 2 To 9
        If j <> 4 Then
            ws2pdf.Cells(2, colIndex).Value = wsAllMembers.Cells(1, j).Value
            colIndex = colIndex + 1
        End If
    Next j
    ws2pdf.Rows(2).Font.Bold = True

    Dim cb As CheckBox
    Dim cnt As Long
    cnt = 0
    Dim row As Long
    row = 3
    colIndex = 2
    ' 全体のチェックボックスを見て参加メンバーを取り出す
    For i = 2 To lastRow
        ' レッスン参加メンバー
        ' チェックボックスは2行目からなので-1する
        If wsAllMembers.CheckBoxes(i - 1).Value = 1 Then
            For j = 2 To 9
                If j <> 4 Then
                    ws2pdf.Cells(row, colIndex).Value = wsAllMembers.Cells(i, j).Value
                    colIndex = colIndex + 1
                End If
            Next j
            colIndex = 2
            row = row + 1
        End If
    Next i
    
    ' ws2pdf.Columns("A").AutoFit
    ws2pdf.Columns("B").AutoFit
    ws2pdf.Columns("C").AutoFit
    ws2pdf.Columns("D").AutoFit
    ws2pdf.Columns("E").AutoFit
    ws2pdf.Columns("F").AutoFit
    ws2pdf.Columns("G").AutoFit
    
    ' MsgBox myName 'Debug
    ' Call writeOverview(wsAllMembers, ws2pdf)
    
    ' MsgBox "end"
End Sub

' 料金や参加人数の書き込みを行うサブルーチン
Sub writeOverview(ByVal wsAllMembers As Worksheet, ByVal ws2pdf As Worksheet)
    ' 最終行の取得
    Dim lastRow As Long, i As Long
    lastRow = ws2pdf.Cells(ws2pdf.Rows.Count, 2).End(xlUp).row
    
    ' 会員と非会員の参加人数
    Dim cntMember As Long
    Dim cntNonMember As Long
    cntMember = 0
    cntNonMember = 0

    ' 会員と非会員の参加人数をカウントする
    For i = 3 To lastRow
        ' 会員No.が書いていなければ非会員
        If IsEmpty(ws2pdf.Cells(i, 6).Value) Then
            cntNonMember = cntNonMember + 1
        ' 会員番号に「休」と書いてあったら非会員扱い
        ElseIf InStr(ws2pdf.Cells(i, 6).Value, "休") > 0 Then
            cntNonMember = cntNonMember + 1
        ' 会員！
        Else
            cntMember = cntMember + 1
        End If
    Next i

    ' 全体の料金を求める
    Dim sumMoney As Long
    sumMoney = cntNonMember * 1100
    ' 今日の日付
    Dim today As Date
    today = Date
    Dim todayString As String
    todayString = Format(today, "yyyy年mm月dd日")
    ' レッスン日の日付
    Dim lessonDay As Date
    lessonDay = wsAllMembers.Cells(4, 12).Value 'レッスン日
    Dim lessonDayString As String
    lessonDayString = Format(lessonDay, "yyyy年mm月dd日")
    ' 全体の参加人数
    Dim players As Long
    players = cntMember + cntNonMember
    ' MsgBox cntMember

    ' ここからセルへの書き込み
    ' 罫線を引く
    With Range(Cells(2, 2), Cells(lastRow, 8))
        '⑤範囲内の縦線
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        '⑥範囲内の横線
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        '周り
        .BorderAround LineStyle:=xlContinuous
        '外周と一行目の線を太くする
        '①上部
        .Borders(xlEdgeTop).Weight = xlMedium
        '②左
        .Borders(xlEdgeLeft).Weight = xlMedium
        '③下部
        .Borders(xlEdgeBottom).Weight = xlMedium
        '④右
        .Borders(xlEdgeRight).Weight = xlMedium
    End With
    '２行目
    Range(Cells(3, 2), Cells(lastRow, 8)).Borders(xlEdgeTop).Weight = xlMedium

    ' セルに文字を記入
    Dim titleString As String
    ' titleString = "本日（" + todayString + "）の利用者及び精算金額"
    titleString = "本日（" + lessonDayString + "）の利用者及び精算金額"
    ws2pdf.Cells(lastRow + 3, 2).Value = titleString
    ws2pdf.Cells(lastRow + 4, 3).Value = "非会員（1100円）"
    ws2pdf.Cells(lastRow + 5, 3).Value = "会員 サークル利用"
    ws2pdf.Cells(lastRow + 6, 3).Value = "利用総数"
    ws2pdf.Cells(lastRow + 4, 4).Value = Str(cntNonMember) + "名"
    ws2pdf.Cells(lastRow + 5, 4).Value = Str(cntMember) + "名"
    ws2pdf.Cells(lastRow + 6, 4).Value = Str(players) + "名"
    
    ws2pdf.Cells(lastRow + 6, 5).Value = "本日の精算金額"
    ws2pdf.Cells(lastRow + 6, 6).Value = Str(sumMoney) + "円"

    ' 代表者指名の書き込み
    Dim myName As String
    myName = wsAllMembers.Cells(3, 12).Value
    ws2pdf.Cells(lastRow + 8, 2).Value = "代表者氏名：" + myName

    ' 太文字
    ws2pdf.Cells(lastRow + 3, 2).Font.Bold = True
    ws2pdf.Cells(lastRow + 6, 5).Font.Bold = True
    ws2pdf.Cells(lastRow + 6, 6).Font.Bold = True
    ws2pdf.Cells(lastRow + 8, 2).Font.Bold = True
    ' 中央寄せ
    ws2pdf.Cells(lastRow + 4, 4).HorizontalAlignment = xlCenter
    ws2pdf.Cells(lastRow + 5, 4).HorizontalAlignment = xlCenter
    ws2pdf.Cells(lastRow + 6, 4).HorizontalAlignment = xlCenter
    ws2pdf.Cells(lastRow + 6, 6).HorizontalAlignment = xlCenter
    ws2pdf.Columns(2).HorizontalAlignment = xlCenter
    ws2pdf.Columns(3).HorizontalAlignment = xlCenter
    ws2pdf.Columns(4).HorizontalAlignment = xlCenter
    ws2pdf.Columns(5).HorizontalAlignment = xlCenter
    ws2pdf.Columns(6).HorizontalAlignment = xlCenter
    ws2pdf.Columns(7).HorizontalAlignment = xlCenter
    ws2pdf.Columns(8).HorizontalAlignment = xlCenter
    ' 右寄せ
    ws2pdf.Cells(lastRow + 4, 3).HorizontalAlignment = xlRight
    ws2pdf.Cells(lastRow + 5, 3).HorizontalAlignment = xlRight
    ws2pdf.Cells(lastRow + 6, 3).HorizontalAlignment = xlRight
    ws2pdf.Cells(lastRow + 6, 5).HorizontalAlignment = xlRight
    ' 左寄せ
    ws2pdf.Cells(lastRow + 3, 2).HorizontalAlignment = xlLeft
    ws2pdf.Cells(lastRow + 8, 2).HorizontalAlignment = xlLeft

    ' 文字の大きさ変更
    ws2pdf.Cells(lastRow + 3, 2).Font.Size = 15
    ws2pdf.Cells(lastRow + 8, 2).Font.Size = 15
    ws2pdf.Cells(lastRow + 6, 5).Font.Size = 13
    ws2pdf.Cells(lastRow + 6, 6).Font.Size = 13
    ws2pdf.Cells(lastRow + 4, 3).Font.Size = 12
    ws2pdf.Cells(lastRow + 5, 3).Font.Size = 12
    ws2pdf.Cells(lastRow + 6, 3).Font.Size = 12
    ws2pdf.Cells(lastRow + 4, 4).Font.Size = 13
    ws2pdf.Cells(lastRow + 5, 4).Font.Size = 13
    ws2pdf.Cells(lastRow + 6, 4).Font.Size = 13
    
    ' ボタン作成
    Call makeSavePDFButton(ws2pdf)

End Sub

' 現在の状態のPDFを保存
Sub savePDF(ByVal ws2pdf As Worksheet)
    ' ページ設定を調整して1ページに収める
    With ws2pdf.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintArea = Range(Cells(2, 2), Cells(lastRow + 8, 8)).Address
        .CenterHorizontally = True ' 水平方向の中央揃え
    End With

    ' PDFに保存
    Dim folderPath As String
    Dim filePath As String
    ' ファイルパス（これは外部から指定できるようにしたい）
    ' 可能なら文字列じゃなくてフォルダを選ぶ形式にしたい
    ' フォルダ選択ダイアログを表示

    On Error Resume Next 'エラーハンドリングスタート，エラーが起こっても停止しない
    Application.DisplayAlerts = False 'エラーメッセージの表示をOFFにする
    folderPath = MacFolderPicker()
    If folderPath = "" Then
        MsgBox "フォルダが選択されませんでした．" + vbLf + "PDF化を中止します．"  'vbLf は改行を表す
        Exit Sub 'サブルーチンをこの時点で止める
    End If
    Application.DisplayAlerts = True 'エラーメッセージの表示をONにする
    On Error GoTo 0 'エラーハンドリングストップ，以降は普通にエラーで停止する

    ' レッスン日
    lessonDayString = ws2pdf.Name
    
    ' PDFファイルのパスを設定
    filePath = folderPath & lessonDayString & "_参加者.pdf"

    ' MsgBox ThisWorkbook.Path + "/提出書類"
    ' Mkdir ThisWorkbook.Path + "/提出書類" ' 新規フォルダ作成
    ' filePath = ThisWorkbook.Path + "/提出書類/" + lessonDayString + "_参加者.pdf"
    ws2pdf.ExportAsFixedFormat Type:=xlTypePDF, FileName:=filePath, Quality:=xlQualityStandard

    ' メッセージを表示
    MsgBox "PDFが保存されました: " & filePath

End Sub

' Macのフォルダ選択を行う
Function MacFolderPicker() As String
    Dim folderPath As String
    Dim script As String
    Dim folderDialog As String
    
    ' AppleScriptを使用してフォルダ選択ダイアログを表示
    script = "set folderPath to POSIX path of (choose folder with prompt ""フォルダを選択してください:"") as string"
    folderDialog = MacScript(script)
    If folderDialog <> "" Then
        MacFolderPicker = folderDialog
    Else
        MacFolderPicker = ""
    End If
End Function

' ボタンを作成
Sub makeSavePDFButton(ws2pdf As Worksheet)
    Dim btnSavePDF As Shape
    ' ボタンの位置とサイズを設定
    Set btnSavePDF = ws2pdf.Shapes.AddShape(msoShapeRectangle, 650, 100, 150, 80)

    ' ボタンのテキストを設定
    btnSavePDF.TextFrame.Characters.Text = "現在の表を" + vbLf + "PDFに保存"
    ' テキストを中央寄せに設定
    btnSavePDF.TextFrame.HorizontalAlignment = xlHAlignCenter
    ' ボタンの背景色を設定
    btnSavePDF.Fill.ForeColor.RGB = RGB(235, 0, 0)  ' 赤色
    ' ボタンのフォント色を設定
    btnSavePDF.TextFrame.Characters.Font.Color = RGB(255, 255, 255)  ' 白色
    ' ボタンのフォントを太字に設定
    btnSavePDF.TextFrame.Characters.Font.Bold = True
    ' ボタンのフォントサイズを設定
    btnSavePDF.TextFrame.Characters.Font.Size = 20
    ' ボタンがクリックされたときに実行されるマクロを設定
    btnSavePDF.OnAction = "savePDFbyButton"

End Sub

' 後からボタンを押してPDF化する際の関数
Sub savePDFbyButton()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim filePath As String

    ' 押されたボタンに付いている名前を取得
    sheetName = ActiveSheet.Name
    ' Worksheetを紐づけ
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' PDF保存用のフォルダパスを取得
    On Error Resume Next 'エラーハンドリングスタート，エラーが起こっても停止しない
    Application.DisplayAlerts = False 'エラーメッセージの表示をOFFにする
    folderPath = MacFolderPicker()
    If folderPath = "" Then
        MsgBox "フォルダが選択されませんでした．" + vbLf + "PDF化を中止します．"  'vbLf は改行を表す
        Exit Sub 'サブルーチンをこの時点で止める
    End If
    Application.DisplayAlerts = True 'エラーメッセージの表示をONにする
    On Error GoTo 0 'エラーハンドリングストップ，以降は普通にエラーで停止する

    ' PDFファイルのパスを設定
    filePath = folderPath & sheetName & "_参加者.pdf"
    
    ' シートをPDFに保存
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, Quality:=xlQualityStandard
    
    ' 結果を通知
    MsgBox "シート '" & sheetName & "' をPDFに保存しました。"

End Sub

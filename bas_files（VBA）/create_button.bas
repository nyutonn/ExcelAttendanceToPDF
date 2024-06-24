' Excel上ボタンを出現させるサブルーチン
Sub createButtons()
    Dim ws As Worksheet
    Dim btnSavePDF As Shape
    Dim btnMakeTable As Shape
    Set ws = ThisWorkbook.Sheets("メンバーリスト")
    
    ' PDF保存ボタン
    ' ボタンの位置とサイズを設定
    Set btnSavePDF = ws.Shapes.AddShape(msoShapeRectangle, 713, 100, 150, 40)
    ' ボタンのテキストを設定
    btnSavePDF.TextFrame.Characters.Text = "表作成&PDF保存"
    ' テキストを中央寄せに設定
    btnSavePDF.TextFrame.HorizontalAlignment = xlHAlignCenter
    ' ボタンの背景色を設定
    btnSavePDF.Fill.ForeColor.RGB = RGB(235, 0, 0)  ' 赤色
    ' ボタンのフォント色を設定
    btnSavePDF.TextFrame.Characters.Font.Color = RGB(255, 255, 255)  ' 白色
    ' ボタンのフォントを太字に設定
    btnSavePDF.TextFrame.Characters.Font.Bold = True
    ' ボタンのフォントサイズを設定
    btnSavePDF.TextFrame.Characters.Font.Size = 18
    ' ボタンがクリックされたときに実行されるマクロを設定
    btnSavePDF.OnAction = "main"

    ' 他シートに表作成ボタン
    ' ボタンの位置とサイズを設定
    Set btnMakeTable = ws.Shapes.AddShape(msoShapeRectangle, 713, 150, 150, 40)
    ' ボタンのテキストを設定
    btnMakeTable.TextFrame.Characters.Text = "表を作成"
    ' テキストを中央寄せに設定
    btnMakeTable.TextFrame.HorizontalAlignment = xlHAlignCenter
    ' ボタンの背景色を設定
    btnMakeTable.Fill.ForeColor.RGB = RGB(0, 180, 0)  ' 緑色
    ' ボタンのフォント色を設定
    btnMakeTable.TextFrame.Characters.Font.Color = RGB(255, 255, 255)  ' 白色
    ' ボタンのフォントを太字に設定
    btnMakeTable.TextFrame.Characters.Font.Bold = True
    ' ボタンのフォントサイズを設定
    btnMakeTable.TextFrame.Characters.Font.Size = 20
    ' ボタンがクリックされたときに実行されるマクロを設定
    btnMakeTable.OnAction = "makeTable"

    ' 時を進めるボタン
    ' ボタンの位置とサイズを設定
    Set btnMakeTable = ws.Shapes.AddShape(msoShapeRectangle, 1050, 110, 70, 20)
    ' ボタンのテキストを設定
    btnMakeTable.TextFrame.Characters.Text = "時を進める"
    ' テキストを中央寄せに設定
    btnMakeTable.TextFrame.HorizontalAlignment = xlHAlignCenter
    ' ボタンの背景色を設定
    btnMakeTable.Fill.ForeColor.RGB = RGB(245, 245, 245)  ' 白
    ' ボタンのフォント色を設定
    btnMakeTable.TextFrame.Characters.Font.Color = RGB(0, 0, 0)  ' 黒
    ' ボタンのフォントを太字に設定
    btnMakeTable.TextFrame.Characters.Font.Bold = True
    ' ボタンのフォントサイズを設定
    btnMakeTable.TextFrame.Characters.Font.Size = 10
    ' ボタンがクリックされたときに実行されるマクロを設定
    btnMakeTable.OnAction = "advanceTime"

    ' 時を戻すボタン
    ' ボタンの位置とサイズを設定
    Set btnMakeTable = ws.Shapes.AddShape(msoShapeRectangle, 1050, 160, 70, 20)
    ' ボタンのテキストを設定
    btnMakeTable.TextFrame.Characters.Text = "時を戻す"
    ' テキストを中央寄せに設定
    btnMakeTable.TextFrame.HorizontalAlignment = xlHAlignCenter
    ' ボタンの背景色を設定
    btnMakeTable.Fill.ForeColor.RGB = RGB(200, 200, 200)  ' 白
    ' ボタンのフォント色を設定
    btnMakeTable.TextFrame.Characters.Font.Color = RGB(0, 0, 0)  ' 黒
    ' ボタンのフォントを太字に設定
    btnMakeTable.TextFrame.Characters.Font.Bold = True
    ' ボタンのフォントサイズを設定
    btnMakeTable.TextFrame.Characters.Font.Size = 10
    ' ボタンがクリックされたときに実行されるマクロを設定
    btnMakeTable.OnAction = "backTime"

    ' チェックマークを外すボタン
    ' ボタンの位置とサイズを設定
    Set btnMakeTable = ws.Shapes.AddShape(msoShapeRectangle, 900, 130, 100, 25)
    ' ボタンのテキストを設定
    btnMakeTable.TextFrame.Characters.Text = "チェック全解除"
    ' テキストを中央寄せに設定
    btnMakeTable.TextFrame.HorizontalAlignment = xlHAlignCenter
    ' ボタンの背景色を設定
    btnMakeTable.Fill.ForeColor.RGB = RGB(200, 200, 255)  ' 水色
    ' ボタンのフォント色を設定
    btnMakeTable.TextFrame.Characters.Font.Color = RGB(0, 0, 0)  ' 黒
    ' ボタンのフォントを太字に設定
    btnMakeTable.TextFrame.Characters.Font.Bold = True
    ' ボタンのフォントサイズを設定
    btnMakeTable.TextFrame.Characters.Font.Size = 12
    ' ボタンがクリックされたときに実行されるマクロを設定
    btnMakeTable.OnAction = "ClearCheckboxes"

End Sub





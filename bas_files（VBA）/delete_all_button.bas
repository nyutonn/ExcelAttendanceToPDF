' エクセル上の全てのボタンを削除する（デバッグ用）
Sub DeleteAllButtonsInWorkbook()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim oleObj As OLEObject
    
    ' すべてのシートをループ
    For Each ws In ThisWorkbook.Sheets
        ' 各シート上のすべてのShapeをループ
        For Each shp In ws.Shapes
            On Error Resume Next
            ' フォームコントロールのボタンかどうかをチェック
            If shp.FormControlType = xlButtonControl Then
                shp.Delete
            ElseIf shp.Type = msoOLEControlObject Then
                ' ActiveXボタンかどうかをチェック
                If TypeName(shp.OLEFormat.Object.Object) = "CommandButton" Then
                    shp.Delete
                End If
            End If
            On Error GoTo 0
        Next shp
        
        ' 各シート上のすべてのOLEObjectをループ
        For Each oleObj In ws.OLEObjects
            On Error Resume Next
            ' ActiveXコントロールのボタンかどうかをチェック
            If TypeName(oleObj.Object) = "CommandButton" Then
                oleObj.Delete
            End If
            On Error GoTo 0
        Next oleObj
    Next ws

    MsgBox "すべてのシートのすべてのボタンが削除されました。"
End Sub


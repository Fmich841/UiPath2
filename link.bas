Sub wslinks()
    Dim ws As Worksheet
    Dim cell As Range
    Dim targetSheetName As String
    Dim targetSheet As Worksheet
    Dim sourceWorkbook As Workbook
    Dim currentDir As String

    ' マクロが実行されるファイルを設定
    Set sourceWorkbook = ThisWorkbook
    currentDir = sourceWorkbook.Path ' フルパスを文字列として取得

    ' 操作対象のシートを指定（現在のブックのシート）
    Set ws = sourceWorkbook.Sheets("EmployeeInfo")

    ' A列のすべてのセルを処理
    For Each cell In ws.Range("A1:A" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)
        If cell.Value <> "" Then
            targetSheetName = cell.Value

            ' シートが存在するか確認
            On Error Resume Next
            Set targetSheet = sourceWorkbook.Sheets(targetSheetName)
            On Error GoTo 0

            If Not targetSheet Is Nothing Then
                ' ハイパーリンクを追加
                ws.Hyperlinks.Add Anchor:=cell, Address:="", SubAddress:="'" & targetSheetName & "'!A1", TextToDisplay:=cell.Value
            End If
        End If
    Next cell

    ' 処理が終わった後の必要な処理（保存や閉じるなど）を追加することができます
End Sub

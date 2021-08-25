Attribute VB_Name = "ModEventAutoInput"
Option Explicit

'単語登録しながら単語の自動入力を行うためのイベントプロシージャ
'ModFileと一緒に使うこと

Const TextFileName$ = "単語登録.txt" '←←←←←←←←←←←←←←←←←←←←←←←

Private Function 入力セル範囲取得()
    
    Set 入力セル範囲取得 = Sheet1.Range("B3:B50") '←←←←←←←←←←←←←←←←←←←←←←←

End Function

Private Function 出力セル範囲取得()
    
    Set 出力セル範囲取得 = Sheet1.Range("C3:C50") '←←←←←←←←←←←←←←←←←←←←←←←

End Function

Sub セルの値変更時に登録単語出力と単語登録(ByVal Target As Range) 'Worksheet_Changeプロシージャで実行
'セルの値変更時に登録単語出力と単語登録

    If VarType(Target.Value) >= vbArray Then
        '変更セルが複数の場合は1番目のセルだけ対象とする。
        Set Target = Target(1)
    End If
    
    If Not Intersect(入力セル範囲取得, Target) Is Nothing Then
        '変更したセルが入力セル範囲の場合
        Application.EnableEvents = False '関連語句出力後のイベント発生を停止
        Call 入力値から関連語句出力(Target)
        Application.EnableEvents = True
    ElseIf Not Intersect(出力セル範囲取得, Target) Is Nothing Then
        '変更したセルが出力セル範囲の場合
        Call 関連語句登録(Target.Offset(0, -1).Value, Target.Value)
    End If
    
End Sub

Private Sub 入力値から関連語句出力(TargetCell As Range)
    
    Dim InputValue$
    InputValue = TargetCell.Value
        
     '入力範囲と出力範囲からオフセット量計算
    Dim InputStartCell As Range, OutputStartCell As Range
    Dim OffsetRow&, OffsetCol&
    Set InputStartCell = 入力セル範囲取得(1)
    Set OutputStartCell = 出力セル範囲取得(1)
    OffsetRow = OutputStartCell.Row - InputStartCell.Row
    OffsetCol = OutputStartCell.Column - InputStartCell.Column
    
    If InputValue = "" Then
        TargetCell.Offset(OffsetRow, OffsetCol) = ""
        Exit Sub
    End If
    
    '登録済みの関連語句をテキストファイルから読み込み
    Dim TextList
    TextList = InputText(ThisWorkbook.Path & "\" & TextFileName, Chr(9))
    
    Dim KanrenStr$ '関連語句
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    

    '関連語句出力
    If UBound(TextList, 1) = 1 Then
        '登録した単語が一つもない場合は何もしない
    Else
        For I = 2 To UBound(TextList, 1)
            If InputValue = TextList(I, 1) Then
                KanrenStr = TextList(I, 2)
                Exit For
            End If
        Next I
        
        TargetCell.Offset(OffsetRow, OffsetCol).Value = KanrenStr
        
    End If
    
End Sub
Private Sub 関連語句登録(InputStr$, KanrenStr$)
    
    '空白は登録しない
    If KanrenStr = "" Or InputStr = "" Then Exit Sub
    
    Dim TextList
    TextList = InputText(ThisWorkbook.Path & "\" & TextFileName, Chr(9))
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim OutputList
    Dim TorokuzumiNaraTrue As Boolean
    Dim TorokuZumiNum%
    If UBound(TextList, 1) = 1 Then
        '最初から登録
        ReDim OutputList(1 To 2, 1 To 2)
        OutputList(1, 1) = "入力"
        OutputList(1, 2) = "関連語句"
        OutputList(2, 1) = InputStr
        OutputList(2, 2) = KanrenStr
        
    Else
        TorokuzumiNaraTrue = False
        For I = 2 To UBound(TextList, 1)
            '既に登録済みであるか確認
            If InputStr = TextList(I, 1) Then
                TorokuzumiNaraTrue = True
                TorokuZumiNum = I
                Exit For
            End If
        Next I
        
        If TorokuzumiNaraTrue Then
            OutputList = TextList
            
            If OutputList(TorokuZumiNum, 2) = KanrenStr Then
                '登録済みが同じなので何もしない
                Exit Sub
            End If
            
            OutputList(TorokuZumiNum, 2) = KanrenStr '登録情報変更
        Else
            N = UBound(TextList, 1)
            ReDim OutputList(1 To N + 1, 1 To 2)
            For J = 1 To N
                OutputList(J, 1) = TextList(J, 1)
                OutputList(J, 2) = TextList(J, 2)
            Next J
            
            OutputList(N + 1, 1) = InputStr
            OutputList(N + 1, 2) = KanrenStr
        End If
                
    End If
    
    '編集後のテキストデータを出力
    Call OutputText(ThisWorkbook.Path, TextFileName, OutputList, Chr(9))
    
End Sub


Attribute VB_Name = "ModEventAutoInput"
Option Explicit

'単語登録しながら単語の自動入力を行うためのイベントプロシージャ
'イベントプロシージャはコピーして使う

Const TextFileName$ = "単語登録.txt" '←←←←←←←←←←←←←←←←←←←←←←←

Sub セルの値変更時に登録単語出力と単語登録() 'Worksheet_Changeプロシージャに中身コピー
'セルの値変更時に登録単語出力と単語登録

    Dim InputCell As Range, OutputCell As Range '入力出力セル
    Set InputCell = 入力セル範囲取得
    Set OutputCell = 出力セル範囲取得
    
    If VarType(Target.Value) >= vbArray Then Exit Sub
    
    If Not Intersect(InputCell, Target) Is Nothing Then
        Application.EnableEvents = False
        Call 入力値から関連語句出力(Target, Target.Offset(0, 1)) '←←←←←←←←←←←←←←←←←←←←←←←
        Application.EnableEvents = True
    ElseIf Not Intersect(OutputCell, Target) Is Nothing Then
        Call 関連語句登録(Target.Offset(0, -1).Value, Target.Value)
    End If

End Sub

Function 入力セル範囲取得()
    
    Set 入力セル範囲取得 = Range("B3:B1000") '←←←←←←←←←←←←←←←←←←←←←←←

End Function

Function 出力セル範囲取得()
    
    Set 出力セル範囲取得 = Range("C3:C1000") '←←←←←←←←←←←←←←←←←←←←←←←

End Function

Sub 入力値から関連語句出力(TargetCell As Range)
    
    Dim InputValue$
    InputValue = TargetCell.Value
        
    If InputValue = "" Then
        TargetCell.Offset(0, 1) = ""
        Exit Sub
    End If
        
    Dim TextList
    TextList = InputText(ThisWorkbook.Path & "\" & TextFileName, ",")
    
    Dim KanrenStr$ '関連語句
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    
    If UBound(TextList, 1) = 1 Then
        '何もしない
    Else
        For I = 2 To UBound(TextList, 1)
            If InputValue = TextList(I, 1) Then
                KanrenStr = TextList(I, 2)
                Exit For
            End If
        Next I
        
        TargetCell.Offset(0, 1).Value = KanrenStr
        
    End If
    
End Sub

Sub 関連語句登録(InputStr$, KanrenStr$)

    If KanrenStr = "" Or InputStr = "" Then Exit Sub
    
    Dim TextList
    TextList = InputText(ThisWorkbook.Path & "\" & TextFileName, ",")
    
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
    
    Call OutputText(ThisWorkbook.Path, TextFileName, OutputList, ",")
    
End Sub

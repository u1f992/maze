Attribute VB_Name = "Solve"
Option Explicit

'四方を探索
'探索するセルのリストを作成
Function SearchSet(ByVal Target As Range) As Boolean
    
    SearchSet = False
    
    '左右上下の進行可能セルをマーキング
    
    Dim Directions(NORTH To WEST) As Range
    Set Directions(NORTH) = Target.Cells(0, 1)
    Set Directions(EAST) = Target.Cells(1, 2)
    Set Directions(SOUTH) = Target.Cells(2, 1)
    Set Directions(WEST) = Target.Cells(1, 0)
    
    Dim i As Long
    
    For i = NORTH To WEST
        If Directions(i).Interior.Color = GOAL.Interior.Color Then 'ゴールにたどりついた場合
            AllCellsChecked
            SearchSet = True
            Exit Function
        End If
        If IsAvailable(Directions(i)) Then '壁か探索済みかスタートではない場合
            
            Directions(i).Interior.Color = SEARCHING
            Directions(i).Value = Target.Value + 1
            
            '添え字は1から始まる、0には空白のデータ
            '要素数が0になるとまずいため
            
            ReDim Preserve Candidates(0 To UBound(Candidates) + 1)
            Set Candidates(UBound(Candidates)) = Directions(i)
            
        End If
    Next i
    
    If Target.Interior.Color <> START.Interior.Color Then
        Target.Interior.Color = SEARCHED
        Candidates = ArrDelete(Candidates, Target)
    End If
    
End Function

'次に探索するセルを選定
Function SearchGet() As Range
    
    Dim i As Long
    Dim minimum As Long
    minimum = 2147483647
    
    Dim vTemp As Long
    Dim hTemp As Long
    
    For i = LBound(Candidates) + 1 To UBound(Candidates)
        'If minimum >= Candidates(i).Value + CLng(Sqr((Abs(Candidates(i).Row - Goal.Row)) ^ 2 + (Abs(Candidates(i).Row - Goal.Row)) ^ 2)) Then 'スタートからの距離+ゴールまでの距離(直線距離)
        If minimum >= Candidates(i).Value + Abs(Candidates(i).Row - GOAL.Row) + Abs(Candidates(i).Row - GOAL.Row) Then 'スタートからの距離+ゴールまでの距離(辺の合計)
            minimum = Candidates(i).Value + CLng(Sqr((Abs(Candidates(i).Row - GOAL.Row)) ^ 2 + (Abs(Candidates(i).Row - GOAL.Row)) ^ 2))
            
            Set SearchGet = Candidates(i)
            
        End If
    Next i
    
End Function

'スタートまで帰る
Function Back2Start(ByVal Target As Range) As Range
    
    Dim Directions(NORTH To WEST) As Range
    Set Directions(NORTH) = Target.Cells(0, 1)
    Set Directions(EAST) = Target.Cells(1, 2)
    Set Directions(SOUTH) = Target.Cells(2, 1)
    Set Directions(WEST) = Target.Cells(1, 0)
    
    Dim i As Long
    Dim val(NORTH To WEST) As Long
    
    For i = NORTH To WEST '4方向にスタートがあれば終了
        If Directions(i).Interior.Color = START.Interior.Color Then
            Solved = True
            Exit Function
        End If
        
        val(i) = CastValue(Directions(i).Value) '各セルの値(スタートまでの距離)を格納
        
    Next i
    
    Dim s As Long
    s = Smallest(val) '各セルの値が最も小さいものを選択
    Directions(s).Interior.Color = ROUTE
    
    Set Back2Start = Directions(s)
    
End Function



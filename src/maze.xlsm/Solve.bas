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
            
            TempSearch = ArrAdd(TempSearch, Directions(i))
            
        End If
    Next i
    
    '探索済みのセルをリストから削除
    If Target.Interior.Color <> START.Interior.Color Then
        Target.Interior.Color = SEARCHED
        TempSearch = ArrDelete(TempSearch, Target)
    End If
    
End Function

'次に探索するセルを選定
Function SearchGet() As Range
    
    Dim i As Long
    Dim Minimum As Long
    Minimum = 2147483647
    Dim Maximum As Long
    Maximum = 0
    
    Dim Minimums() As Range
    ReDim Minimums(0 To 0)
    
    Dim Evaluation As Long
    
    Dim vTemp As Long
    Dim hTemp As Long
    
    For i = LBound(TempSearch) + 1 To UBound(TempSearch)
    
        Evaluation = TempSearch(i).Value + Abs(TempSearch(i).Row - GOAL.Row) + Abs(TempSearch(i).Row - GOAL.Row)  'スタートからの距離+ゴールまでの距離(辺の合計)
        'TempSearch(i).Value + CLng(Sqr((Abs(TempSearch(i).Row - GOAL.Row)) ^ 2 + (Abs(TempSearch(i).Row - GOAL.Row)) ^ 2)) 'スタートからの距離+ゴールまでの距離(直線距離)
        
        If Minimum >= Evaluation Then
            
            If Minimum > Evaluation Then '最小を更新する場合
                Minimum = Evaluation
                ReDim Minimums(0 To 0)
                Minimums = ArrAdd(Minimums, TempSearch(i))
            Else '最小に並ぶものを見つけた場合
                Minimums = ArrAdd(Minimums, TempSearch(i))
            End If
            
        End If
    Next i
    
    For i = LBound(Minimums) + 1 To UBound(Minimums)
        
        Evaluation = Minimums(i).Value
        
        If Maximum < Evaluation Then 'スタートからの位置が一番遠いものを探す(等しい場合は最初に見つけたものになる)
            Maximum = Evaluation
            
            Set SearchGet = Minimums(i)
            
        End If
        
    Next i
    
End Function

Function SolveSet(ByVal Target As Range) As Boolean
    
    SolveSet = False
    
    '左右上下のセルから、各セルの値(スタートまでの距離)が小さいものを選択
    
    Dim Directions(NORTH To WEST) As Range
    Set Directions(NORTH) = Target.Cells(0, 1)
    Set Directions(EAST) = Target.Cells(1, 2)
    Set Directions(SOUTH) = Target.Cells(2, 1)
    Set Directions(WEST) = Target.Cells(1, 0)
    
    Dim i As Long
    Dim val(NORTH To WEST) As Long
    
    For i = NORTH To WEST '4方向にスタートがあれば終了
        If Directions(i).Interior.Color = START.Interior.Color Then
            SolveSet = True
            Exit Function
        End If
        
        val(i) = CastValue(Directions(i).Value) '各セルの値(スタートまでの距離)を格納
        
    Next i
    
    Dim s As Long
    s = Smallest(val) '各セルの値が最も小さいものを選択
    Set TempRoute = Directions(s)
    
End Function

Function SolveGet() As Range
    
    TempRoute.Interior.Color = ROUTE
    Set SolveGet = TempRoute
    
End Function

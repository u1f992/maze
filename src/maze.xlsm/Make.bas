Attribute VB_Name = "Make"
Option Explicit

Function MakeMaze()

    '壁候補を管理
    Dim Knots() As Range
    ReDim Knots(0 To 0)
    '新規壁を管理
    Dim TempWalls() As Range
    ReDim TempWalls(0 To 0)
    
    '迷路の初期化
    RangeMaze.Rows.RowHeight = 5 * 0.75
    RangeMaze.Columns.ColumnWidth = 5 * 0.07
    
    Cells(SIZE + 1, SIZE + 1).Select
    
    '外周を壁(既存壁)に
    RangeMaze.Interior.Color = BUILT
    Range(RangeMaze.Cells(2, 2), RangeMaze.Cells(RangeMaze.Rows.Count - 1, RangeMaze.Columns.Count - 1)).ClearFormats
    
    Dim i As Long
    Dim j As Long
    
    '壁候補(x,yとも奇数かつ枠ではない)のリストを作成
    For i = 3 To SIZE - 2 Step 2
        For j = 3 To SIZE - 2 Step 2
            
            ReDim Preserve Knots(0 To UBound(Knots) + 1)
            Set Knots(UBound(Knots)) = Cells(i, j)
            
        Next j
    Next i
    
    Dim prev As Long
    prev = UBound(Knots) - 1
    Application.StatusBar = "迷路を生成しています... - 0%"
        
    Dim Selected As Range
    
    Dim temp As Long
    Dim Direction As Long
    Dim CantEnter(NORTH To WEST) As Boolean '進めない場合にtrue
    
    Dim vTemp As Long
    Dim hTemp As Long
    
    Dim Target As Range '進行先
    
    Do While True
        
        i = 0
        
        '壁候補からランダムに取り出す
        'Int((最大値 - 最小値 + 1) * Rnd + 最小値)
        Randomize
        temp = Int((UBound(Knots) - (LBound(Knots) + 1) + 1) * Rnd + (LBound(Knots) + 1))
        
        Set Selected = Knots(temp)
        
        vTemp = Selected.Row
        hTemp = Selected.Column
        
        '◎壁伸ばし処理
        '壁候補を新規壁に
        Selected.Interior.Color = BUILDING
        ReDim Preserve TempWalls(0 To UBound(TempWalls) + 1)
        Set TempWalls(UBound(TempWalls)) = Selected
        
        '進行方向をランダムに決定
        BoolReset CantEnter
        
        Do While True
        
            Randomize
            Direction = Int((WEST - NORTH + 1) * Rnd + NORTH)
            
            '進行先の状況を取得
            Select Case Direction
                Case NORTH
                    Set Target = Cells(vTemp - 2, hTemp)
                Case EAST
                    Set Target = Cells(vTemp, hTemp + 2)
                Case SOUTH
                    Set Target = Cells(vTemp + 2, hTemp)
                Case WEST
                    Set Target = Cells(vTemp, hTemp - 2)
            End Select
            
            If Target.Interior.Color = BUILT Then '進行先が既存壁の場合
                '進行先を新規壁にして確定
                Range(Cells(vTemp, hTemp), Target).Interior.Color = BUILDING
                
                ReDim Preserve TempWalls(0 To UBound(TempWalls) + 1)
                Set TempWalls(UBound(TempWalls)) = Range(Cells(vTemp, hTemp), Target)
                
                Exit Do
                
            ElseIf Target.Interior.Color = BUILDING Then '新規壁の場合
                '進行不可フラグを立てる
                Select Case Direction
                    Case NORTH
                        CantEnter(NORTH) = True
                    Case EAST
                        CantEnter(EAST) = True
                    Case SOUTH
                        CantEnter(SOUTH) = True
                    Case WEST
                        CantEnter(WEST) = True
                End Select
            
            Else '通路の場合
                '進行先を新規壁にする
                Range(Cells(vTemp, hTemp), Target).Interior.Color = BUILDING
                
                ReDim Preserve TempWalls(0 To UBound(TempWalls) + 1)
                Set TempWalls(UBound(TempWalls)) = Range(Cells(vTemp, hTemp), Target)
                
                '使用された壁候補はリストから削除
                If UBound(Knots) <> 1 Then
                    Knots = ArrDelete(Knots, Target)
                    Application.StatusBar = "迷路を生成しています... - " & Round((prev - UBound(Knots)) / prev, 2) * 100 & "%"
                Else
                    Knots = ArrDelete(Knots, Target)
                    Application.StatusBar = "迷路を生成しています... - 100%"
                    Exit Function
                End If
                
                vTemp = Target.Row
                hTemp = Target.Column
                
                BoolReset CantEnter
                
            End If
            
            If CantEnter(NORTH) = True And CantEnter(EAST) = True And CantEnter(SOUTH) = True And CantEnter(WEST) = True Then '全ての方向に進行不可の場合
                '新規壁を確定して新たな候補を取る
                Exit Do
            End If
            
            DoEvents
        Loop
        
        
        '新規壁を確定
        For i = LBound(TempWalls) + 1 To UBound(TempWalls)
            TempWalls(i).Interior.Color = BUILT
        Next i
        ReDim TempWalls(0 To 0)
        
        '使用された壁候補はリストから削除
        If UBound(Knots) <> 1 Then
            Knots = ArrDelete(Knots, Selected)
            Application.StatusBar = "迷路を生成しています... - " & Round((prev - UBound(Knots)) / prev, 2) * 100 & "%"
        Else
            Knots = ArrDelete(Knots, Selected)
            Application.StatusBar = "迷路を生成しています... - 100%"
            Exit Do
        End If
        
        DoEvents
    Loop
    
End Function

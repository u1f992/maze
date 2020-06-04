Attribute VB_Name = "Module1"
Option Explicit

'迷路の1辺
'必ず奇数
Const SIZE As Long = 103
Public RangeMaze As Range

Public START As Range
Public GOAL As Range

Public Get2Start As Boolean
Public Get2Goal As Boolean

'進行候補(水色の部分)を管理
Public listCandidate() As Range

'色定数
Const BUILT As Long = 0
Const BUILDING As Long = 255

Const SEARCHING As Long = 16776960
Const SEARCHED As Long = 12632256

Const ROUTE As Long = 16711680

'方角定数
Const NORTH As Long = 0
Const EAST As Long = 1
Const SOUTH As Long = 2
Const WEST As Long = 3

Sub main()
    
    DefaultHeightWidth
    
    ReDim listCandidate(0 To 0)
    
    Get2Start = False
    Get2Goal = False

    Set RangeMaze = Range(Cells(1, 1), Cells(SIZE, SIZE))
    
    Application.StatusBar = "迷路を生成しています..."
    
    MakeMaze
    
    Application.StatusBar = "スタート/ゴール地点を設定しています..."

    Set START = RangeMaze.Cells(2, 2) '左上端
    START.Interior.Color = RGB(0, 255, 0)
    
    Set GOAL = RangeMaze.Cells(RangeMaze.Rows.Count - 1, RangeMaze.Columns.Count - 1) '右下端
    GOAL.Interior.Color = RGB(255, 0, 0)
    
    Application.StatusBar = "最短経路探索を行います..."
        
    SetNext START
    Do While Get2Start = False
        SetNext GetNext
        DoEvents
    Loop
    
    Dim Target As Range
    Set Target = Back2Start(GOAL)
    Do While Get2Goal = False
        Set Target = Back2Start(Target)
        DoEvents
    Loop
    
    Application.StatusBar = "最短経路を探索しました。"
    
End Sub

Function MakeMaze()

    '壁候補を管理
    Dim listMakeMaze() As Range
    ReDim listMakeMaze(0 To 0)
    '新規壁を管理
    Dim listBuilding() As Range
    ReDim listBuilding(0 To 0)
    
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
            
            ReDim Preserve listMakeMaze(0 To UBound(listMakeMaze) + 1)
            Set listMakeMaze(UBound(listMakeMaze)) = Cells(i, j)
            
        Next j
    Next i
        
    Dim prev As Long
    prev = UBound(listMakeMaze) - 1
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
        temp = Int((UBound(listMakeMaze) - (LBound(listMakeMaze) + 1) + 1) * Rnd + (LBound(listMakeMaze) + 1))
        
        Set Selected = listMakeMaze(temp)
        
        vTemp = Selected.Row
        hTemp = Selected.Column
        
        '◎壁伸ばし処理
        '壁候補を新規壁に
        Selected.Interior.Color = BUILDING
        ReDim Preserve listBuilding(0 To UBound(listBuilding) + 1)
        Set listBuilding(UBound(listBuilding)) = Selected
        
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
                
                ReDim Preserve listBuilding(0 To UBound(listBuilding) + 1)
                Set listBuilding(UBound(listBuilding)) = Range(Cells(vTemp, hTemp), Target)
                
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
                
                ReDim Preserve listBuilding(0 To UBound(listBuilding) + 1)
                Set listBuilding(UBound(listBuilding)) = Range(Cells(vTemp, hTemp), Target)
                
                '使用された壁候補はリストから削除
                If UBound(listMakeMaze) <> 1 Then
                    listMakeMaze = ArrDelete(listMakeMaze, Target)
                    Application.StatusBar = "迷路を生成しています... - " & Round((prev - UBound(listMakeMaze)) / prev, 2) * 100 & "%"
                Else
                    listMakeMaze = ArrDelete(listMakeMaze, Target)
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
        For i = LBound(listBuilding) + 1 To UBound(listBuilding)
            listBuilding(i).Interior.Color = BUILT
        Next i
        ReDim listBuilding(0 To 0)
        
        '使用された壁候補はリストから削除
        If UBound(listMakeMaze) <> 1 Then
            listMakeMaze = ArrDelete(listMakeMaze, Selected)
            Application.StatusBar = "迷路を生成しています... - " & Round((prev - UBound(listMakeMaze)) / prev, 2) * 100 & "%"
        Else
            listMakeMaze = ArrDelete(listMakeMaze, Selected)
            Application.StatusBar = "迷路を生成しています... - 100%"
            Exit Do
        End If
        
        DoEvents
    Loop
    
End Function

Function DefaultHeightWidth()
    ActiveSheet.Cells.Clear
    ActiveSheet.Rows.RowHeight = 18.75
    ActiveSheet.Columns.ColumnWidth = 8.38
    ActiveSheet.Range(Cells(1, 1), Cells(Rows.Count, Columns.Count)).ClearFormats
    ActiveSheet.Cells(1, 1).Select
End Function

Function BoolReset(ByRef arr() As Boolean)
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        arr(i) = False
    Next i
    
End Function

'常に0を残しておく必要がある
'添え字は1から始まる
Function ArrDelete(ByRef arr() As Range, ByVal Target As Range) As Range()
    
    Dim i As Long
    Dim copy() As Range
    Dim flag As Boolean
    flag = False
    
    For i = LBound(arr) + 1 To UBound(arr)
        If arr(i).Row = Target.Row And arr(i).Column = Target.Column Then
            flag = True
        Else
            If flag = False Then
                ReDim Preserve copy(LBound(arr) To i)
                Set copy(i) = arr(i)
            Else
                ReDim Preserve copy(LBound(arr) To i - 1)
                Set copy(i - 1) = arr(i)
            End If
        End If
    Next i
    
    ArrDelete = copy
    
End Function

Function SetNext(ByVal Target As Range) As Range
    
    '左右上下の進行可能セルをマーキング
    
    Dim Directions(NORTH To WEST) As Range
    Set Directions(NORTH) = Target.Cells(0, 1)
    Set Directions(EAST) = Target.Cells(1, 2)
    Set Directions(SOUTH) = Target.Cells(2, 1)
    Set Directions(WEST) = Target.Cells(1, 0)
    
    Dim i As Long
    
    For i = NORTH To WEST '四方を探索
        If Directions(i).Interior.Color = GOAL.Interior.Color Then 'ゴールにたどりついた場合
            AllCellsChecked
            Get2Start = True
            Exit Function
        End If
        If IsAvailable(Directions(i)) Then '壁か探索済みかスタートではない場合
            
            Directions(i).Interior.Color = SEARCHING
            Directions(i).Value = Target.Value + 1
            
            '添え字は1から始まる、0には空白のデータ
            '要素数が0になるとまずいため
            
            ReDim Preserve listCandidate(0 To UBound(listCandidate) + 1)
            Set listCandidate(UBound(listCandidate)) = Directions(i)
            
        End If
    Next i
    
    If Target.Interior.Color <> START.Interior.Color Then
        Target.Interior.Color = SEARCHED
        listCandidate = ArrDelete(listCandidate, Target)
    End If
    
End Function

Function GetNext() As Range
    
    Dim i As Long
    Dim minimum As Long
    minimum = 2147483647
    
    Dim vTemp As Long
    Dim hTemp As Long
    
    For i = LBound(listCandidate) + 1 To UBound(listCandidate)
        'If minimum >= listcandidate(i).Value + CLng(Sqr((Abs(listcandidate(i).Row - Goal.Row)) ^ 2 + (Abs(listcandidate(i).Row - Goal.Row)) ^ 2)) Then 'スタートからの距離+ゴールまでの距離(直線距離)
        If minimum >= listCandidate(i).Value + Abs(listCandidate(i).Row - GOAL.Row) + Abs(listCandidate(i).Row - GOAL.Row) Then 'スタートからの距離+ゴールまでの距離(辺の合計)
            minimum = listCandidate(i).Value + CLng(Sqr((Abs(listCandidate(i).Row - GOAL.Row)) ^ 2 + (Abs(listCandidate(i).Row - GOAL.Row)) ^ 2))
            
            Set GetNext = listCandidate(i)
            
        End If
    Next i
    
End Function

'壁か探索済みかスタートではない場合true
Function IsAvailable(ByVal Target As Range) As Boolean
    If Target.Interior.Color <> BUILT And Target.Interior.Color <> SEARCHING And Target.Interior.Color <> SEARCHED And Target.Interior.Color <> START.Interior.Color Then
        IsAvailable = True
    End If
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
            Get2Goal = True
            Exit Function
        End If
        
        val(i) = CastValue(Directions(i).Value) '各セルの値(スタートまでの距離)を格納
        
    Next i
    
    Dim s As Long
    s = Smallest(val) '各セルの値が最も小さいものを選択
    Directions(s).Interior.Color = ROUTE
    
    Set Back2Start = Directions(s)
    
End Function

'探索中を探索済みにする
Function AllCellsChecked()
    Dim i As Long
    For i = LBound(listCandidate) + 1 To UBound(listCandidate)
        listCandidate(i).Interior.Color = SEARCHED
    Next i
End Function

'配列のうち最小の値が入っている添え字を返す
Function Smallest(ByRef arr() As Long) As Long
    Dim i As Long
    
    Dim sNum As Long
    sNum = 2147483647
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) < sNum Then
            Smallest = i
            sNum = arr(i)
        End If
    Next i
End Function

'空白文字をLong型の最大値にして返す
'空白文字のセル(壁か未探索)を辿らないようにするため
Function CastValue(ByVal str As String) As Long
    If str = "" Or str = vbNullString Then
        CastValue = 2147483647
    Else
        CastValue = CLng(str)
    End If
End Function

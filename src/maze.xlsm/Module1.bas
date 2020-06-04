Attribute VB_Name = "Module1"
Option Explicit

'trueにするとスタートとゴールを入力できる
Public GlobalToolSwitch As Boolean
'設定し終えたらtrue
Public SetStart As Boolean
Public SetGoal As Boolean

Public size As Long

Public Start As Range
Public Goal As Range

Public Get2Start As Boolean
Public Get2Goal As Boolean


Public RangeMaze As Range

'進行候補(水色の部分)を管理
Public listCandidate() As Range

'色定数
Const BUILT As Long = 0 '建設済み
Const BUILDING As Long = 255 '建設中

Const NORTH As Long = 1
Const EAST As Long = 2
Const SOUTH As Long = 3
Const WEST As Long = 4


Sub main()
    
    ReDim listCandidate(0 To 0)
    
    Get2Start = False
    Get2Goal = False
    
    DefaultHeightWidth
    
    GlobalToolSwitch = False
    SetStart = False
    SetGoal = False
    
    '壁に1
    '通路に0
    '建設中の壁を2
    '要素数5以上
    Dim maze() As Integer
    
    Dim Target As Range
    
'    Dim temp As String
'    temp = InputBox("サイズを3以上の奇数で入力してください。")
'
'    If temp <> "" Then
'        If IsNumeric(temp) = True Then
'            size = CLng(temp)
'        Else
'            Exit Sub
'        End If
'    End If
'
'    If size < 3 Then
'        size = 3
'    End If
'    If size Mod 2 = 0 Then
'        size = size + 1
'    End If
    
    '自動でサイズを決定する
'    Do While True
'        Randomize
'        size = Int((201 - 181 + 1) * Rnd + 181) 'Int((51 - 3 + 1) * Rnd + 3)
'        If size Mod 2 <> 0 Then
'            Exit Do
'        End If
'    Loop

    size = 501
    Set RangeMaze = Range(Cells(1, 1), Cells(size + 2, size + 2))
    
'    Dim t As Double
'    t = Timer

    Application.StatusBar = "迷路を生成しています..."
    
    'MakeMaze
    MakeMaze
    Cells(size + 3, size + 3).Select
    
    'MsgBox Round(Timer - t, 2) & " sec."
    
    'SetNext
'    MsgBox "スタート地点を選択してください。"
'    GlobalToolSwitch = True
'    Do While GlobalToolSwitch = True
'        DoEvents
'    Loop
    
    'ランダムに開始地点と終了地点を指定する
    
    Application.StatusBar = "スタート/ゴール地点を設定しています..."
    
'    Randomize
'    Set Target = Cells(Int((size + 2 - 1 + 1) * Rnd + 1), Int((size + 2 - 1 + 1) * Rnd + 1))
'    Do While Target.Interior.Color = RGB(0, 0, 0)
'        Randomize
'        Set Target = Cells(Int((size + 2 - 1 + 1) * Rnd + 1), Int((size + 2 - 1 + 1) * Rnd + 1))
'    Loop
'
'    Set Start = Target
'    Start.Interior.Color = RGB(0, 255, 0)
'
'    Randomize
'    Set Target = Cells(Int((size + 2 - 1 + 1) * Rnd + 1), Int((size + 2 - 1 + 1) * Rnd + 1))
'    Do While Target.Interior.Color = RGB(0, 0, 0)
'        Randomize
'        Set Target = Cells(Int((size + 2 - 1 + 1) * Rnd + 1), Int((size + 2 - 1 + 1) * Rnd + 1))
'    Loop
'    Set Goal = Target

    Set Start = Cells(2, 2)
    Start.Interior.Color = RGB(0, 255, 0)
    
    Set Goal = Cells(size + 1, size + 1)
    Goal.Interior.Color = RGB(255, 0, 0)
    Stop
    Application.StatusBar = "最短経路探索を行います..."
        
    SetNext Start
    Do While Get2Start = False
        SetNext GetNext
        DoEvents
    Loop
        
    Set Target = Back2Start(Goal)
    Do While Get2Goal = False
        Set Target = Back2Start(Target)
        DoEvents
    Loop
    
    Application.StatusBar = "最短経路を探索しました。"
    
End Sub

Function MakeMaze()

    '壁候補(x,yとも奇数かつ枠ではない)を管理
    Dim listMakeMaze() As Range
    ReDim listMakeMaze(0 To 0)
    '建設中の壁(赤)を管理
    Dim listBuilding() As Range
    ReDim listBuilding(0 To 0)
    
    '迷路の初期化
    RangeMaze.Rows.RowHeight = 5 * 0.75
    RangeMaze.Columns.ColumnWidth = 5 * 0.07
    '外周を壁に
    RangeMaze.Interior.Color = BUILT
    Range(RangeMaze.Cells(2, 2), RangeMaze.Cells(RangeMaze.Rows.Count - 1, RangeMaze.Columns.Count - 1)).ClearFormats
    
    Dim i As Long
    Dim j As Long
    '壁候補座標のリストを作成
    For i = 3 To size Step 2
        For j = 3 To size Step 2
            
            '壁候補(x,yとも奇数かつ枠ではない)を追加
            ReDim Preserve listMakeMaze(0 To UBound(listMakeMaze) + 1)
            Set listMakeMaze(UBound(listMakeMaze)) = Cells(i, j)
            
        Next j
    Next i
        
    Dim prev As Long
    prev = UBound(listMakeMaze) - 1
    Application.StatusBar = "迷路を生成しています... - 0%"
        
    Dim Selected As Range
    
    Dim temp As Long
    Dim Direction As Long '方向を決定 '1北・2東・3南・4西
    Dim CantEnter(1 To 4) As Boolean '進めない場合にtrue
    
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
        '指定座標を壁に
        Selected.Interior.Color = BUILDING
        ReDim Preserve listBuilding(0 To UBound(listBuilding) + 1)
        Set listBuilding(UBound(listBuilding)) = Selected
        
        '進行方向をランダムに決定
        BoolReset CantEnter
        
        Do While True
        
            Randomize
            Direction = Int(4 * Rnd + 1)
            
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
            
            If Target.Interior.Color = BUILT Then '進行先が既存の壁の場合
                Range(Cells(vTemp, hTemp), Target).Interior.Color = BUILDING
                
                ReDim Preserve listBuilding(0 To UBound(listBuilding) + 1)
                Set listBuilding(UBound(listBuilding)) = Range(Cells(vTemp, hTemp), Target)
                
                Exit Do
                
            ElseIf Target.Interior.Color = BUILDING Then '建設中の壁の場合
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
                Range(Cells(vTemp, hTemp), Target).Interior.Color = BUILDING
                
                ReDim Preserve listBuilding(0 To UBound(listBuilding) + 1)
                Set listBuilding(UBound(listBuilding)) = Range(Cells(vTemp, hTemp), Target)
                
                '使用された壁候補はリストから削除
                If UBound(listMakeMaze) <> 1 Then
                    listMakeMaze = RangeArrDelete(listMakeMaze, Target)
                    Application.StatusBar = "迷路を生成しています... - " & Round((prev - UBound(listMakeMaze)) / prev, 2) * 100 & "%"
                Else
                    listMakeMaze = RangeArrDelete(listMakeMaze, Target)
                    Application.StatusBar = "迷路を生成しています... - 100%"
                    Exit Function
                End If
                
                vTemp = Target.Row
                hTemp = Target.Column
                
                BoolReset CantEnter
                
            End If
            
            If CantEnter(NORTH) = True And CantEnter(EAST) = True And CantEnter(SOUTH) = True And CantEnter(WEST) = True Then 'どの方向にも進めなくなったら新たな候補を取る
                Exit Do
            End If
            
            DoEvents
        Loop
        
        
        '建設中のリストを全て建設済みに
        For i = LBound(listBuilding) + 1 To UBound(listBuilding)
            listBuilding(i).Interior.Color = BUILT
        Next i
        ReDim listBuilding(0 To 0)
        
        '使用された壁候補はリストから削除
        If UBound(listMakeMaze) <> 1 Then
            listMakeMaze = RangeArrDelete(listMakeMaze, Selected)
            Application.StatusBar = "迷路を生成しています... - " & Round((prev - UBound(listMakeMaze)) / prev, 2) * 100 & "%"
        Else
            listMakeMaze = RangeArrDelete(listMakeMaze, Selected)
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

Function BoolSwitch(ByRef arr() As Boolean, ByVal num As Long)
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If i = num Then
            arr(i) = True
        Else
            arr(i) = False
        End If
    Next i
    
End Function

Function str2coordinate(ByVal str As String, ByRef vertical As Long, ByRef horizontal As Long)

    vertical = CLng(Left(str, InStr(str, ",") - 1))
    horizontal = CLng(Mid(str, InStr(str, ",") + 1))
    
End Function

'添え字numを削除
'ただし要素数が1(添え字0)の場合、内容をvbNullStringにする
Function CloseGap(ByRef arr() As String, ByVal num As Long) As String()
    
    Dim i As Long
    Dim flag As Boolean
    flag = False '消すまでfalse、過ぎたらTrue
    Dim copy() As String
    
    If UBound(arr) = 0 Then
        ReDim copy(0 To 0)
        copy(0) = vbNullString
        CloseGap = copy
        Exit Function
    End If
    
    For i = LBound(arr) To UBound(arr)
        If i <> num Then
            If flag = False Then
                ReDim Preserve copy(LBound(arr) To i)
                copy(i) = arr(i)
            Else
                ReDim Preserve copy(LBound(arr) To i - 1)
                copy(i - 1) = arr(i)
            End If
        Else
            flag = True
        End If
    Next i
    
    CloseGap = copy
    
End Function

'ただし要素数が1(添え字0)の場合、内容をvbNullStringにする
Function CloseGapByVal(ByRef arr() As String, ByVal val As String) As String()
    
    Dim i As Long
    Dim flag As Boolean
    flag = False '消すまでfalse、過ぎたらTrue
    Dim copy() As String
    
    If UBound(arr) = 0 Then
        ReDim copy(0 To 0)
        copy(0) = vbNullString
        CloseGapByVal = copy
        Exit Function
    End If
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) <> val Then
            If flag = False Then
                ReDim Preserve copy(LBound(arr) To i)
                copy(i) = arr(i)
            Else
                ReDim Preserve copy(LBound(arr) To i - 1)
                copy(i - 1) = arr(i)
            End If
        Else
            flag = True
        End If
    Next i
    
    CloseGapByVal = copy
    
End Function

'常に0を残しておく必要がある
'添え字は1から始まる
'listCandidateからRangeを削除
Function RangeArrDelete(ByRef arr() As Range, ByVal Target As Range) As Range()
    
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
    
    RangeArrDelete = copy
    
End Function

Function SetNext(ByVal Target As Range) As Range
    
    
    '選択中の色 : 0,255,255
    '探索済みの色 : 192,192,192
    
    '左右上下の進行可能セルをマーキング
    
    Dim Directions(0 To 3) As Range
    Set Directions(0) = Target.Cells(0, 1)
    Set Directions(1) = Target.Cells(1, 2)
    Set Directions(2) = Target.Cells(2, 1)
    Set Directions(3) = Target.Cells(1, 0)
    
    Dim i As Long
    
    For i = 0 To 3
        If Directions(i).Interior.Color = RGB(255, 0, 0) Then 'ゴールにたどりついた場合
            AllCellsChecked
            Get2Start = True
            Exit Function
        End If
        If IsAvailable(Directions(i)) Then '壁か探索済みではない場合
            
            Directions(i).Interior.Color = RGB(0, 255, 255)
            Directions(i).Value = Target.Value + 1
            
            '添え字は1から始まる、0には空白のデータ
            '要素数が0になるとまずいため
            
            ReDim Preserve listCandidate(0 To UBound(listCandidate) + 1)
            Set listCandidate(UBound(listCandidate)) = Directions(i)
            
        End If
    Next i
    
    If Target.Interior.Color <> Start.Interior.Color Then
        Target.Interior.Color = RGB(192, 192, 192)
        listCandidate = RangeArrDelete(listCandidate, Target)
    End If
    
End Function

Function GetNext() As Range
    Dim i As Long
    Dim minimum As Long
    minimum = (size + 2) ^ 2
    
    Dim vTemp As Long
    Dim hTemp As Long
    
    For i = LBound(listCandidate) + 1 To UBound(listCandidate)
        'If minimum >= listcandidate(i).Value + CLng(Sqr((Abs(listcandidate(i).Row - Goal.Row)) ^ 2 + (Abs(listcandidate(i).Row - Goal.Row)) ^ 2)) Then 'スタートからの距離+ゴールまでの距離(直線距離)
        If minimum >= listCandidate(i).Value + Abs(listCandidate(i).Row - Goal.Row) + Abs(listCandidate(i).Row - Goal.Row) Then 'スタートからの距離+ゴールまでの距離(辺の合計)
            minimum = listCandidate(i).Value + CLng(Sqr((Abs(listCandidate(i).Row - Goal.Row)) ^ 2 + (Abs(listCandidate(i).Row - Goal.Row)) ^ 2))
            
            Set GetNext = listCandidate(i)
            
        End If
    Next
    
End Function

Function IsAvailable(ByVal Target As Range) As Boolean
    
    If Target.Interior.Color <> RGB(0, 0, 0) And Target.Interior.Color <> RGB(0, 255, 255) And Target.Interior.Color <> RGB(192, 192, 192) And Target.Interior.Color <> RGB(0, 255, 0) Then
        IsAvailable = True
    End If
    
End Function

'再帰でスタートまで帰る
Function Back2Start(ByVal Target As Range) As Range
    
    Dim Directions(0 To 3) As Range
    Set Directions(0) = Target.Cells(0, 1)
    Set Directions(1) = Target.Cells(1, 2)
    Set Directions(2) = Target.Cells(2, 1)
    Set Directions(3) = Target.Cells(1, 0)
    
    Dim i As Long
    
    For i = 0 To 3 '4方向にスタートがあれば終了
        If Directions(i).Interior.Color = RGB(0, 255, 0) Then
            Get2Goal = True
            Exit Function
        End If
    Next i
    
    Dim val(0 To 3) As Long '各セルの値(スタートまでの距離)が最も短いものを選択
    For i = 0 To 3
        val(i) = CastValue(Directions(i).Value)
    Next i
    
    Dim s As Long
    s = Smallest(val)
    Directions(s).Interior.Color = RGB(0, 0, 255)
    
    Set Back2Start = Directions(s)
    
End Function

'選択中の水色セルを全て灰色にする
Function AllCellsChecked()
    Dim cell As Range
    
    For Each cell In RangeMaze
        If cell.Interior.Color = RGB(0, 255, 255) Then
            cell.Interior.Color = RGB(192, 192, 192)
        End If
    Next
    
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
Function CastValue(ByVal str As String) As Long
    If str = "" Or str = vbNullString Then
        CastValue = 2147483647
    Else
        CastValue = CLng(str)
    End If
End Function

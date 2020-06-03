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

Sub main()
    
    ReDim listCandidate(0 To 0)
    
'    On Error GoTo StackOverflow
    
    Dim flag As Boolean
    flag = False
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
    
    Dim target As Range
    
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
    Do While True
        Randomize
        size = Int((201 - 181 + 1) * Rnd + 181) 'Int((51 - 3 + 1) * Rnd + 3)
        If size Mod 2 <> 0 Then
            Exit Do
        End If
    Loop

'    size = 501

    
'    Dim t As Double
'    t = Timer
    
    ReDim maze(0 To size + 1, 0 To size + 1)
    
    'MakeMaze
    MakeMaze maze
    
    '描画
    DrawMaze maze
    
    Set RangeMaze = Range(Cells(1, 1), Cells(size + 2, size + 2))
    
    
    'MsgBox Round(Timer - t, 2) & " sec."
    
    'SetNext
'    MsgBox "スタート地点を選択してください。"
'    GlobalToolSwitch = True
'    Do While GlobalToolSwitch = True
'        DoEvents
'    Loop
    
    'ランダムに開始地点と終了地点を指定する
    
    Randomize
    Set target = Cells(Int((size + 2 - 1 + 1) * Rnd + 1), Int((size + 2 - 1 + 1) * Rnd + 1))
    Do While target.Interior.Color = RGB(0, 0, 0)
        Randomize
        Set target = Cells(Int((size + 2 - 1 + 1) * Rnd + 1), Int((size + 2 - 1 + 1) * Rnd + 1))
    Loop
    
'    If target.Row Mod 2 <> 0 Then
'        Set target = Cells(target.Row + 1, target.Column)
'    End If
'    If target.Column Mod 2 <> 0 Then
'        Set target = Cells(target.Row, target.Column + 1)
'    End If
    
    Set Start = target
    Start.Interior.Color = RGB(0, 255, 0)
    
    Randomize
    Set target = Cells(Int((size + 2 - 1 + 1) * Rnd + 1), Int((size + 2 - 1 + 1) * Rnd + 1))
    Do While target.Interior.Color = RGB(0, 0, 0)
        Randomize
        Set target = Cells(Int((size + 2 - 1 + 1) * Rnd + 1), Int((size + 2 - 1 + 1) * Rnd + 1))
    Loop
    
'    If target.Row Mod 2 <> 0 Then
'        Set target = Cells(target.Row + 1, target.Column)
'    End If
'    If target.Column Mod 2 <> 0 Then
'        Set target = Cells(target.Row, target.Column + 1)
'    End If
    
    Set Goal = target
    Goal.Interior.Color = RGB(255, 0, 0)
    
    MsgBox "最短経路探索を行います。"
        
    SetNext Start
    Do While Get2Start = False
        SetNext GetNext
        DoEvents
    Loop
        
    Set target = Back2Start(Goal)
    Do While Get2Goal = False
        Set target = Back2Start(target)
    Loop
    
    MsgBox "最短経路を探索しました。"
    flag = True
    
StackOverflow:
    If flag <> True Then
        MsgBox "スタック領域が不足しています。"
    End If
    
End Sub

Function MakeMaze(ByRef maze() As Integer)
    
    Dim i As Long
    Dim j As Long
    
    Dim vTemp As Long
    Dim hTemp As Long
    
    Dim vStart As Long
    Dim hStart As Long
    
    Dim Direction As Long
    
    'それぞれ北東南西に対応
    '進行不能になるとTrue
    Dim bool(3) As Boolean
    '同上
    '詰む直前回進んだ方向
    Dim flag(3) As Boolean
    
    '迷路を初期化
    '外周を壁とし、それ以外を通路とする
    For i = LBound(maze, 1) To UBound(maze, 1)
        For j = LBound(maze, 2) To UBound(maze, 2)
            If i = LBound(maze, 1) Or j = LBound(maze, 2) Or i = UBound(maze, 1) Or j = UBound(maze, 2) Then
                maze(i, j) = 1
            Else
                maze(i, j) = 0
            End If
        Next j
    Next i
    
    'Int((最大値 - 最小値 + 1) * Rnd + 最小値)
    
    '壁伸ばし候補座標
    Dim tempAvailable() As String
    ReDim tempAvailable(0 To (UBound(maze, 1) - 2) / 2 - 1, 0 To (UBound(maze, 2) - 2) / 2 - 1)
    For i = 0 To UBound(tempAvailable, 1)
        For j = 0 To UBound(tempAvailable, 2)
            tempAvailable(i, j) = (i + 1) * 2 & "," & (j + 1) * 2
        Next j
    Next i
    
    Dim listAvailable() As String
    ReDim listAvailable(0 To ((UBound(tempAvailable, 1) + 1) * (UBound(tempAvailable, 2) + 1)) - 1)
    For i = 0 To UBound(listAvailable)
        listAvailable(i) = tempAvailable(i Mod ((UBound(maze, 1) - 2) / 2), i \ ((UBound(maze, 1) - 2) / 2))
    Next i
    
    Do While listAvailable(0) <> vbNullString '要素数が0になるまで
        
        j = 0
        
        'ランダムに取り出す
        Randomize
        i = Int((UBound(listAvailable) - 0 + 1) * Rnd + 0)
        str2coordinate listAvailable(i), vTemp, hTemp
        
        
        '一度使用された要素は切り取る
        '要素数が0になるとき、listAvailable(0)にvbNullStringが代入される
        listAvailable = CloseGap(listAvailable, i)

        If maze(vTemp, hTemp) = False Then
            'まだ通路の場合、壁伸ばし処理
            
            '指定座標を壁に
            maze(vTemp, hTemp) = 2
'            DrawMaze maze

                
            Do While True
                
                Randomize
                Direction = Int((4 - 1 + 1) * Rnd + 1)
                
                If vTemp = 0 Or hTemp = 0 Then 'たまに0になるバグが発生するので対症療法
                    Exit Do
                End If
                
                Select Case Direction
                    Case 1 '北
                        If maze(vTemp - 2, hTemp) = 1 Then '進行先が既存の壁の場合
                            maze(vTemp - 1, hTemp) = 2
                            Exit Do
                        ElseIf maze(vTemp - 2, hTemp) = 0 Then '通路の場合
                            maze(vTemp - 1, hTemp) = 2
                            maze(vTemp - 2, hTemp) = 2
                            
                            vTemp = vTemp - 2
                            hTemp = hTemp
                            
                            BoolReset bool
                            
                            If flag(1) = True Or flag(2) = True Or flag(3) = True Then
                                BoolSwitch flag, 0
                                j = 0
                            Else
                                bool(0) = True
                                j = j + 1
                            End If
                            
                        ElseIf maze(vTemp - 2, hTemp) = 2 Then '建設中の壁の場合
                            bool(0) = True
                        End If
                        
                    Case 2 '東
                        If maze(vTemp, hTemp + 2) = 1 Then '進行先が既存の壁の場合
                            maze(vTemp, hTemp + 1) = 2
                            Exit Do
                        ElseIf maze(vTemp, hTemp + 2) = 0 Then '通路の場合
                            maze(vTemp, hTemp + 1) = 2
                            maze(vTemp, hTemp + 2) = 2
                            
                            vTemp = vTemp
                            hTemp = hTemp + 2
                            
                            BoolReset bool
                                                        
                            If flag(0) = True Or flag(2) = True Or flag(3) = True Then
                                BoolSwitch flag, 1
                                j = 0
                            Else
                                flag(1) = True
                                j = j + 1
                            End If
                            
                        ElseIf maze(vTemp, hTemp + 2) = 2 Then '建設中の壁の場合
                            bool(1) = True
                        End If
                    
                    Case 3 '南
                        If maze(vTemp + 2, hTemp) = 1 Then '進行先が既存の壁の場合
                            maze(vTemp + 1, hTemp) = 2
                            Exit Do
                        ElseIf maze(vTemp + 2, hTemp) = 0 Then '通路の場合
                            maze(vTemp + 1, hTemp) = 2
                            maze(vTemp + 2, hTemp) = 2
                            
                            vTemp = vTemp + 2
                            hTemp = hTemp
                            
                            BoolReset bool
                                                        
                            If flag(0) = True Or flag(1) = True Or flag(3) = True Then
                                BoolSwitch flag, 2
                                j = 0
                            Else
                                flag(2) = True
                                j = j + 1
                            End If
                            
                        ElseIf maze(vTemp + 2, hTemp) = 2 Then '建設中の壁の場合
                            bool(2) = True
                        End If
                        
                    Case 4 '西
                        If maze(vTemp, hTemp - 2) = 1 Then '進行先が既存の壁の場合
                            maze(vTemp, hTemp - 1) = 2
                            Exit Do
                        ElseIf maze(vTemp, hTemp - 2) = 0 Then '通路の場合
                            maze(vTemp, hTemp - 1) = 2
                            maze(vTemp, hTemp - 2) = 2
                            
                            vTemp = vTemp
                            hTemp = hTemp - 2
                            
                            BoolReset bool
                                                        
                            If flag(0) = True Or flag(1) = True Or flag(2) = True Then
                                BoolSwitch flag, 3
                                j = 0
                            Else
                                flag(3) = True
                                j = j + 1
                            End If
                            
                        ElseIf maze(vTemp, hTemp - 2) = 2 Then '建設中の壁の場合
                            bool(3) = True
                        End If
                    
                End Select
                
                If bool(0) = True And bool(1) = True And bool(2) = True And bool(3) = True Then
                    
                    
                    If flag(0) = True Then
                        vTemp = vTemp + 2 * (j + 1)
                    ElseIf flag(1) = True Then
                        hTemp = hTemp - 2 * (j + 1)
                    ElseIf flag(2) = True Then
                        vTemp = vTemp - 2 * (j + 1)
                    ElseIf flag(3) = True Then
                        hTemp = hTemp + 2 * (j + 1)
                    Else
                        
                        '渦巻き状の中心になってしまった時など
                        '現状ではどう復帰したらいいかわからないので、
                        '始点に戻ってもらって運良くどこかの壁に接地するのを願う
                        'この線を諦めて新たな点から始める
                        
                        Randomize
                        Select Case Int(2 * Rnd + 1)
                            Case 1
                                vTemp = vStart
                                hTemp = hStart
                            Case 2
                                Exit Do
                        End Select
                        
                    End If
                    
                    BoolReset bool
                    BoolReset flag
                End If
                
                DrawMaze maze
                
                DoEvents
                
            Loop
            
            j = 0
            BoolReset bool
            BoolReset flag
            
            '建設中=2の壁を1に
            For i = LBound(maze, 1) To UBound(maze, 1)
                For j = LBound(maze, 2) To UBound(maze, 2)
                    If maze(i, j) = 2 Then
                        maze(i, j) = 1
                        
'                        '候補点のうち既に壁になったものをlistAvailableから削除
'                        If i Mod 2 = 0 And j Mod 2 = 0 Then
'                            listAvailable = CloseGapByVal(listAvailable, i & "," & j)
'                        End If
                        
                    End If
                Next j
            Next i
            
            DrawMaze maze
            
        End If
        
    Loop
    
End Function

Function DrawMaze(ByRef maze() As Integer)
    
    Range(Cells(LBound(maze, 1) + 1, LBound(maze, 2) + 1), Cells(UBound(maze, 1) + 1, UBound(maze, 2) + 1)).Rows.RowHeight = 5 * 0.75
    Range(Cells(LBound(maze, 1) + 1, LBound(maze, 2) + 1), Cells(UBound(maze, 1) + 1, UBound(maze, 2) + 1)).Columns.ColumnWidth = 5 * 0.07
    
    Dim i As Long
    Dim j As Long
    
    '0 : 何もしない
    '1 : RGB(0,0,0)
    
    For i = LBound(maze, 1) To UBound(maze, 1)
        For j = LBound(maze, 2) To UBound(maze, 2)
            If maze(i, j) = 1 Then
                Cells(i + 1, j + 1).Interior.Color = RGB(0, 0, 0)
            ElseIf maze(i, j) = 2 Then
                Cells(i + 1, j + 1).Interior.Color = RGB(255, 0, 0)
            End If
        Next j
    Next i
    
    Cells(UBound(maze, 1) + 2, UBound(maze, 2) + 2).Select
    
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
Function RangeArrDelete(ByRef arr() As Range, ByVal target As Range) As Range()
    
    Dim i As Long
    Dim copy() As Range
    Dim flag As Boolean
    flag = False
    
    For i = LBound(arr) + 1 To UBound(arr)
        If arr(i).Row = target.Row And arr(i).Column = target.Column Then
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

Function SetNext(ByVal target As Range) As Range
    
    
    '選択中の色 : 0,255,255
    '探索済みの色 : 192,192,192
    
    '左右上下の進行可能セルをマーキング
    
    Dim Directions(0 To 3) As Range
    Set Directions(0) = target.Cells(0, 1)
    Set Directions(1) = target.Cells(1, 2)
    Set Directions(2) = target.Cells(2, 1)
    Set Directions(3) = target.Cells(1, 0)
    
    Dim i As Long
    
    For i = 0 To 3
        If Directions(i).Interior.Color = RGB(255, 0, 0) Then 'ゴールにたどりついた場合
            AllCellsChecked
            Get2Start = True
            Exit Function
        End If
        If IsAvailable(Directions(i)) Then '壁か探索済みではない場合
            
            Directions(i).Interior.Color = RGB(0, 255, 255)
            Directions(i).Value = target.Value + 1
            
            '添え字は1から始まる、0には空白のデータ
            '要素数が0になるとまずいため
            
            ReDim Preserve listCandidate(0 To UBound(listCandidate) + 1)
            Set listCandidate(UBound(listCandidate)) = Directions(i)
            
        End If
    Next i
    
    If target.Interior.Color <> Start.Interior.Color Then
        target.Interior.Color = RGB(192, 192, 192)
        listCandidate = RangeArrDelete(listCandidate, target)
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

Function IsAvailable(ByVal target As Range) As Boolean
    
    If target.Interior.Color <> RGB(0, 0, 0) And target.Interior.Color <> RGB(0, 255, 255) And target.Interior.Color <> RGB(192, 192, 192) And target.Interior.Color <> RGB(0, 255, 0) Then
        IsAvailable = True
    End If
    
End Function

'再帰でスタートまで帰る
Function Back2Start(ByVal target As Range) As Range
    
    Dim Directions(0 To 3) As Range
    Set Directions(0) = target.Cells(0, 1)
    Set Directions(1) = target.Cells(1, 2)
    Set Directions(2) = target.Cells(2, 1)
    Set Directions(3) = target.Cells(1, 0)
    
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

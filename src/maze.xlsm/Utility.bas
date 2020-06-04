Attribute VB_Name = "Utility"
Option Explicit

'Boolean型の配列の全てをFalseに
Function BoolReset(ByRef arr() As Boolean)
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        arr(i) = False
    Next i
End Function

'探索中を探索済みにする
Function AllCellsChecked()
    Dim i As Long
    For i = LBound(TempSearch) + 1 To UBound(TempSearch)
        TempSearch(i).Interior.Color = SEARCHED
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

'壁か探索済みかスタートではない場合True
Function IsAvailable(ByVal Target As Range) As Boolean
    If Target.Interior.Color <> BUILT And Target.Interior.Color <> SEARCHING And Target.Interior.Color <> SEARCHED And Target.Interior.Color <> START.Interior.Color Then
        IsAvailable = True
    End If
End Function

'常に0を残しておく必要がある
'添え字は0から始まるが、1から処理をする
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

'arr()をひとつ広げて入れるだけ
Function ArrAdd(ByRef arr() As Range, ByVal Target As Range) As Range()
    ReDim Preserve arr(0 To UBound(arr) + 1)
    Set arr(UBound(arr)) = Target
    
    ArrAdd = arr
    
End Function

Attribute VB_Name = "Main"
Option Explicit

Sub Main()
    
    DefaultHeightWidth
    
    ReDim TempSearch(0 To 0)

    Set RangeMaze = Range(Cells(1, 1), Cells(SIZE, SIZE))
    
    Application.StatusBar = "���H�𐶐����Ă��܂�..."
    
    MakeMaze
    
    Application.StatusBar = "�X�^�[�g/�S�[���n�_��ݒ肵�Ă��܂�..."

    Set START = RangeMaze.Cells(2, 2) '����[
    START.Interior.Color = RGB(0, 255, 0)
    
    Set GOAL = RangeMaze.Cells(RangeMaze.Rows.Count - 1, RangeMaze.Columns.Count - 1) '�E���[
    GOAL.Interior.Color = RGB(255, 0, 0)
    
    Application.StatusBar = "�ŒZ�o�H�T�����s���܂�..."
        
    SearchSet START
    Do While SearchSet(SearchGet) = False
        DoEvents
    Loop

    SolveSet GOAL
    Do While SolveSet(SolveGet) = False
        DoEvents
    Loop
    
    Application.StatusBar = "�ŒZ�o�H��T�����܂����B"
    
End Sub

Function DefaultHeightWidth()
    ActiveSheet.Cells.Clear
    ActiveSheet.Rows.RowHeight = 18.75
    ActiveSheet.Columns.ColumnWidth = 8.38
    ActiveSheet.Range(Cells(1, 1), Cells(Rows.Count, Columns.Count)).ClearFormats
    ActiveSheet.Cells(1, 1).Select
End Function

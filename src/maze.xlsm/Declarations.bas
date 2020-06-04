Attribute VB_Name = "Declarations"
Option Explicit

'迷路の1辺
'必ず奇数
Public Const SIZE As Long = 103
Public RangeMaze As Range

Public START As Range
Public GOAL As Range

'色定数
Public Const BUILT As Long = 0
Public Const BUILDING As Long = 255

Public Const SEARCHING As Long = 16776960
Public Const SEARCHED As Long = 12632256

Public Const ROUTE As Long = 16711680

'方角定数
Public Const NORTH As Long = 0
Public Const EAST As Long = 1
Public Const SOUTH As Long = 2
Public Const WEST As Long = 3

'進行候補(水色の部分)を管理
Public TempSearch() As Range

'辿ってきた道
Public TempRoute As Range

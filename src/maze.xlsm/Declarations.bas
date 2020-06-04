Attribute VB_Name = "Declarations"
Option Explicit

'–À˜H‚Ì1•Ó
'•K‚¸Šï”
Public Const SIZE As Long = 53
Public RangeMaze As Range

Public START As Range
Public GOAL As Range

'F’è”
Public Const BUILT As Long = 0
Public Const BUILDING As Long = 255

Public Const SEARCHING As Long = 16776960
Public Const SEARCHED As Long = 12632256

Public Const ROUTE As Long = 16711680

'•ûŠp’è”
Public Const NORTH As Long = 0
Public Const EAST As Long = 1
Public Const SOUTH As Long = 2
Public Const WEST As Long = 3

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMath"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function fSum(ByVal X As Long, ByVal Y As Long) As Long
    fSum = X + Y
End Function

Public Function fSub(ByVal X As Long, Y As Long) As Long
    fSub = X - Y
End Function

Public Function fMult(ByVal X As Long, Y As Long) As Long
    fMult = X * Y
End Function

Public Function fDiv(ByVal X As Long, Y As Long) As Long
    If Y <> 0 Then
       fDiv = X / Y
    Else
       MsgBox " The divider must be different of zero.!"
    End If
End Function


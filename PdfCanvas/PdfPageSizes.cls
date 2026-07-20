VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfPageSizes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get A5() As Long
    A5 = 0
End Property

Public Property Get A4() As Long
    A4 = 1
End Property

Public Property Get A3() As Long
    A3 = 2
End Property

Public Property Get Letter() As Long
    Letter = 3
End Property

Public Property Get Legal() As Long
    Legal = 4
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfPageObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPageObject As Long
Private mContentObject As Long


Public Property Get PageObject() As Long

    PageObject = mPageObject

End Property

Public Property Let PageObject(ByVal Value As Long)

    mPageObject = Value

End Property

Public Property Get ContentObject() As Long

    ContentObject = mContentObject

End Property

Public Property Let ContentObject(ByVal Value As Long)

    mContentObject = Value

End Property

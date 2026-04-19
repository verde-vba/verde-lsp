VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassBasic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private myData As Object
Private myPrev As Object
Private myNext As Object

Public Property Get Data() As Object
    Set Data = myData
End Property

Public Property Let Data(ByRef newData As Object)
    Set myData = newData
End Property

Public Property Get Prev() As ClassBasic
    Set Prev = myPrev
End Property

Public Property Let Prev(ByRef newPrev As ClassBasic)
    Set myPrev = newPrev
End Property

Public Property Get Next() As ClassBasic
    Set Next = myNext
End Property

Public Property Let Next(ByRef newNext As ClassBasic)
    Set myNext = newNext
End Property

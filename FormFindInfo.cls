VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormFindInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIndex As Long
Private mName As String
Private mCaption As String
Private mIsCaptionLoaded As Boolean
Private mIsDisplayed As Boolean

Public Property Let Index(Value As Long)
    mIndex = Value
End Property

Public Property Let Name(Value As String)
    mName = Value
End Property

Public Property Let Caption(Value As String)
    mCaption = Value
End Property

Public Property Let IsDisplayed(Value As Boolean)
    mIsDisplayed = Value
End Property

Public Property Get Index() As Long
    Index = mIndex
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Get Caption() As String
    Caption = mCaption
    mIsCaptionLoaded = True
End Property

Public Property Get IsCaptionLoaded() As Boolean
    IsCaptionLoaded = mIsCaptionLoaded
End Property

Public Property Get IsDisplayed() As Boolean
    IsDisplayed = mIsDisplayed
End Property

Public Property Get Text() As String
    Text = mCaption & " " & "(" & mName & ")"
End Property

Private Sub Class_Initialize()
    mIsCaptionLoaded = False
    mIsDisplayed = False
End Sub

VERSION 5.00
Begin VB.Form frmFormFinder 
   Caption         =   "Find Forms"
   ClientHeight    =   5070
   ClientLeft      =   2190
   ClientTop       =   1950
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopyCaption 
      Caption         =   "Copy Item Caption"
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopyName 
      Caption         =   "Copy Item Name"
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpenExplorer 
      Caption         =   "Open Containing Folder"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdLoadCaptions 
      Caption         =   "Load Forms Captions"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox lstForms 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   5175
   End
   Begin VB.CommandButton CmdShowCode 
      Caption         =   "Open Form Code"
      Default         =   -1  'True
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtFilter 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdShowDesigner 
      Caption         =   "Open Form Designer"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.Label ListCaption 
      Caption         =   "Form List"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Enter part of a Form name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmFormFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' limit maximum number of loaded Form Captions
' reading form caption requires loading the form. And loading too many forms in memory can lead to IDE instability
Private Const MAX_CAPTIONS_LOAD = 200

Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private FormList() As FormFindInfo

Private Sub Form_Load()
    InitializeFormList
    FillList lstForms, sFilter:=""
End Sub

Private Sub Form_Resize()
    Dim newWidth As Single
    Dim newHeight As Single
    
    newWidth = Me.ScaleWidth - lstForms.Left - 10
    newHeight = Me.ScaleHeight - lstForms.Top - 10
    If newWidth > 0 And newHeight > 0 Then
        lstForms.Move lstForms.Left, lstForms.Top, newWidth, newHeight
    End If
End Sub

Private Sub CmdShowCode_Click()
    Dim i As Integer
    
    i = lstForms.listIndex
    If i <> -1 Then
        VBInstance.ActiveVBProject.VBComponents.Item(lstForms.ItemData(i)).CodeModule.CodePane.Show
    ElseIf lstForms.ListCount = 1 Then
        VBInstance.ActiveVBProject.VBComponents.Item(lstForms.ItemData(0)).CodeModule.CodePane.Show
    Else
        MsgBox "Select an item in the list", vbExclamation
    End If
End Sub

Private Sub cmdShowDesigner_Click()
    Dim i As Integer
    
    i = lstForms.listIndex
    If i <> -1 Then
        VBInstance.ActiveVBProject.VBComponents.Item(lstForms.ItemData(i)).Activate
    ElseIf lstForms.ListCount = 1 Then
        VBInstance.ActiveVBProject.VBComponents.Item(lstForms.ItemData(0)).Activate
    Else
        MsgBox "Select an item in the list", vbExclamation
    End If
End Sub

Private Sub cmdLoadCaptions_Click()
    If lstForms.ListCount > 0 Then
        LoadCaptions lstForms
    End If
End Sub

Private Sub cmdOpenExplorer_Click()
    Dim i As Integer
    
    If lstForms.ListCount = 1 Then
        ' There is only one item in the Listbox. Navigate to its file location
        i = 0
    Else
        i = lstForms.listIndex
    End If
    
    If i <> -1 Then
        With VBInstance.ActiveVBProject.VBComponents.Item(lstForms.ItemData(i))
            OpenContainingFolder VBInstance, .Name, .Type
        End With
    End If
End Sub

Private Sub cmdCopyName_Click()
    Dim i As Integer
    
    i = lstForms.listIndex
    If i <> -1 Then
        Clipboard.Clear
        Clipboard.SetText VBInstance.ActiveVBProject.VBComponents.Item(lstForms.ItemData(i)).Name
    ElseIf lstForms.ListCount = 1 Then
        Clipboard.Clear
        Clipboard.SetText VBInstance.ActiveVBProject.VBComponents.Item(lstForms.ItemData(i)).Name
    Else
        MsgBox "Select an item in the list", vbExclamation
    End If
End Sub

Private Sub cmdCopyCaption_Click()
    Dim i As Integer
    
    i = lstForms.listIndex
    If i <> -1 Then
        Clipboard.Clear
        Clipboard.SetText VBInstance.ActiveVBProject.VBComponents.Item(lstForms.ItemData(i)).Properties("Caption").Value
    ElseIf lstForms.ListCount = 1 Then
        Clipboard.Clear
        Clipboard.SetText VBInstance.ActiveVBProject.VBComponents.Item(lstForms.ItemData(i)).Properties("Caption").Value
    Else
        MsgBox "Select an item in the list", vbExclamation
    End If
End Sub

Private Sub lstForms_DblClick()
    CmdShowCode_Click
End Sub

Private Sub txtFilter_Change()
    FillList lstForms, txtFilter.Text
End Sub

Private Sub FillList(list As ListBox, sFilter As String)
    Dim i As Integer
    list.Clear
    
    If Len(Trim(sFilter)) = 0 Then
        For i = 0 To UBound(FormList)
            list.AddItem FormList(i).Text, i
            list.ItemData(i) = FormList(i).Index
            FormList(i).IsDisplayed = True
        Next
    Else
        Dim s As String
        s = UCase(sFilter)
        Dim idx As Integer
        idx = 0
        
        For i = 0 To UBound(FormList)
            If InStr(UCase(FormList(i).Text), s) > 0 Then
                list.AddItem FormList(i).Text, idx
                list.ItemData(idx) = FormList(i).Index
                FormList(i).IsDisplayed = True
                idx = idx + 1
            Else
                FormList(i).IsDisplayed = False
            End If
        Next
    End If
    
    ListCaption.Caption = "Found forms (" & list.ListCount & " items)"
End Sub

Private Sub InitializeFormList()
    Dim Components As VBComponents
    Dim component As VBComponent
    Dim i As Integer
    Dim idx As Integer
    Dim info As FormFindInfo
    
    Set Components = VBInstance.ActiveVBProject.VBComponents
    
    ReDim FormList(Components.Count)
    idx = 0
    
    For i = 1 To Components.Count
        Set component = Components.Item(i)
        If component.Type = vbext_ct_VBForm Then
            Set info = New FormFindInfo
            info.Name = component.Name
            info.Index = i
            Set FormList(idx) = info
            idx = idx + 1
        End If
    Next
    
    If idx < Components.Count Then
        ReDim Preserve FormList(idx - 1)
    End If
End Sub

Private Sub LoadCaptions(list As ListBox)
    Dim i As Integer
    Dim Components As VBComponents
    Dim component As VBComponent
    Dim info As FormFindInfo
    Dim listIndex As Integer
    Dim captionsDisplayed As Integer
    
    On Error GoTo err1
    
    Set Components = VBInstance.ActiveVBProject.VBComponents
    listIndex = 0
    captionsDisplayed = 0
    
    For i = 0 To UBound(FormList)
        Set info = FormList(i)
        If info.IsDisplayed Then
            If Not info.IsCaptionLoaded Then
                Set component = Components.Item(FormList(i).Index)
                info.Caption = component.Properties("Caption").Value
                list.list(listIndex) = FormList(i).Text
                list.ItemData(listIndex) = FormList(i).Index
            End If
            listIndex = listIndex + 1
        End If
        
        If info.IsCaptionLoaded Then
            captionsDisplayed = captionsDisplayed + 1
            If captionsDisplayed >= MAX_CAPTIONS_LOAD Then
                Dim msg As String
                msg = "Too many forms read. This can lead to IDE instability." & vbNewLine & "Do you still want to continue Caption loading?"
                If MsgBox(msg, vbYesNo) = vbYes Then
                    Exit Sub
                End If
            End If
        End If
    Next
    
    Exit Sub
    
err1:
    MsgBox Err.Description, vbExclamation
End Sub

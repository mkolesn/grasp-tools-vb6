VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9945
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   17542
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "Grasp Tools for VB6"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private FormDisplayed         As Boolean
Private VBInstance            As VBIDE.VBE

Public WithEvents CloseAllCodeHandler As CommandBarEvents
Attribute CloseAllCodeHandler.VB_VarHelpID = -1
Private CloseAllCodeCBar      As Office.CommandBarControl

Public WithEvents FindFormsHandler As CommandBarEvents
Attribute FindFormsHandler.VB_VarHelpID = -1
Private FindFormsCBar         As Office.CommandBarControl

Public WithEvents OpenStartUpObjectHandler As CommandBarEvents
Attribute OpenStartUpObjectHandler.VB_VarHelpID = -1
Private OpenStartUpObjectCBar As Office.CommandBarControl

Public WithEvents OpenMDIFormHandler As CommandBarEvents
Attribute OpenMDIFormHandler.VB_VarHelpID = -1
Private OpenMDIFormCBar As Office.CommandBarControl

Public WithEvents OpenContainingFolderHandler As CommandBarEvents
Attribute OpenContainingFolderHandler.VB_VarHelpID = -1
Private ExploreInFolderCBar   As Office.CommandBarControl

Public WithEvents ProjFormsOpenContainingFolderHandler As CommandBarEvents
Attribute ProjFormsOpenContainingFolderHandler.VB_VarHelpID = -1
Private ProjFormsExploreInFolderCBar   As Office.CommandBarControl

Public WithEvents ProjModulesOpenContainingFolderHandler As CommandBarEvents
Attribute ProjModulesOpenContainingFolderHandler.VB_VarHelpID = -1
Private ProjModulesExploreInFolderCBar   As Office.CommandBarControl

Private mfrmFormFinder        As New frmFormFinder

Sub Hide()
    On Error Resume Next
    
    FormDisplayed = False
    mfrmFormFinder.Hide
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    Set VBInstance = Application
    
    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
    Else
        Set CloseAllCodeCBar = AddToCommandBar("Close All C&ode Windows Except the Active One", "fvba_CloseAllCodePanes", "Window")
        Set Me.CloseAllCodeHandler = VBInstance.Events.CommandBarEvents(CloseAllCodeCBar)
        CloseAllCodeCBar.BeginGroup = True
        
        Set FindFormsCBar = AddToCommandBar("Find forms", "fvba_FindForm", "Tools")
        Set Me.FindFormsHandler = VBInstance.Events.CommandBarEvents(FindFormsCBar)
        
        Set OpenStartUpObjectCBar = AddToCommandBar("Open StartUp Object (If Present)", "fvba_OpenStartUpObject", "Tools")
        Set Me.OpenStartUpObjectHandler = VBInstance.Events.CommandBarEvents(OpenStartUpObjectCBar)
        
        Set OpenMDIFormCBar = AddToCommandBar("Open MDI Form (If Present)", "fvba_OpenMDIForm", "Tools")
        Set Me.OpenMDIFormHandler = VBInstance.Events.CommandBarEvents(OpenMDIFormCBar)
        
        Set ExploreInFolderCBar = AddToCommandBar("Open Containing Folder of Active CodePane", "fvba_OpenContainingFolder", "Window")
        Set Me.OpenContainingFolderHandler = VBInstance.Events.CommandBarEvents(ExploreInFolderCBar)
        
        Set ProjFormsExploreInFolderCBar = AddToCommandBar("Open Containing Folder", "fvba_ProjFormsOpenContainingFolder", "Project Window Form Folder")
        Set Me.ProjFormsOpenContainingFolderHandler = VBInstance.Events.CommandBarEvents(ProjFormsExploreInFolderCBar)
        
        Set ProjModulesExploreInFolderCBar = AddToCommandBar("Open Containing Folder", "fvba_ProjModulesOpenContainingFolder", "Project Window Module/Class Folder")
        Set Me.ProjModulesOpenContainingFolderHandler = VBInstance.Events.CommandBarEvents(ProjModulesExploreInFolderCBar)
    End If

    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    CloseAllCodeCBar.Delete
    FindFormsCBar.Delete
    OpenStartUpObjectCBar.Delete
    OpenMDIFormCBar.Delete
    ExploreInFolderCBar.Delete
    ProjFormsExploreInFolderCBar.Delete
    ProjModulesExploreInFolderCBar.Delete

    'shut down the Add-In
    If FormDisplayed Then
        FormDisplayed = False
    End If
    
    Unload mfrmFormFinder
    Set mfrmFormFinder = Nothing
End Sub

Private Sub CloseAllCodeHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    CloseAllCodeWindowsExceptActive VBInstance
End Sub

Private Sub FindFormsHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    ShowFormFinder
End Sub

Private Sub OpenStartUpObjectHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    OpenStartUpObject VBInstance.ActiveVBProject
End Sub

Private Sub OpenMDIFormHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    OpenMDIForm VBInstance.ActiveVBProject
End Sub

Private Sub OpenContainingFolderHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    ActiveCodePaneOpenContainingFolder VBInstance
End Sub

Private Sub ProjFormsOpenContainingFolderHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    If Not VBInstance.SelectedVBComponent Is Nothing Then
        With VBInstance.SelectedVBComponent
            OpenContainingFolder VBInstance, .Name, .Type
        End With
    End If
End Sub

Private Sub ProjModulesOpenContainingFolderHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    If Not VBInstance.SelectedVBComponent Is Nothing Then
        With VBInstance.SelectedVBComponent
            OpenContainingFolder VBInstance, .Name, .Type
        End With
    End If
End Sub

Private Function AddToCommandBar(sCaption As String, sTag As String, CommandBarName As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl
    Dim cbMenu As Object
  
    On Error GoTo AddToCommandBarErr
    
    'see if we can find the menu
    Set cbMenu = VBInstance.CommandBars(CommandBarName)
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add item to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(MsoControlType.msoControlButton)
    cbMenuCommandBar.Caption = sCaption
    cbMenuCommandBar.Tag = sTag
    Set AddToCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToCommandBarErr:

End Function

Private Sub ShowFormFinder()
    On Error Resume Next
    
    If mfrmFormFinder Is Nothing Then
        Set mfrmFormFinder = New frmFormFinder
    End If
    
    Set mfrmFormFinder.VBInstance = VBInstance
    Set mfrmFormFinder.Connect = Me
    FormDisplayed = True
    mfrmFormFinder.Show
End Sub


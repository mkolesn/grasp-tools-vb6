Attribute VB_Name = "Commands"
Option Explicit

Public Sub CloseAllCodeWindowsExceptActive(VBInstance As VBIDE.VBE)
    Dim ActiveCodePane As CodePane
    Set ActiveCodePane = VBInstance.ActiveCodePane
    
    If ActiveCodePane Is Nothing Then
        MsgBox "There are no open Code Panes", vbInformation
        Exit Sub
    End If
    
    Dim Wind As Window
    Dim ActiveCodePaneWindow As Window
    
    Set ActiveCodePaneWindow = ActiveCodePane.Window
    
    For Each Wind In VBInstance.Windows
        If Wind.Type = vbext_wt_CodeWindow Then
            If Not Wind Is ActiveCodePaneWindow Then
                Wind.Close
            End If
        End If
    Next
End Sub

Public Sub OpenMDIForm(Project As VBProject)
    Dim component As VBComponent
    
    For Each component In Project.VBComponents
        If component.Type = vbext_ct_VBMDIForm Then
            component.CodeModule.CodePane.Show
            Exit Sub
        End If
    Next
    
    MsgBox "An MDI Form Is not found in the Project", vbInformation
End Sub

Public Sub ExploreToPath(Path As String)
    Shell PathName:="Explorer.exe /select,""" & Path & """", WindowStyle:=vbNormalFocus
End Sub

Public Function PathExists(strPath As String) As Boolean
    PathExists = (Dir$(strPath) <> "")
End Function

Public Sub ActiveCodePaneOpenContainingFolder(VBInstance As VBIDE.VBE)
    Dim ActiveCodePane As CodePane
    Set ActiveCodePane = VBInstance.ActiveCodePane

    If Not ActiveCodePane Is Nothing Then
        With ActiveCodePane.CodeModule
            OpenContainingFolder VBInstance, .Name, .Parent.Type
        End With
    End If
End Sub

Public Sub OpenContainingFolder(VBInstance As VBIDE.VBE, ComponentName As String, ComponentType As Integer)
    Dim ProjectPath As String
    ProjectPath = VBInstance.ActiveVBProject.FileName

    Dim i As Integer

    Dim Path As String
    Dim pos As Integer
    Dim FileName As String
    Dim FileExtension As String

    Select Case ComponentType
        Case vbext_ComponentType.vbext_ct_ClassModule
            FileExtension = "cls"
        Case vbext_ComponentType.vbext_ct_PropPage
            FileExtension = "pag"
        Case vbext_ComponentType.vbext_ct_ResFile
            FileExtension = "res"
        Case vbext_ComponentType.vbext_ct_StdModule
            FileExtension = "bas"
        Case vbext_ComponentType.vbext_ct_UserControl
            FileExtension = "ctl"
        Case vbext_ComponentType.vbext_ct_VBForm
            FileExtension = "frm"
        Case Else
            FileExtension = ""
    End Select
    
    FileName = ComponentName & "." & FileExtension
    
    pos = InStrRev(ProjectPath, "\")
    ' Replace project file name with the file name of the component
    Path = Left(ProjectPath, pos) & FileName

    If PathExists(Path) Then
        ExploreToPath Path
    Else
        ' By default - navigate to the Project file location in the file system
        ExploreToPath ProjectPath
    End If
End Sub

Public Sub OpenStartUpObject(Project As VBProject)
    If IsObject(Project.VBComponents.StartUpObject) Then
        Dim StartUpForm As VBComponent
        Set StartUpForm = Project.VBComponents.StartUpObject
        StartUpForm.CodeModule.CodePane.Show
    Else
        Dim StartUpObject As Integer
        StartUpObject = Project.VBComponents.StartUpObject
        Select Case StartUpObject
            Case vbext_StartupObject.vbext_so_SubMain
                OpenStartupModule Project
            Case vbext_StartupObject.vbext_so_None
                MsgBox "The Project does not have any StartUp Object", vbInformation
        End Select
    End If
End Sub

Private Sub OpenStartupModule(Project As VBProject)
    Dim component As VBComponent
    Dim ModuleMember As Member
    
    For Each component In Project.VBComponents
        If component.Type = vbext_ct_StdModule Then
            For Each ModuleMember In component.CodeModule.Members
                If ModuleMember.Type = vbext_mt_Method And ModuleMember.Name = "Main" Then
                    component.CodeModule.CodePane.Show
                    component.CodeModule.CodePane.TopLine = ModuleMember.CodeLocation
                    Exit Sub
                End If
            Next
            
        End If
    Next
    
    MsgBox "Could not find the Main Subroutine in the Project", vbInformation
End Sub

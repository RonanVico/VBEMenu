Attribute VB_Name = "modCreateVBEMenuItems"
Option Explicit




Private MenuEvent As CVBECommandHandler
Private CmdBarItem As CommandBarControl
Public EventHandlers As New Collection

Private Const C_TAG = "MY_VBE_TAG"
Private Const C_TECNUN_BAR As String = "TECNUN"

Sub AddNewVBEControls()

Dim Ctrl As Office.CommandBarControl
Dim cmbar As Office.CommandBar

'''''''''''''''''''''''''''''''''''''''''''''''''
' Delete any existing controls with our Tag.
'''''''''''''''''''''''''''''''''''''''''''''''''
Set Ctrl = Application.VBE.CommandBars.FindControl(Tag:=C_TAG)

Do Until Ctrl Is Nothing
    Ctrl.Delete
    Set Ctrl = Application.VBE.CommandBars.FindControl(Tag:=C_TAG)
Loop

'''''''''''''''''''''''''''''''''''''''''''''''''
' Delete any existing event handlers.
'''''''''''''''''''''''''''''''''''''''''''''''''
Do Until EventHandlers.Count = 0
    EventHandlers.Remove 1
Loop

'''''''''''''''''''''''''''''''''''''''''''''''''
' add the first control to the Tools menu.
'''''''''''''''''''''''''''''''''''''''''''''''''
Set MenuEvent = New CVBECommandHandler


Set cmbar = Application.VBE.CommandBars("Barra de menus")

If cmbar Is Nothing Then
    Set cmbar = Application.VBE.CommandBars(1)
End If


If cmbar.FindControl(Tag:="TECNUN") Is Nothing Then
    With cmbar.Controls.Add(10, , , cmbar.Controls.Count + 1, False)
        .Tag = "TECNUN"
        .Caption = "TECNUN"
        .BeginGroup = True
        .Visible = True
    End With
End If

With Application.VBE.CommandBars("Barra de menus").FindControl(Tag:="TECNUN")
    Set CmdBarItem = .Controls.Add
End With
With CmdBarItem
    .Caption = "FAZ ALGO AQUI"
    .BeginGroup = True
    .OnAction = "'" & ThisWorkbook.Name & "'!Procedure_One"
    .Tag = C_TAG
End With

Set MenuEvent.EvtHandler = Application.VBE.Events.CommandBarEvents(CmdBarItem)
EventHandlers.Add MenuEvent


Set MenuEvent = New CVBECommandHandler
With Application.VBE.CommandBars("Barra de menus").FindControl(Tag:="TECNUN")
    Set CmdBarItem = .Controls.Add
End With
With CmdBarItem
    .Caption = "FAÇA ALGO AQUI"
    .BeginGroup = False
    .OnAction = "'" & ThisWorkbook.Name & "'!Procedure_Two"
    .Tag = C_TAG
End With

Set MenuEvent.EvtHandler = Application.VBE.Events.CommandBarEvents(CmdBarItem)
EventHandlers.Add MenuEvent

End Sub

Sub DeleteMenuItems()
'''''''''''''''''''''''''''''''''''''''''''''''''''''
' This procedure deletes all controls that have a
' tag of C_TAG.
'''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Ctrl As Office.CommandBarControl
    Set Ctrl = Application.VBE.CommandBars.FindControl(Tag:=C_TAG)
    Do Until Ctrl Is Nothing
        Ctrl.Delete
        Set Ctrl = Application.VBE.CommandBars.FindControl(Tag:=C_TAG)
    Loop
    
    Set Ctrl = Application.VBE.CommandBars.FindControl(Tag:=C_TECNUN_BAR)
    Do Until Ctrl Is Nothing
        Ctrl.Delete
        Set Ctrl = Application.VBE.CommandBars.FindControl(Tag:=C_TECNUN_BAR)
    Loop
End Sub





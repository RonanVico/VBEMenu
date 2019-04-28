Attribute VB_Name = "modCreateVBEMenuItems"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Feito Por: Ronan Vico
'Descrição: Este módulo possui Rotinas para criação do botão na Barra de Comandos do VBE (Visual Basic Editor)
'           é necessario toda vez que iniciar a aplicação instanciar a barra novamente ,pois ela funciona com eventos
'           Também é possivel rodar manualmente a rotina InitVBRVTools.
'Como usar?: Apenas rode InitVBRVTools e ela instanciara a barra de comando.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private MenuEvent As CVBECommandHandler
Private CmdBarItem As CommandBarButton 'CommandBarControl
Private cbBarTOOL As Office.CommandBarPopup
Public EventHandlers As New Collection
Private cmbar As Office.CommandBar

Private Const C_TAG = "MY_VBE_TAG"
Private Const C_TECNUN_BAR As String = "TECNUN"

Sub InitVBRVTools()

Dim Ctrl As Office.CommandBarControl


'''''''''''''''''''''''''''''''''''''''''''''''''
' Delete any existing controls with our Tag.
'''''''''''''''''''''''''''''''''''''''''''''''''
Set Ctrl = Application.VBE.CommandBars.FindControl(tag:=C_TAG)

Do Until Ctrl Is Nothing
    Ctrl.Delete
    Set Ctrl = Application.VBE.CommandBars.FindControl(tag:=C_TAG)
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
On Error Resume Next
Set cmbar = Application.VBE.CommandBars("Barra de menus")
If cmbar Is Nothing Then
    Set cmbar = Application.VBE.CommandBars(1)
End If
On Error GoTo 0

Set cbBarTOOL = cmbar.FindControl(tag:=C_TECNUN_BAR)
If cbBarTOOL Is Nothing Then
    With cmbar.Controls.Add(10, , , cmbar.Controls.Count + 1, False)
        .tag = "TECNUN"
        .CAption = "&TECNUN_RV_TOOLS"
        .BeginGroup = True
        .Visible = True
    End With
End If

Set cbBarTOOL = cmbar.FindControl(tag:=C_TECNUN_BAR)

Call AddMenuButton("Inserir &Cabeçalho", True, "inserirCabeçalhoNaProcedure", 12)
Call AddMenuButton("Inserir &Error Treatment", False, "inserirTratamentoDeErro", 464)
Call AddMenuButton("Identar &Variaveis", True, "IdentaVariaveis", 123)
Call AddMenuButton("Desbloquear All VBE's", True, "Hook", 650)
Call AddMenuButton("About Creator", True, "aboutme", 111)

End Sub

Sub AddMenuButton(ByVal CAption As String, BeginGroup As Boolean, OnACtion As String, FaceId As Long)

    With cbBarTOOL
        Set CmdBarItem = .Controls.Add
        With CmdBarItem
            '.Type = 1
            .FaceId = FaceId
            .CAption = CAption
            .BeginGroup = BeginGroup
            '.OnAction = "'" & ThisWorkbook.Name & "'!Procedure_One"
            .OnACtion = OnACtion
            .tag = C_TAG
            .TooltipText = "(Ctrl+8)"
        End With
    End With
    
    Set MenuEvent = New CVBECommandHandler
    Set MenuEvent.EvtHandler = Application.VBE.Events.CommandBarEvents(CmdBarItem)
    EventHandlers.Add MenuEvent
End Sub


Sub DeleteMenuItems()
'''''''''''''''''''''''''''''''''''''''''''''''''''''
' This procedure deletes all controls that have a
' tag of C_TAG.
'''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Ctrl As Office.CommandBarControl
    Set Ctrl = Application.VBE.CommandBars.FindControl(tag:=C_TAG)
    Do Until Ctrl Is Nothing
        Ctrl.Delete
        Set Ctrl = Application.VBE.CommandBars.FindControl(tag:=C_TAG)
    Loop
    
    Set Ctrl = Application.VBE.CommandBars.FindControl(tag:=C_TECNUN_BAR)
    Do Until Ctrl Is Nothing
        Ctrl.Delete
        Set Ctrl = Application.VBE.CommandBars.FindControl(tag:=C_TECNUN_BAR)
    Loop
End Sub







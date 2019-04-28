Attribute VB_Name = "zModExemplo"
Option Explicit

Public Sub MinhaSubZoada1()
    Dim a As Long, b As String: Dim c, e() As String, f()
    
 End Sub
 

Public Sub MinhaSubZoada3()
    Dim a As String 'Variavel do capeta
    Dim b   As Long ' Variavel de deus
    Dim heineken  As Double: Dim skol  As Single ' se beber dirija e bata num poste
    Dim x As Object
 
 
     a = b + x
     a = 123
     x = 1
     For b = 1 To 10
         x = a + b
     Next b
TratarErro:
         Select Case Err.Number
                 Case 0
                 Case Else
                         MsgBox Err.Description & " " & Err.Number, vbCritical
         End Select
 End Sub

Public Sub tst()
    Debug.Print VBA.Now()
End Sub

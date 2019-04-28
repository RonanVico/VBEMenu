Attribute VB_Name = "modRVTools"
Option Explicit
'--------------------------------------------------------------------------------------------------------------
'Criado por: Ronan Raphael Vico // ronanvico@hotmail.com // https://br.linkedin.com/in/ronan-vico
'   Descrição:Módulo utilizado para programar o IDE VBE, facilitando criação de rotinas e manutenção de códigos.
'       as rotinas serão utilizadas em botões programaveis na barra de comandos do VBE dentro do EXCEL
'
'
'
' Códigos utilizados como ajuda e fornecedores
' 1- GetProcedureDeclaration  e ProcedureInfo - Made By CPearson
' 2- IsLinhaMatch - made By TECNUN - www.tecnun.com.br
'----------------------------------------------------------------------------------------------------------------------------
  Public Enum ProcScope
     ScopePrivate = 1
 
    ScopePublic = 2
    ScopeFriend = 3
    ScopeDefault = 4
End Enum
 
Public Enum LineSplits
    LineSplitRemove = 0
    LineSplitKeep = 1
    LineSplitConvert = 2
End Enum
 
Public Type ProcInfo
    ProcName                As String
    procKind                As VBIDE.vbext_ProcKind
    ProcStartLine           As Long
    ProcBodyLine            As Long
    ProcCountLines          As Long
    ProcScope               As ProcScope
    ProcDeclaration         As String
End Type
 

 
Private pInfo           As ProcInfo
       
       
        
'-------- PARAMETROS
'-------- ABAIXO ESTÃO OS PARAMETROS UTILIZADOS
Private Const PARAM_PROGRAMADOR              As String = "RONAN VICO"
Private Const PARAM_EMPRESA                  As String = "Tecnun Tecnologia em Informática"
Private Const PARAM_PROGRAMADOR_MAIL         As String = "RONANVICO@hotmail.com"
Private Const PARAM_CHAR_IDENTAÇÃO           As String = vbTab '"  "
Private Const PARAM_ERROR_HANDLER_NAME       As String = "TratarErro"
Private Const PARAM_TABULACAO_VARIAVEIS      As Long = 20
 
 
Private Const QUEBRA_DE_LINHA               As String = "_VBNEWLINE!"
Private Const tagVarInit                    As String = "[V@_"
Private Const tagVarEnd                     As String = "@]"
 
Private Property Get PARAM_HEADER_DEFAULT() As String
pInfo = ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
PARAM_HEADER_DEFAULT = "'---------------------------------------------------------------------------------------" & _
    vbNewLine & "' Modulo....: " & Application.VBE.ActiveCodePane.CodeModule & " \ " & VBA.TypeName(Application.VBE.ActiveCodePane.CodeModule) & _
    vbNewLine & "' Rotina....: " & pInfo.ProcDeclaration & _
    vbNewLine & "' Autor.....: " & PARAM_PROGRAMADOR & _
    vbNewLine & "' Contato...: " & PARAM_PROGRAMADOR_MAIL & _
    vbNewLine & "' Data......: " & VBA.CStr(VBA.Date) & _
    vbNewLine & "' Empresa...: " & PARAM_EMPRESA & _
    vbNewLine & "' Descrição.: " & _
    vbNewLine & "'---------------------------------------------------------------------------------------"
End Property
 
 
 
Private Property Get PARAM_ERROR_HANDLER_DEFAULT() As String
    PARAM_ERROR_HANDLER_DEFAULT = PARAM_ERROR_HANDLER_NAME & ":" & vbNewLine _
        & VBA.String(2, PARAM_CHAR_IDENTAÇÃO) & "select case err.number " & vbNewLine _
            & VBA.String(4, PARAM_CHAR_IDENTAÇÃO) & "case 0 " & vbNewLine _
            & VBA.String(4, PARAM_CHAR_IDENTAÇÃO) & "case else " & vbNewLine _
                & VBA.String(6, PARAM_CHAR_IDENTAÇÃO) & "msgbox err.description  & "" "" & Err.number , vbCritical " & vbNewLine _
        & VBA.String(2, PARAM_CHAR_IDENTAÇÃO) & "end Select"
End Property
'/\--------PARAMETROS------/\------------/\------------/\--------------/\-------------/\

Public Sub addFromGuidVBEPRoject()
    On Error Resume Next
    'Muda o registro do windows pra liberar acesso ao VBE ;) WE ARE HACKERS
    Call ChangeRegistry_AccessVBOM
    'Adiciona biblioteca do VBE
    Call Application.VBE.ActiveVBProject.References.AddFromGuid("{0002E157-0000-0000-C000-000000000046}", 2, 0)
End Sub
 
'--------PARTE 1 DAS PROPRIEDADES -------------------
Private Property Get ActiveProcedure() As String
    ActiveProcedure = Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(ActiveStartCodeLine, pInfo.procKind)
End Property
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
 
'-------------parte 2 das propriedades ------------------------------------
'--- Properties para pegar as linhas e colunas ja selecionadas no codemodule
Private Property Get ActiveStartCodeLine() As Long
    Application.VBE.ActiveCodePane.GetSelection ActiveStartCodeLine, 0, 0, 0
End Property
Private Property Get ActiveStartCodeColumn() As Long
    Application.VBE.ActiveCodePane.GetSelection 0, ActiveStartCodeColumn, 0, 0
End Property
Private Property Get ActiveEndCodeLine() As Long
    Application.VBE.ActiveCodePane.GetSelection 0, 0, ActiveEndCodeLine, 0
End Property
Private Property Get ActiveEndCodeColumn() As Long
    Application.VBE.ActiveCodePane.GetSelection 0, 0, 0, ActiveEndCodeColumn
End Property
'/\------------/\------------/\--------------/\-------------/\---------------------
 
 
'-------------Parte 3 Funções que irão mudar o mundo do VBA ------------------------------------
Public Sub inserirTratamentoDeErro()
    Dim nlinha          As Long
    Dim sLinha          As String
    Dim sSplit
   
    On Error GoTo t
    'Verifica se está numa procedure
    If ActiveProcedure = "" Then Exit Sub
 
    pInfo = ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
    
    'Insere na primeira linha o on error e na ultima o texto padrão
    For nlinha = pInfo.ProcStartLine + pInfo.ProcCountLines - 1 To pInfo.ProcStartLine + 2 Step -1
        sLinha = Application.VBE.ActiveCodePane.CodeModule.Lines(nlinha, 1)
        For Each sSplit In VBA.Split(sLinha, ":")
            sSplit = VBA.Split(sSplit, "'")(0)
           If IsLinhaMatch(sSplit, "(End (Function|Sub|Property))") Then
                Call Application.VBE.ActiveCodePane.CodeModule.InsertLines _
                                (nlinha, _
                             PARAM_ERROR_HANDLER_DEFAULT)
                Call Application.VBE.ActiveCodePane.CodeModule.InsertLines _
                     (pInfo.ProcBodyLine + 1, _
                    "on error goto " & PARAM_ERROR_HANDLER_NAME)
                Exit Sub
            End If
        Next sSplit
    Next nlinha
   
    
    Exit Sub
    Resume
t:
    End
End Sub
 
 
 
'-------------Parte 3 Funções que irão mudar o mundo do VBA ------------------------------------
'---------------------------------------------------------------------------------------
' Modulo....: vbProjcet / Módulo
' Rotina....: Public Sub inserirCabeçalhoNaProcedure()
' Autor.....: RONAN VICO
' Contato...: RONANVICO@hotmail.com
' Data......: 23/04/2019
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Insere cabeçalho na procedure (Utilizei ela mesmo para criar esse cabeçalho aqui ;)
'---------------------------------------------------------------------------------------
Public Sub inserirCabeçalhoNaProcedure()
    Dim nlinha          As Long
    Dim sLinha          As String
    Dim sSplit
    On Error GoTo t
    'Verifica se está numa procedure
    If ActiveProcedure = "" Then Exit Sub
    Debug.Print ActiveProcedure
    pInfo = ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
    Call Application.VBE.ActiveCodePane.CodeModule.InsertLines _
                    (pInfo.ProcBodyLine, _
                    PARAM_HEADER_DEFAULT)
 

    Exit Sub
    Resume
t:
    End
End Sub
 
 
 
Public Sub teste()
'    Dim ActiveLine As Long
'    Dim Procedure As String
'    'Get the line who the cursor in VBE is selected
'     Application.VBE.ActiveCodePane.GetSelection ActiveLine, 0, 0, 0
'    'With the number of the line we can search for the proc selected
'    If ActiveLine <= Application.VBE.ActiveCodeModule.CountOfDeclarationLines Then
'        MsgBox "Você esta não esta em uma procedure e sim nas declarações ", vbInformation, "hey boy"
'    Else
'        Procedure = Application.VBE.ActiveCodeModule.ProcOfLine(ActiveLine, VBIDE.vbext_ProcKind.vbext_pk_Proc)
'        MsgBox "Você esta na procedure -> " & Procedure, vbInformation, "hey boy"
'    End If
End Sub
 
 
Public Sub Identar_Modulo(Optional modulo)
    On Error GoTo TratarErro
    Dim md As VBIDE.CodeModule
    If VBA.IsError(modulo) Then
        Set md = Application.VBE.ActiveCodePane.CodeModule
    ElseIf TypeOf modulo Is VBIDE.CodeModule Then
        Set md = modulo
    ElseIf VBA.VarType(modulo) = vbString Then
   
    End If
    Exit Sub
TratarErro:
End Sub
 
'---------------------------------------------------------------------------------------
' Modulo....: Publicas / Módulo
' Rotina....: IsLinhaMatch() / Function
' Autor.....: Jefferson
' Contato...: jefferson@tecnun.com.br
' Data......: 09/11/2012
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para verificar se uma linha Corresponde a um padrao, usando regex
'---------------------------------------------------------------------------------------
 
Public Function IsLinhaMatch(ByVal linha As String, ParamArray Padroes() As Variant) As Boolean
 Dim resultado           As Boolean
 Dim Contador            As Byte
 Dim regExp              As Object
 On Error GoTo TratarErro
                                          'New VBScript_RegExp_55.RegExp
    If regExp Is Nothing Then Set regExp = VBA.CreateObject("VBScript.RegExp")
    With regExp
        'Padroes = TFW_AuxArray.Acertar_Array_Parametros(Padroes)
        For Contador = 0 To UBound(Padroes) Step 1
            If Not Padroes(Contador) = VBA.vbNullString Then
                .Pattern = Padroes(Contador)
                If .test(linha) Then
                    resultado = True
                    Exit For
                End If
            End If
        Next Contador
    End With
    IsLinhaMatch = resultado
 Exit Function
TratarErro:
    'Call TFW_Excecoes.tratarerro(VBA.Err.Description, VBA.Err.Number, "TFW_AuxRegex.IsLinhaMatch()", Erl)
 End Function
 

 
 
 
Public Sub ChangeRegistry_AccessVBOM()
    'Made by Ronan Vico
    'helped by Rabaquim
    'helpde by Fernando
    'MADE AT TECNUN www.tecnun.com.br
    Dim shl
    Dim key As String
    key = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Excel\Security\AccessVBOM"
    Set shl = CreateObject("WScript.Shell")
 
     Debug.Print shl.regRead(key)
     Call shl.regWrite(key, 1, "REG_DWORD")
End Sub
 
 
 
 
Public Sub IdentaVariaveis()
    'Cara , não vou explicar essa função porque nem eu sei oq eu fiz , _
                essa função segue as leis de software internacional e de Deus , pois só ele _
                deve saber como isso funciona
                
            
    Dim Proc                    As ProcInfo
    Dim linha                   As String
    Dim linhaFormatada          As String
    Dim linhas
    Dim contLine                As Long
    Dim textoProc               As String
    Dim ArrVars
    Dim TAntComment             As String
    Dim TPosComment             As String
    Dim contSubLines            As Long
    Dim subLines                As Variant
    Dim subLine                 As String
    Dim DimLines
    Dim contDimLines            As Long
    Dim DimLine
    Dim ProcBodyStart           As Long
    Dim NovoTexto               As String
    Dim contVar                 As Long
    
    
    textoProc = PegarProcedureSemQuebraDeLinha(ActiveProcedure)
    
    linhas = VBA.Split(textoProc, vbNewLine)
   
    For contLine = 0 To UBound(linhas)
        linha = linhas(contLine)
        linhaFormatada = formataTexto(linha, ArrVars)
        If PosComentario(linha) <> 0 Then
            TAntComment = VBA.Left(linhaFormatada, VBA.InStr(linhaFormatada, "'") - 1)
            TPosComment = VBA.Mid(linhaFormatada, VBA.InStr(linhaFormatada, "'"))
        Else
            TAntComment = linhaFormatada
            TPosComment = ""
        End If
        
        subLines = VBA.Split(TAntComment, ":")
        
        For contSubLines = 0 To UBound(subLines)
            subLine = VBA.Trim(subLines(contSubLines))
            
            If VBA.Left((subLine), 4) = "Dim " Then
                subLine = VBA.Replace(subLine, ",", vbNewLine & " Dim ")
                subLine = SingleSpace(subLine)
                DimLines = VBA.Split(subLine, vbNewLine)
                For contDimLines = 0 To UBound(DimLines)
                    DimLine = DimLines(contDimLines)
                    'Colocando identação das variaveis
                    If VBA.InStr(DimLine, " As ") = 0 Then
                        DimLine = DimLine & VBA.Strings.Space$(PARAM_TABULACAO_VARIAVEIS - (VBA.Len(DimLine) - 4)) & "As Variant"
                    ElseIf contDimLines = 0 Then
                        DimLine = VBA.Left(DimLine, VBA.InStr(DimLine, " As ") - 1) & VBA.Strings.Space$(PARAM_TABULACAO_VARIAVEIS - (VBA.InStr(DimLine, " As ") - 5)) & VBA.Mid(DimLine, VBA.InStr(DimLine, " As ") + 1)
                    Else
                        DimLine = VBA.Left(DimLine, VBA.InStr(DimLine, " As ") - 1) & VBA.Strings.Space$(PARAM_TABULACAO_VARIAVEIS + 1 - (VBA.InStr(DimLine, " As ") - 5)) & VBA.Mid(DimLine, VBA.InStr(DimLine, " As ") + 1)
                    End If
                    DimLines(contDimLines) = DimLine
                Next
                'Stop
                'If(UBound(DimLines) > 0, vbNewLine & " ", "") &
                subLines(contSubLines) = VBA.Join(DimLines, vbNewLine)
            End If
            'Stop
        Next contSubLines
        'Stop
        

        NovoTexto = VBA.Join(subLines, ":") & TPosComment & vbNewLine

        If VBA.Join(ArrVars) <> "" Then
            For contVar = 0 To UBound(ArrVars)
                NovoTexto = VBA.Replace(NovoTexto, tagVarInit & contVar & tagVarEnd, ArrVars(contVar))
            Next
        End If
        
        'É necessario retirar um espaço em branco do canto esquerdo pois ele é inserido sem qerer
        NovoTexto = VBA.Replace(NovoTexto, QUEBRA_DE_LINHA, " _" & vbNewLine)
        If VBA.Left(NovoTexto, 2) = vbNewLine And VBA.Len(NovoTexto) <> 2 Then
            NovoTexto = VBA.Mid(NovoTexto, 3)
        End If
        If VBA.Left(NovoTexto, 1) = " " Then
            NovoTexto = VBA.Mid(NovoTexto, 2)
        End If
        linhas(contLine) = NovoTexto
    Next contLine
    
    If linhas(UBound(linhas)) = "" Then
         ReDim Preserve linhas(LBound(linhas) To UBound(linhas) - 1)
    End If
    
    With pInfo
        ProcBodyStart = .ProcBodyLine
        Call Application.VBE.ActiveCodePane.CodeModule.DeleteLines(ProcBodyStart, .ProcCountLines - (ProcBodyStart - .ProcStartLine))
        Call Application.VBE.ActiveCodePane.CodeModule.InsertLines(ProcBodyStart, VBA.Join(linhas))
    End With
End Sub
 
Public Sub asdeasd()
    Dim a, b            As Long: a = "My:String": Dim c As Integer, d, e, f, G _
            , s 'COMMENTARIO BACANNA!""'""
'asd
   
End Sub


 
Public Function PegarProcedureSemQuebraDeLinha(ProcedureName As String) As String
    'Se tiver _ no final da linha de código , significa que a proxima linha pertence ao mesmo comando do VBA _
               Exemplo este comentario , a linha de baixo esta comentada de vido a esse Underline localizado /\ aqui _
               Logo , Iremos transformar tudo em "1 linha" para conseguirmos rodar funções de formatação de texto _
               e posteriormente plotar de volta as qebras de llinhas no seus devidos lugares !
 
 
 Dim Proc                As ProcInfo
 Dim textAntComment      As String
 Dim textComment         As String
 Dim line                As String
 Dim texto               As String
 Dim i                   As Long
 
 
    Proc = ProcedureInfo(ProcedureName, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
    With Proc
        i = .ProcBodyLine
        While i <= (.ProcCountLines + .ProcStartLine) - 1
            line = VBA.RTrim(Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1))
            While VBA.Right(line, 1) = "_"
                i = i + 1
                line = VBA.Left(line, VBA.Len(line) - 1) & QUEBRA_DE_LINHA & VBA.RTrim(Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1))
            Wend
            'textAntComment = VBA.Mid(line, 1, PosComentario(line) - 1)
            'textComment = VBA.Mid(line, PosComentario(line))
            texto = texto & line & vbNewLine
            i = i + 1
        Wend
        'Debug.Print Application.VBE.ActiveCodePane.CodeModule.Lines(.ProcBodyLine + 1, .ProcCountLines - 1)
    End With
    PegarProcedureSemQuebraDeLinha = texto
 End Function
 

 
 
Public Function formataTexto(ByVal TextoOriginal As String, Optional ByRef MyArrVar) As String
    'Transforma todas as strings em Variaveis dentro de um array , assim podemos manipular _
            todo o texto sem medo de estar mechendo com dados dentro de string , por exemplo _
            uma string pode ser MyString = "OLA : ' "  , Logo os caracteres ":" e "'" , são _
            importantes na nossa formatação de texto ,e não podemos consideralo na formatação, _
            por isso monto um array para posteriormente após a formatação pegar as strings e jogar de volta.
            
On Error GoTo TratarErro

    Dim ArrVars()
    Dim contVar                 As Long
    Dim i                       As Long
    Dim y                       As Long
    Dim isString                As Boolean
    Dim c                       As String
    Dim Var                     As String
    Dim LenMax                  As Long
    Dim PosEndQuote             As Long
    Dim texto                   As String
    Dim tag                     As String
    Dim ValorVariavel           As String
   
    texto = TextoOriginal
   
    LenMax = VBA.Len(texto)
    'Debug.Print Texto
    While (i <= LenMax)
        i = i + 1
        c = VBA.Mid$(texto, i, 1)
        'Verifica se é um caracter Aspas Dupla "
        If c = VBA.Chr(34) Then
            'Percorre as proximas letras até que a string seja fechada MyString= "abc" <Fim da string
            For y = i + 1 To VBA.Len(texto)
                If VBA.Mid$(texto, y, 1) = VBA.Chr(34) Then
                    If VBA.Mid$(texto, y + 1, 1) = VBA.Chr(34) Then
                        y = y + 1
                    Else
                        Exit For
                    End If
                End If
            Next y
           'Adcionando a variavel nova ao vetor
            ReDim Preserve ArrVars(0 To contVar)
            ValorVariavel = VBA.Mid(texto, i + 1, (y) - (i + 1))
            'Aspas duplas de comentario estavam sendo apagadass
            ArrVars(contVar) = VBA.Chr(34) & ValorVariavel & VBA.Chr(34)
            tag = tagVarInit & contVar & tagVarEnd
            texto = VBA.Mid(texto, 1, i - 1) & tag & VBA.Mid(texto, y + 1, VBA.Len(texto))
            contVar = contVar + 1
            i = i + VBA.Len(tag) - 1
            LenMax = VBA.Len(texto)
            'contVar = contVar + 1
            'VBA.Mid(TEXTO,I+1,Y-1) <- string
            'VBA.Mid(TEXTO,1,I) & VBA.Mid(TEXTO,Y,VBA.Len(TEXTO))
        End If
       
    Wend
   
    MyArrVar = ArrVars
    formataTexto = texto
    Exit Function
    Stop
   
TratarErro:
        Stop
        Resume
        Select Case Err.Number
                Case 0
                Case Else
                        MsgBox Err.Description & " " & Err.Number, vbCritical
        End Select
End Function
 
 
 
 
Public Function PosComentario(ByVal text As String) As Long
    Dim auxtexto        As String
    auxtexto = formataTexto(text)
    PosComentario = (VBA.InStr(1, auxtexto, "'", vbTextCompare))
End Function
 
 
'Public Function PossuiComentario(ByVal text As String) As Boolean
'    Dim auxtexto        As String
'    auxtexto = formataTexto(text)
'    PossuiComentario = (VBA.InStr(1, auxtexto, "'", vbTextCompare) <> 0)
'End Function
 
 
 
'- --------------------------------------------\/  --------------------------------------------\/'- --------------------------------------------\/  --------------------------------------------\/
'PARTE 4 --------------------------------------------\/ FUNCTIONS UTILIZADAS PARA PEGAR OS DADOS DO VBE
'- --------------------------------------------\/  --------------------------------------------\/'- --------------------------------------------\/  --------------------------------------------\/
 Public Function ProcedureInfo(ProcName As String, CodeMod As VBIDE.CodeModule, procKind As VBIDE.vbext_ProcKind) As ProcInfo
    Dim BodyLine As Long
    Dim Declaration As String
    Dim FirstLine As String
   
    BodyLine = CodeMod.ProcStartLine(ProcName, pInfo.procKind)
    If BodyLine > 0 Then
        With CodeMod
            pInfo.ProcName = ProcName
            pInfo.procKind = procKind
            pInfo.ProcBodyLine = .ProcBodyLine(ProcName, procKind)
            pInfo.ProcCountLines = .ProcCountLines(ProcName, procKind)
            pInfo.ProcStartLine = .ProcStartLine(ProcName, procKind)
           
            FirstLine = .Lines(pInfo.ProcBodyLine, 1)
            If VBA.Strings.StrComp(VBA.Left(FirstLine, VBA.Len("Public")), "Public", vbBinaryCompare) = 0 Then
                pInfo.ProcScope = ScopePublic
            ElseIf VBA.Strings.StrComp(VBA.Left(FirstLine, VBA.Len("Private")), "Private", vbBinaryCompare) = 0 Then
                pInfo.ProcScope = ScopePrivate
            ElseIf VBA.Strings.StrComp(VBA.Left(FirstLine, VBA.Len("Friend")), "Friend", vbBinaryCompare) = 0 Then
                pInfo.ProcScope = ScopeFriend
            Else
                pInfo.ProcScope = ScopeDefault
            End If
            pInfo.ProcDeclaration = GetProcedureDeclaration(CodeMod, ProcName, LineSplitKeep)
        End With
    End If
   
    ProcedureInfo = pInfo
 
End Function
 
 
Public Function GetProcedureDeclaration(CodeMod As VBIDE.CodeModule, _
    ProcName As String, _
    Optional LineSplitBehavior As LineSplits = LineSplitRemove)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetProcedureDeclaration
' This return the procedure declaration of ProcName in CodeMod. The LineSplitBehavior
' determines what to do with procedure declaration that span more than one line using
' the "_" line continuation character. If LineSplitBehavior is LineSplitRemove, the
' entire procedure declaration is converted to a single line of text. If
' LineSplitBehavior is LineSplitKeep the "_" characters are retained and the
' declaration is split with vbNewLine into multiple lines. If LineSplitBehavior is
' LineSplitConvert, the "_" characters are removed and replaced with vbNewLine.
' The function returns vbNullString if the procedure could not be found.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim LineNum As Long
    Dim s As String
    Dim Declaration As String
   
    On Error Resume Next
    LineNum = CodeMod.ProcBodyLine(ProcName, pInfo.procKind)
    If Err.Number <> 0 Then
        Exit Function
    End If
    s = CodeMod.Lines(LineNum, 1)
    Do While VBA.Right(s, 1) = "_"
        Select Case True
            Case LineSplitBehavior = LineSplitConvert
                s = VBA.Left(s, VBA.Len(s) - 1) & vbNewLine
            Case LineSplitBehavior = LineSplitKeep
                s = s & vbNewLine
            Case LineSplitBehavior = LineSplitRemove
                s = VBA.Left(s, VBA.Len(s) - 1) & " "
        End Select
        Declaration = Declaration & s
        LineNum = LineNum + 1
        s = CodeMod.Lines(LineNum, 1)
    Loop
    Declaration = SingleSpace(Declaration & s)
    GetProcedureDeclaration = Declaration
   
 
End Function
 
Private Function SingleSpace(ByVal text As String) As String
    Dim Pos As String
    Pos = InStr(1, text, VBA.Space(2), vbBinaryCompare)
    Do Until Pos = 0
        text = VBA.Replace(text, VBA.Space(2), VBA.Space(1))
        Pos = InStr(1, text, VBA.Space(2), vbBinaryCompare)
    Loop
    SingleSpace = VBA.Trim$(text)
End Function
 
Sub ShowProcedureInfo()
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim CompName As String
    Dim ProcName As String
    Dim procKind As VBIDE.vbext_ProcKind
    Dim pInfo As ProcInfo
   
    CompName = "md_VBACrack"
    ProcName = "Hook"
    procKind = vbext_pk_Proc
   
    Set VBProj = ActiveWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(CompName)
    Set CodeMod = VBComp.CodeModule
   
    pInfo = ProcedureInfo(ProcName, CodeMod, pInfo.procKind)
   
    Debug.Print "ProcName: " & pInfo.ProcName
    Debug.Print "ProcKind: " & CStr(pInfo.procKind)
    Debug.Print "ProcStartLine: " & CStr(pInfo.ProcStartLine)
    Debug.Print "ProcBodyLine: " & CStr(pInfo.ProcBodyLine)
    Debug.Print "ProcCountLines: " & CStr(pInfo.ProcCountLines)
    Debug.Print "ProcScope: " & CStr(pInfo.ProcScope)
    Debug.Print "ProcDeclaration: " & pInfo.ProcDeclaration
End Sub
 
 
Public Function formataTexto2(ByVal TextoOriginal As String, Optional ByRef MyArrVar) As String
'Transforma todas as strings em Variaveis dentro de um array , assim podemos manipular _
            todo o texto sem medo de estar mechendo com dados dentro de string , por exemplo _
            uma string pode ser MyString = "OLA : ' "  , Logo os caracteres ":" e "'" , são _
            importantes na nossa formatação de texto ,e não podemos consideralo na formatação, _
            por isso monto um array para posteriormente após a formatação pegar as strings e jogar de volta.
 
 On Error GoTo TratarErro
 
 Dim ArrVars()           As Variant: Dim contVar            As Long: Dim i                  As Long:
 Dim y                   As Long
 Dim isString            As Boolean
 Dim c                   As String
 Dim Var                 As String
 Dim LenMax              As Long
 Dim PosEndQuote         As Long
 Dim texto               As String
 Dim tag                 As String
 Dim ValorVariavel       As String
 
     texto = TextoOriginal
 
     LenMax = VBA.Len(texto)
     'Debug.Print Texto
     While (i <= LenMax)
         i = i + 1
         c = VBA.Mid$(texto, i, 1)
         'Verifica se é um caracter Aspas Dupla ""
         If c = VBA.Chr(34) Then
             'Percorre as proximas letras até que a string seja fechada MyString= "abc" <Fim da string
             For y = i + 1 To VBA.Len(texto)
                 If VBA.Mid$(texto, y, 1) = VBA.Chr(34) Then
                     If VBA.Mid$(texto, y + 1, 1) = VBA.Chr(34) Then
                         y = y + 1
                     Else
                         Exit For
                     End If
                 End If
             Next y
            'Adcionando a variavel nova ao vetor
             ReDim Preserve ArrVars(0 To contVar)
             ValorVariavel = VBA.Mid(texto, i + 1, (y) - (i + 1))
             'Aspas duplas de comentario estavam sendo apagadass
             ArrVars(contVar) = VBA.Chr(34) & ValorVariavel & VBA.Chr(34)
             tag = tagVarInit & contVar & tagVarEnd
             texto = VBA.Mid(texto, 1, i - 1) & tag & VBA.Mid(texto, y + 1, VBA.Len(texto))
             contVar = contVar + 1
             i = i + VBA.Len(tag) - 1
             LenMax = VBA.Len(texto)
             'contVar = contVar + 1
             'VBA.Mid(TEXTO,I+1,Y-1) <- string
             'VBA.Mid(TEXTO,1,I) & VBA.Mid(TEXTO,Y,VBA.Len(TEXTO))
         End If
 
     Wend
 
     MyArrVar = ArrVars
     formataTexto2 = texto
     Exit Function
     Stop
 
TratarErro:
         Stop
         Resume
         Select Case Err.Number
                 Case 0
                 Case Else
                         MsgBox Err.Description & " " & Err.Number, vbCritical
         End Select
 End Function
 
 
Private Function AboutMe()
    MsgBox "Feito por Ronan Vico", vbInformation, "TECNUN"
End Function

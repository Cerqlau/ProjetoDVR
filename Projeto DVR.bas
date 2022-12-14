Attribute VB_Name = "Module1"

Function captura_caminho(ByVal caminho_inicial As String, ByVal tipo As String) As String
Dim frase As String
Dim posicao As Integer

frase = caminho_inicial
posicao = InStr(frase, ".")

nome = Left(frase, posicao - 1)
extensao = Right(frase, Len(frase) - posicao + 1)

captura_caminho = nome & "-" & tipo & extensao
End Function

'Seleciona os arquivos
Public Sub lsSelecionaArquivo()
    Dim Caminho As String
    Dim NomeBase As String
    Dim OS As String
    Dim Data As String
    Dim Hinicial As String
    Dim NomeFim As String

    
    Caminho = InputBox("Informe o local dos arquivos a serem renomeados:", "Pasta", "C:\Users\ss7v30631\Documents\VIDEOS")
    NomeBase = InputBox("Informe o nome do relat?rio:", "Nome do relat?rio", "AAMLS20")
    NomeBase = UCase(NomeBase)
    OS = InputBox("Informe a OS de servi?o:", "Ordem de servi?o", "OS123654")
    Data = InputBox("Digite a data inicial de grava??o do v?deo", "Data Inicial", "2022/05/02")
    Hinicial = CDate(InputBox("Digite a hora inicial de grava??o do v?deo", "Hora Inicial", "23:00:00"))
    NomeFim = InputBox("Informe o tipo de c?mera", "Tipo de C?mera", "BW ou COLOR")
    
    'Chama a fun??o para renomear os arquivos
    lsRenomearArquivos Caminho, NomeBase, Hinicial, OS, Data, NomeFim
End Sub
 
'Fun??o que renomea os arquivos
Public Sub lsRenomearArquivos(Caminho As String, NomeBase As String, Hinicial As String, OS As String, Data As String, NomeFim As String)
 
    Dim FSO As Object, Pasta As Object, Arquivo As Object, Arquivos As Object
    Dim Linha As Long
    Dim Data_troca As Integer
    Dim Hora As String
    Dim Min As String
    Dim Seg As String
    Dim hifen As String
    Dim hifenf As String
    Dim lNovoNome As String
    Dim Hcalc As Variant
    Dim cont As Integer
    Dim Data_dia As String
    Dim Data_mes As String
    Dim Data_ano As String
    Dim Path() As String
    Dim extensao() As String
    Dim arrDates() As Date
    Dim arrDates2() As Date
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim strTmp As String
    Dim extTmp As String
    Dim dtmTmp As Date
    Dim dtmTmp2 As Date
    Dim Check As Boolean
    Dim FimNome As String
    

    Set FSO = CreateObject("Scripting.FileSystemObject")

    If Not FSO.FolderExists(Caminho) Then
        MsgBox "A pasta '" & Caminho & "' n?o existe.", vbCritical, "Erro"
        Exit Sub
    End If
    
    Hcalc = TimeValue(Hinicial)

    'Bloco de compara??o de data e hora'
    Hora = Hour(Hcalc)
    If CInt(Hora) < 10 Then
        Hora = "0" + Hora
    End If
    Min = Minute(Hcalc)
    If CInt(Min) < 10 Then
        Min = "0" + Min
    End If
    Seg = Second(Hcalc)
    If CInt(Seg) < 10 Then
        Seg = "0" + Seg
    End If
    hifen = "-"
    hifenf = "-D030"
    FimNome = NomeFim
    Data_dia = Day(CDate(Data))
    If CInt(Data_dia) < 10 Then
        Data_dia = "0" + Data_dia
    End If
    Data_mes = Month(CDate(Data))
    If CInt(Data_mes) < 10 Then
        Data_mes = "0" + Data_mes
    End If
    Data_ano = Year(CDate(Date))

    'Itera??o com os arquivos da pasta'

    Set Pasta = FSO.GetFolder(Caminho)
    Set Arquivos = Pasta.Files
    
    
    ' Redimensionamento dos Arrays
    n = Pasta.Files.Count
    ReDim Path(1 To n)
    ReDim extensao(1 To n)
    ReDim arrDates(1 To n)
    ReDim arrDates2(1 To n)
    
    Dim frase As String
    Dim posicao As Integer
    

    ' Preenchimentos dos arrays
    For Each Arquivo In Arquivos
        i = i + 1
        Path(i) = Arquivo.Path
        arrDates(i) = Arquivo.DateLastModified
        arrDates2(i) = Arquivo.DateLastModified
        frase = Arquivo.Path
        posicao = InStr(frase, ".")
        extensao(i) = Right(frase, Len(frase) - posicao + 1)
    Next Arquivo

    ' Organizar caminho pela ultima data de modifica??o
    For i = 1 To n - 1
        For j = i + 1 To n
            If arrDates(i) > arrDates(j) Then
                dtmTmp = arrDates(i)
                arrDates(i) = arrDates(j)
                arrDates(j) = dtmTmp
                strTmp = Path(i)
                Path(i) = Path(j)
                Path(j) = strTmp
            End If
        Next j
    Next i
    
    
    ' Organizar extens?o pela ultima data de modifica??o
    For i = 1 To n - 1
        For j = i + 1 To n
            If arrDates2(i) > arrDates2(j) Then
                dtmTmp = arrDates2(i)
                arrDates2(i) = arrDates2(j)
                arrDates2(j) = dtmTmp
                strTmp = extensao(i)
                extensao(i) = extensao(j)
                extensao(j) = strTmp
            End If
        Next j
    Next i
     
       
    'calculo para verifica??o de quantidades de steps necess?rios para terminar um dia'
    Data_troca = (DateDiff("n", Hcalc, "23:59:00") + 1) / 30
    
    'zeragem de vari?veis contadoras'
    cont = 0
    Linha = 2
    Check = False
    i = 0

    For i = 1 To n

        'Cells(Linha, 1) = UCase$(Arquivo.Path)'
        lNovoNome = Caminho & "\" & NomeBase & hifen & OS & hifen & Data_ano & hifen & Data_mes & hifen & Data_dia & hifen & "T" & hifen & Hora & hifen & Min & hifen & Seg & hifenf & hifen & FimNome & extensao(i)
        Name FSO.GetFile(Path(i)) As lNovoNome

        'Cells(Linha, 2) = lNovoNome'
        Hcalc = CDate(Hcalc) + CDate("00:30:00") 'soma de mais 30min'
        
        'Bloco de compara??o de data e hora'
        Hora = Hour(Hcalc)
        If CInt(Hora) < 10 Then
            Hora = "0" + Hora
        End If
        Min = Minute(Hcalc)
        If CInt(Min) < 10 Then
            Min = "0" + Min
        End If
        Seg = Second(Hcalc)
        If CInt(Seg) < 10 Then
            Seg = "0" + Seg
        End If
        hifen = "-"
        hifenf = "-D030"
        Data_dia = Day(CDate(Data))
        If CInt(Data_dia) < 10 Then
            Data_dia = "0" + Data_dia
        End If
        Data_mes = Month(CDate(Data))
        If CInt(Data_mes) < 10 Then
            Data_mes = "0" + Data_mes
        End If
        Data_ano = Year(CDate(Date))
        
        
        'verifica??o de data'
        If Check = False Then
            If cont >= Data_troca - 2 Then
            Data = DateAdd("d", 1, Data)
            cont = 0
            Check = True
            End If
        End If
        
        If Check = True Then
            If cont >= 48 Then
            Data = DateAdd("d", 1, Data)
            cont = 0
            End If
        End If
                
        'contadores'
        Linha = Linha + 1
        cont = cont + 1
        

    Next i
End Sub
'Seleciona caminhos para copy
Public Sub lsSelecionaArquivo_Fonte2()
    Dim Caminho As String, NomeBase As String, tipo As String
    
    Caminho = InputBox("Informe o local dos arquivos a serem renomeados:", "Pasta destino para arquivos", "C:\Users\ss7v30631\Documents\VIDEOS")
    NomeBase = InputBox("Informe o local dos arquivos de onde os nomes ser?o copiados:", "Pasta fonte de arquivos", "C:\Users\ss7v30631\Documents\VIDEOS2")
    tipo = InputBox("Informe a sigla para o nome", "Tipo de renomeio", "BW ou COLOR")
    
    'Caminho = "D:\lauro\Engenharia\Programa??o\VBA Projetos\Projeto DVR\BW"
    'NomeBase = "D:\lauro\Engenharia\Programa??o\VBA Projetos\Projeto DVR\COLOR"
    'tipo = "BW"
    
    'Chama a fun??o para renomear os Arquivo_Fontes
    lsRenomearArquivo_Fontes2 Caminho, NomeBase, tipo
End Sub

'Fun??o para copiar nome de Arquivo_Fontes de acordo com o nome
Public Sub lsRenomearArquivo_Fontes2(Caminho1 As String, Caminho2 As String, tipo As String)

    Dim FSO_Fonte As Object, Pasta_Fonte As Object, objeto_Fonte As Object, Arquivo_Fonte As Object
    Dim Path_Fonte() As String, Extensao_Fonte() As String, Nomes_Fonte() As String
    Dim arrDates_Fonte() As Date, arrDates2_Fonte() As Date
    
    Dim FSO_Fonte2 As Object, Pasta_Fonte2 As Object, objeto_Fonte2 As Object, Arquivo_Fonte2 As Object
    Dim FSO_Fonte3 As Object, Pasta_Fonte3 As Object, objeto_Fonte3 As Object, Arquivo_Fonte3 As Object
    Dim Path_Fonte2() As String, Extensao_Fonte2() As String, Nomes_Fonte2() As String
    Dim arrDates_Fonte2() As Date, arrDates2_Fonte2() As Date
    
    Dim n As Integer, Table_Fonte() As Integer, n2 As Integer, Table_Fonte2() As Integer
    Dim item As Variant, item2 As Variant, nome As Variant
    Dim novonome As String, nomenovocompleto As String
    
Begin:
        
    Set FSO_Fonte = CreateObject("Scripting.FileSystemObject")
    Set FSO_Fonte2 = CreateObject("Scripting.FileSystemObject")
    
    If Not FSO_Fonte.FolderExists(Caminho2) Then
        MsgBox "A pasta '" & Caminho2 & "' n?o existe.", vbCritical, "Erro"
        Exit Sub
    End If
    
    If Not FSO_Fonte.FolderExists(Caminho1) Then
        MsgBox "A pasta '" & Caminho1 & "' n?o existe.", vbCritical, "Erro"
        Exit Sub
    End If
    
    'Itera??o com os Arquivo_Fontes da Pasta_Fonte'

    Set Pasta_Fonte = FSO_Fonte.GetFolder(Caminho2)
    Set Arquivo_Fonte = Pasta_Fonte.Files
    
    Set Pasta_Fonte2 = FSO_Fonte2.GetFolder(Caminho1)
    Set Arquivo_Fonte2 = Pasta_Fonte2.Files
    
    
    Save = Caminho1 & "\" & "r"
    If Not FSO_Fonte2.FolderExists(Save) Then
            FSO_Fonte2.CreateFolder (Save)
    End If
    
    ' Redimensionamento dos Arrays
    n = Pasta_Fonte.Files.Count
    n2 = Pasta_Fonte2.Files.Count
    If n2 = 0 Then GoTo Ending 'Sa?da para fim do programa
    
    ReDim Table_Fonte(1 To n)
    ReDim Path_Fonte(1 To n)
    ReDim Extensao_Fonte(1 To n)
    ReDim arrDates_Fonte(1 To n)
    ReDim arrDates2_Fonte(1 To n)
    ReDim Nomes_Fonte(1 To n)
    
    ReDim Table_Fonte2(1 To n2)
    ReDim Path_Fonte2(1 To n2)
    ReDim Extensao_Fonte2(1 To n2)
    ReDim arrDates_Fonte2(1 To n2)
    ReDim arrDates2_Fonte2(1 To n2)
    ReDim Nomes_Fonte2(1 To n2)
    
    
    
    ' Preenchimentos dos arrays
    For Each objeto_Fonte In Arquivo_Fonte
        i = i + 1
        Path_Fonte(i) = objeto_Fonte.Path
        arrDates_Fonte(i) = objeto_Fonte.DateLastModified
        arrDates2_Fonte(i) = objeto_Fonte.DateLastModified
        Extensao_Fonte(i) = Right(objeto_Fonte, 4)
        Nomes_Fonte(i) = objeto_Fonte.Name
    Next
    
    For Each objeto_Fonte2 In Arquivo_Fonte2
        x = x + 1
        Path_Fonte2(x) = objeto_Fonte2.Path
        arrDates_Fonte2(x) = objeto_Fonte2.DateLastModified
        arrDates2_Fonte2(x) = objeto_Fonte2.DateLastModified
        Extensao_Fonte2(x) = Right(objeto_Fonte2, 4)
        Nomes_Fonte2(x) = objeto_Fonte2.Name
    Next
    
    
    'Loop de verifica??o para extens?o e data de ultima modifica??o
    
   
    For Each objeto_Fonte In Arquivo_Fonte
        cont = cont + 1
        For Each objeto_Fonte2 In Arquivo_Fonte2
            cont2 = cont2 + 1
            If arrDates_Fonte(cont) = arrDates_Fonte2(cont2) And Extensao_Fonte(cont) = Extensao_Fonte2(cont2) Then
            nome = FSO_Fonte2.GetFile(Path_Fonte2(cont2))
            novonome = Save & "\" & Nomes_Fonte(cont)
            Name FSO_Fonte2.GetFile(Path_Fonte2(cont2)) As novonome
            
            End If
        Next objeto_Fonte2
        cont2 = 0
    Next objeto_Fonte

    i = 0
    x = 0
    cont = 0
    cont2 = 0
    
    GoTo Begin
    
    
Ending:

    'Copiar arquivos renomeados da pasta tempor?ria para pasta principal
    
    Set FSO_Fonte3 = CreateObject("Scripting.FileSystemObject")
    If Not FSO_Fonte.FolderExists(Save) Then
        MsgBox "A pasta '" & Caminho2 & "' n?o existe.", vbCritical, "Erro"
        Exit Sub
    End If
    
    Set Pasta_Fonte3 = FSO_Fonte3.GetFolder(Save)
    Set Arquivo_Fonte3 = Pasta_Fonte3.Files
   
    
    For Each objeto_Fonte3 In Arquivo_Fonte3
        cont3 = cont3 + 1
        nomeantigo = Caminho1 & "\" & Nomes_Fonte(cont3)
        nomenovocompleto = captura_caminho(nomeantigo, tipo)
        Name FSO_Fonte3.GetFile(objeto_Fonte3.Path) As nomenovocompleto
    Next objeto_Fonte3
    
    'deletar pasta r
    FSO_Fonte3.DeleteFolder (Save)
    
    End Sub


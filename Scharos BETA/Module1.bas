Attribute VB_Name = "Module1"
Global CPF, StrConf, strCampo As String
'''''''''''''''Variaveis para conectar a base Scharos''''''''''
Global conectar As New ADODB.Connection
Global caminho As String
Global capsula As String
''''''''''''''Variaveis para conectar a base Ceps'''''''''''''''''''''
Global conectar2 As New ADODB.Connection
Global caminho2 As String
Global capsula2 As String
''''''''''''''Variaveis para conectar a base Configuracoes'''''''''''''''''''''
Global con As New ADODB.Connection
Global cam As String
Global cap As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Global status As String
Global img As String
'''''''''''''variaveis para manipular as configurações'''''''''
Global numero As String


Option Explicit
Function abrir_banco()
            If conectar.State = 1 Then conectar.Close
            capsula = "Provider=microsoft.jet.oledb.4.0;data source="
            caminho = capsula + App.Path & "\Dados\Scharos.mdb"
            conectar.Open (caminho)
End Function
Function abrir_banco2()
            If conectar2.State = 1 Then conectar2.Close
            capsula2 = "Provider=microsoft.jet.oledb.4.0;data source="
            caminho2 = capsula2 + App.Path & "\Dados\Ceps.mdb"
            conectar2.Open (caminho2)
End Function
Function configu()
            If con.State = 1 Then con.Close
            cap = "Provider=microsoft.jet.oledb.4.0;data source="
            cam = cap + App.Path & "\Dados\Configuracoes.mdb"
            con.Open (cam)
End Function
Function box()
            MsgBox "As informações foram " & status & " com sucesso", vbInformation, "Scharos BETA"
End Function
Public Function CalcularIdade(DTNasc As Date) As String
            Dim Anos As Single, Meses As String, Dias As Single
            Dim UTDTNasc As Date
                If Month(DTNasc) <= Month(Date) Then
                    If Month(DTNasc) <> Month(Date) Then
                        UTDTNasc = Day(DTNasc) & "/" & Month(DTNasc) & "/" & Year(Format(Date, "dd/mm/yyyy"))
                Else
                If Day(DTNasc) <= Day(Date) Then
                    UTDTNasc = Day(DTNasc) & "/" & Month(DTNasc) & "/" & Year(Format(Date, "dd/mm/yyyy"))
                Else
                    GoTo NPassou
                End If
                    End If
                Else
NPassou:
                    UTDTNasc = Day(DTNasc) & "/" & Month(DTNasc) & "/" & Year(Format(Date, "dd/mm/yyyy")) - 1
                End If
                    Anos = DateDiff("yyyy", DTNasc, UTDTNasc)
                    Meses = DateDiff("m", UTDTNasc, Date)
                If Day(Date) < Day(UTDTNasc) Then
                Meses = Meses - 1
      Dias = DateDiff("d", DateAdd("m", -1, Day(DTNasc) & "/" & Month(Date) & "/" & Year(Format(Date, "dd/mm/yyyy"))), Date)
   ElseIf Day(Date) = Day(UTDTNasc) Then
      Dias = 0
   ElseIf Day(Date) > Day(UTDTNasc) Then
      Dias = DateDiff("d", Day(DTNasc) & "/" & Month(Date) & "/" & Year(Format(Date, "dd/mm/yyyy")), Date)
   End If
   CalcularIdade = Anos & " Ano(s) " & Meses & " Mês(es) " & Dias & " Dia(s)"
   
    
End Function
Function CalculaCPF()
         Dim I As Integer
         Dim strCaracter As String
         Dim intNumero As Integer
         Dim intMais As Integer
         Dim lngSoma As Long
         Dim dblDivisao As Double
         Dim lngInteiro As Long
         Dim intResto As Integer
         Dim intDig1 As Integer
         Dim intDig2 As Integer

         lngSoma = 0
         intNumero = 0
         intMais = 0
         
         'Inicia cálculos do 1º dígito
         For I = 2 To 10
             strCaracter = Right(strCampo, I - 1)
             intNumero = Left(strCaracter, 1)
             intMais = intNumero * I
             lngSoma = lngSoma + intMais
        Next I
        dblDivisao = lngSoma / 11

        lngInteiro = Int(dblDivisao) * 11
        intResto = lngSoma - lngInteiro
        If intResto = 0 Or intResto = 1 Then
           intDig1 = 0
        Else
           intDig1 = 11 - intResto
        End If

        strCampo = strCampo & intDig1
        lngSoma = 0
        intNumero = 0
        intMais = 0

        'Inicia cálculos do 2º dígito
        For I = 2 To 11
            strCaracter = Right(strCampo, I - 1)
            intNumero = Left(strCaracter, 1)
            intMais = intNumero * I
            lngSoma = lngSoma + intMais
        Next I
        dblDivisao = lngSoma / 11
        lngInteiro = Int(dblDivisao) * 11
        intResto = lngSoma - lngInteiro
        If intResto = 0 Or intResto = 1 Then
           intDig2 = 0
        Else
           intDig2 = 11 - intResto
        End If
        StrConf = intDig1 & intDig2
End Function
Function par_impar()
            If (numero And 1) = 0 Then
                MsgBox "Será par"
            Else
                MsgBox "Será Impar"
End If
End Function
Function carrega_imagem()
On Error Resume Next
            
End Function
Function imagem()
            
            img = App.Path & "\Imagens\Padrão.jpg"
            Call carrega_imagem
End Function
Function desculpa()
            MsgBox "desculpe, para melhor atendê-lo, estamos em processo de melhoria", vbInformation, "Scharos BETA"
End Function

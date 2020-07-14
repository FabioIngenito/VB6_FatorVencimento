VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFatorVencimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mDtHOJE As Date

'http://boletonet.codeplex.com/sourcecontrol/list/changesets?ViewAll=true&ProjectName...
'87353 By Stiven
'Ajuste Fator Vencimento com base na proposta feita pela CENEGESC de _
   acordo com o comunicado FEBRABAN de n° 082/2012 de 14/06/2012 segue _
   regra para implantação. No dia 21/02/20025 o fator vencimento chegará _
   em 9999 assim atigindo o tempo de utilização, para contornar esse _
   problema foi definido com uma nova regra de utilizaçao criando um _
   range de uso o range funcionara controlando a emissão dos boletos.

Public Function CalculaFatorDeVencimento(strFatorDeVencimento As String) As String
    CalculaFatorDeVencimento = DateDiff("d", "07/10/1997", strFatorDeVencimento)
    If Len(CalculaFatorDeVencimento) < 3 Then CalculaFatorDeVencimento = String$(4 - Len(Trim(CalculaFatorDeVencimento)), "0") & Trim(CalculaFatorDeVencimento)
    
    'Ajuste para nova legislação
    If CalculaFatorDeVencimento > 9999 Then
        CalculaFatorDeVencimento = CalculaFatorDeVencimento - 9000
        'Se mesmo assim estourar '9999' ... ZERA!
        If CalculaFatorDeVencimento > 9999 Then CalculaFatorDeVencimento = 0

    End If

    'Para a próxima 'Gambiware' substitua o "IF" acima pelo while abaixo:
    'Exemplo: Coloque a data de: "13/10/2490"
'    Do While CalculaFatorDeVencimento > 9999
'        CalculaFatorDeVencimento = CalculaFatorDeVencimento - 9000
'    Loop

End Function


Property Let DtHoje(Value As Date)
    mDtHOJE = Value
End Property



Public Function fuFatorVencimento(Fator As String) As String
Dim vltDataCalcVelho As Date
Dim vltDataCalcNovo As Date

Dim strFatorHoje As String
Dim strFatorVencido As String
Dim strFatorAVencer As String
Dim strFatorTransicao As String
Dim dtaVencida As Date
Dim dtaAVencer As Date
Dim dtaHOJE As Date
Dim teste1 As String
Dim teste2 As String
Dim strDataCalculo As String

    fuFatorVencimento = ""
        
    dtaHOJE = Date
    dtaHOJE = mDtHOJE
    'dtaHOJE = "12/03/2014"
    'dtaHOJE = "13/03/2014"
    'dtaHOJE = "14/03/2014"
    'dtaHOJE = "15/03/2014"
    'dtaHOJE = "16/03/2014"
    'dtaHOJE = "25/07/2015"
    'dtaHOJE = "23/02/2035"

    If Val(Left(Fator, 1)) > 0 Then
        '28/02/2014 -
        dtaVencida = DateAdd("d", -3000, dtaHOJE)
        dtaAVencer = DateAdd("d", 5500, dtaHOJE)
        
        strFatorHoje = CalculaFatorDeVencimento(Conversion.CStr(dtaHOJE))
        strFatorVencido = strFatorHoje - 3000
        strFatorAVencer = strFatorHoje + 5500
        
        If strFatorVencido < 0 Then strFatorVencido = strFatorVencido + 9000
        If strFatorAVencer > 9999 Then strFatorAVencer = strFatorAVencer - 9000
        
        If strFatorVencido > strFatorAVencer Then
            strFatorTransicao = strFatorVencido
            strFatorVencido = strFatorAVencer
            strFatorAVencer = strFatorTransicao
        End If
        
        vltDataCalcVelho = CDate("03/07/2000") + (Val(Fator) - 1000)
        vltDataCalcNovo = CDate("22/02/2025") + (Val(Fator) - 1000)
        
        If dtaVencida <= vltDataCalcVelho And dtaAVencer >= vltDataCalcVelho Then
            fuFatorVencimento = vltDataCalcVelho
        End If

        If dtaVencida <= vltDataCalcNovo And dtaAVencer >= vltDataCalcNovo Then
            fuFatorVencimento = vltDataCalcNovo
        End If
        
        If fuFatorVencimento <> "" Then fuFatorVencimento = Format(fuFatorVencimento, "DDMMYYYY")

        'Eliminar esta faixa: Range de Segurança nessa data.
        'Exemplo: em "12/03/2014" -> fator 2501 ao fator 2999 (499 fatores)
        If Fator > strFatorVencido And Fator < strFatorAVencer Then
        'If Fator > 2502 And Fator < 3000 Then
            fuFatorVencimento = ""
        End If
    
    End If

End Function


'Public Function fuFatorVencimentoTESTE2(Fator As String) As String
'Dim vltDataCalcVelho As Date
'Dim vltDataCalcNovo As Date
'
'Dim strFatorHoje As String
'Dim strFatorVencido As String
'Dim dtaVencida As Date
'Dim strFatorAVencer As String
'Dim dtaAVencer As Date
'Dim dtaHOJE As Date
'Dim teste1 As String
'Dim teste2 As String
'Dim strDataCalculo As String
'
'
'    If Val(Left(Fator, 1)) = 0 Then
'        fuFatorVencimento = ""
'    Else
'        dtaHOJE = Date
'        dtaHOJE = "12/03/2014"
'        'dtaHOJE = "13/03/2014"
'        'dtaHOJE = "14/03/2014"
'        'dtaHOJE = "23/02/2025"
'
''        strFatorHoje = CalculaFatorDeVencimento(Conversion.CStr(dtaHOJE))
''        strFatorVencido = strFatorHoje - 3000
''        strFatorAVencer = strFatorHoje + 5500
'
'
'        vltDataCalcVelho = CDate("03/07/2000") + (Val(Fator) - 1000)
'        vltDataCalcNovo = CDate("22/02/2025") + (Val(Fator) - 1000)
'
'        '28/02/2014 -
'        dtaVencida = DateAdd("d", -3000, dtaHOJE)
'        dtaAVencer = DateAdd("d", 5500, dtaHOJE)
'        fuFatorVencimento = ""
'
'        If strFatorVencido > Fator Then
'            fuFatorVencimento = vltDataCalcNovo
'        Else
'
'            If strFatorAVencer > Fator Then
'                fuFatorVencimento = vltDataCalcVelho
'            End If
'
'        End If
'
'
'        If dtaHOJE < "22/02/2025" Then
'
'            If dtaVencida <= vltDataCalcVelho And dtaAVencer >= vltDataCalcVelho Then
'                fuFatorVencimento = vltDataCalcVelho
'            End If
'
'        Else
'
'            If dtaVencida <= vltDataCalcNovo And dtaAVencer >= vltDataCalcNovo Then
'                fuFatorVencimento = vltDataCalcNovo
'            End If
'
'        End If
'
'    End If
'
'End Function
'
'
'Public Function fuFatorVencimentoTESTE(Fator As String) As String
'Dim vltDataCalc As Date
'Dim strFatorHoje As String
'Dim strFatorVencido As String
'Dim dtaVencida As Date
'Dim strFatorAVencer As String
'Dim dtaAVencer As Date
'Dim dtaHOJE As Date
'Dim teste1 As String
'Dim teste2 As String
'Dim strDataCalculo As String
'
'    If Val(Left(Fator, 1)) = 0 Then
'        fuFatorVencimento = ""
'    Else
'
'        dtaHOJE = Date
'        'dtaHOJE = "23/02/2025"
'
'        If dtaHOJE <= "22/02/2025" Then
'            strDataCalculo = "03/07/2000"
'        Else
'            strDataCalculo = "22/02/2025"
'        End If
'
'        '28/02/2014 -
'        strFatorHoje = CalculaFatorDeVencimento(Conversion.CStr(dtaHOJE))
'        strFatorVencido = strFatorHoje - 3000
'        strFatorAVencer = strFatorHoje + 5500
'
'
'        If strFatorVencido < 0 Then strFatorVencido = 0
'        If strFatorAVencer > 9999 Then strFatorAVencer = strFatorAVencer - 9000
'
'        dtaVencida = DateAdd("d", -3000, dtaHOJE)
'        dtaAVencer = DateAdd("d", 5500, dtaHOJE)
'
'        teste1 = CalculaFatorDeVencimento(Conversion.CStr(dtaVencida))
'        teste2 = CalculaFatorDeVencimento(Conversion.CStr(dtaAVencer))
'
'        vltDataCalc = CDate(strDataCalculo) + (Val(Fator) - 1000)
'
'        'If vltDataCalc < Date Then
'        If dtaVencida > vltDataCalc Then
'            fuFatorVencimento = ""
'        ElseIf dtaAVencer < vltDataCalc Then
'            fuFatorVencimento = Format(vltDataCalc, "DDMMYYYY")
'        Else
'            fuFatorVencimento = Format(vltDataCalc, "DDMMYYYY")
'        End If
'
'
'    End If
'
'End Function



'FUNÇÃO ORIGINAL:

'Public Function fuFatorVencimentoVELHO(Fator As String) As String
'Dim vltDataCalc As Date
'
'    If Val(Left(Fator, 1)) = 0 Then
'        fuFatorVencimento = ""
'    Else
'        vltDataCalc = CDate("03/07/2000") + (Val(Fator) - 1000)
'        If vltDataCalc < Date Then
'            fuFatorVencimento = ""
'        Else
'            fuFatorVencimento = Format(vltDataCalc, "DDMMYYYY")
'        End If
'    End If
'
'End Function

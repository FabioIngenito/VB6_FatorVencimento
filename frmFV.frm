VERSION 5.00
Begin VB.Form frmFV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fator de Vencimento = Alta Tecnolgia Gambiare Brasileira!"
   ClientHeight    =   6975
   ClientLeft      =   540
   ClientTop       =   1725
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   5325
   Begin VB.TextBox txtDataHoje 
      Height          =   315
      Left            =   120
      TabIndex        =   72
      Text            =   "12/03/2014"
      Top             =   6540
      Width           =   1035
   End
   Begin VB.CommandButton cmdLimparFatorVencimento 
      Caption         =   "&Limpar Fator de Vencimento"
      Height          =   315
      Left            =   1680
      TabIndex        =   71
      Top             =   6600
      Width           =   2235
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar Data"
      Height          =   315
      Left            =   4200
      TabIndex        =   70
      Top             =   6600
      Width           =   1035
   End
   Begin VB.CommandButton cmdDataDeVencimento 
      Caption         =   "&2º) Dt.Valida"
      Height          =   315
      Left            =   4200
      TabIndex        =   69
      Top             =   6240
      Width           =   1035
   End
   Begin VB.TextBox txtFVDigitadoDataFun1 
      Height          =   315
      Left            =   4200
      TabIndex        =   53
      Top             =   5760
      Width           =   1035
   End
   Begin VB.TextBox txtFVDigitadoData 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Text            =   "13/10/2049"
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdFatorDeVencimento 
      Caption         =   "&1º) Fator de Vencimento"
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   6240
      Width           =   2235
   End
   Begin VB.Label lblDataHoje 
      Caption         =   "Data Hoje:"
      Height          =   255
      Left            =   120
      TabIndex        =   73
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label lblFV1000DataFun1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4260
      TabIndex        =   68
      Top             =   780
      Width           =   915
   End
   Begin VB.Label lblFV0001DataFun1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4260
      TabIndex        =   67
      Top             =   420
      Width           =   915
   End
   Begin VB.Label lblFV9999DataFun1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4260
      TabIndex        =   66
      Top             =   5100
      Width           =   915
   End
   Begin VB.Label lblFV2500DataFun1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4260
      TabIndex        =   65
      Top             =   1140
      Width           =   915
   End
   Begin VB.Label lblFV2501DataFun1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4260
      TabIndex        =   64
      Top             =   1500
      Width           =   915
   End
   Begin VB.Label lblFV2502DataFun1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4260
      TabIndex        =   63
      Top             =   1860
      Width           =   915
   End
   Begin VB.Label lblFV0000DataFun1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4260
      TabIndex        =   62
      Top             =   5460
      Width           =   915
   End
   Begin VB.Label lblFV3002DataFun1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4260
      TabIndex        =   61
      Top             =   3660
      Width           =   915
   End
   Begin VB.Label lblFV3001DataFun1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4260
      TabIndex        =   60
      Top             =   3300
      Width           =   915
   End
   Begin VB.Label lblFV3000DataFun1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4260
      TabIndex        =   59
      Top             =   2940
      Width           =   915
   End
   Begin VB.Label lblFV2503DataFun1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4260
      TabIndex        =   58
      Top             =   2220
      Width           =   915
   End
   Begin VB.Label lblFV2999DataFun1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4260
      TabIndex        =   57
      Top             =   2580
      Width           =   915
   End
   Begin VB.Label lblFV6002DataFun1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4260
      TabIndex        =   56
      Top             =   4740
      Width           =   915
   End
   Begin VB.Label lblFV6001DataFun1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4260
      TabIndex        =   55
      Top             =   4380
      Width           =   915
   End
   Begin VB.Label lblFV6000DataFun1 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4260
      TabIndex        =   54
      Top             =   4020
      Width           =   915
   End
   Begin VB.Label lblDataVencimento1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data Fun1"
      Height          =   195
      Left            =   4260
      TabIndex        =   52
      Top             =   120
      Width           =   915
   End
   Begin VB.Line Line1 
      X1              =   4080
      X2              =   4080
      Y1              =   60
      Y2              =   6540
   End
   Begin VB.Label lblFV 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fator de Vencimento"
      Height          =   195
      Left            =   120
      TabIndex        =   51
      Top             =   120
      Width           =   2235
   End
   Begin VB.Label lblFun1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fun1"
      Height          =   195
      Left            =   3480
      TabIndex        =   50
      Top             =   120
      Width           =   435
   End
   Begin VB.Label lblData 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data"
      Height          =   195
      Left            =   2460
      TabIndex        =   49
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lblFVDigitadoResposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   48
      Top             =   5820
      Width           =   435
   End
   Begin VB.Label lblFV0000Resposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   47
      Top             =   5460
      Width           =   435
   End
   Begin VB.Label lblFV9999Resposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   46
      Top             =   5100
      Width           =   435
   End
   Begin VB.Label lblFV6002Resposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   45
      Top             =   4740
      Width           =   435
   End
   Begin VB.Label lblFV6001Resposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   44
      Top             =   4380
      Width           =   435
   End
   Begin VB.Label lblFV6000Resposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   43
      Top             =   4020
      Width           =   435
   End
   Begin VB.Label lblFV3002Resposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   42
      Top             =   3660
      Width           =   435
   End
   Begin VB.Label lblFV3001Resposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   41
      Top             =   3300
      Width           =   435
   End
   Begin VB.Label lblFV3000Resposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   40
      Top             =   2940
      Width           =   435
   End
   Begin VB.Label lblFV2999Resposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   39
      Top             =   2580
      Width           =   435
   End
   Begin VB.Label lblFV2503Resposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   38
      Top             =   2220
      Width           =   435
   End
   Begin VB.Label lblFV2502Resposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   37
      Top             =   1860
      Width           =   435
   End
   Begin VB.Label lblFV2501Resposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   36
      Top             =   1500
      Width           =   435
   End
   Begin VB.Label lblFV2500Resposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   35
      Top             =   1140
      Width           =   435
   End
   Begin VB.Label lblFV1000Resposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   34
      Top             =   780
      Width           =   435
   End
   Begin VB.Label lblFV0001Resposta1 
      Alignment       =   2  'Center
      Caption         =   "0000"
      Height          =   255
      Left            =   3480
      TabIndex        =   33
      Top             =   420
      Width           =   435
   End
   Begin VB.Label lblFV6000Data 
      Caption         =   "12/03/2014"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2460
      TabIndex        =   32
      Top             =   4020
      Width           =   915
   End
   Begin VB.Label lblFV6001Data 
      Caption         =   "13/03/2014"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2460
      TabIndex        =   31
      Top             =   4380
      Width           =   915
   End
   Begin VB.Label lblFV6002Data 
      Caption         =   "14/03/2014"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2460
      TabIndex        =   30
      Top             =   4740
      Width           =   915
   End
   Begin VB.Label lblFV2999Data 
      Caption         =   "23/12/2005"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2460
      TabIndex        =   29
      ToolTipText     =   "Mil noites de gambiarra!"
      Top             =   2580
      Width           =   915
   End
   Begin VB.Label lblFV2503Data 
      Caption         =   "05/04/2029"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2460
      TabIndex        =   28
      ToolTipText     =   "Primeiro dia de gambiarra! Viva a tecnologia Gambiware!"
      Top             =   2220
      Width           =   915
   End
   Begin VB.Label lblFV3000Data 
      Caption         =   "24/12/2005"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2460
      TabIndex        =   27
      Top             =   2940
      Width           =   915
   End
   Begin VB.Label lblFV3001Data 
      Caption         =   "25/12/2005"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2460
      TabIndex        =   26
      Top             =   3300
      Width           =   915
   End
   Begin VB.Label lblFV3002Data 
      Caption         =   "26/12/2005"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2460
      TabIndex        =   25
      Top             =   3660
      Width           =   915
   End
   Begin VB.Label lblFV0000Data 
      Caption         =   "22/02/2025"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2460
      TabIndex        =   24
      Top             =   5460
      Width           =   915
   End
   Begin VB.Label lblFV0000Label 
      Caption         =   "Fator de Vencimento ""-1"":"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   5460
      Width           =   2235
   End
   Begin VB.Label lblFV6001Label 
      Caption         =   "Fator de Vencimento 6001:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4380
      Width           =   2235
   End
   Begin VB.Label lblFV6002Label 
      Caption         =   "Fator de Vencimento 6002:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4740
      Width           =   2235
   End
   Begin VB.Label lblFV3002Label 
      Caption         =   "Fator de Vencimento 3002:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3660
      Width           =   2235
   End
   Begin VB.Label lblFV3001Label 
      Caption         =   "Fator de Vencimento 3001:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3300
      Width           =   2235
   End
   Begin VB.Label lblFV3000Label 
      Caption         =   "Fator de Vencimento 3000:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2940
      Width           =   2235
   End
   Begin VB.Label lblFV2501Label 
      Caption         =   "Fator de Vencimento 2501:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1500
      Width           =   2235
   End
   Begin VB.Label lblFV2999Label 
      Caption         =   "Fator de Vencimento 2999:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2580
      Width           =   2235
   End
   Begin VB.Label lblFV2503Label 
      Caption         =   "Fator de Vencimento 2503:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2220
      Width           =   2235
   End
   Begin VB.Label lblFV2502Label 
      Caption         =   "Fator de Vencimento 2502:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1860
      Width           =   2235
   End
   Begin VB.Label lblFV2500Label 
      Caption         =   "Fator de Vencimento 2500:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1140
      Width           =   2235
   End
   Begin VB.Label lblFV2502Data 
      Caption         =   "04/04/2029"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2460
      TabIndex        =   12
      Top             =   1860
      Width           =   915
   End
   Begin VB.Label lblFV2501Data 
      Caption         =   "03/04/2029"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2460
      TabIndex        =   11
      Top             =   1500
      Width           =   915
   End
   Begin VB.Label lblFV2500Data 
      Caption         =   "02/04/2029"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2460
      TabIndex        =   10
      Top             =   1140
      Width           =   915
   End
   Begin VB.Label lblFV6000Label 
      Caption         =   "Fator de Vencimento 6000:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4020
      Width           =   2235
   End
   Begin VB.Label txtFVDigitadoLabel 
      Caption         =   "Digite um Fator de Vencimento:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5820
      Width           =   2235
   End
   Begin VB.Label lblFV9999Label 
      Caption         =   "Último Fator de Vencimento:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5100
      Width           =   2235
   End
   Begin VB.Label lblFV1000Label 
      Caption         =   "Fator de Vencimento 1000:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "O VERDADEIRO 1º FATOR DE VENCIMENTO!"
      Top             =   780
      Width           =   2235
   End
   Begin VB.Label lblFV0001Label 
      Caption         =   "Primeiro Fator de Vencimento:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   420
      Width           =   2235
   End
   Begin VB.Label lblFV9999Data 
      Caption         =   "21/02/2025"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2460
      TabIndex        =   4
      ToolTipText     =   "Boletos Futuros Financiados à Perder de Vista (Exemplo: minha casa, meu barraco )"
      Top             =   5100
      Width           =   915
   End
   Begin VB.Label lblFV0001Data 
      Caption         =   "07/10/1997"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2460
      TabIndex        =   3
      ToolTipText     =   "Viva a tecnologia Gambiware!"
      Top             =   420
      Width           =   915
   End
   Begin VB.Label lblFV1000Data 
      Caption         =   "03/07/2000"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2460
      TabIndex        =   2
      ToolTipText     =   "O VERDADEIRO 1º FATOR DE VENCIMENTO!"
      Top             =   780
      Width           =   915
   End
End
Attribute VB_Name = "frmFV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clsFV As New clsFatorVencimento

'Explicações:
'-------------------------------------------------------------------------------------------------------------
'1     Informamos que o comunicado Febraban 082/2012 alterou a forma de utilização _
   do campo "FATOR DE VENCIMENTO" constante nos códigos de barras dos boletos de _
   cobrança bancária, sendo estes utilizados arrecadados pelo SIACC no compromisso _
   de pagamento à forncedores.

'2     No referido comunicado foi informado que o fator de vencimento no dia _
   21/02/2025 atingirá o valor limite de "9999" e, a partir do dia 22/02/2025, _
   este número deverá reiniciar para "1000(mil fatores)".

'2.1   Foi definido também que os bancos deveriam para parametrizar os seus _
   sistemas de forma a não permitir o recebimento de boletos fora dos limites _
   definidos de fator de vencimento.

'2.1.1  Assim, inicialmente ficou definido o limite máximo de tolerância para _
   recebimento, entre os documentos "Vencidos" e "a Vencer", da seguinte forma:
'      => Parâmetro Inicial:
'    Vencido: 3000 fatores;
'    A Vencer: 5500 fatores;
'    Range de Segurança: 499 fatores.

'2.1.2  Exemplos:

'       Exemplo 1 - considerando como Data de Pagamento 12/03/2014 (fator 6000):
'   Nessa data receber boletos VENCIDOS até a data limite 24/12/2005 (fator 3000), _
 que corresponde a Data de Pagamento menos 3000 fatores (parâmetro vencido);
'   Nessa data NÃO receber boletos VENCIDOS de 23/12/2005 (fator 2999) ou data _
 anterior - Range de Segurança;
'   Nessa data receber boletos A VENCER até a data limite 02/04/2029 (fator 2500), _
 que corresponde a Data de Pagamento mais 5500 fatores (parâmetro a vencer);
'   Nessa data NÃO receber boletos A VENCER de 03/04/2029 (fator 2501) ou data _
 posterior - Range de Segurança;
'   Range de Segurança nessa data: fator 2501 ao fator 2999 (499 fatores)

'       Exemplo 2 - considerando como Data de Pagamento 13/03/2014 (fator 6001):
'   Nessa data receber boletos VENCIDOS até a data limite 25/12/2005 (fator 3001), _
 que corresponde a Data de Pagamento menos 3000 fatores (parâmetro vencido);
'   Nessa data NÃO receber boletos VENCIDOS de 24/12/2005 (fator 3000) ou data _
 anterior - Range de Segurança;
'   Nessa data receber boletos A VENCER até a data limite 03/04/2029 (fator 2501), _
 que corresponde a Data de Pagamento mais 5500 fatores (parâmetro a vencer);
'   Nessa data NÃO receber boletos A VENCER de 04/04/2029 (fator 2502) ou data _
 posterior - Range de Segurança;
'   Range de Segurança nessa data: fator 2502 ao fator 3000 (499 fatores)
 
'    Exemplo 3 - considerando como Data de Pagamento 14/03/2014 (fator 6002):
'   Nessa data receber boletos VENCIDOS até a data limite 26/12/2005 (fator 3002), _
 que corresponde a Data de Pagamento menos 3000 fatores (parâmetro vencido);
'   Nessa data NÃO receber boletos VENCIDOS de 25/12/2005 (fator 3001) ou data _
 anterior - Range de Segurança;
'   Nessa data receber boletos A VENCER até a data limite 04/04/2029 (fator 2502), _
 que corresponde a Data de Pagamento mais 5500 fatores (parâmetro a vencer);
'   Nessa data NÃO receber boletos A VENCER de 05/04/2029 (fator 2503) ou data _
 posterior - Range de Segurança;
'    Range de Segurança nessa data: fator 2503 ao fator 3001 (499 fatores)
 
'3     Assim, considerando todo o exposto, solicitamos alterar o Caixa Programado _
 para que passe a validar o fator de vencimento conforme exposto nos itens 2 e _
 subitens, na tela de cadastro de agendamento de pagamento a fornecedor para as _
 formas de pagamento "Cobrança Eletronica Caixa" e "Cobrança Eletronica Outros _
 Bancos", conforme tela anexa.

'3.1   A alteração dessa regra de cálculo deve refletir na crítica da tela de _
 documentos vencidos e no preenchimento do campo "Data de Vencimento".

'4      Estas adequações devem estar em produção até o dia 12/03/2014.
'-------------------------------------------------------------------------------------------------------------

Private Sub cmdFatorDeVencimento_Click()
    
    lblFV0001Resposta1.Caption = clsFV.CalculaFatorDeVencimento(lblFV0001Data.Caption)
    lblFV1000Resposta1.Caption = clsFV.CalculaFatorDeVencimento(lblFV1000Data.Caption)
    lblFV2500Resposta1.Caption = clsFV.CalculaFatorDeVencimento(lblFV2500Data.Caption)
    lblFV2501Resposta1.Caption = clsFV.CalculaFatorDeVencimento(lblFV2501Data.Caption)
    lblFV2502Resposta1.Caption = clsFV.CalculaFatorDeVencimento(lblFV2502Data.Caption)
    lblFV2503Resposta1.Caption = clsFV.CalculaFatorDeVencimento(lblFV2503Data.Caption)
    lblFV2999Resposta1.Caption = clsFV.CalculaFatorDeVencimento(lblFV2999Data.Caption)
    lblFV3000Resposta1.Caption = clsFV.CalculaFatorDeVencimento(lblFV3000Data.Caption)
    lblFV3001Resposta1.Caption = clsFV.CalculaFatorDeVencimento(lblFV3001Data.Caption)
    lblFV3002Resposta1.Caption = clsFV.CalculaFatorDeVencimento(lblFV3002Data.Caption)
    lblFV6000Resposta1.Caption = clsFV.CalculaFatorDeVencimento(lblFV6000Data.Caption)
    lblFV6001Resposta1.Caption = clsFV.CalculaFatorDeVencimento(lblFV6001Data.Caption)
    lblFV6002Resposta1.Caption = clsFV.CalculaFatorDeVencimento(lblFV6002Data.Caption)
    lblFV9999Resposta1.Caption = clsFV.CalculaFatorDeVencimento(lblFV9999Data.Caption)
    lblFV0000Resposta1.Caption = clsFV.CalculaFatorDeVencimento(lblFV0000Data.Caption)
    
    If fuValidaData(txtFVDigitadoData.Text) Then
        lblFVDigitadoResposta1.Caption = clsFV.CalculaFatorDeVencimento(txtFVDigitadoData.Text)
    Else
        txtFVDigitadoData.Text = ""
    End If

End Sub

Private Sub cmdDataDeVencimento_Click()

    If fuValidaData(txtDataHoje.Text) Then
        clsFV.DtHoje = txtDataHoje.Text
    Else
        clsFV.DtHoje = Date
        txtDataHoje.Text = Date
    End If

    lblFV0001DataFun1.Caption = clsFV.fuFatorVencimento(lblFV0001Resposta1.Caption)
    lblFV1000DataFun1.Caption = clsFV.fuFatorVencimento(lblFV1000Resposta1.Caption)
    lblFV2500DataFun1.Caption = clsFV.fuFatorVencimento(lblFV2500Resposta1.Caption)
    lblFV2501DataFun1.Caption = clsFV.fuFatorVencimento(lblFV2501Resposta1.Caption)
    lblFV2502DataFun1.Caption = clsFV.fuFatorVencimento(lblFV2502Resposta1.Caption)
    lblFV2503DataFun1.Caption = clsFV.fuFatorVencimento(lblFV2503Resposta1.Caption)
    lblFV2999DataFun1.Caption = clsFV.fuFatorVencimento(lblFV2999Resposta1.Caption)
    lblFV3000DataFun1.Caption = clsFV.fuFatorVencimento(lblFV3000Resposta1.Caption)
    lblFV3001DataFun1.Caption = clsFV.fuFatorVencimento(lblFV3001Resposta1.Caption)
    lblFV3002DataFun1.Caption = clsFV.fuFatorVencimento(lblFV3002Resposta1.Caption)
    lblFV6000DataFun1.Caption = clsFV.fuFatorVencimento(lblFV6000Resposta1.Caption)
    lblFV6001DataFun1.Caption = clsFV.fuFatorVencimento(lblFV6001Resposta1.Caption)
    lblFV6002DataFun1.Caption = clsFV.fuFatorVencimento(lblFV6002Resposta1.Caption)
    lblFV9999DataFun1.Caption = clsFV.fuFatorVencimento(lblFV9999Resposta1.Caption)
    lblFV0000DataFun1.Caption = clsFV.fuFatorVencimento(lblFV0000Resposta1.Caption)
    txtFVDigitadoDataFun1.Text = clsFV.fuFatorVencimento(lblFVDigitadoResposta1.Caption)
    
End Sub


Public Function fuValidaData(sData As String) As Boolean
    fuValidaData = True
    
    If Val(Mid(sData, 4, 2)) > 12 Then
        MsgBox "Data inválida!", vbInformation, App.Title
        fuValidaData = False
    Else
    
        If Not IsDate(sData) Then
            MsgBox "Data inválida!", vbInformation, App.Title
            fuValidaData = False
        End If
        
    End If

End Function

Private Sub cmdLimpar_Click()
    lblFV0001DataFun1.Caption = ""
    lblFV1000DataFun1.Caption = ""
    lblFV2500DataFun1.Caption = ""
    lblFV2501DataFun1.Caption = ""
    lblFV2502DataFun1.Caption = ""
    lblFV2503DataFun1.Caption = ""
    lblFV2999DataFun1.Caption = ""
    lblFV3000DataFun1.Caption = ""
    lblFV3001DataFun1.Caption = ""
    lblFV3002DataFun1.Caption = ""
    lblFV6000DataFun1.Caption = ""
    lblFV6001DataFun1.Caption = ""
    lblFV6002DataFun1.Caption = ""
    lblFV9999DataFun1.Caption = ""
    lblFV0000DataFun1.Caption = ""
    txtFVDigitadoDataFun1.Text = ""
End Sub

Private Sub cmdLimparFatorVencimento_Click()
    lblFV0001Resposta1.Caption = ""
    lblFV1000Resposta1.Caption = ""
    lblFV2500Resposta1.Caption = ""
    lblFV2501Resposta1.Caption = ""
    lblFV2502Resposta1.Caption = ""
    lblFV2503Resposta1.Caption = ""
    lblFV2999Resposta1.Caption = ""
    lblFV3000Resposta1.Caption = ""
    lblFV3001Resposta1.Caption = ""
    lblFV3002Resposta1.Caption = ""
    lblFV6000Resposta1.Caption = ""
    lblFV6001Resposta1.Caption = ""
    lblFV6002Resposta1.Caption = ""
    lblFV9999Resposta1.Caption = ""
    lblFV0000Resposta1.Caption = ""
    lblFVDigitadoResposta1.Caption = ""
End Sub

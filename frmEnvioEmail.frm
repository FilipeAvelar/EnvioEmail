VERSION 5.00

Begin VB.Form frmEnvioEmail 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Envio e-mail"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmEnvioEmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEnvioEmail.frx":000C
   ScaleHeight     =   6000
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraEnvioEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5430
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6630
      Begin Text_USR.Text txtFrom 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   503
         MaxLength       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SempreVisivel   =   0   'False
         MudarFocoComSetas=   0   'False
         Text            =   ""
      End
      Begin Button_USR.Button cmdEnviar 
         Height          =   375
         Left            =   4920
         TabIndex        =   2
         Top             =   4815
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   661
         Caption         =   "Enviar"
         ButtonStyle     =   3
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
      End
      Begin Button_USR.Button cmdCancelar 
         Height          =   375
         Left            =   5730
         TabIndex        =   3
         Top             =   4815
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   661
         Caption         =   "Cancelar"
         ButtonStyle     =   3
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
      End
      Begin Text_USR.Text txtTo 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   600
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   503
         MaxLength       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SempreVisivel   =   0   'False
         MudarFocoComSetas=   0   'False
         Text            =   ""
      End
      Begin Text_USR.Text txtCC 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   960
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   503
         MaxLength       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SempreVisivel   =   0   'False
         MudarFocoComSetas=   0   'False
         Text            =   ""
      End
      Begin Text_USR.Text txtSubject 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   1320
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   503
         MaxLength       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SempreVisivel   =   0   'False
         MudarFocoComSetas=   0   'False
         Text            =   ""
      End
      Begin Text_USR.Text txtAnexo 
         Height          =   285
         Left            =   960
         TabIndex        =   14
         Top             =   1680
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   503
         Locked          =   -1  'True
         MaxLength       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SempreVisivel   =   0   'False
         MudarFocoComSetas=   0   'False
         Text            =   ""
      End
      Begin Text_USR.Text txtBody 
         Height          =   2565
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4524
         Linhas          =   10
         MaxLength       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SempreVisivel   =   0   'False
         MudarFocoComSetas=   0   'False
         MudarFocoComENTER=   0   'False
         Text            =   ""
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   4800
         Width           =   4695
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Axexo:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Para:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "CC:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "De:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Assunto:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.Label lblTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Novo E-mail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   165
      TabIndex        =   0
      Top             =   45
      Width           =   6690
   End
   Begin VB.Shape shpForm 
      Height          =   6000
      Left            =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmEnvioEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mAssunto                                 As String
Dim mAnexo                                   As String
Dim mEmailOrigem                             As String
Dim mEmaildestino                            As String
Dim mRetorno                                 As Boolean

Dim Movendo                                  As Boolean
Dim sX                                       As Integer
Dim sY                                       As Integer

Const BODY_FORMAT_HTML                       As Byte = 1
Const LICENSE_KEY                            As String = "MBC500-4A42932875-E5DC65976D41CB62D51A804BF5E27B76"

Const CHAVE_CAMPO_SERVIDOR_SMTP              As String = "Email.Param.ServidorSMTP"
Const CHAVE_CAMPO_REMETENTE                  As String = "Email.Param.Remetente"
Const CHAVE_CAMPO_EMAIL_REMETENTE            As String = "Email.Param.EmailRemetente"
Const CHAVE_CAMPO_IND_AUTENTICAR_NO_LOGON    As String = "Email.Param.IndAutenticarNoLogon"
Const CHAVE_CAMPO_IND_USAR_EMAIL_REMETENTE_NO_LOGON As String = "Email.Param.IndUsarEmailRemetenteNoLogon"
Const CHAVE_CAMPO_USUARIO_SERVIDOR_SMTP      As String = "Email.Param.UsuarioServidorSMTP"
Const CHAVE_CAMPO_SENHA_SERVIDOR_SMTP        As String = "Email.Param.SenhaServidorSMTP"
Const CHAVE_CAMPO_INTERVALO_MINUTOS          As String = "Email.Param.IntervaloMinutos"
Const CHAVE_CAMPO_PORTA_SERVIDOR_SMTP        As String = "Email.Param.PortaServidorSMTP"
Const CHAVE_CAMPO_HABILITAR_START_TLS        As String = "Email.Param.HabilitarStartTLS"
Const CHAVE_CAMPO_IND_USAR_PROTOCOLO_LOGON   As String = "Email.Param.IndUsarProtocoloLogon"
Const CHAVE_CAMPO_PROTOCOLO_LOGON            As String = "Email.Param.ProtocoloLogon"
Const CHAVE_CAMPO_IND_USAR_CONEXAO_SEGURA    As String = "Email.Param.IndUsarSSL"
Const CHAVE_CAMPO_PROTOCOLO_SSL              As String = "Email.Param.ProtocoloSSL"

Private Const PROTOCOLO_SSL_TLS10            As Byte = 4
Private Const PROTOCOLO_LOGON_NO_SERVIDOR    As Byte = 2




Public Property Get retorno() As String
1   retorno = mRetorno
End Property

Public Property Let Assunto(ByVal vNewValue As String)
1   mAssunto = vNewValue
End Property
Public Property Get Assunto() As String
1   Assunto = mAssunto
End Property
Public Property Let Anexo(ByVal vNewValue As String)
1   mAnexo = vNewValue
End Property
Public Property Get Anexo() As String
1   Anexo = mAnexo
End Property
Public Property Let EmailOrigem(ByVal vNewValue As String)
1   mEmailOrigem = vNewValue
End Property
Public Property Get EmailOrigem() As String
1   EmailOrigem = mEmailOrigem
End Property
Public Property Let EmailDestino(ByVal vNewValue As String)
1   mEmaildestino = vNewValue
End Property
Public Property Get EmailDestino() As String
1   EmailDestino = mEmaildestino
End Property

Private Sub cmdCancelar_Click()

1   mRetorno = False

2   Unload Me

End Sub

Private Sub cmdEnviar_Click()

Dim msg                                      As Object

1   On Error GoTo TrataErro

2   Set msg = CreateObject("MailBee.Message")

3   If Trim(txtFrom.Text) = vbNullString Then
4       txtFrom.Invalido = "Campo preenchimento obrigatório."
5       Exit Sub
6   End If

7   If Not msg.ValidateEmailAddress(Trim(txtFrom.Text)) Then
8       txtFrom.Invalido = "Endereço de e-mail inválido."
9       Exit Sub
10  End If

11  If Trim(txtTo.Text) = vbNullString Then
12      txtTo.Invalido = "Campo preenchimento obrigatório."
13      Exit Sub
14  End If

15  If Not msg.ValidateEmailAddress(Trim(txtTo.Text)) Then
16      txtTo.Invalido = "Endereço de e-mail inválido."
17      Exit Sub
18  End If

19  If Trim(txtSubject.Text) = vbNullString Then
20      txtSubject.Invalido = "Campo preenchimento obrigatório"
21      Exit Sub
22  End If

23  Call EnviarEmail

24  mRetorno = True

25  Exit Sub
TrataErro:
26  mRetorno = False

27  Call MostraExcecao(App.EXEName, "frmEnvioEmail", "cmdEnviar_Click", Err.Number, Err.Description, Err.Source, Erl)

End Sub

Private Sub Form_Load()

1   On Error GoTo TrataErro

2   shpForm.Width = Me.Width
3   shpForm.Top = Me.Top
4   shpForm.Height = Me.Height

5   txtFrom.Text = GetParametro("", "Email.Param.EmailRemetente", mEmailOrigem)'mEmailOrigem
6   txtTo.Text = mEmaildestino
7   txtSubject.Text = mAssunto
8   txtAnexo.Text = RetornaNomeArq(mAnexo)
9   txtAnexo.Tag = mAnexo

10  Exit Sub
TrataErro:

End Sub

Public Function RetornaNomeArq(ByVal arq As String) As String
On Error GoTo Trata_Erro_RetornaNomeArq

    Dim i                                    As Integer

For i = Len(arq) To 1 Step -1
  If Mid(arq, i, 1) = "\" Then Exit For

Next
RetornaNomeArq = Right(arq, Len(arq) - i)


Exit Function

Trata_Erro_RetornaNomeArq:
End Function

Private Sub lblTitulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1   Movendo = True
2   sX = X
3   sY = Y
End Sub

Private Sub lblTitulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
1   On Error Resume Next
2   If Movendo Then
3       Me.Left = Me.Left + (X - sX)
4       Me.Top = Me.Top + (Y - sY)
5   End If
End Sub

Private Sub lblTitulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
1   Movendo = False
End Sub


Private Sub EnviarEmail()

Dim objSMTP                                  As Object    'MailBee.SMTP
Dim objSSL                                   As Object    'MailBee.SSL
Dim strErro                                  As String
Dim ServidorSMTP                             As String  '= "Email.Param.ServidorSMTP"
Dim Remetente                                As String  '= "Email.Param.Remetente"
Dim EmailRemetente                           As String  '= "Email.Param.EmailRemetente"
Dim IndAutenticarNoLogon                     As String  '= "Email.Param.IndAutenticarNoLogon"
Dim IndUsarEmailRemetenteNoLogon             As String  '= "Email.Param.IndUsarEmailRemetenteNoLogon"
Dim UsuarioServidorSMTP                      As String  '= "Email.Param.UsuarioServidorSMTP"
Dim SenhaServidorSMTP                        As String  '= "Email.Param.SenhaServidorSMTP"
Dim IntervaloMinutos                         As String  '= "Email.Param.IntervaloMinutos"
Dim PortaServidorSMTP                        As String  '= "Email.Param.PortaServidorSMTP"
Dim HabilitarStartTLS                        As String  '= "Email.Param.HabilitarStartTLS"
Dim IndUsarProtocoloLogon                    As String  '= "Email.Param.IndUsarProtocoloLogon"
Dim ProtocoloLogon                           As String  '= "Email.Param.ProtocoloLogon"
Dim IndUsarSSL                               As String  '= "Email.Param.IndUsarSSL"
Dim ProtocoloSSL                             As String  '= "Email.Param.ProtocoloSSL"

1   On Error GoTo finally

2   ServidorSMTP = GetParametro("", CHAVE_CAMPO_SERVIDOR_SMTP, "")
3   UsuarioServidorSMTP = GetParametro("", CHAVE_CAMPO_USUARIO_SERVIDOR_SMTP, "")
4   SenhaServidorSMTP = UnCript(GetParametro("", CHAVE_CAMPO_SENHA_SERVIDOR_SMTP, ""))
5   IndAutenticarNoLogon = GetParametro("", CHAVE_CAMPO_IND_AUTENTICAR_NO_LOGON, "")
6   ProtocoloLogon = GetParametro("", CHAVE_CAMPO_PROTOCOLO_LOGON, "")
7   PortaServidorSMTP = GetParametro("", CHAVE_CAMPO_PORTA_SERVIDOR_SMTP, "")

8   IndUsarEmailRemetenteNoLogon = GetParametro("", CHAVE_CAMPO_IND_USAR_EMAIL_REMETENTE_NO_LOGON, "")
9   IntervaloMinutos = GetParametro("", CHAVE_CAMPO_INTERVALO_MINUTOS, "")
10  IndUsarProtocoloLogon = GetParametro("", CHAVE_CAMPO_IND_USAR_PROTOCOLO_LOGON, "")

11  IndUsarSSL = GetParametro("", CHAVE_CAMPO_IND_USAR_CONEXAO_SEGURA, "")
12  ProtocoloSSL = GetParametro("", CHAVE_CAMPO_PROTOCOLO_SSL, "")
13  HabilitarStartTLS = GetParametro("", CHAVE_CAMPO_HABILITAR_START_TLS, "")

14  Remetente = GetParametro("", CHAVE_CAMPO_REMETENTE, "")
15  EmailRemetente = GetParametro("", CHAVE_CAMPO_EMAIL_REMETENTE, "")

16  If Trim(ServidorSMTP) = vbNullString Then
17      Call MsgBox("Parâmetros do servidors de e-mail não está configurado." & vbCrLf & "verifique os dados configurados em 'Administração-->Serviços--Parâmetro de e-mail'", vbOKOnly)
18      Exit Sub
19  End If

20  Set objSMTP = CreateObject("MailBee.SMTP")

21  objSMTP.LicenseKey = LICENSE_KEY

22  If (Not objSMTP.Licensed) Then
23      Call Err.Raise(-97, "Envio de e-mail", "Chave de Licença do objeto SMTP inválida")
24      GoTo finally
25  End If

26  objSMTP.BodyFormat = BODY_FORMAT_HTML
27  objSMTP.EnableLogging = False
28  objSMTP.ServerName = ServidorSMTP
29  objSMTP.PortNumber = PortaServidorSMTP
30  objSMTP.EnableEvents = True
31  objSMTP.FromAddr = frmEnvioEmail.txtFrom.Text

    ' Use SMTP authentication?
32  If CBool(IndUsarProtocoloLogon) Then
33      objSMTP.AuthMethod = ProtocoloLogon
34      objSMTP.UserName = UsuarioServidorSMTP
35      objSMTP.Password = SenhaServidorSMTP
36  Else
37      objSMTP.AuthMethod = 0
38  End If

39  If CBool(IndAutenticarNoLogon) Then
40      If CBool(IndUsarEmailRemetenteNoLogon) Then
41          objSMTP.UserName = EmailRemetente
42      Else
43          objSMTP.UserName = UsuarioServidorSMTP
44      End If
45      objSMTP.Password = SenhaServidorSMTP
46  End If

47  objSMTP.message.AddAttachment (txtAnexo.Tag)
48  objSMTP.message.Locked = False
49  objSMTP.message.FromAddr = txtFrom.Text
50  objSMTP.message.ToAddr = txtTo.Text
51  objSMTP.message.CCAddr = txtCC.Text
52  objSMTP.message.Subject = txtSubject.Text
53  objSMTP.message.BodyFormat = BODY_FORMAT_HTML
54  objSMTP.message.BodyText = Replace(txtBody.Text, vbCrLf, "<br/>")
55  objSMTP.message.ReplyToAddr = frmEnvioEmail.txtFrom.Text

    'Config SSL
56  If CBool(IndUsarSSL) Then

57      Set objSSL = CreateObject("MailBee.SSL")
58      objSSL.LicenseKey = LICENSE_KEY
59      If (Not objSSL.Licensed) Then
60          Call Err.Raise(-98, "Envio de e-mail", "Chave de Licença do objeto SSL inválida")
61          GoTo finally
62      End If

63      objSSL.UseStartTLS = CBool(HabilitarStartTLS)
64      objSSL.Protocol = ProtocoloSSL

65      Set objSMTP.SSL = objSSL

66  End If

    '-----------------------------------------------------------------------------

67  cmdEnviar.Enabled = False
68  lblStatus.Caption = "Conectando ao servidor " & ServidorSMTP & " ... "
69  DoEvents
70  If objSMTP.Connect Then
71      lblStatus.Caption = "Enviando Mensagem" & " ... "
72      DoEvents
73      objSMTP.send
74      If Not objSMTP.IsError Then
75          If objSMTP.Connected Then
76              objSMTP.Disconnect
77          End If
78          lblStatus.Caption = "Mensagem enviada com sucesso"
79          cmdCancelar.Caption = "Fechar"
80      Else
81          cmdEnviar.Enabled = True
82          strErro = "Erro ao enviar mensagem : " & descricaoErroSMTP(objSMTP.ErrCode, objSMTP.ErrDesc, objSMTP.ServerResponse)
83      End If
84  Else
85      cmdEnviar.Enabled = True
86      strErro = "Erro ao enviar mensagem : " & descricaoErroSMTP(objSMTP.ErrCode, objSMTP.ErrDesc, objSMTP.ServerResponse)
87  End If

88  If strErro <> vbNullString Then
89      lblStatus.Caption = strErro
90  End If

91  Exit Sub
finally:
92  Set objSSL = Nothing
93  Set objSMTP = Nothing

94  Call EmpilhaExcecao(App.EXEName, "frmEnvioEmail", "EnviarEmail", Err.Number, Err.Description, Err.Source, Erl)

End Sub


Private Function descricaoErroSMTP(intIdErr, strErrDesc, strServerResponse) As String
1   On Error GoTo TrataErro

2   Select Case intIdErr

        Case 1
3           descricaoErroSMTP = " não conectado"
4       Case 2
5           descricaoErroSMTP = " Já estava conectado"
6       Case 3
7           descricaoErroSMTP = " Endereço do servidor remoto não pôde ser resolvido"
8       Case 4
9           descricaoErroSMTP = " Servidor não encontrado"
10      Case 5
11          descricaoErroSMTP = " Nenhuma resposta do servidor no intervalo especificado pelo tempo de espera"
12      Case 6
13          descricaoErroSMTP = " Conexão encerrada pelo servidor"
14      Case 7
15          descricaoErroSMTP = " Conexão fechada pelo cliente"
16      Case 8
17          descricaoErroSMTP = " Erro de conexão (Outros)"
18      Case 11
19          descricaoErroSMTP = " A conexão SSL pela porta padrão não é suportada no servidor"
20      Case 12
21          descricaoErroSMTP = " Falha de inicialização da conexão via SSL"
22      Case 13
23          descricaoErroSMTP = " Falha na encriptação pelo protocolo SSL"
24      Case 14
25          descricaoErroSMTP = " Falha da decriptação pelo protocolo SSL"
26      Case 101
27          descricaoErroSMTP = " Erro de licença do componente"
28      Case 102
29          descricaoErroSMTP = " Operação cancelada"
30      Case 111
31          descricaoErroSMTP = " O servidor não trabalha com autenticação pelo protocolo ESMTP"
32      Case 112
33          descricaoErroSMTP = " O servidor não dá suporte ao protocolo de autenticação ESMTP"
34      Case 113
35          descricaoErroSMTP = " Falha no protocolo de autenticação (motivo desconhecido pelo componente de e-mail)"
36      Case 114
37          descricaoErroSMTP = " Usuário e/ou senha incorreto(s)"
38      Case 115
39          descricaoErroSMTP = " Nome do servidor de origem não pôde ser resolvido"
40      Case 116
41          descricaoErroSMTP = " Falha de comando (outros)"
42      Case 121
43          descricaoErroSMTP = " Remetente e pelo menos um destinatário deve ser especificado"
44      Case 122
45          descricaoErroSMTP = " Remetente não permitido pelo servidor de e-mail"
46      Case 123
47          descricaoErroSMTP = " Destinatário não permitido pelo servidor ou o número de destinatários excedeu o máximo permitido pelo servidor"
48      Case 124
49          descricaoErroSMTP = " O servidor rejeitou os dados da mensagem"
50      Case 131
51          descricaoErroSMTP = " Arquivo da mensagem não pôde ser criado na pasta especificada"
52      Case 132
53          descricaoErroSMTP = " Arquivo da mensagem não pôde ser salvo na pasta especificada. Verifique espaço em disco, permissões etc"
54      Case 133
55          descricaoErroSMTP = " Arquivo da mensagem não pôde ser lido na pasta especificada"
56      Case Else
57          descricaoErroSMTP = " " & strErrDesc

58  End Select

59  descricaoErroSMTP = intIdErr & " - " & descricaoErroSMTP

60  If (strServerResponse <> "") Then
        'descricaoErroSMTP = descricaoErroSMTP & vbLf & "Mensagem do Servidor SMTP: " & strServerResponse
61  End If

62  Exit Function

TrataErro:
End Function

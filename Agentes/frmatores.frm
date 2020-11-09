VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form frmatores 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Exemplo - Agentes Microsoft"
   ClientHeight    =   2520
   ClientLeft      =   5445
   ClientTop       =   3135
   ClientWidth     =   5370
   Icon            =   "frmatores.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "&Limpa Texto"
      Height          =   390
      Left            =   3375
      TabIndex        =   7
      Top             =   150
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Sair"
      Height          =   540
      Left            =   4650
      TabIndex        =   6
      Top             =   975
      Width           =   690
   End
   Begin VB.CommandButton command4 
      Caption         =   "&Fazer Mágica"
      Height          =   375
      Left            =   3435
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton command2 
      Caption         =   "&Gesticular"
      Height          =   375
      Left            =   1230
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton command3 
      Caption         =   "&Movimentos"
      Height          =   375
      Left            =   2325
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton command1 
      Caption         =   "&Falar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4485
   End
   Begin VB.Label Label1 
      Caption         =   "Digite aqui o que você quer que o Genio Fale e Pressione o botão Falar ..."
      Height          =   390
      Left            =   150
      TabIndex        =   5
      Top             =   150
      Width           =   3165
      WordWrap        =   -1  'True
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   4725
      Top             =   1650
      _cx             =   847
      _cy             =   847
   End
End
Attribute VB_Name = "frmatores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Genie As IAgentCtlCharacterEx 'inicializa o ator
Const DataAtor = "genie.acs" 'define o caminho para o agente que voce vai usar
Const dataAtor2 = "merlin.acs"
Dim Merlin As IAgentCtlCharacter
Private Sub command4_Click()
  'o ator agora faz as magicas
  Genie.Play ("DoMagic1")
  Genie.Play ("DoMagic2")
End Sub
Private Sub command2_Click()
 Genie.Play ("GestureRight")  'aciona os gestos do ator
 Genie.Play ("GestureLeft")
 Genie.Play ("GestureUp")
 Genie.Play ("GestureDown")
 Genie.Play ("LookRight")
 Genie.Play ("LookLeft")
 Genie.Play ("LookUp")
 Genie.Play ("LookDown")
End Sub
Private Sub command3_Click()
 Genie.MoveTo 100, 10 'move o ator para direita
 Genie.MoveTo 500, 50 'move o ator um pouco mais
 Genie.MoveTo 600, 20 'continua a mover o ator
 Genie.MoveTo 10, 10   'move o ator a sua posicao inicial
End Sub
Private Sub command1_Click()
    If Text1 = "" Then Text1 = "Digite algo na caixa de texto para que eu possa falar..."
    Genie.Speak Text1
End Sub
Private Sub Command6_Click()
    Dim fala As String
    Agent1.Characters.Load "Merlin", dataAtor2 'carrega o ator
    Set Merlin = Agent1.Characters("Merlin") 'define o ator que vai atuar
    Merlin.LanguageID = &H409 'define a linguagem
    Merlin.MoveTo 500, 240
    Merlin.Show   'faz o ator aparecer
    Merlin.Play ("Greet") 'ator atua dando saudacoes
    fala = "Bem , então por hoje é só \pau=400\ "
    fala = fala & "hasta la vista baby ... e não se esqueça visite \pau=400\"
    fala = fala & "\emp\geocities.com/SiliconValley/Bay/3994 " 'mensagem saida
    Merlin.Speak fala 'aqui o ator 'fala' mensagem de saida
    Merlin.Play "wave"
    Genie.Play "wave"
    Merlin.Hide
    Genie.Hide
    MsgBox "Adeus , Merlin ... Adeus Genie...!", vbOKOnly, "Companhia de teatro - MS"
    Set Merlin = Nothing
    Set Genie = Nothing
    Unload Me
End Sub
Private Sub Command5_Click()
  Text1.Text = ""
End Sub
Private Sub Form_Load()
    Dim fala As String
    Agent1.Characters.Load "Genie", DataAtor 'carrega o ator
    Set Genie = Agent1.Characters("Genie") 'define o ator que vai atuar
    Genie.LanguageID = &H409 'define a linguagem
    Genie.MoveTo 100, 240
    Genie.Show   'faz o ator aparecer
    Genie.Play ("Greet") 'ator atua dando saudacoes
    fala = "\pit=400\Benvindo , Eu fui programado para ser seu escravo,\pau=400\ "
    fala = fala & "visite o site do meu criador em \pau=400\"
    fala = fala & "\emp\geocities.com/SiliconValley/Bay/3994 " 'mensagem de boas vindas
    Genie.Speak fala 'aqui o ator 'fala' as mensagens boas de vindas
End Sub

'Modificadores do discurso :
'   1.  \emp\       enfatiza a palavra
'   2.  \pau = m\   pause de m milisegundos
'   3.  \pit = p\    voz para p Hertz (1 - 400)
'   4.  \spd = s\  define a velocidade para s palavras por minuto


 

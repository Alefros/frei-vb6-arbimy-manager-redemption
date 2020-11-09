VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_loca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localizações"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5385
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Comandos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   1680
      Width           =   5175
      Begin VB.CommandButton cmd_alterar 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   3960
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_novo 
         Caption         =   "Novo"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_excluir 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   2760
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_salvar 
         Caption         =   "Salvar"
         Height          =   375
         Left            =   1560
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fme_logradouro 
      Caption         =   "Logradouros"
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.ComboBox cbo_bairro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   18
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txt_logradouro 
         Height          =   345
         Left            =   1200
         TabIndex        =   17
         Top             =   1200
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "Form1.frx":030A
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "Form1.frx":036E
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Frame fme_bairro 
      Caption         =   "Bairros"
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txt_bairro 
         Height          =   345
         Left            =   1200
         TabIndex        =   13
         Top             =   1200
         Width           =   3855
      End
      Begin VB.ComboBox cbo_cidade 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "Form1.frx":03DA
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Form1.frx":043E
         TabIndex        =   15
         Top             =   1320
         Width           =   735
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3120
      OleObjectBlob   =   "Form1.frx":04A2
      Top             =   2400
   End
   Begin VB.CommandButton cmd_consultar 
      Caption         =   "Consultar"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmd_anterior 
      Caption         =   "Anterior"
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmd_proximo 
      Caption         =   "Próximo"
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.Frame fme_cidade 
      Caption         =   "Cidades"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txt_cidade 
         Height          =   345
         Left            =   1200
         TabIndex        =   10
         Top             =   1200
         Width           =   3855
      End
      Begin VB.ComboBox cbo_uf 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "Form1.frx":1C6F5
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Form1.frx":1C751
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame fme_estado 
      Caption         =   "Estados"
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txt_uf 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   0
         Top             =   1200
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "Form1.frx":1C7B5
         TabIndex        =   6
         Top             =   1200
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm_loca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim foco As Integer
Option Explicit

Private Sub cmd_anterior_Click()
            If foco = 4 Then
                foco = 3
                cmd_proximo.Enabled = True
            ElseIf foco = 3 Then
                    foco = 2
                ElseIf foco = 2 Then
                        foco = 1
                        cmd_anterior.Enabled = False
            End If
            Call frame
End Sub

Private Sub cmd_cidade_Click()
            
End Sub

Private Sub cmd_cidade_GotFocus()
             foco = 2
            Call frame
End Sub

Private Sub cmd_estado_GotFocus()
           
End Sub

Private Sub cmd_loca_Click()
            
End Sub

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub cmd_loca_GotFocus()
            foco = 3
            Call frame
End Sub

Private Sub cmd_consultar_Click()
            MsgBox " :( este controle está em manutenção", vbInformation, "Manutenção"
End Sub

Private Sub cmd_novo_Click()
            If foco = 1 Then
                txt_uf = Empty
                txt_uf.SetFocus
            ElseIf foco = 2 Then
                    cbo_uf = Empty
                    txt_cidade = Empty
                    txt_cidade.SetFocus
                ElseIf foco = 3 Then
                        cbo_cidade = Empty
                        txt_bairro = Empty
                        txt_bairro.SetFocus
                    ElseIf foco = 4 Then
                            cbo_bairro = Empty
                            txt_logradouro = Empty
                            txt_logradouro.SetFocus
            End If
End Sub

Private Sub cmd_proximo_Click()
            If foco = 1 Then
                foco = 2
                cmd_anterior.Enabled = True
            ElseIf foco = 2 Then
                    foco = 3
                ElseIf foco = 3 Then
                        foco = 4
                        cmd_proximo.Enabled = False
                        cmd_anterior.SetFocus
            End If
            Call frame
            
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
            Call dimensoes
            foco = "1"
            Call frame
            Skin1.ApplySkin Me.hWnd
            Call abrir_banco
            Call abrir_banco2
           
End Sub
Private Sub frame()

            If foco = "1" Then
                fme_estado.Visible = True
                fme_cidade.Visible = False
                fme_bairro.Visible = False
                fme_logradouro.Visible = False
            ElseIf foco = "2" Then
                    fme_estado.Visible = False
                    fme_cidade.Visible = True
                    fme_bairro.Visible = False
                    fme_logradouro.Visible = False
                    txt_cidade.SetFocus
                ElseIf foco = "3" Then
                    fme_estado.Visible = False
                    fme_cidade.Visible = False
                    fme_bairro.Visible = True
                    fme_logradouro.Visible = False
                    txt_bairro.SetFocus
                    ElseIf foco = "4" Then
                            fme_estado.Visible = False
                            fme_cidade.Visible = False
                            fme_bairro.Visible = False
                            fme_logradouro.Visible = True
                            txt_logradouro.SetFocus
            End If
End Sub
Private Sub dimensoes()
            frm_loca.Height = 3435
            frm_loca.ScaleHeight = 2970
            frm_loca.ScaleWidth = 5430
            frm_loca.Width = 5520
End Sub

Private Sub txt_uf_LostFocus()
            txt_uf = UCase(txt_uf)
End Sub

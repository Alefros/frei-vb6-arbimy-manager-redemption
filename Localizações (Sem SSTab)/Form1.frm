VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_loca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localizações"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4905
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
   ScaleHeight     =   2250
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2880
      OleObjectBlob   =   "Form1.frx":030A
      Top             =   1680
   End
   Begin VB.CommandButton cmd_consultar 
      Caption         =   "Consultar"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmd_anterior 
      Caption         =   "Anterior"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmd_proximo 
      Caption         =   "Próximo"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame fme_logradouro 
      Caption         =   "Logradouro"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Frame fme_cidade 
      Caption         =   "Cidade"
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Frame fme_estado 
      Caption         =   "Estado"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "frm_loca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim foco As Integer
Option Explicit

Private Sub Command1_Click()
        
End Sub

Private Sub cmd_anterior_Click()
            If foco = "3" Then
                foco = "2"
                cmd_proximo.Enabled = True
            ElseIf foco = "2" Then
                    foco = "1"
                    cmd_anterior.Enabled = False
            End If
            Call frame
End Sub

Private Sub cmd_cidade_Click()
            
End Sub

Private Sub cmd_cidade_GotFocus()
             foco = "2"
            Call frame
End Sub

Private Sub cmd_estado_GotFocus()
           
End Sub

Private Sub cmd_loca_Click()
            
End Sub

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub cmd_loca_GotFocus()
            foco = "3"
            Call frame
End Sub

Private Sub cmd_consultar_Click()
            MsgBox " :( este controle está em manutenção", vbInformation, "Manutenção"
End Sub

Private Sub cmd_proximo_Click()
            If foco = "1" Then
                foco = "2"
                cmd_anterior.Enabled = True
            ElseIf foco = "2" Then
                    foco = "3"
                    cmd_proximo.Enabled = False
                    cmd_anterior.SetFocus
            End If
            Call frame
            
End Sub

Private Sub Form_Load()
            foco = "1"
            Call frame
            Skin1.ApplySkin Me.hWnd
           
End Sub
Private Sub frame()
            If foco = "1" Then
                fme_estado.Visible = True
                fme_cidade.Visible = False
                fme_logradouro.Visible = False
            ElseIf foco = "2" Then
                    fme_estado.Visible = False
                    fme_cidade.Visible = True
                    fme_logradouro.Visible = False
                ElseIf foco = "3" Then
                        fme_estado.Visible = False
                        fme_cidade.Visible = False
                        fme_logradouro.Visible = True
            End If
End Sub

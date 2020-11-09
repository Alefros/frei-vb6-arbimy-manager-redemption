VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   3465
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "frm_login.frx":030A
      Top             =   600
   End
   Begin VB.CommandButton cmd_entrar 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Entrar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txt_senha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox txt_login 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   0
      Width           =   2535
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "frm_login.frx":1C55D
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "frm_login.frx":1C5BF
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tabpraga As New ADODB.Recordset
Option Explicit
Private Sub msg()
            MsgBox "O Login/Senha estão incorretos", vbExclamation, "Arbimy Manager 2.0"
End Sub
Private Sub cmd_entrar_Click()
''''''''''''senha/login não digitados''''''''''''''''''
            If txt_login = Empty Then
                Call msg
                txt_login = Empty
                txt_senha = Empty
                txt_login.SetFocus
                Exit Sub
            Else
            If txt_senha = Empty Then
a:
                Call msg
                txt_login = Empty
                txt_senha = Empty
                txt_login.SetFocus
                Exit Sub
            End If
            End If
''''''''Verificar Login e Senha'''''''''''''''''''''''''''''''''''''''''''''''''''''
            tabpraga.Close
            tabpraga.Open "Select * from Usuarios where Login = '" & txt_login & "' and Senha like '" & txt_senha & "'"
            If tabpraga.RecordCount = 1 Then
                    Unload Me
                    frmSplash.Show
            ElseIf tabpraga.RecordCount = 0 Then
                    GoTo a:
            End If
End Sub
Private Sub Form_Load()
            Skin1.ApplySkin Me.hWnd
            Call abrir
End Sub

Private Sub txt_senha_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                Call cmd_entrar_Click
            End If
End Sub
Private Sub abrir()
            Call abrir_banco
            Call fechar
            tabpraga.Open "Usuarios", conectar, adOpenKeyset, adLockOptimistic
End Sub
Private Sub fechar()
            If tabpraga.State = 1 Then tabpraga.Close
End Sub

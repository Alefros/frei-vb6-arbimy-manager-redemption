VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4830
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   4320
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
      Min             =   1e-4
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   360
      Top             =   3720
   End
   Begin VB.Timer Timer2 
      Interval        =   60
      Left            =   1320
      Top             =   4320
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6720
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4050
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7200
      Begin VB.TextBox txt_status 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   3720
         Width           =   6855
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copyright Todos os direitos reservados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label lbl_carregando 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Carregando"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edition Redemption"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         TabIndex        =   4
         Top             =   1800
         Width           =   2970
      End
      Begin VB.Image imgLogo 
         Height          =   1785
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   2295
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "2.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         TabIndex        =   1
         Top             =   2160
         Width           =   330
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Arbimy manager"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   660
         Left            =   2400
         TabIndex        =   3
         Top             =   1140
         Width           =   4395
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "MicroConnect"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   2400
         TabIndex        =   2
         Top             =   705
         Width           =   2415
      End
   End
   Begin VB.Label lbl_data 
      BackColor       =   &H00FFFFFF&
      Caption         =   "00/00/0000"
      BeginProperty Font 
         Name            =   "Digital-7"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label lbl_hora 
      BackColor       =   &H00FFFFFF&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Digital-7"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim max, value As Integer
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
            
            lbl_data = Date
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
            lbl_hora = Time
'            Call carregando
            
End Sub
Private Sub carregando()
            If lbl_carregando.Caption = "Carregando" Then
                lbl_carregando.Caption = "Carregando."
            ElseIf lbl_carregando.Caption = "Carregando." Then
                lbl_carregando.Caption = "Carregando.."
                ElseIf lbl_carregando.Caption = "Carregando.." Then
                        lbl_carregando.Caption = "Carregando..."
                    ElseIf lbl_carregando.Caption = "Carregando..." Then
                            lbl_carregando.Caption = "Carregando"
            End If
End Sub
Private Sub Timer2_Timer()
            
            max = 100
            value = 1
            ProgressBar1.max = max
            If ProgressBar1.value = max Then GoTo A
            ProgressBar1.value = Int(ProgressBar1.value) + value
            txt_status.Text = ProgressBar1.value & " %"
                If ProgressBar1.value > ProgressBar1.max Then
A:
                    MDIForm1.Show
                    Unload Me
                End If
                Exit Sub
                
End Sub

Private Sub Timer3_Timer()
            Call carregando
End Sub

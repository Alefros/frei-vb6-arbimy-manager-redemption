VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_categorias 
   Caption         =   "Cadastro de categorias"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "frm_categorias.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3240
      OleObjectBlob   =   "frm_categorias.frx":0064
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "frm_categorias.frx":1C2B7
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "frm_categorias.frx":1C321
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frm_categorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
            Skin1.ApplySkin Me.hWnd
End Sub

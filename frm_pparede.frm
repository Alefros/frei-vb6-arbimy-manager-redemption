VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "Skin.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_pparede 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alterar Papel de Parede"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8040
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4200
      TabIndex        =   18
      Top             =   3600
      Width           =   3735
      Begin VB.CommandButton cmd_novo 
         Caption         =   "Novo evento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmd_verificar 
         Caption         =   "Verificar eventos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmd_gravar 
         Caption         =   "Gravar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmd_alterar 
         Caption         =   "Alterar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   19
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      TabIndex        =   15
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton cmd_padrao 
         Caption         =   "Imagem Padrão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmd_aplicar 
         Caption         =   "Aplicar agora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Eventos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   8040
      TabIndex        =   12
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton cmd_voltar 
         Caption         =   "Voltar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   14
         Top             =   4440
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid mfg_eventos 
         Height          =   3975
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   7011
         _Version        =   393216
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2400
      OleObjectBlob   =   "frm_pparede.frx":0000
      Top             =   2160
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecionar imagem"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   3975
      Begin VB.TextBox txt_imagem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   4080
         Width           =   3735
      End
      Begin VB.CommandButton cmd_procurar 
         Caption         =   "Procurar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   0
         Top             =   4440
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H000000FF&
         Height          =   3615
         Left            =   120
         ScaleHeight     =   3585
         ScaleWidth      =   3705
         TabIndex        =   10
         Top             =   360
         Width           =   3735
         Begin VB.Image img_pparede 
            Height          =   3345
            Left            =   120
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3480
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2280
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Intervalo de data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4200
      TabIndex        =   4
      Top             =   1920
      Width           =   3735
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frm_pparede.frx":1C253
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtp_inicio 
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8421504
         Format          =   16908289
         CurrentDate     =   40554
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frm_pparede.frx":1C2B5
         TabIndex        =   7
         Top             =   1080
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtp_fim 
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8421504
         Format          =   16908289
         CurrentDate     =   40554
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informações do evento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      TabIndex        =   1
      Top             =   960
      Width           =   3735
      Begin VB.TextBox txt_evento 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   840
         TabIndex        =   3
         Text            =   "(ex: Natal, Ano novo, etc...)"
         Top             =   360
         Width           =   2775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frm_pparede.frx":1C311
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm_pparede"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_Colunas, l_linha As Long
Dim L_codp As Integer

Dim imagem As String
Dim prop As String
Dim desvio As String
Dim tabparede As New ADODB.Recordset
Option Explicit

Private Sub cmd_alterar_Click()
            status = "alteradas"
            desvio = "xD"
            Call gravar
            
'            If txt_imagem.Text <> tabparede!imagem Then
'                        imagem = tabparede!imagem
'                        FileCopy CommonDialog1.FileName, App.Path & "\Imagens\" & imagem
'            End If
            
            Call carregar_lista
            
                
                
End Sub

Private Sub cmd_aplicar_Click()

            If img_pparede = Empty Then GoTo b
                If imagem <> "" Then
                    img = imagem
                ElseIf imagem = "" Then
b:
            
                        MsgBox "Não há imagem a ser aplicada", vbInformation, "Arbimy Manager 2.0"
                Exit Sub
                End If
            Call carrega_imagem
End Sub

Private Sub cmd_gravar_Click()
            Call gravar
            Call cmd_novo_Click
End Sub

Private Sub cmd_novo_Click()
            img_pparede.Picture = LoadPicture(Empty)
                If cmd_alterar.Enabled = True Then
                    cmd_alterar.Enabled = False
                End If
            txt_imagem = Empty
            txt_evento = "(ex: Natal, Ano novo, etc...)"
            dtp_inicio = Date
            dtp_fim = Date + 1
            txt_evento.SetFocus
            
End Sub


Private Sub cmd_padrao_Click()
            img = App.Path & "\Imagens\Padrão.jpg"
                Call carrega_imagem
End Sub

Private Sub cmd_procurar_Click()
            On Error GoTo a
            
            CommonDialog1.DialogTitle = "Procurar imagem - Arbimy manager 2.0" ' define o titulo da caixa de dialogo
            prop = "Bitmap (*bmp; *.dib)|*bmp; *.dib|JPEG (*.jpg; *.jpeg; *.jpe; *.jfif)|*.jpg; *.jpeg; *.jpe; *.jfif|GIF (*.gif)|*.gif|PNG (*.png)|*.png|TIFF (*.tif; *.tiff)|*.tif; *.tiff|Todas as Imagens|*bmp; *.dib; *.jpg; *.jpeg; *.jpe; *.jfif; *.gif; *.png; *.tif; *.tiff;" ' define as extensões de arquivos
            CommonDialog1.Filter = prop ' cria um filtro para as extensões definidas
            CommonDialog1.FilterIndex = 6
            CommonDialog1.ShowOpen
            imagem = CommonDialog1.FileTitle
            
            txt_imagem = CommonDialog1.FileTitle
            img_pparede.Picture = LoadPicture(imagem)

a:          If Err.Description = "Invalid picture" Then
                MsgBox "Esta imagem é inválida, selecione outra imagem", vbInformation, "Arbimy Manager 2.0"
                    txt_imagem = Empty
                    Call cmd_procurar_Click
                Exit Sub
            End If
End Sub

Private Sub cmd_verificar_Click()
            Skin1.RemoveSkin Me.hWnd
            frm_pparede.Width = 13755
            cmd_verificar.Visible = False
'            cmd_ok.Visible = False
            Skin1.ApplySkin Me.hWnd
End Sub
Private Sub cmd_voltar_Click()
            Skin1.RemoveSkin Me.hWnd
                frm_pparede.Width = 8130
                cmd_verificar.Visible = True
                cmd_alterar.Enabled = False
            Skin1.ApplySkin Me.hWnd
End Sub

Private Sub Form_Load()
            Call dimensoes
            dtp_inicio.value = Date
            dtp_fim.value = Date + 1
            Skin1.ApplySkin Me.hWnd
            Call configu
            Call abrir
            Call leventos
            Call carregar_lista
End Sub

Private Sub mfg_eventos_Click()
            cmd_alterar.Enabled = True
            l_linha = mfg_eventos.Row
            L_codp = mfg_eventos.TextMatrix(l_linha, 0)
                Call abrir
            tabparede.Close
            tabparede.Open "Select * From pparede Where cod = " & L_codp
            Call mostrar
End Sub

Private Sub txt_evento_Click()
            If txt_evento.Text = "(ex: Natal, Ano novo, etc...)" Then
                txt_evento = Empty
            End If
End Sub
Private Sub dimensoes()
            frm_pparede.Height = 5490
            frm_pparede.Width = 8130
End Sub
Private Sub txt_evento_GotFocus()
            If txt_evento.Text = "(ex: Natal, Ano novo, etc...)" Then
                txt_evento = Empty
            End If
End Sub
Private Sub abrir()
            Call fechar
            tabparede.Open "pparede", con, adOpenKeyset, adLockOptimistic
End Sub
Private Sub fechar()
            If tabparede.State = 1 Then tabparede.Close
End Sub
Private Sub gravar()
            On Error Resume Next
            If desvio = "xD" Then GoTo a
                status = "salvas"
''''''''Testes posteriores a gravação do evento'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If dtp_inicio > dtp_fim Then
                MsgBox "A data de inicio não pode ser posterior a data de fim", vbInformation, "Arbimy Manager 2.0"
                Exit Sub
            End If
            If txt_evento = Empty Or txt_evento = "(ex: Natal, Ano novo, etc...)" Then
                MsgBox "Este não é um evento válido, por favor verificar", vbInformation, "Arbimy Manager 2.0"
                Exit Sub
            End If
                tabparede.Close
                tabparede.Open "select * from pparede where evento = '" & txt_evento & "'"
                    If tabparede.RecordCount <> 0 Then
                        MsgBox "Este evento já esta cadastrado, verifique na lista de eventos ou cadastre um novo evento", vbInformation, "Arbimy Manager 2.0"
                        Exit Sub
                    End If
                        tabparede.Close
                        
'                        tabparede.Open "select * from pparede where inicio = '" & dtp_inicio & "'"
'                        tabparede.Open "select * from pparede where inicio >= " & Date & " and fim >= " & Date & ""
                        tabparede.Open "select * from pparede where inicio >= '" & dtp_inicio.value & "' and fim >= '" & dtp_fim & "'"
                        If tabparede.RecordCount > 0 Then
                            MsgBox "Já existe um evento cadastrado neste intervalo de data", vbInformation, "Arbimy Manager 2.0"
                            Exit Sub
                        End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            tabparede.Close
            tabparede.Open "select * from pparede where evento = '" & txt_evento & "'"
                If tabparede.RecordCount = 0 Then
                tabparede.AddNew
                End If
a:
                    tabparede!inicio = dtp_inicio.value
                    tabparede!fim = dtp_fim.value
                    tabparede!evento = txt_evento
                    tabparede!imagem = txt_imagem
                    tabparede.Update
                    Call box
                    FileCopy CommonDialog1.FileName, App.Path & "\Imagens\" & imagem
               
End Sub
Private Sub leventos()
            Call abrir
                tabparede.Close
                tabparede.Open "select * from pparede where fim <  " & Date & ""
'                tabparede.Open "select * from pparede where imagem = " & a
                    If tabparede.RecordCount > 0 Then
                        con.Execute "delete * from pparede where fim < '" & Date & "'"
                    ElseIf tabparede.RecordCount = 0 Then
                        Exit Sub
                    End If
End Sub

Private Sub txt_evento_LostFocus()
            If txt_evento.Text = Empty Then
                txt_evento.Text = "(ex: Natal, Ano novo, etc...)"
            End If
End Sub
Private Sub mostrar()
            txt_imagem = tabparede!imagem
            txt_evento = tabparede!evento
            dtp_inicio.value = tabparede!inicio
            dtp_fim.value = tabparede!fim
            imagem = tabparede!imagem
            imagem = App.Path & "\Imagens\" & imagem & ""
            img_pparede.Picture = LoadPicture(imagem)

End Sub
Private Sub carregar_lista()
            
        On Error Resume Next
            
            Call abrir
                If tabparede.BOF = False Or tabparede.EOF = False Then
                    tabparede.MoveFirst
                    mfg_eventos.Rows = 2
                    mfg_eventos.Clear
                    mfg_eventos.FormatString = "Código        |Evento                          |Início            |Fim                 |"
                Do Until tabparede.EOF
                
                    mfg_eventos.TextMatrix(mfg_eventos.Rows - 1, 0) = tabparede!cod
                    mfg_eventos.TextMatrix(mfg_eventos.Rows - 1, 1) = tabparede!evento
                    mfg_eventos.TextMatrix(mfg_eventos.Rows - 1, 2) = tabparede!inicio
                    mfg_eventos.TextMatrix(mfg_eventos.Rows - 1, 3) = tabparede!fim
                        mfg_eventos.Rows = mfg_eventos.Rows + 1
                        tabparede.MoveNext
                Loop
                    mfg_eventos.Rows = mfg_eventos.Rows - 1
                Else
                    mfg_eventos.Rows = 2
                    mfg_eventos.Clear
                    mfg_eventos.FormatString = "Evento             |Início        |Fim                |"
                End If
                
                
                
End Sub

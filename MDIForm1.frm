VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Arbimy manager"
   ClientHeight    =   10815
   ClientLeft      =   -120
   ClientTop       =   630
   ClientWidth     =   15240
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   10695
      Left            =   0
      ScaleHeight     =   10635
      ScaleWidth      =   15180
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   0
         Top             =   1440
      End
      Begin VB.Image Image1 
         Height          =   4200
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   4215
      End
   End
   Begin VB.Menu mnu_cadastro 
      Caption         =   "Cadastro"
      Begin VB.Menu mnu_clientes 
         Caption         =   "Clientes"
      End
   End
   Begin VB.Menu mnu_config 
      Caption         =   "Configurações"
      Begin VB.Menu mnu_pparede 
         Caption         =   "Papel de parede"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mensagem As String
Dim nome As String
Dim tabcli As New ADODB.Recordset
Dim tabparede As New ADODB.Recordset
Option Explicit

Private Sub MDIForm_Load()
            
            Call abrir_banco
            Call configu
            Call apagar_velha
            Call aplicar
            
End Sub
Private Sub mnu_clientes_Click()
            frm_clientes.Show
            
End Sub
Private Sub MDIForm_Resize()
            Call carrega_imagem
End Sub

Private Sub Timer1_Timer()
            If Left(Time$, 2) > 18 And Left(Time$, 2) < 24 Then mensagem = " - Boa Noite"
            If Left(Time$, 2) >= 0 And Left(Time$, 2) < 12 Then mensagem = " - Bom Dia"
            If Left(Time$, 2) > 12 And Left(Time$, 2) < 18 Then mensagem = " - Boa Tarde"
            MDIForm1.Caption = "Arbimy manager 2.0 - " & Date & "  " & mensagem
End Sub
Private Sub apagar_velha()
           Call abrir
                tabparede.Close
                tabparede.Open "select * from pparede where fim <  " & Date + 5 & ""
                    If tabparede.RecordCount > 0 Then
                        con.Execute "delete * from pparede where fim < '" & Date + 5 & "'"
                    ElseIf tabparede.RecordCount = 0 Then
                        Exit Sub
                    End If
End Sub
Private Sub aplicar()
            Call abrir
            tabparede.Close
            tabparede.Open "select * from pparede where inicio <= " & Date & " "
                If tabparede.RecordCount = 1 Then
                    nome = tabparede!imagem
                    img = App.Path & "\Imagens\" & nome & ""
                        tabparede.Close
                        tabparede.Open "select * from pparede where fim >= " & Date & ""
                            If tabparede.RecordCount = 0 Then
                                Call carrega_imagem
                                Exit Sub
                            End If
                ElseIf tabparede.RecordCount = 0 Then
                End If
            Call imagem
End Sub
Private Sub fechar()
            If tabparede.State = 1 Then tabparede.Open
End Sub
Private Sub abrir()
            Call configu
            Call fechar
            tabparede.Open "pparede", con, adOpenKeyset, adLockOptimistic
End Sub

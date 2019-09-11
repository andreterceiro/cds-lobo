VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ricardo Lobo lista generator Tabajara versão 0.000000000001"
   ClientHeight    =   7680
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7245
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Saída - Local e nome do arquivo gerado"
      Height          =   3255
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   6975
      Begin VB.CheckBox chkSubdiretorios 
         Caption         =   "Incluir subdiretórios"
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   840
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.OptionButton optTexto 
         Caption         =   "Texto"
         Height          =   255
         Left            =   4560
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optHTML 
         Caption         =   "HTML"
         Height          =   255
         Left            =   5640
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkNumeros 
         Caption         =   "Incluir números na lista"
         Height          =   375
         Left            =   3840
         TabIndex        =   12
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   4440
         TabIndex        =   10
         Text            =   "lista"
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CommandButton cmdTexto 
         Caption         =   "Gerar arquivo"
         Height          =   615
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2400
         Width           =   2655
      End
      Begin VB.DriveListBox drvSaida 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   3135
      End
      Begin VB.DirListBox dirSaida 
         Height          =   2115
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo:"
         Height          =   375
         Left            =   3840
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nome:"
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   1920
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Entrada - Arquivos a serem incluídos na lista"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.DriveListBox drvEntrada 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
      Begin VB.DirListBox dirEntrada 
         Height          =   2115
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   2535
      End
      Begin VB.FileListBox filEntrada 
         Height          =   2625
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   3615
      End
      Begin VB.CheckBox chkFiltrar 
         Caption         =   "Filtro:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txtFiltrar 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Text            =   "*.mp3;*.wav;*.cda"
         Top             =   3240
         Width           =   5055
      End
   End
   Begin VB.Menu cmdArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu cmdSair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu cmdAjuda 
      Caption         =   "&Ajuda"
      Begin VB.Menu cmdTopicos 
         Caption         =   "Tópicos de ajuda"
         Shortcut        =   {F1}
      End
      Begin VB.Menu cmdSobre 
         Caption         =   "Sobre"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strDrvEntradaAnterior As String
Dim strDrvSaidaAnterior As String

Private Sub chkFiltrar_Click()
   Filtrar
End Sub

Private Sub cmdSair_Click()
   End
End Sub

Private Sub cmdSobre_Click()
   frmAbout.Show
End Sub

Private Sub cmdTexto_Click()
   Dim booVerificacao As Boolean
   Dim strCaminho As String
   Dim strExtensao As String
   Dim strTipo As String
   Dim strBarra As String
   
   'Verificando se o diretório escolhido tem arquivos
   If filEntrada.ListCount = 0 And chkSubdiretorios.Value = 0 Then
      MsgBox "A lista de entrada está em branco. Verifique se não há um filtro incorreto aplicado", vbInformation, "Diretório vazio"
      Exit Sub
   End If
   
   
   If optHTML.Value = True Then
      strExtensao = ".htm"
      strTipo = "html"
   Else
      strExtensao = ".txt"
      strTipo = "texto"
   End If
   
   If Right(dirSaida.Path, 1) = "\" Then
      strBarra = ""
   Else
      strBarra = "\"
   End If
   strCaminho = dirSaida.Path & strBarra & txtNome.Text & strExtensao
  
   'strCaminho = "c:\lsit.txt"
   If Dir(strCaminho) <> "" Then
      If MsgBox("O arquivo '" & strCaminho & "' já existe. Deseja sobreescrever?", vbDefaultButton2 + vbYesNo + vbQuestion, "Arquivo já existe") = vbYes Then
         If chkSubdiretorios.Value = 0 Then
            GerarArquivo strCaminho, strTipo
         Else
            GerarArquivoHierarquico strCaminho, strTipo
         End If
      End If
   Else
      If chkSubdiretorios.Value = 0 Then
         GerarArquivo strCaminho, strTipo
      Else
         GerarArquivoHierarquico strCaminho, strTipo
      End If
   End If
End Sub

Private Sub cmdTopicos_Click()
   MsgBox "Sabia que eu tinha esquecido de algo... da ajuda..." & Chr(10) & Chr(10) & "Favor mande um e-mail para andreterceiro@yahoo.com.br para que eu tire sua dúvida.", vbExclamation, "Puuuuuutz..."
End Sub

Private Sub DirEntrada_Change()
   filEntrada.Path = dirEntrada.Path
End Sub

Private Sub drvEntrada_Change()
   On Error GoTo TrataErro
   dirEntrada.Path = drvEntrada.Drive
   strDrvEntradaAnterior = drvEntrada.Drive
   
TrataErro:
   If Err.Number = 68 Then
      MsgBox "Deu merda... o dispositivo está inacessível. Se for drive de CD, o CD tá no porta copos? Se for drive de disquete, tem disquete lá?", vbCritical, "i i i i i i..."
      drvEntrada.Drive = strDrvEntradaAnterior
   End If
End Sub

Private Sub drvSaida_Change()
   On Error GoTo TrataErro
   dirSaida.Path = drvSaida.Drive
   strDrvSaidaAnterior = drvSaida.Drive
   
TrataErro:
   If Err.Number = 68 Then
      MsgBox "Deu merda... o dispositivo está inacessível. Se for drive de CD, o CD tá no porta copos? Se for drive de disquete, tem disquete lá?", vbCritical, "i i i i i i..."
      drvSaida.Drive = strDrvSaidaAnterior
   End If
End Sub

Sub Filtrar()
   If chkFiltrar.Value = 1 Then
      filEntrada.Pattern = txtFiltrar.Text
   Else
      filEntrada.Pattern = "*.*"
   End If
End Sub

Private Sub Form_activate()
   strDrvEntradaAnterior = drvEntrada.Drive
   strDrvSaidaAnterior = drvSaida.Drive
End Sub

Private Sub txtFiltrar_Change()
   Filtrar
End Sub

Sub GerarArquivo(caminho As String, tipo As String)
   Dim cont As Double
   Open caminho For Output As #1
   
   On Error GoTo TrataErro
   
   If tipo = "html" Then
      Dim strHTML As String
      strHTML = "<!--Gerado por RicardoLoboListaGenerator Tabajara versão 0.000000000001--><html><head><style type=text/css>body{font-family:tahoma,arial,helvetica;}</style></head><body>"
      Print #1, strHTML
      
      If chkNumeros.Value = 0 Then
         For cont = 0 To filEntrada.ListCount - 1
            Print #1, filEntrada.List(cont) & "<br>"
         Next
      Else
         For cont = 0 To filEntrada.ListCount - 1
            Print #1, cont & "- " & filEntrada.List(cont) & "<br>"
         Next
      End If
      
      strHTML = "</body></html>"
      Print #1, strHTML
   Else 'tipo="texto"
      If chkNumeros.Value = 0 Then
         For cont = 0 To filEntrada.ListCount - 1
            Print #1, filEntrada.List(cont)
         Next
      Else
         For cont = 0 To filEntrada.ListCount - 1
            Print #1, cont & "- " & filEntrada.List(cont)
         Next
      End If
   End If
   Close #1
   MsgBox caminho & " gerado com sucesso!", vbInformation, "Arquivo gerado"
   
TrataErro:
   If Err.Number = 75 Then
      MsgBox "O acesso para escrita ao local indicado não é permitido", vbExclamation, "Erro"
   ElseIf Err.Number <> 0 Then
      MsgBox "Ocorreu um erro ao tentar realizar a operação", vbExclamation, "Erro"
   End If
   
End Sub

Sub GerarArquivoHierarquico(caminho As String, tipo As String)
   Dim cont As Integer
   Dim recont As Integer
   Dim indiceDoDiretorio(1 To 255) As Integer
   Dim booMostrar As Boolean
   
   On Error GoTo TrataErro
   
   booMostrar = True
      
   For cont = 1 To 255
      indiceDoDiretorio(cont) = 0
   Next
   cont = 1
   
   Open caminho For Output As #1
         
   If tipo = "html" Then
      Dim strHTML As String
      strHTML = "<!--Gerado por RicardoLoboListaGenerator Tabajara versão 0.000000000001--><html><head><style type=text/css>body{font-family:tahoma,arial,helvetica;}</style></head><body>"
      Print #1, strHTML
      
      If chkNumeros = 1 Then
        While True
           'indiceDoDiretorio(cont) = dirEntrada.ListCount
           If booMostrar = True Then
              Print #1, "<u><b>" & dirEntrada.Path & "</u></b><br>"
              For recont = 0 To filEntrada.ListCount - 1
                 Print #1, recont + 1 & "- " & filEntrada.List(recont) & "<br>"
              Next
              Print #1, "<br><br>"
           End If
           booMostrar = True
           'If dirEntrada.ListCount > 0 Then
           If dirEntrada.ListCount > indiceDoDiretorio(cont) Then
              dirEntrada.Path = dirEntrada.List(indiceDoDiretorio(cont))
              indiceDoDiretorio(cont) = indiceDoDiretorio(cont) + 1
              cont = cont + 1
           Else
              indiceDoDiretorio(cont) = 0
              cont = cont - 1
              If cont = 0 Then
                 Close #1
                 MsgBox caminho & " gerado com sucesso!", vbInformation, "Arquivo gerado"
                 Exit Sub
              End If
              dirEntrada.Path = ".."
              booMostrar = False
           End If
        Wend
      Else
        While True
           'indiceDoDiretorio(cont) = dirEntrada.ListCount
           If booMostrar = True Then
              Print #1, "<u><b>" & dirEntrada.Path & "</u></b><br>"
              For recont = 0 To filEntrada.ListCount - 1
                 Print #1, filEntrada.List(recont) & "<br>"
              Next
              Print #1, "<br><br>"
           End If
           booMostrar = True
           'If dirEntrada.ListCount > 0 Then
           If dirEntrada.ListCount > indiceDoDiretorio(cont) Then
              dirEntrada.Path = dirEntrada.List(indiceDoDiretorio(cont))
              indiceDoDiretorio(cont) = indiceDoDiretorio(cont) + 1
              cont = cont + 1
           Else
              indiceDoDiretorio(cont) = 0
              cont = cont - 1
              If cont = 0 Then
                 Close #1
                 MsgBox caminho & " gerado com sucesso!", vbInformation, "Arquivo gerado"
                 Exit Sub
              End If
              dirEntrada.Path = ".."
              booMostrar = False
           End If
        Wend
      End If
      strHTML = "</body></html>"
      Print #1, strHTML
   Else 'tipo=texto
      If chkNumeros = 1 Then
        While True
           'indiceDoDiretorio(cont) = dirEntrada.ListCount
           If booMostrar = True Then
              Print #1, dirEntrada.Path
              For recont = 0 To filEntrada.ListCount - 1
                 Print #1, "  " & recont + 1 & "- " & filEntrada.List(recont)
              Next
              Print #1, ""
           End If
           booMostrar = True
           'If dirEntrada.ListCount > 0 Then
           If dirEntrada.ListCount > indiceDoDiretorio(cont) Then
              dirEntrada.Path = dirEntrada.List(indiceDoDiretorio(cont))
              indiceDoDiretorio(cont) = indiceDoDiretorio(cont) + 1
              cont = cont + 1
           Else
              indiceDoDiretorio(cont) = 0
              cont = cont - 1
              If cont = 0 Then
                 Close #1
                 MsgBox caminho & " gerado com sucesso!", vbInformation, "Arquivo gerado"
                 Exit Sub
              End If
              dirEntrada.Path = ".."
              booMostrar = False
           End If
        Wend
      Else
        While True
           'indiceDoDiretorio(cont) = dirEntrada.ListCount
           If booMostrar = True Then
              Print #1, dirEntrada.Path
              For recont = 0 To filEntrada.ListCount - 1
                 Print #1, "  " & filEntrada.List(recont)
              Next
              Print #1, ""
           End If
           booMostrar = True
           'If dirEntrada.ListCount > 0 Then
           If dirEntrada.ListCount > indiceDoDiretorio(cont) Then
              dirEntrada.Path = dirEntrada.List(indiceDoDiretorio(cont))
              indiceDoDiretorio(cont) = indiceDoDiretorio(cont) + 1
              cont = cont + 1
           Else
              indiceDoDiretorio(cont) = 0
              cont = cont - 1
              If cont = 0 Then
                 Close #1
                 MsgBox caminho & " gerado com sucesso!", vbInformation, "Arquivo gerado"
                 Exit Sub
              End If
              dirEntrada.Path = ".."
              booMostrar = False
           End If
        Wend
      End If
   End If
   
TrataErro:
   If Err.Number = 75 Then
      MsgBox "O acesso para escrita ao local indicado não é permitido", vbExclamation, "Erro"
   Else
      MsgBox "Ocorreu um erro ao tentar realizar a operação", vbExclamation, "Erro"
   End If
   
End Sub

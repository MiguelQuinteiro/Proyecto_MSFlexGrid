VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrincipal 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   13350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDirectorio 
      Caption         =   "Directorio"
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   9600
      TabIndex        =   11
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtCampo4 
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Text            =   "true"
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtCampo3 
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Text            =   "17/08/2019"
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtCampo2 
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Text            =   "Miguel"
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtCampo1 
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Text            =   "10"
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Carga los Datos"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3975
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7011
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.Label Label4 
      Caption         =   "Campo 01"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Campo 01"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Campo 01"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Campo 01"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDirectorio_Click()
  Dim miLinea As String
  Dim miContador As Integer

  Open "Prueba.txt" For Input As #1
  miContador = 0
  Do While Not EOF(1) And miContador < 20
    Line Input #1, miLinea
    List1.AddItem miLinea
    miContador = miContador + 1
  Loop
  Close #1
End Sub

' Al cargar el formulario
Private Sub Form_Load()
' Configura el MSFlexGrid
  Call ConfiguraMSFlexGrid(MSFlexGrid1, 2, 4)
End Sub

' Cargar registro en el MSFlexGrid
Private Sub Command1_Click()
' Declara las variables globales
  Dim miRegistro As tRegistro
  ' Asigna los textbox al tRegistro
  Call LeeTextBox(miRegistro)
  ' Muestra los datos del tRegistro en el MSFlexGrid
  Call CargaFG(MSFlexGrid1, miRegistro)
  ' Aumenta en 1 para diferenciar los refistros
  txtCampo1.Text = Str(Val(txtCampo1.Text) + 1)
  ' Cambia  el color de fila
  Call CambiaColorFila(MSFlexGrid1, MSFlexGrid1.Row)
End Sub

' Al hacer click sobre el MSFlexGrid
Private Sub MSFlexGrid1_Click()
' Declara las variables globales
  Dim miRegistro As tRegistro
  ' Descarga el MSFlexgrid en el registro
  Call DescargaFG(MSFlexGrid1, miRegistro, MSFlexGrid1.Row)
  ' Asigna los textbox al tRegistro
  Call EscribeTextBox(miRegistro)
End Sub

' Al hacer dobleclick sobre el MSFlexGrid
Private Sub MSFlexGrid1_DblClick()
  Call ActivaDesactiva(MSFlexGrid1, MSFlexGrid1.Row)
  Call MSFlexGrid1_Click
  ' Cambia  el color de fila
  Call CambiaColorFila(MSFlexGrid1, MSFlexGrid1.Row)
End Sub

' Botón de Prueba
Private Sub Command2_Click()
  MSFlexGrid1.BackColorSel = vbYellow
  MSFlexGrid1.BackColorFixed = vbRed
  MSFlexGrid1.Refresh
End Sub



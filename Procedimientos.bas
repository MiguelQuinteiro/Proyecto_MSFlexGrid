Attribute VB_Name = "Procedimientos"

' Configura el MSFlexGrid
Public Sub ConfiguraMSFlexGrid(ByRef FG As MSFlexGrid, ByVal fil As Integer, ByVal col As Integer)
  FG.Rows = fil
  FG.Cols = col
  Dim x As Integer
  Dim y As Integer
  'For x = 0 To fil
  For y = 0 To col - 1
    FG.ColWidth(y) = 1100
    FG.ColAlignment(y) = 4
  Next y
  'Next x
  ' Colores del MSFlexgrid
  FG.BackColorBkg = vbCyan
  'FG.BackColorSel = vbWhiteBlue
End Sub

' Carga el tRegistro en el MSFlexGrid
Public Sub CargaFG(ByRef FG As MSFlexGrid, ByRef dato As tRegistro)
' Mueve a la siguiente fila
  FG.Rows = FG.Rows + 1
  ' Inserta el dato
  FG.col = 0
  FG.Text = dato.campo1
  FG.col = 1
  FG.Text = dato.campo2
  FG.col = 2
  FG.Text = dato.campo3
  FG.col = 3
  FG.Text = dato.campo4
  ' Agrega una fila
  FG.Row = FG.Row + 1
End Sub

' Descarga el MSFlexGrid en el tRegistro
Public Sub DescargaFG(ByRef FG As MSFlexGrid, ByRef dato As tRegistro, ByRef fil As Integer)
  FG.Row = fil
  FG.col = 0
  dato.campo1 = FG.Text
  FG.col = 1
  dato.campo2 = FG.Text
  FG.col = 2
  dato.campo3 = FG.Text
  FG.col = 3
  dato.campo4 = FG.Text
End Sub

' Activa y desactiva el registro
Public Sub ActivaDesactiva(ByRef FG As MSFlexGrid, ByRef fil As Integer)
  FG.Row = fil
  FG.col = 3
  If FG.Text = True Then
    FG.Text = False
  Else
    FG.Text = True
  End If
End Sub

' Cambia el color de la fila en la grilla
Public Sub CambiaColorFila(ByRef FG As MSFlexGrid, ByRef fil As Integer)
  Dim miColor As Variant
  FG.col = 3
  If FG.Text <> "" Then
    If FG.Text = True Then
      miColor = vbGreen
    Else
      miColor = vbRed
    End If
    Dim x As Integer
    For x = 0 To FG.Cols - 1
      FG.col = x
      FG.CellBackColor = miColor
    Next x
  End If
End Sub

' Lee los textbox y almacena en el tRegistro
Public Sub LeeTextBox(ByRef dato As tRegistro)
  dato.campo1 = Val(frmPrincipal.txtCampo1.Text)
  dato.campo2 = frmPrincipal.txtCampo2.Text
  dato.campo3 = frmPrincipal.txtCampo3.Text
  dato.campo4 = frmPrincipal.txtCampo4.Text
End Sub

' Escribe en los textbox lo que está almacenado en el tRegistro
Public Sub EscribeTextBox(ByRef dato As tRegistro)
  frmPrincipal.txtCampo1.Text = dato.campo1
  frmPrincipal.txtCampo2.Text = dato.campo2
  frmPrincipal.txtCampo3.Text = dato.campo3
  frmPrincipal.txtCampo4.Text = dato.campo4
End Sub


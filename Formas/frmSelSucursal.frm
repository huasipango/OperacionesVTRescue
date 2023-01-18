VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmSelSucursal 
   Caption         =   "Selección de Sucursales"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin FPSpread.vaSpread spBodegas 
      Height          =   4455
      Left            =   120
      OleObjectBlob   =   "frmSelSucursal.frx":0000
      TabIndex        =   2
      Top             =   1200
      Width           =   4215
   End
   Begin VB.CheckBox chktodas 
      Caption         =   "Todas las plazas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Continuar"
            Object.ToolTipText     =   "Continuar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelSucursal.frx":104C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelSucursal.frx":498E6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSelSucursal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sucurs As Integer
Dim opcion_todas As Integer, servip As String

Private Sub Form_Load()
On Error GoTo ERR:
   spBodegas.MaxRows = 0
   opcion_todas = 0
   sqls = "SELECT * FROM Derechos With (Nolock) Where Modulo='MST'"
   sqls = sqls & " And Usuario='" & Trim(gstrUsuario) & "'"
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   If rsBD.EOF Then
      sqls = "Select B.* from Bodegas B With (Nolock)"
      sqls = sqls & " Inner Join Usuarios U With (Nolock) on U.Bodega=B.Bodega"
      sqls = sqls & " Where U.Usuario='" & Trim(gstrUsuario) & "'"
      sqls = sqls & " And U.Status=1"
      Set rsBD2 = New ADODB.Recordset
      rsBD2.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
      If Not rsBD2.EOF Then
         opcion_todas = 0
         sucurs = rsBD2!Bodega
         llena_bodegas
      Else
         opcion_todas = 0
         sucurs = 1
         llena_bodegas
      End If
   Else
     sqls = "SELECT * FROM USUARIOS WITH (NOLOCK) WHERE USUARIO='" & Trim(gstrUsuario) & "'"
     Set rsBD2 = New ADODB.Recordset
     rsBD2.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
     If Not rsBD2.EOF Then
        If rsBD2!Bodega = 1 Then ' son todas las sucursales
           opcion_todas = 1
           llena_bodegas
        Else
           sqls = "SELECT SERVIDOR FROM BODEGAS WITH (NOLOCK) WHERE BODEGA=" & Val(rsBD2!Bodega)
           Set rsBD3 = New ADODB.Recordset
           rsBD3.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
           If Not rsBD3.EOF Then
              servip = UCase(rsBD3!Servidor)
           Else
              servip = UCase(rsBD3!Servidor)
           End If
           opcion_todas = 2
           llena_bodegas
        End If
     Else
        MsgBox "Numero de Empleado no existe", vbCritical, "Empleado no válido"
     End If
   End If
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Sub llena_bodegas()
Dim i As Integer, j As Double
If opcion_todas = 0 Then
   sql = "SELECT * FROM BODEGAS WITH (NOLOCK) WHERE Bodega=" & Val(sucurs)
ElseIf opcion_todas = 1 Then
   sql = "SELECT * FROM BODEGAS WITH (NOLOCK)"
ElseIf opcion_todas = 2 Then
   sql = "SELECT * FROM BODEGAS WITH (NOLOCK) WHERE SERVIDOR='" & servip & "'"
End If
sql = sql & " Order by Bodega"
Set Consultaint = New ADODB.Recordset
Consultaint.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly

If Not Consultaint.EOF Then
   With spBodegas
     .Col = -1
     .Row = -1
     .Action = 12
     .MaxRows = 0
     i = 1

     Do While Not Consultaint.EOF
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Consultaint!descripcion
        .Col = 2
        .Text = Consultaint!Bodega
        i = i + 1
        Consultaint.MoveNext
     Loop
   End With
   Call chktodas_Click
Else
  MsgBox "Su Numero de Empleado no tiene asignado una sucursal", vbCritical, "Empleado no válido"
  Exit Sub
End If
End Sub

Private Sub chktodas_Click()
Dim j As Integer
If chktodas.value = 1 Then
   For j = 1 To spBodegas.MaxRows
       spBodegas.Row = j
       spBodegas.Col = 3
       spBodegas.value = 1
   Next
End If
If chktodas.value = 0 Then
   For j = 1 To spBodegas.MaxRows
       spBodegas.Row = j
       spBodegas.Col = 3
       spBodegas.value = 0
   Next
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
       Case "Salir": Unload Me
       Case "Continuar"
            Call Seleccion
            If plazasuc <> "" Then
               Unload Me
               frmRepDispersion.Show
            End If
End Select
End Sub

Sub Seleccion()
Dim j As Integer, band As Boolean
On Error GoTo ERR:
band = False
If spBodegas.MaxRows > 0 Then
   For j = 1 To spBodegas.MaxRows
       spBodegas.Col = 3
       spBodegas.Row = j
       If spBodegas.value = 1 Then
          band = True
          Exit For
       End If
   Next
   If band = False Then
      plazasuc = ""
      MsgBox "No ha seleccionado ninguna sucursal para consultar", vbExclamation, "Falta seleccionar una sucursal"
      Exit Sub
   End If
End If

    plazasuc = ""
    For j = 1 To spBodegas.MaxRows
        spBodegas.Col = 3
        spBodegas.Row = j
        If spBodegas.value = 1 Then
           spBodegas.Col = 2
           plazasuc = plazasuc & Val(spBodegas.Text) & ","
        End If
    Next
    If Len(plazasuc) > 0 Then
       plazasuc = Mid(plazasuc, 1, Len(plazasuc) - 1)
    End If
ERR:
Exit Sub
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

VERSION 5.00
Begin VB.Form frmLibro 
   Caption         =   "Agrega un libro"
   ClientHeight    =   11520
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9375
   LinkTopic       =   "Form2"
   ScaleHeight     =   11520
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   15
      Top             =   10080
      Width           =   1935
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   14
      Top             =   10080
      Width           =   1935
   End
   Begin VB.TextBox txtPrestadoA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   8160
      Width           =   4815
   End
   Begin VB.CheckBox chkPrestado 
      Caption         =   "Prestado actualmente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   7320
      Width           =   4815
   End
   Begin VB.CheckBox chkRecomendado 
      Caption         =   "Recomendado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   6360
      Width           =   4815
   End
   Begin VB.CheckBox chkPorLeer 
      Caption         =   "Quiero Leer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   5280
      Width           =   4815
   End
   Begin VB.CheckBox chkLeido 
      Caption         =   "Ya leído"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   4200
      Width           =   4815
   End
   Begin VB.TextBox txtCalificacion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   3120
      Width           =   735
   End
   Begin VB.ComboBox cboGenero 
      Height          =   315
      Left            =   3360
      TabIndex        =   4
      Top             =   2400
      Width           =   4815
   End
   Begin VB.TextBox txtAutor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   1440
      Width           =   4815
   End
   Begin VB.TextBox txtTitulo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label5 
      Caption         =   "Prestado a:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Calificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Genero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Autor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Título"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmLibro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EditandoID As Integer

Private Sub chkLeido_Click()
    If chkLeido.Value = 1 Then
        chkPorLeer.Value = 0
        txtCalificacion.Enabled = True
    Else
        txtCalificacion.Enabled = False
    End If
End Sub

Private Sub chkPorLeer_Click()
    If chkPorLeer.Value = 1 Then
        chkLeido.Value = 0
    End If
Option Explicit
Public EditandoID As Integer



Private Sub chkPrestado_Click()
    If chkPrestado.Value = 1 Then
        txtPrestadoA.Enabled = True
    Else
        txtPrestadoA.Enabled = False
        txtPrestadoA.Text = ""
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    ' Valiadación básica
    
    If Trim(txtTitulo.Text) = "" Or Trim(txtAutor.Text) = "" Then
        MsgBox "El titulo y el autor son obligatorios", vbExclamation, "Datos incompletos"
        Exit Sub
    End If
    
    If cboGenero.ListIndex = -1 Then
        MsgBox "Selecciona un género", vbExclamation, "Datos incompletos"
        Exit Sub
    End If
    
    If chkLeido.Value = 1 And Trim(txtCalificacion.Text) = "" Then
        MsgBox "Por favor ingrese una calificación (1-5) para el libro leído", vbInformation
    End If
    
    ' Validar calificación 1-5
    Dim calif As Variant
    If Trim(txtCalificacion.Text) <> "" Then
        calif = Val(txtCalificacion.Text)
        If (calif < 1 Or calif > 5) Then
            MsgBox "Calificación debe ser un número del 1 al 5.", vbExclamation
            Exit Sub
        End If
    Else
        calif = "NULL"
    End If
    
    ' Preparar los datos
    Dim titulo As String, autor As String, generoID As Long
    titulo = Replace(txtTitulo.Text, "'", "''")
    autor = Replace(txtAutor.Text, "'", "''")
    generoID = cboGenero.ItemData(cboGenero.ListIndex)
    
    Dim leido As Integer, porLeer As Integer, recom As Integer, prestado As Integer
    leido = IIf(chkLeido.Value = 1, 1, 0)
    porLeer = IIf(chkPorLeer.Value = 1, 1, 0)
    recom = IIf(chkRecomendado.Value = 1, 1, 0)
    prestado = IIf(chkPrestado.Value = 1, 1, 0)
    
    Dim prestadoA As String, fechaPrestamo As String
    If prestado = 1 Then
        prestadoA = Replace(txtPrestadoA.Text, "'", "''")
        fechaPrestamo = Format$(Now, "yyyy-mm-dd")
    Else
        prestadoA = ""
        fechaPrestamo = ""
    End If
    
    On Error GoTo ErrSave
        
        ' Falta lo de EditandoID
        Dim sqlInsert As String
        sqlInsert = "INSERT INTO Libros (Titulo, Autor, GeneroID, Calificacion, Leido, PorLeer, Recomendado, Prestado, PrestadoA, FechaPrestamo) VALUES ('" & titulo & "', '" & autor & "', " & CStr(generoID) & ", "
        
        If calif = "NULL" Then
            sqlInsert = sqlInsert & "NULL"
        Else
            sqlInsert = sqlInsert & CStr(calif)
        End If
            
        sqlInsert = sqlInsert & ", " & CStr(leido) & ", " & CStr(porLeer) & ", " & CStr(recom) & ", " & CStr(prestado) & ", "
        
        If prestado = 1 Then
            sqlInsert = sqlInsert & "'" & prestadoA & "', '" & fechaPrestamo & "')"
        Else
            sqlInsert = sqlInsert & "NULL, NULL)"
        End If
        
        MsgBox sqlInsert, vbInformation
        
        conn.Execute sqlInsert
        MsgBox "Libro agregado exitosamente", vbInformation
    
ErrSave:
    MsgBox "Ocuirió un error al guardar: " & Err.Description, vbCritical
    
    
End Sub

Private Sub Form_Load()
    Dim rsG As ADODB.Recordset
    Set rsG = New ADODB.Recordset
    rsG.Open "SELECT GeneroID, Nombre FROM Generos ORDER BY Nombre", conn, adOpenStatic, adLockReadOnly
    cboGenero.Clear
    Do Until rsG.EOF
        cboGenero.AddItem rsG!Nombre
        cboGenero.ItemData(cboGenero.NewIndex) = rsG!generoID
        rsG.MoveNext
    Loop
    
    rsG.Close: Set rsG = Nothing
    
    If EditandoID = 0 Then
        ' Modo agregar, limpiar campos
        txtTitulo.Text = ""
        txtAutor = ""
        cboGenero.ListIndex = -1 ' no hay nada seleccionado
        txtCalificacion = ""
        chkLeido.Value = 0
        txtPrestadoA.Enabled = False
        Me.Caption = "Agregar Libro"
    Else
    
    End If
    
End Sub

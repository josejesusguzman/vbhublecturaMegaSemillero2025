VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15690
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   15690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnEditar 
      Caption         =   "Editar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   9
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton btn_favoritos 
      Caption         =   "Libros favoritos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton btn_generos_favoritos 
      Caption         =   "Generos favoritos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton btn_no_gustar 
      Caption         =   "No me gustaron"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton btn_quiero 
      Caption         =   "Quiero leer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton btn_agregar 
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   6600
      Width           =   1815
   End
   Begin MSComctlLib.ListView list_libros 
      Height          =   6015
      Left            =   3120
      TabIndex        =   3
      Top             =   360
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   10610
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Height          =   7935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      Begin VB.CommandButton btn_leiste 
         Caption         =   "Ya leíste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton btn_catalogo 
         Caption         =   "Cátalogo MEGA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CargarLibros(filtroSQL As String)
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT L.LibroID, L.Titulo, L.Autor, G.Nombre As Genero, L.Calificacion, L.Prestado, L.PrestadoA FROM Libros L INNER JOIN Generos G ON L.GeneroID = G.GeneroID"
        
    If filtroSQL <> "" Then
        sql = sql & " WHERE " & filtroSQL
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    list_libros.ListItems.Clear
    
    If Not rs.EOF Then
        Dim item As ListItem
        Do Until rs.EOF
        
            Set item = list_libros.ListItems.Add(, , rs!titulo)
            item.SubItems(1) = rs!autor
            item.SubItems(2) = rs!Genero
            item.SubItems(3) = IIf(IsNull(rs!Calificacion), "", rs!Calificacion)
            If rs!prestado = True Then
                item.SubItems(4) = rs!prestadoA
            Else
                item.SubItems(4) = ""
            End If
            
            item.Tag = rs!LibroID
            
            rs.MoveNext
        Loop
    End If
    
    rs.Close: Set rs = Nothing
    
End Sub

Private Sub btn_agregar_Click()
    frmLibro.EditandoID = 0
    frmLibro.Show vbModal
End Sub

Private Sub btn_catalogo_Click()
    CargarLibros ""
End Sub


Private Sub btn_favoritos_Click()
    CargarLibros "L.Recomendado = 1"
End Sub

Private Sub btn_generos_favoritos_Click()
    CargarLibros "G.EsFavorito = 1"
End Sub

Private Sub btn_leiste_Click()
    CargarLibros "L.Leido = 1"
End Sub

Private Sub btn_no_gustar_Click()
    CargarLibros "L.Leido = 1 AND L.Calificacion <= 2"
End Sub

Private Sub btn_quiero_Click()
    CargarLibros "L.PorLeer = 1"
End Sub

Private Sub btnEditar_Click()
    frmLibro.EditandoID = list_libros.SelectedItem.Tag
    frmLibro.Show vbModal
End Sub

Private Sub Form_Load()
    Set conn = New ADODB.Connection
    conn.CursorLocation = adUseClient
    
    Dim connString As String
    connString = "Provider=SQLOLEDB.1;Data Source=.\MYSERVER;Initial Catalog=LibreriaMega;Integrated Security=SSPI;"
        
        
    conn.Open connString
    
    With list_libros
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Título", 2000
        .ColumnHeaders.Add , , "Autor", 1500
        .ColumnHeaders.Add , , "Género", 1000
        .ColumnHeaders.Add , , "Calificación", 800
        .ColumnHeaders.Add , , "Prestado a", 1500
    End With
    
End Sub

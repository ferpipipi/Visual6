VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7740
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   14595
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBuscar 
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   6840
      Width           =   3975
   End
   Begin VB.CommandButton btn_eliminar 
      BackColor       =   &H008080FF&
      Caption         =   "Eliminar libro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   2175
   End
   Begin MSComctlLib.ListView list_libros 
      Height          =   5895
      Left            =   3240
      TabIndex        =   3
      Top             =   360
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   10398
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
   Begin VB.CommandButton btn_leiste 
      Caption         =   "Leido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      Begin VB.CommandButton btn_recomendados 
         Caption         =   "Recomendados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Top             =   6240
         Width           =   1695
      End
      Begin VB.CommandButton btn_gen_fav 
         Caption         =   "Generos Favoritos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Top             =   5040
         Width           =   1695
      End
      Begin VB.CommandButton btn_no_gustaron 
         Caption         =   "No te gustaron"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Top             =   3840
         Width           =   1695
      End
      Begin VB.CommandButton btn_quiero 
         Caption         =   "Quiero leer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton btn_catalogo 
         Caption         =   "Catalogo MEGA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Busqueda por autor o libro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   6480
      Width           =   3975
   End
   Begin VB.Menu btnarchivo 
      Caption         =   "Archivo"
      Begin VB.Menu btnagregar 
         Caption         =   "Agregar libro"
         Shortcut        =   ^N
      End
      Begin VB.Menu btnedit 
         Caption         =   "Editar libro"
         Shortcut        =   ^E
      End
      Begin VB.Menu btnsalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu btnacerca 
      Caption         =   "Acerca de..."
      Begin VB.Menu btndesarrollador 
         Caption         =   "Desarrolladores"
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
    
    sql = "SELECT L.LibroID, L.Titulo, L.Autor, G.Nombre As Genero, L.Calificacion, L.Prestado, L.PrestadoA FROM Libros L INNER JOIN GENEROS G ON L.GeneroID = G.GeneroID"
    
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
                item.SubItems(4) = rs!PrestadoA
            Else
                item.SubItems(4) = ""
            End If
            
            item.Tag = rs!libroID
            rs.MoveNext
        
        Loop
    End If

    rs.Close: Set rs = Nothing

    
    
End Sub

Private Sub Archivo_Click()

End Sub



Private Sub btn_catalogo_Click()
    CargarLibros ""
End Sub

Private Sub btn_eliminar_Click()
    Dim item As ListItem
    Set item = list_libros.SelectedItem
    
    If item Is Nothing Then
        MsgBox "Selecciona algo para eliminar", vbExclamation
        Exit Sub
    End If
    
    Dim titulo As String
    titulo = item.Text
    Dim resp As Integer
    resp = MsgBox("¿Estas seguro de eliminar el libro '" & titulo & "'?", vbYesNo + vbQuestion, "Confirmar eliminacion")
    
    If resp = vbYes Then
        Dim libroID As Long
        libroID = item.Tag
        On Error GoTo ErrorDelete
        conn.Execute "DELETE FROM Libros WHERE LibroID =" & CStr(libroID)
        MsgBox "Libro eliminado", , vbInformation
        CargarLibros ""
    End If
    Exit Sub
        
ErrorDelete:
MsgBox "Error al eliminar el libro" & Err.Description, vbCritical


End Sub

Private Sub btn_gen_fav_Click()
    CargarLibros "G.EsFavorito = 1"
End Sub

Private Sub btn_leiste_Click()
    CargarLibros "L.Leido = 1"
End Sub

Private Sub btn_no_gustaron_Click()
CargarLibros "L.Leido = 1 AND L.Calificacion <= 5"
End Sub

Private Sub btn_quiero_Click()
CargarLibros "L.PorLeer = 1"
End Sub

Private Sub btn_recomendados_Click()
CargarLibros "L.Recomendado = 1"
End Sub

Private Sub btnagregar_Click()
    Frm_libro.EditandoID = 0
    Frm_libro.Show vbModal
End Sub



Private Sub btndesarrollador_Click()
    MsgBox _
        "   Desarrolladores: Humberto Rivera y Cynthia Diaz" & vbCrLf & _
        "   Proyecto: HUBLectura" & vbCrLf & _
        "   Fecha: Junio 2025", _
        vbInformation + vbOKOnly, _
        "Información del Desarrollador"
End Sub

Private Sub btnedit_Click()
    Frm_libro.EditandoID = list_libros.SelectedItem.Tag
    Frm_libro.Show vbModal
End Sub

Private Sub btnsalir_Click()
    End
End Sub


Private Sub Form_Load()
    Set conn = New ADODB.Connection
    conn.CursorLocation = adUseClient
    
    Dim connString As String
    connString = "Provider=SQLOLEDB.1;Data Source =LAPTOP-TEOUBOL8;Initial Catalog=LibreriaMega; Integrated Security=SSPI;"
        
        
    conn.Open connString
    
    With list_libros
         .View = lvwReport
         .GridLines = True
         .FullRowSelect = True
         .ColumnHeaders.Clear
         .ColumnHeaders.Add , , "Titulo", 2000
         .ColumnHeaders.Add , , "Autor", 1500
         .ColumnHeaders.Add , , "Genero", 1000
         .ColumnHeaders.Add , , "Calificacion", 1500
         .ColumnHeaders.Add , , "Prestado a", 1500
         
    End With
    
End Sub

Private Sub txtBuscar_Change()
    Dim texto As String
    texto = Replace(txtBuscar.Text, "'", "''")
    
    If Trim(texto) = "" Then
        CargarLibros ""
    Else
        CargarLibros "L.Titulo LIKE '%" & texto & "%' OR L.Autor LIKE '%" & texto & "%'"
    End If
End Sub


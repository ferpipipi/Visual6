VERSION 5.00
Begin VB.Form Frm_libro 
   Caption         =   "Agregar libro"
   ClientHeight    =   10875
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10845
   LinkTopic       =   "Form2"
   ScaleHeight     =   10875
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   15
      Top             =   9480
      Width           =   2415
   End
   Begin VB.TextBox TxtPrestadoA 
      Height          =   615
      Left            =   3720
      TabIndex        =   13
      Top             =   7800
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Prestamo"
      Height          =   2175
      Left            =   360
      TabIndex        =   11
      Top             =   7080
      Width           =   10095
      Begin VB.CheckBox ChkPrestado 
         Caption         =   "Prestado actualmente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   12
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label PrestadoA 
         Caption         =   "Prestado a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.CheckBox ChkRecomendado 
      Caption         =   "Lo recomiendo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   10
      Top             =   6000
      Width           =   3375
   End
   Begin VB.CheckBox ChkQuieroLeer 
      Caption         =   "Quiero leerlo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   9
      Top             =   5040
      Width           =   3375
   End
   Begin VB.CheckBox ChkLeido 
      Caption         =   "Ya lo leí"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   8
      Top             =   3960
      Width           =   3375
   End
   Begin VB.TextBox TxtCalificacion 
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin VB.ComboBox CboGenero 
      Height          =   315
      Left            =   3480
      TabIndex        =   4
      Text            =   "Genero"
      Top             =   2880
      Width           =   6735
   End
   Begin VB.TextBox TxtAutor 
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   1680
      Width           =   6735
   End
   Begin VB.TextBox TxtTitulo 
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Calificacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   7
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      Height          =   735
      Left            =   2280
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Frm_libro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EditandoID As Integer

Private Sub ChkLeido_Click()
    If ChkLeido.Value = 1 Then
        ChkQuieroLeer.Value = 0
        TxtCalificacion.Enabled = True
    Else
        TxtCalificacion.Enabled = False
    End If
End Sub

Private Sub ChkPrestado_Click()
    If ChkPrestado.Value = 1 Then
        TxtPrestadoA.Enabled = True
    Else
        TxtPrestadoA.Enabled = False
        TxtPrestadoA.Text = ""
    End If
        
End Sub

Private Sub ChkQuieroLeer_Click()
    If ChkQuieroLeer.Value = 1 Then
        ChkLeido.Value = 0
    End If
End Sub



Private Sub cmdGuardar_Click()
    If Trim(TxtTitulo.Text) = "" Or Trim(TxtAutor.Text) = "" Then
        MsgBox "El título y el autor son obligatorios", vbExclamation, "Datos incompletos"
        Exit Sub
    End If
    
    If CboGenero.ListIndex = -1 Then
        MsgBox "Selecciona un género", vbExclamation, "Datos incompletos"
        Exit Sub
    End If

    ' Validación de calificación si fue leído
    If ChkLeido.Value = 1 Then
        If Trim(TxtCalificacion.Text) = "" Then
            MsgBox "Ingresa una calificación (1 a 10) para el libro leído", vbInformation
            Exit Sub
        End If
        If Not IsNumeric(TxtCalificacion.Text) Then
            MsgBox "La calificación debe ser un número entre 1 y 10", vbExclamation
            Exit Sub
        End If
        If Val(TxtCalificacion.Text) < 1 Or Val(TxtCalificacion.Text) > 10 Then
            MsgBox "La calificación debe estar entre 1 y 10", vbExclamation
            Exit Sub
        End If
    End If

    ' Variables de entrada
    Dim calif As Variant
    If Trim(TxtCalificacion.Text) <> "" And IsNumeric(TxtCalificacion.Text) Then
        calif = Val(TxtCalificacion.Text)
    Else
        calif = "NULL"
    End If

    Dim titulo As String, autor As String, generoID As Long
    titulo = Replace(TxtTitulo.Text, "'", "''")
    autor = Replace(TxtAutor.Text, "'", "''")
    generoID = CboGenero.ItemData(CboGenero.ListIndex)

    Dim leido As Integer, porLeer As Integer, recom As Integer, prestado As Integer
    leido = IIf(ChkLeido.Value = 1, 1, 0)
    porLeer = IIf(ChkQuieroLeer.Value = 1, 1, 0)
    recom = IIf(ChkRecomendado.Value = 1, 1, 0)
    prestado = IIf(ChkPrestado.Value = 1, 1, 0)

    Dim PrestadoA As String, fechaPrestamo As String
    If prestado = 1 Then
        PrestadoA = Replace(TxtPrestadoA.Text, "'", "''")
        fechaPrestamo = Format$(Now, "yyyy-mm-dd")
    Else
        PrestadoA = ""
        fechaPrestamo = ""
    End If

    On Error GoTo ErrSave

    If EditandoID = 0 Then
        ' INSERTAR NUEVO LIBRO
        Dim sqlInsert As String
        sqlInsert = "INSERT INTO Libros (Titulo, Autor, GeneroID, Calificacion, Leido, PorLeer, Recomendado, Prestado, PrestadoA, FechaPrestamo) VALUES ('" & titulo & "', '" & autor & "', " & CStr(generoID) & ", "

        If calif = "NULL" Then
            sqlInsert = sqlInsert & "NULL"
        Else
            sqlInsert = sqlInsert & CStr(calif)
        End If

        sqlInsert = sqlInsert & ", " & CStr(leido) & ", " & CStr(porLeer) & ", " & CStr(recom) & ", " & CStr(prestado)

        If prestado = 1 Then
            sqlInsert = sqlInsert & ", '" & PrestadoA & "', '" & fechaPrestamo & "')"
        Else
            sqlInsert = sqlInsert & ", NULL, NULL)"
        End If

        conn.Execute sqlInsert
        MsgBox "Libro agregado correctamente", vbInformation

    Else
        ' ACTUALIZAR LIBRO EXISTENTE
        Dim sqlUpdate As String
        sqlUpdate = "UPDATE Libros SET " & _
                    "Titulo = '" & titulo & "', " & _
                    "Autor = '" & autor & "', " & _
                    "GeneroID = " & CStr(generoID) & ", " & _
                    "Calificacion = "

        If calif = "NULL" Then
            sqlUpdate = sqlUpdate & "NULL, "
        Else
            sqlUpdate = sqlUpdate & CStr(calif) & ", "
        End If

        sqlUpdate = sqlUpdate & _
                    "Leido = " & CStr(leido) & ", " & _
                    "PorLeer = " & CStr(porLeer) & ", " & _
                    "Recomendado = " & CStr(recom) & ", " & _
                    "Prestado = " & CStr(prestado) & ", "

        If prestado = 1 Then
            sqlUpdate = sqlUpdate & _
                "PrestadoA = '" & PrestadoA & "', " & _
                "FechaPrestamo = '" & fechaPrestamo & "' "
        Else
            sqlUpdate = sqlUpdate & _
                "PrestadoA = NULL, " & _
                "FechaPrestamo = NULL "
        End If

        sqlUpdate = sqlUpdate & "WHERE LibroID = " & EditandoID

        conn.Execute sqlUpdate
        MsgBox "Libro actualizado correctamente", vbInformation
    End If

    Unload Me
    Exit Sub

ErrSave:
    MsgBox "Ocurrió un error al guardar: " & Err.Description, vbCritical
End Sub



Private Sub Form_Load()
    Dim rsG As ADODB.Recordset
    Set rsG = New ADODB.Recordset
    rsG.Open "SELECT GeneroID, Nombre FROM GENEROS ORDER BY Nombre", conn, adOpenStatic, adLockReadOnly
    CboGenero.Clear
    Do Until rsG.EOF
        CboGenero.AddItem rsG!Nombre
        CboGenero.ItemData(CboGenero.NewIndex) = rsG!generoID
        rsG.MoveNext
    Loop
    
    rsG.Close: Set rsG = Nothing
    
    
    
    If EditandoID = 0 Then
    
        TxtTitulo.Text = ""
        TxtAutor.Text = ""
        CboGenero.ListIndex = -1
        TxtCalificacion = ""
        ChkLeido.Value = 0
        TxtPrestadoA.Enabled = False
        Me.Caption = "Agregar Libro"
    Else
    
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Dim resp As Integer
            resp = MsgBox("¿Seguro que quieres salir sin guardar?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar salir")
            If resp = vbNo Then
                Cancel = True
            End If
    End If
End Sub

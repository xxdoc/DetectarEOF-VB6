VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Detectar EOF by Blau"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   4320
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtEOF 
      Enabled         =   0   'False
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton btnBuscar 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtArchivo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBuscar_Click()
    With CD
        .DialogTitle = "Selecciona un archivo"
        .Filter = "EXE|*.exe"
        .ShowOpen
    End With
    If CD.FileName <> vbNullString Then
        txtArchivo.Text = CD.FileName
    End If
    Dim sEOF As String
    sEOF = ReadEOFData(txtArchivo.Text)
    If sEOF = vbNullString Then
        txtEOF.Text = "No se ha detectado EOF"
    Else
        txtEOF.Text = "Se ha detectado EOF:" & vbCrLf
        txtEOF.Text = txtEOF.Text & sEOF
    End If
End Sub

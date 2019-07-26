VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Frmlauncher 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frmlauncher.frx":0000
   ScaleHeight     =   2430
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4680
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   8160
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   2566
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Frmlauncher.frx":F58E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ONLINE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   240
      Picture         =   "Frmlauncher.frx":F612
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2880
      Picture         =   "Frmlauncher.frx":128EC
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "Frmlauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Sub Command1_Click()
Dim tx As String
tx = GetVar(App.path & "\INIT\Update.ini", "INIT", "X")
Label1.Caption = tx
End Sub

Private Sub Command2_Click()
Dim ix As String
ix = Inet1.OpenURL("http://hispanoao.ucoz.net/Update.txt")
Label1.Caption = ix
End Sub

Private Sub Command3_Click()
Analizar
End Sub

Private Sub Image1_Click()


If EnProceso Then Exit Sub
Call addConsole("Conectando...", 0, 0, 0, True, True) '>> Informacion
EnProceso = True
Analizar 'Iniciamos la función Analizar =).

'Call Main
End Sub

Private Sub Image2_Click()
End
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''

        Function Analizar()
            On Error Resume Next
           
            Dim ix As Integer
            Dim tx As Integer
            Dim DifX As Integer
            Dim strsX As String
           
'LINK1            'Variable que contiene el numero de actualización correcto del servidor
                ix = Inet1.OpenURL("http://hispanoao.ucoz.net/Update.txt")
                Label1.Caption = ix
            'Variable que contiene el numero de actualización del cliente
                tx = GetVar(App.path & "\INIT\Update.ini", "INIT", "X")
            'Variable con la diferencia de actualizaciones servidor-cliente
                DifX = ix - tx
           
            If Not (DifX = 0) Then 'Si la diferencia no es nula,
            Call addConsole("Iniciando, se descargarán " & DifX & " actualizaciones.", 255, 255, 255, True, True)   '>> Informacion
                For i = 1 To DifX 'Descargamos todas las versiones de diferencia
'LINK2
                    strURL = "http://hispanoao.ucoz.net/Parche" & CStr(i + tx) & ".zip" 'URL del parche .zip
                    Darchivo = App.path & "\INIT\Parche" & i + tx & ".zip" 'Directorio del parche
                        Call addConsole("   Descargando parche nº " & i, 0, 255, 255, False, True)    '>> Informacion
                    Call AutoDownload(i + tx) 'Descargamos todas las versiones faltantes a partir de la nuestra
                        Call addConsole("   Parche nº " & i & " descargado satisfactoriamente.", 0, 0, 255, True, True)    '>> Informacion
               
                  Call addConsole(" Actualizaciones: " & i & "/" & DifX, 100, 100, 100, True, True)   '>> Informacion
                Next i
            Else
                Call addConsole("No hay actualizaciones pendientes", 0, 0, 0, True, True)    '>> Informacion
            End If
           
           
            'Call WriteVar(App.path & "\INIT\Update.ini", "INIT", "X", CStr(iX)) 'Avisamos al cliente que está actualizado
            WriteVar App.path & "\INIT\Update.ini", "INIT", "X", CStr(ix)
            EnProceso = False
           
            Call addConsole("El cliente ya está listo para jugar", 255, 1, 1, True, True)  '>> Informacion
            sRGY.Picture = sG.Picture
           
            Me.Visible = False
            Call Main
        End Function
       
        Public Sub AutoDownload(Numero As Integer)
            On Error Resume Next
           
            sRGY.Picture = SR.Picture
           
            Inet1.AccessType = icUseDefault
            Dim B() As Byte
           
           
            B() = Inet1.OpenURL(strURL, icByteArray)
           
            'Descargamos y guardamos el archivo
            Open Darchivo For Binary Access _
            Write As #1
            Put #1, , B()
            Close #1
           
            'Informacion
            Call addConsole("   Instalando actualización.", 0, 100, 255, False, False)    '>> Informacion
           
            sRGY.Picture = sY.Picture
           
            'Unzipeamos
            UnZip Darchivo, App.path & "\"
           
            'Borramos el zip
            Kill Darchivo
        End Sub



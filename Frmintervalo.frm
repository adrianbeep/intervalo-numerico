VERSION 5.00
Begin VB.Form Frmintervalo 
   Appearance      =   0  'Flat
   BackColor       =   &H00181818&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intervalo"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4095
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdrei 
      Caption         =   "Reiniciar todo"
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame frabotones 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   3855
      Begin VB.CommandButton cmdreiniciar 
         Caption         =   "Reiniciar"
         Height          =   495
         Left            =   1380
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdcalc 
         Caption         =   "Calcular Intervalo"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2640
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdnum3 
         Caption         =   "Tercer numero"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   3840
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame franum3 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   3855
      Begin VB.TextBox txtnum3 
         Appearance      =   0  'Flat
         BackColor       =   &H00636363&
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   840
         MaxLength       =   6
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tercer numero:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame franumeros 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3855
      Begin VB.TextBox txtnum2 
         Appearance      =   0  'Flat
         BackColor       =   &H00636363&
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtnum1 
         Appearance      =   0  'Flat
         BackColor       =   &H00636363&
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Segundo numero:"
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Primer numero:"
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraresultado 
      Appearance      =   0  'Flat
      BackColor       =   &H00181818&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   3855
      Begin VB.Label lblintervalo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "500 esta en el intervalo entre 100 y 1500"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   3615
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   3960
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Intervalo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Frmintervalo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim num1 As Long, num2 As Long, num3 As Long
Private Sub cmdcalc_Click()
    ' verifica que num3 sea un numero luego lo transforma.
    ' si no lo bloquea
    If (IsNumeric(txtnum3.Text)) Then
        num3 = CLng(txtnum3.Text)
        If (num3 > 32768) Or (num3 < -32767) Then
            MsgBox "El numero no puede ser mayor a 32,768 o menor a -32,767.", vbCritical, "Aviso"
            txtnum3.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Solo puede ingresar caracteres numericos y signo negativo.", vbCritical, "Aviso"
        txtnum3.SetFocus
        Exit Sub
    End If
    ' evalua si esta dentro del intervalo de num1 y num2.
    ' si no evalua si num3 es igual a num1 o num2.
    ' si no devuelve que num3 no esta dentro del intervalo.
    If (num3 > num1) And (num3 < num2) Then
        lblintervalo.Caption = num3 & " esta dentro del intervalo de " & num1 & " y " & num2 & "."
    Else
        If (num3 = num1) Or (num3 = num2) Then
            MsgBox "El tercer numero ingresado no puede ser igual a ninguno de los numeros ingresados anteriormente: " & num1 & " , " & num2 & ".", vbCritical, "Aviso"
            txtnum3.SetFocus
            Exit Sub
        Else
            lblintervalo.Caption = num3 & " esta fuera del intervalo de " & num1 & " y " & num2 & "."
        End If
    End If
    
    franumeros.Visible = False
    franum3.Visible = False
    fraresultado.Visible = True
    cmdcalc.Enabled = False
    Frmintervalo.Height = 4245
    cmdreiniciar.Enabled = True
    cmdreiniciar.TabStop = True
    cmdrei.TabStop = True
    cmdcerrar.TabStop = True
    cmdreiniciar.SetFocus
End Sub
Private Sub cmdcerrar_Click()
    Unload Me
End Sub
Private Sub cmdnum3_Click()
    ' si txtnum1 es numerico transforma en long, luego evaluo si el numero esta en valores de integer y si esta fuera lo bloqueo.
    ' si no es numerico tambien bloqueo.
    If (IsNumeric(txtnum1.Text)) Then
        num1 = CLng(txtnum1.Text)
        If (num1 > 32768) Or (num1 < -32767) Then
            MsgBox "El numero no puede ser mayor a 32,768 o menor a -32,767.", vbCritical, "Aviso"
            txtnum1.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Solo puede ingresar caracteres numericos y signo negativo.", vbCritical, "Aviso"
        txtnum1.SetFocus
        Exit Sub
    End If
    ' si txtnum2 es numerico transforma en long, luego evaluo si el numero esta en valores de integer y si esta fuera lo bloqueo.
    ' si no es numerico tambien bloqueo.
    If (IsNumeric(txtnum2.Text)) Then
        num2 = CLng(txtnum2.Text)
        If (num2 > 32768) Or (num2 < -32767) Then
            MsgBox "El numero no puede ser mayor a 32,768 o menor a -32,767.", vbCritical, "Aviso"
            txtnum2.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Solo puede ingresar caracteres numericos y signo negativo.", vbCritical, "Aviso"
        txtnum2.SetFocus
        Exit Sub
    End If
    ' si num1 es mayor a num2 no te permite pasar al ingreso del tercer numero.
    ' bloquea si no hay un numero entero de distancia entre num1 y num2.
    ' bloquea si num1 y num2 son iguales.
    If (num1 > num2) Then
        MsgBox "El primer numero debe ser menor al segundo numero", vbCritical, "Aviso"
        txtnum1.SetFocus
        Exit Sub
    Else
        If ((num1 = num2 - 1)) Then
            MsgBox "Entre el primer numero y el segundo debe haber un numero de distancia. Ej.: 2 y 4.", vbCritical, "Aviso"
            txtnum2.SetFocus
            Exit Sub
        Else
            If (num1 = num2) Then
                MsgBox "El primer numero no puede ser igual al segundo numero.", vbCritical, "Aviso"
                txtnum1.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    franumeros.Visible = False
    franum3.Visible = True
    fraresultado.Visible = False
    cmdnum3.Enabled = False
    cmdnum3.TabStop = False
    cmdcalc.Enabled = True
    cmdcalc.TabStop = True
    cmdcalc.TabIndex = 5
    cmdreiniciar.TabIndex = 6
    txtnum3.Text = ""
    txtnum3.SetFocus
    
'    If (IsNumeric(txtnum1.Text)) Then
'        num1 = CLng(txtnum1.Text)
'    Else
'        MsgBox "Solo puedo ingresar caracteres numericos, puntos y signo negativo.", vbCritical, "Aviso"
'        txtnum1.SetFocus
'        Exit Sub
'    End If
End Sub

Private Sub cmdrei_Click()
    Form_Load
    Form_Activate
    
    lblintervalo.Caption = ""
    txtnum1.Text = ""
    txtnum2.Text = ""
    txtnum3.Text = ""
End Sub
Private Sub cmdreiniciar_Click()
    Form_Load
    Form_Activate

    cmdnum3.Enabled = True
    cmdnum3.TabStop = True
    cmdcalc.Enabled = False
    cmdcalc.TabStop = False
End Sub
Private Sub Form_Activate()
    txtnum1.SetFocus
End Sub
Private Sub Form_Load()
    franumeros.Visible = True
    franum3.Visible = False
    fraresultado.Visible = False
    cmdcalc.Enabled = False
    cmdcalc.TabStop = False
    cmdnum3.TabIndex = 5
    cmdnum3.Enabled = True
    cmdnum3.TabStop = True
    cmdreiniciar.TabIndex = 6
    cmdreiniciar.Enabled = False
    cmdcalc.TabIndex = 7
    Frmintervalo.Height = 3645
    cmdrei.TabStop = False
    cmdcerrar.TabStop = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Frmintervalo = Nothing
End Sub
Private Sub txtnum1_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 45) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtnum2_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 45) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtnum3_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 45) Then
        KeyAscii = 0
    End If
End Sub

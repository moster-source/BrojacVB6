VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleMode       =   0  'User
   ScaleWidth      =   1800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   720
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   300
      Width           =   225
   End
   Begin VB.Label lblPotvrdi 
      AutoSize        =   -1  'True
      Caption         =   "Unesi minute"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "15:00"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   26.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   1410
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private mlngX As Long  'za pomicanje forme pomocu klikanja na label
Private mlngY As Long
Dim sec
Dim min
Dim minEntered 'pamti ukucano vrijeme

Private Sub Form_Activate()
    Call SetWindowPos(Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub Form_Load()
    min = 15
    minEntered = 15
    sec = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Label1_DblClick()
        If (Form1.BackColor = vbRed) Then
            Pocrveni (False)
        End If

    
        Label1.Visible = False
        'Label2.Visible = False 'ovo je exit button
        lblStart.Visible = False
        
        Text1.Visible = True
        lblPotvrdi.Visible = True
        Timer1.Enabled = False
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'za pomicanje forme
    If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'za pomicanje forme
    Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'nova minutaža na desni klik, neke kontrole sakrij
    If Button = 2 Then
        If (Form1.BackColor = vbRed) Then
            Pocrveni (False)
        End If

    
        Label1.Visible = False
        'Label2.Visible = False
        lblStart.Visible = False
        
        Text1.Visible = True
        lblPotvrdi.Visible = True
        Timer1.Enabled = False
    End If
End Sub

Private Sub Label2_Click() 'gasenje programa
    Timer1.Enabled = False
    Unload Me
End Sub


Private Sub lblStart_Click() 'klik na start, i stop
    If (Timer1.Enabled = True) Then  'ako je vec pokrenut stopiraj
        Timer1.Enabled = False
        sec = 0
        min = minEntered
        Exit Sub
    Else                             'ako timer ne radi
        If Me.BackColor = vbRed Then
            Pocrveni (False)
        End If
        
        If min = 0 Then Exit Sub     'izadji ako je start na 0 minuta ili neispravan unos
        
        sec = 59
        min = minEntered - 1
        Label1.Caption = Format(min, "00") + ":" + Format(sec, "00")
        Timer1.Enabled = True
    End If
    
End Sub

Private Sub lblPotvrdi_Click() 'stavi u novu variablu kod reseta
    min = Val(Text1.Text)
    minEntered = min
    sec = 0
    Label1.Caption = Format(min, "00") + ":" + Format(sec, "00")
    
        Label1.Visible = True
        'Label2.Visible = True 'exit button
        lblStart.Visible = True
        
        Text1.Visible = False
        lblPotvrdi.Visible = False

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer) 'key enter za potvrdu vremena

    If KeyAscii = 13 Then
        lblPotvrdi_Click
    End If
    
        If Me.BackColor = vbRed Then
            Pocrveni (False)
        End If
        
End Sub

Private Sub Timer1_Timer() 'svake sekunde pokreni

    If sec = 0 Then
        If min = 1 Then 'beep na minutu
            MessageBeep (MB_DEFAULTBEEP)
        End If
        If min = 0 Then 'beep na kraj
            MessageBeep (MB_DEFAULTBEEP)
            
            min = minEntered
            Label1.Caption = Format(min, "00") + ":" + Format(sec, "00")
            Timer1.Enabled = False
            Pocrveni (True)
            Exit Sub
        End If
    End If
        
        Label1.Caption = Format(min, "00") + ":" + Format(sec, "00") 'ako nije ni kraj ni minuta do kraja
        sec = sec - 1
        If sec = -1 Then
            sec = 59
            min = min - 1
        End If

End Sub



Private Sub Pocrveni(blnCrveni As Boolean) 'pocrveni za kraj ili popravi ako treba ponovo
If blnCrveni = True Then
    Form1.BackColor = vbRed
    Label1.BackColor = vbRed
    Label2.BackColor = vbRed
    lblStart.BackColor = vbRed
    Me.BorderStyle = 2
    Me.Caption = Me.Caption
    Me.BorderStyle = 0
    Me.Caption = Me.Caption
    
Else
    Form1.BackColor = vbButtonFace
    Label1.BackColor = vbButtonFace
    Label2.BackColor = vbButtonFace
    lblStart.BackColor = vbButtonFace
    Me.BorderStyle = 2
    Me.Caption = Me.Caption
    Me.BorderStyle = 0
    Me.Caption = Me.Caption

End If

End Sub





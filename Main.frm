VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program For Beginners"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   3600
      Top             =   2880
   End
   Begin VB.Frame Frame15 
      Height          =   615
      Left            =   6480
      TabIndex        =   59
      Top             =   4800
      Width           =   1695
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Planet Source Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame14 
      Height          =   1335
      Left            =   6480
      TabIndex        =   55
      Top             =   3360
      Width           =   1695
      Begin VB.CommandButton Command19 
         Caption         =   "Wordpad"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Solitaire"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Notepad"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame13 
      Height          =   735
      Left            =   6480
      TabIndex        =   53
      Top             =   2400
      Width           =   1695
      Begin VB.CommandButton Command16 
         Caption         =   "Click Here"
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame12 
      Height          =   735
      Left            =   6480
      TabIndex        =   51
      Top             =   1560
      Width           =   1695
      Begin VB.CommandButton Command15 
         Caption         =   "Mouse Over"
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame11 
      Height          =   615
      Left            =   6480
      TabIndex        =   49
      Top             =   840
      Width           =   1695
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Mouse Over"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame10 
      Height          =   615
      Left            =   6480
      TabIndex        =   47
      Top             =   120
      Width           =   1695
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Mouse Over"
         Height          =   255
         Left            =   120
         MouseIcon       =   "Main.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame9 
      Height          =   2055
      Left            =   4200
      TabIndex        =   40
      Top             =   3360
      Width           =   2055
      Begin VB.CommandButton Command14 
         Caption         =   "Filter Help"
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Font"
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Color"
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Save"
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Open"
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Time"
      Height          =   255
      Left            =   360
      TabIndex        =   36
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame Frame8 
      Height          =   975
      Left            =   120
      TabIndex        =   34
      Top             =   4440
      Width           =   3855
      Begin VB.CommandButton Command9 
         Caption         =   "Both"
         Height          =   255
         Left            =   1680
         TabIndex        =   39
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Date"
         Height          =   255
         Left            =   2400
         TabIndex        =   38
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Date"
         Height          =   255
         Left            =   2160
         TabIndex        =   37
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Time"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame7 
      Height          =   3135
      Left            =   4200
      TabIndex        =   27
      Top             =   120
      Width           =   2055
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   33
         Text            =   "             Text Box"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Text            =   "          Combo Box"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Add"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Text            =   "Enter Text"
         Top             =   2400
         Width           =   1815
      End
      Begin VB.ListBox List2 
         Height          =   840
         ItemData        =   "Main.frx":030A
         Left            =   120
         List            =   "Main.frx":0311
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "                Label"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1560
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "About"
      Height          =   375
      Left            =   720
      TabIndex        =   26
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6120
      TabIndex        =   25
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Frame Frame6 
      Height          =   975
      Left            =   120
      TabIndex        =   24
      Top             =   2280
      Width           =   1815
      Begin VB.CommandButton Command21 
         Caption         =   "Click Here"
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EXTvgamer@Juno.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   0
         TabIndex        =   62
         Top             =   650
         Width           =   1815
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1350
      Left            =   2160
      TabIndex        =   18
      Top             =   120
      Width           =   1815
      Begin VB.CheckBox Check4 
         Caption         =   "UnLine"
         Height          =   195
         Left            =   840
         TabIndex        =   22
         Top             =   480
         Width           =   855
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Strike"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Italic"
         Height          =   195
         Left            =   840
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bold"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hello World"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   3855
      Begin VB.HScrollBar HScroll5 
         Height          =   255
         Left            =   240
         Max             =   100
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         TickStyle       =   3
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1695
      Left            =   2160
      TabIndex        =   10
      Top             =   1560
      Width           =   1815
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Red Background"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Normal Background"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Pick Background"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1815
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   6
         Top             =   240
         Value           =   1
         Width           =   975
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   5
         Top             =   600
         Value           =   1
         Width           =   975
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   4
         Top             =   960
         Value           =   1
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Left            =   1440
         TabIndex        =   46
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   960
         Width           =   135
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************
'*    Beginners Program      *
'*      By Philip Beam       *
'*****************************
'ALL OF THIS SOURCE CODE WAS WRITTEN BY PHILIP BEAM
'-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~
Private Sub Check1_Click()
'makes label5 bold when check1 is checked (value = 1)
    If Check1.Value = 1 Then
        Label5.FontBold = True
        Exit Sub
    End If
    If Check1.Value = 0 Then
        Label5.FontBold = False
        Exit Sub
    End If
End Sub

Private Sub Check2_Click()
'makes label5 italic when check2 is checked (value = 1)
    If Check2.Value = 1 Then
        Label5.FontItalic = True
        Exit Sub
    End If
    If Check2.Value = 0 Then
        Label5.FontItalic = False
        Exit Sub
    End If
End Sub

Private Sub Check3_Click()
'makes label5 Strike Through when check3
'is checked (value = 1)
    If Check3.Value = 1 Then
        Label5.FontStrikethru = True
        Exit Sub
    End If
    If Check3.Value = 0 Then
        Label5.FontStrikethru = False
        Exit Sub
    End If
End Sub

Private Sub Check4_Click()
'makes label5 underlined when check4
'is checked (value = 1)
    If Check4.Value = 1 Then
        Label5.FontUnderline = True
        Exit Sub
    End If
    If Check4.Value = 0 Then
        Label5.FontUnderline = False
        Exit Sub
    End If
End Sub

Private Sub Command1_Click()
'Changes the form's background color to red
    Me.BackColor = &HFF&
End Sub

Private Sub Command10_Click()
'Shows the open dialog
    CommonDialog1.ShowOpen
    MsgBox "You opened " & CommonDialog1.FileName & ""
End Sub

Private Sub Command11_Click()
'Shows the save dialog
    CommonDialog1.ShowSave
    MsgBox "You chose to save as " & CommonDialog1.FileName & ""
End Sub

Private Sub Command12_Click()
'Shows the color dialog
    CommonDialog1.ShowColor
    MsgBox "You picked (VB color code) " & CommonDialog1.Color & ""
End Sub

Private Sub Command13_Click()
'shows the font dialog
    CommonDialog1.ShowFont
    MsgBox "You picked the font " & CommonDialog1.FontName & ""
End Sub

Private Sub Command14_Click()
'This allows only certain file types to be
'viewed in an open or save dialog

'The File Type that you type first will be the
'one that it shows right when you open the dialog

'Do as follows:

'                           |  File Type 1    | ext.|  File Type 2  |ext.
    CommonDialog1.Filter = "Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*"
    CommonDialog1.ShowOpen

'This means that when I first open the dialog text
'files will be the only files being viewed
'but I can also choose to view all files
End Sub

Private Sub Command15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Changes the button caption to "Click Me" when you
'move the mouse over the button.
    Main.Command15.Caption = "Click Me"
End Sub

Private Sub Command16_Click()
    MsgBox "This program was made by Philip Beam on Tuesday, July 25, 2000", vbOKOnly, "Beginners Program"
End Sub

Private Sub Command17_Click()
'Opens notepad on top (not minimized)
    Call Shell("notepad", vbNormalFocus)
End Sub

Private Sub Command18_Click()
'Opens solitaire on top (not minimized)
    Call Shell("sol", vbNormalFocus)
End Sub

Private Sub Command19_Click()
'Opens wordpad on top (not minimized)
    Call Shell("write", vbNormalFocus)
End Sub

Private Sub Command2_Click()
'Changes the form's background color back to default
    Me.BackColor = &H8000000F
End Sub

Private Sub Command21_Click()
'Makes a beep sound
'Like when you hit enter in a textbox
    Beep
End Sub

Private Sub Command3_Click()
'Lets you pick a color for the form's background color
    CommonDialog1.ShowColor
    Me.BackColor = CommonDialog1.Color
End Sub

Private Sub Command4_Click()
'Exit program
    End
End Sub

Private Sub Command5_Click()
'About this program
    MsgBox "This Program was made to help beginners program in Visual Basic." & Chr(13) & "By Philip Beam", vbOKOnly, "Beginners Program"
End Sub

Private Sub Command6_Click()
'Adds Text2.Text to the objects
    List2.AddItem Text2.Text
    Combo1.AddItem Text2.Text
    Text3.Text = Text2.Text
    Label6.Caption = Text2.Text
End Sub

Private Sub Command7_Click()
'Makes label7 the current time
    Main.Label7.Caption = Time
End Sub

Private Sub Command8_Click()
'Makes label8 the current date
    Main.Label8.Caption = Date
End Sub

Private Sub Command9_Click()
'Makes label7 the current time
'and Makes label8 the current date
    Main.Label7.Caption = Time
    Main.Label8.Caption = Date
End Sub

Private Sub Form_Load()
'Shows the loading message (before for appears)
    MsgBox "This program was written for beginner use. Maybe even intermediate use. Notice the clock on the title bar. The things that say Mouse Over you will need to put you cursor over to see what it does." & Chr(13) & "By Philip Beam", vbOKOnly, "Begginer Program"
'Starts the timer that updates the clock on the title bar
    Timer1.Interval = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Main.Command15.Caption = "Mouse Over"
    Main.Label11.ForeColor = &H80000012
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Shows how to use the MsgBox with vbYesNo buttons
  Dim yesno As String
    yesno = MsgBox("Exit Beginner Program?", vbYesNo, "Exit?")
    If yesno = vbYes Then
        Exit Sub
    End If
    Cancel = 1
End Sub

Private Sub HScroll1_Scroll()
    Main.Label1.Caption = Main.HScroll1.Value
End Sub

Private Sub HScroll2_Scroll()
'mixes the 3 scroll bar's values to form one color
    Main.Label10.BackColor = RGB(HScroll2, HScroll3, HScroll4)
'changes label2 to the color value of Hscroll2 (red)
    Main.Label2.BackColor = RGB(HScroll2, 0, 0)
'makes label10's background color the form's
'background color
    Me.BackColor = Label10.BackColor
End Sub

Private Sub HScroll3_Scroll()
'mixes the 3 scroll bar's values to form one color
    Main.Label10.BackColor = RGB(HScroll2, HScroll3, HScroll4)
'changes label3 to the color value of Hscroll3 (green)
    Main.Label3.BackColor = RGB(0, HScroll3, 0)
'makes label10's background color the form's
'background color
    Me.BackColor = Label10.BackColor
End Sub

Private Sub HScroll4_Scroll()
'mixes the 3 scroll bar's values to form one color
    Main.Label10.BackColor = RGB(HScroll2, HScroll3, HScroll4)
'changes label4 to the color value of Hscroll4 (blue)
    Main.Label4.BackColor = RGB(0, 0, HScroll4)
'makes label10's background color the form's
'background color
    Me.BackColor = Label10.BackColor
End Sub

Private Sub HScroll5_Scroll()
    Main.ProgressBar1.Value = Main.HScroll5.Value
    Main.Slider1.Value = Main.HScroll5.Value
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Main.Label11.ForeColor = &HFF&
End Sub

Private Sub Label12_Click()
'Opens the default browser to the planet source code website.
    Call Shell("Start.exe " & "http://www.planetsourcecode.com", 0)
End Sub

Private Sub Label13_Click()
'Opens the default e-mail program and e-mails me
    Call Shell("Start.exe " & "mailto:EXTvgamer@Juno.com", 0)
End Sub

Private Sub Slider1_Scroll()
    ProgressBar1.Value = Slider1.Value
    HScroll5.Value = Slider1.Value
End Sub

Private Sub Timer1_Timer()
'Updates the time on the title bar
    Me.Caption = "Program For Beginners  -  " & Time & ""
End Sub

'-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~
'ALL OF THIS SOURCE CODE WAS WRITTEN BY PHILIP BEAM
'*****************************
'*    Beginners Program      *
'*      By Philip Beam       *
'*****************************



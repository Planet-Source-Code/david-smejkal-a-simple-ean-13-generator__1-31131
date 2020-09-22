VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple EAN"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3600
   ForeColor       =   &H8000000C&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   240
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Tag             =   "Enter 12 digits"
      Text            =   "859200600"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Generate"
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picEan 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000B&
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      ScaleHeight     =   80
      ScaleMode       =   0  'User
      ScaleWidth      =   113
      TabIndex        =   0
      Top             =   1500
      Width           =   3375
      Begin VB.VScrollBar Scroll1 
         Height          =   1755
         LargeChange     =   5
         Left            =   3060
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Label lbCo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   10
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lbCo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lbCo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lbCo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Check Digit:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Item Reference Digits:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Manufacturer Number:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Prefix (ICN):"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1395
   End
   Begin VB.Menu zMenu1 
      Caption         =   "&Ean"
      Begin VB.Menu zEan 
         Caption         =   "&Generate"
         Index           =   0
         Shortcut        =   ^G
      End
      Begin VB.Menu zEan 
         Caption         =   "&Save Ean"
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu zEan 
         Caption         =   "&Print"
         Index           =   2
         Shortcut        =   ^P
      End
      Begin VB.Menu zEan 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu zEan 
         Caption         =   "&Quit"
         Index           =   4
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu zMenu2 
      Caption         =   "&Tools"
      Begin VB.Menu zTool 
         Caption         =   "&Show Digit Modules"
         Index           =   0
         Shortcut        =   ^D
      End
      Begin VB.Menu zTool 
         Caption         =   "S&how Position Modules"
         Index           =   1
         Shortcut        =   ^W
      End
      Begin VB.Menu zTool 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu zTool 
         Caption         =   "&Options"
         Index           =   3
      End
   End
   Begin VB.Menu zMenu3 
      Caption         =   "&About"
      Begin VB.Menu zAbout 
         Caption         =   "...&Simple EAN"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Well, not many functions implemented here. I tried to make it as simple
'as possible even if this app loses any practical usage due it's simplicity,
'exept for the check number calculation and that's just what I needed.
'Be patient with my english,- I am not from this world.
'Any questions? Then mail me: davidsmejkal@hellada.cz

Dim Mdl(2, 9) As String
Dim MdlLeft(9) As String
Dim iMdl As Integer

Private Sub Command1_Click()
Dim i As Integer, m As Integer, n As Integer, s As Integer
On Error GoTo err
If Len(Text1) = 12 Then         'If there's enough digitts for algorithm
    For i = 1 To 11 Step 2      'Sum every number at even position
        m = m + CInt(Right(Left(Text1, i), 1))
    Next i
    For i = 2 To 12 Step 2      'Sum every n° at odd position
        n = n + CInt(Right(Left(Text1, i), 1))
    Next i
    s = 10 - ((n * 3 + m) Mod 10) 'Count the Check digit (s = a number to the nierest multiplicand by 10 from n * 3 + m
    If s = 10 Then s = 0
    lbCo1(0) = Left(Text1, 3)     'International country number
    lbCo1(1) = Left(Right(Text1, Len(Text1) - 3), 6)    'Manufacturer n° (in my case there are 6 digits,- must be from 4 to 6)
    lbCo1(2) = Right(Text1, 3)    'Item reference digit (In my case there are 3 n°s, must be from 3 to 5
    lbCo1(3) = s                  'Check digit
    Text1 = Text1 & lbCo1(3)      'Full EAN code
    Text1.SelStart = Len(Text1) - 1 'make the cursor apear where it's been
 '   Text1.SelLength = 1
    If Scroll1.Visible Then Scroll1.Visible = False
    DrawEan
Else: MsgBox "Enter 12 numbers into text box!", vbExclamation, App.Title
End If
err:
    If err.Number = 13 Then MsgBox "Enter only numbers into text box!", vbExclamation, App.Title    'In case someone puts other characters then numbers into textbox
End Sub

Private Sub Form_Load()
Dim FF As Integer, i As Integer, m As Integer
Dim Str As String
Me.Line (0, 0)-(Me.Width, 0)
i = 0: m = 0
FF = FreeFile(1)
Open App.Path & "\module.txt" For Input As #FF  'All information stored in the txt file
Do While Not EOF(FF)
    Line Input #FF, Str             'Read file line after line
        If m < 3 Then               'There are 3 diferent modules for the number depending on first digit
            Mdl(m, i) = Str
        Else: MdlLeft(i) = Str      'if m > 4 (31'st line), define what module is for the first n°(0-9)
        End If
        i = i + 1                   'Next number(0-9)
        If i = 10 Then
            i = 0: m = m + 1        'First n° (0), next module
        End If
Loop
Close FF
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
Set Form1 = Nothing         'Important!!! = remove form from memory
End Sub

Private Sub Scroll1_Change()
DrawModule Scroll1.Value
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click    'Generate EAN
End Sub

Private Sub DrawEan()
Dim i As Integer, m As Integer, d As Integer, b As Integer, a As Integer
Dim lngX As Long
With picEan
    .Cls
    .BackColor = vbWhite
    .FontSize = 12
    .DrawWidth = 2
    lngX = 11       'X position (11 =must be 11 modules [1 module = usually 0.33 millimeters, in my case picEan.ScaleWidth <bar width> / 113] 11 on left side, 7 on right side
    For i = 1 To 14 '13 digits :-)
        d = CInt(Right(Left(Text1, i), 1))  'Current n°
        If i = 1 Or i = 14 Then             'Draw the guard bars at the begining and end
            picEan.Line (lngX, 5)-(lngX, 72)
            picEan.Line (lngX + 2, 5)-(lngX + 2, 72)
            lngX = lngX + 3
            If i = 1 Then                   'Print first digit
                .CurrentX = 2
                .CurrentY = 66
                picEan.Print d
                b = d                       'Store inf. what's the first n° for the module algorithm
            End If
        Else
            If i < 8 Then                   'On the left side, there are modules 1 or 2 (A, B) depending on the 1st digit = [Mdl(0 - 9, 0 or 1)]...
                a = CInt(Right(Left(MdlLeft(b), i - 1), 1))
            Else: a = 2                     '...on the right side always module 2 (C) = [Mdl(0 - 9, 2)]
            End If
            If i = 8 Then                   'Draw the centre pattern
                picEan.Line (lngX + 1, 5)-(lngX + 1, 72)
                picEan.Line (lngX + 3, 5)-(lngX + 3, 72)
                lngX = lngX + 5
            End If
            For m = 1 To 7                  '7 modules for each n° (System of 7 black or white sprites)
                If CInt(Right(Left(Mdl(a, d), m), 1)) = 1 Then picEan.Line (lngX, 5)-(lngX, 66) 'Draw modules(sprites) for each n°
                lngX = lngX + 1
            Next m
            .CurrentX = lngX - 8
            .CurrentY = 66
            picEan.Print d                  'Print n°s
        End If
    Next i
End With
End Sub

Private Sub zAbout_Click()
MsgBox "Simple EAN 13 Creator, 2002" & vbCr & "Author: David Smejkal & God" & vbCr & "Any questions at: davidsmejkal@hellada.cz" & vbCr & "Copyright: what?", vbInformation, App.Title
End Sub

Private Sub zEan_Click(Index As Integer)
Dim strPath As String
Select Case Index
    Case 0: Command1_Click
    Case 1
        If picEan.BackColor = vbWhite Then              'Only if EAn is drawn
            strPath = App.Path & "\EAN-" & Text1 & ".bmp"
            If Dir(strPath) <> "" Then Kill strPath     'If file exists
            SavePicture picEan.Image, strPath
            MsgBox "Ean saved as: " & Chr(34) & strPath & Chr(34)
        Else: MsgBox "Nothing to save!", vbExclamation, App.Title
        End If
    Case 2
        If picEan.BackColor = vbWhite Then
            PrintEan
        Else: MsgBox "No bar code to print!", vbExclamation, App.Title
        End If
    Case 4: Unload Me
End Select
End Sub

Private Sub PrintEan()
Dim i As Integer
On Error GoTo err
With Printer
    .ColorMode = vbPRCMMonochrome
    .PrintQuality = -2              'Low quality
    .CurrentY = 200
    .CurrentX = 200
    .Font = "Courier New"
    .FontBold = True
    .FontSize = 10
    Printer.Print "EAN 13 Code: " & Text1
    .FontBold = False
    .CurrentY = 600
    For i = 0 To 3
        .CurrentX = 200
        Printer.Print Label1(i).Caption & Space(25 - Len(Label1(i).Caption)) & lbCo1(i).Caption
    Next i
    Printer.PaintPicture picEan.Image, 200, 1600
    Printer.Print
    .EndDoc
    MsgBox "Printing EAN Code: " & Text1 & vbCrLf & "na " & .Port, vbInformation, App.Title
End With
endit:
    Exit Sub
err:
    Printer.KillDoc
    MsgBox "Chyba tisku: " & err.Description, vbExclamation, "Chyba: " & err.Number
    Resume endit
End Sub

Private Sub DrawModule(iStart As Integer)
Dim i As Integer, m As Integer, n As Integer, f As Integer, o As Integer
Dim w As String
Dim tempX As Long
If iMdl = 1 Then
    Scroll1.Max = 5             '10 (0-9) numbers displaying max 5 of them
    picEan.FillColor = vbWhite
Else: Scroll1.Max = 25          '30 (3 x 0-9) numbers
End If
With picEan
    .Cls                        'Clear picture
    .BackColor = &H8000000A
    .FontSize = 8
    .DrawWidth = 1
    n = 0
    For i = iStart To iStart + 4    'To display max. 5 rows
        tempX = 60                  'X position for 2nd column
        .CurrentX = 0
        .CurrentY = n * 16 + 3      'row height = 16
        If iMdl = 0 Then            'module set
            Select Case i           'Get the A,B,C, module for 0-9 number set
                Case Is < 10
                    w = "A": f = 0
                Case Is < 20
                    w = "B": f = 1
                Case Else
                    w = "C": f = 2
            End Select
            picEan.Print Chr(34) & w & Chr(34) & " Module for num.: " & i - f * 10  'first column
            For m = 1 To 7          'second column (each n° contains 7 sprites (modules)
                If CInt(Right(Left(Mdl(f, i - f * 10), m), 1)) = 1 Then
                    .FillColor = vbBlack
                    picEan.Line (tempX, n * 16 + 2)-(tempX + 5, n * 16 + 14), &HE0E0E0, B
                Else
                    .FillColor = vbWhite
                    picEan.Line (tempX, n * 16 + 2)-(tempX + 5, n * 16 + 14), &HE0E0E0, B
                End If
                tempX = tempX + 5
            Next m
        Else                'position set
            picEan.Print "Pos. module for num.: " & i   'First column
            picEan.Line (60, n * 16 + 2)-(100, n * 16 + 14), , B    'Background for second column
            For m = 1 To 7
                If m < 7 Then                           '6 numbers on the left side of the bar code
                    If CInt(Right(Left(Mdl(f, i - f * 10), m), 1)) = 1 Then
                        w = "B"
                    Else: w = "A"
                    End If
                Else: w = " 6×C"                        '6 n° on the right side
                End If
                .CurrentX = tempX + 3
                .CurrentY = n * 16 + 3
                picEan.Print w
                tempX = tempX + 4
            Next m
        End If
        n = n + 1
    Next i
End With
End Sub

Private Sub zTool_Click(Index As Integer)
Select Case Index
    Case 0, 1
        Scroll1.Value = 0
        Scroll1.Visible = True
        iMdl = Index
        DrawModule 0
    Case 3: MsgBox "Not yet implemented.", vbInformation, App.Title
End Select
End Sub

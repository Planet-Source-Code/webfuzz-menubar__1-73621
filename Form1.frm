VERSION 5.00
Begin VB.Form MenuBar 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   -555
   ClientWidth     =   13890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3915
   ScaleWidth      =   13890
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer4 
      Index           =   5000
      Left            =   3060
      Tag             =   "Check Width size"
      Top             =   300
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2100
      Tag             =   "Show Mouse Position"
      Top             =   360
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H000080FF&
      Caption         =   "Back"
      Height          =   435
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton Command17 
      Caption         =   "menu15"
      Height          =   345
      Left            =   7470
      OLEDropMode     =   1  'Manual
      TabIndex        =   22
      Top             =   2040
      Width           =   2265
   End
   Begin VB.CommandButton Command16 
      Caption         =   "menu14"
      Height          =   345
      Left            =   7470
      OLEDropMode     =   1  'Manual
      TabIndex        =   21
      Top             =   1710
      Width           =   2265
   End
   Begin VB.CommandButton Command15 
      Caption         =   "menu13"
      Height          =   345
      Left            =   7470
      OLEDropMode     =   1  'Manual
      TabIndex        =   20
      Top             =   1380
      Width           =   2265
   End
   Begin VB.CommandButton Command14 
      Caption         =   "menu12"
      Height          =   345
      Left            =   5100
      OLEDropMode     =   1  'Manual
      TabIndex        =   19
      Top             =   2370
      Width           =   2265
   End
   Begin VB.CommandButton Command13 
      Caption         =   "menu11"
      Height          =   345
      Left            =   5100
      OLEDropMode     =   1  'Manual
      TabIndex        =   18
      Top             =   2040
      Width           =   2265
   End
   Begin VB.CommandButton Command12 
      Caption         =   "menu10"
      Height          =   345
      Left            =   5100
      OLEDropMode     =   1  'Manual
      TabIndex        =   17
      Top             =   1710
      Width           =   2265
   End
   Begin VB.CommandButton Command11 
      Caption         =   "menu9"
      Height          =   345
      Left            =   5100
      OLEDropMode     =   1  'Manual
      TabIndex        =   16
      Top             =   1380
      Width           =   2265
   End
   Begin VB.CommandButton Command10 
      Caption         =   "menu8"
      Height          =   345
      Left            =   2670
      OLEDropMode     =   1  'Manual
      TabIndex        =   15
      Top             =   2370
      Width           =   2265
   End
   Begin VB.CommandButton Command9 
      Caption         =   "menu7"
      Height          =   345
      Left            =   2670
      OLEDropMode     =   1  'Manual
      TabIndex        =   14
      Top             =   2040
      Width           =   2265
   End
   Begin VB.CommandButton Command8 
      Caption         =   "menu6"
      Height          =   345
      Left            =   2670
      OLEDropMode     =   1  'Manual
      TabIndex        =   13
      Top             =   1710
      Width           =   2265
   End
   Begin VB.CommandButton Command7 
      Caption         =   "menu5"
      Height          =   345
      Left            =   2670
      OLEDropMode     =   1  'Manual
      TabIndex        =   12
      Top             =   1380
      Width           =   2265
   End
   Begin VB.CommandButton Command6 
      Caption         =   "menu4"
      Height          =   345
      Left            =   210
      OLEDropMode     =   1  'Manual
      TabIndex        =   11
      Top             =   2370
      Width           =   2265
   End
   Begin VB.CommandButton Command5 
      Caption         =   "menu3"
      Height          =   345
      Left            =   210
      OLEDropMode     =   1  'Manual
      TabIndex        =   10
      Top             =   2040
      Width           =   2265
   End
   Begin VB.CommandButton Command4 
      Caption         =   "menu2"
      Height          =   345
      Left            =   210
      OLEDropMode     =   1  'Manual
      TabIndex        =   9
      Top             =   1710
      Width           =   2265
   End
   Begin VB.CommandButton Command3 
      Caption         =   "menu1"
      Height          =   345
      Left            =   210
      OLEDropMode     =   1  'Manual
      TabIndex        =   8
      Top             =   1380
      Width           =   2265
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   11700
      Tag             =   "Auto Hide"
      Top             =   60
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1200
      Tag             =   "Mouse movement "
      Top             =   240
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   10140
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   5970
      Width           =   2745
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Save New Item"
      Height          =   375
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2790
      Width           =   9555
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6900
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   5940
      Width           =   3165
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3900
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   870
      Width           =   7215
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   210
      MaxLength       =   29
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   870
      Width           =   3405
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Drag && Drop Here"
      Height          =   585
      Left            =   5550
      OLEDropMode     =   1  'Manual
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5970
      Width           =   1185
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   345
      Left            =   9390
      TabIndex        =   7
      Top             =   4770
      Width           =   2355
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add * to pass parms. EX: *notepad.exe d:\note.txt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3960
      TabIndex        =   6
      Top             =   600
      Width           =   7035
   End
   Begin VB.Menu scott_items 
      Caption         =   "Scott"
      Index           =   0
      Begin VB.Menu mnu_about 
         Caption         =   "v12-04-2010"
      End
      Begin VB.Menu mnu_sep 
         Caption         =   "-"
      End
      Begin VB.Menu reload_menus 
         Caption         =   "Reload Menus"
      End
      Begin VB.Menu stayontop 
         Caption         =   "Stay OnTop"
      End
      Begin VB.Menu DDM 
         Caption         =   "Drag & Drop Mode"
      End
      Begin VB.Menu mnu_view_log 
         Caption         =   "View Log File"
      End
      Begin VB.Menu scott_item1 
         Caption         =   "Exit"
         Index           =   0
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "Menu1"
      Begin VB.Menu Menu1_items 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnu_Menu2 
      Caption         =   "Menu2"
      Begin VB.Menu Menu2_items 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu menu3 
      Caption         =   "Menu3"
      Begin VB.Menu Menu3_items 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu MENU4 
      Caption         =   "MENU4"
      Begin VB.Menu Menu4_items 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu MENU5 
      Caption         =   "MENU5"
      Begin VB.Menu Menu5_items 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu MENU6 
      Caption         =   "MENU6"
      Begin VB.Menu Menu6_items 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu MENU7 
      Caption         =   "MENU7"
      Begin VB.Menu Menu7_items 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu MENU8 
      Caption         =   "MENU8"
      Begin VB.Menu Menu8_items 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu MENU9 
      Caption         =   "MENU9"
      Begin VB.Menu Menu9_items 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu MENU10 
      Caption         =   "MENU10"
      Begin VB.Menu Menu10_items 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu MENU11 
      Caption         =   "MENU11"
      Begin VB.Menu Menu11_items 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu MENU12 
      Caption         =   "MENU12"
      Begin VB.Menu Menu12_items 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu MENU13 
      Caption         =   "MENU13"
      Begin VB.Menu Menu13_items 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu MENU14 
      Caption         =   "MENU14"
      Begin VB.Menu Menu14_items 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu MENU15 
      Caption         =   "MENU15"
      Begin VB.Menu Menu15_items 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu blank1 
      Caption         =   ""
   End
   Begin VB.Menu blank2 
      Caption         =   ""
   End
   Begin VB.Menu blank3 
      Caption         =   ""
   End
   Begin VB.Menu blank4 
      Caption         =   ""
   End
End
Attribute VB_Name = "MenuBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PositionBefore As POINTAPI, PositionNow As POINTAPI
Private TDist As Long
Private BeforeDistance As Long, NowDistance As Long
Private SpeedPixel_Sec As Single, Speedml_Sec As Single

Private Declare Function GetForegroundWindow Lib "user32" () _
        As Long
Private Declare Function GetWindowText Lib "user32" _
    Alias "GetWindowTextA" (ByVal hwnd As Long, _
       ByVal lpString As String, ByVal cch As Long) As Long
'

Dim menu1_caption(150) As String * 29
Dim menu1_command(150) As String
Dim menu2_caption(150) As String * 29
Dim menu2_command(150) As String
Dim menu3_caption(150) As String * 29
Dim menu3_command(150) As String
Dim menu4_caption(150) As String * 29
Dim menu4_command(150) As String
Dim menu5_caption(150) As String * 29
Dim menu5_command(150) As String
Dim menu6_caption(150) As String * 29
Dim menu6_command(150) As String
Dim menu7_caption(150) As String * 29
Dim menu7_command(150) As String
Dim menu8_caption(150) As String * 29
Dim menu8_command(150) As String
Dim menu9_caption(150) As String * 29
Dim menu9_command(150) As String
Dim menu10_caption(150) As String * 29
Dim menu10_command(150) As String
Dim menu11_caption(150) As String * 29
Dim menu11_command(150) As String
Dim menu12_caption(150) As String * 29
Dim menu12_command(150) As String
Dim menu13_caption(150) As String * 29
Dim menu13_command(150) As String
Dim menu14_caption(150) As String * 29
Dim menu14_command(150) As String
Dim menu15_caption(150) As String * 29
Dim menu15_command(150) As String

Dim is_on_top As Boolean







Private Sub EmptyMenu(M As Integer) 'Empty the menu completely but leave the divider (created in design-time)-
    Dim I As Integer
    
    Select Case M
    Case 1
    Menu1_items(0).Visible = True ' Make 'parent' menu item visible.
    For I = 1 To Menu1_items.UBound 'Remove items that were added in runtime
        Unload Menu1_items(I) ' But keep the divider that was created in design-time
    Next I
    Case 2
    Menu2_items(0).Visible = True ' Make 'parent' menu item visible.
    For I = 1 To Menu2_items.UBound 'Remove items that were added in runtime
        Unload Menu2_items(I) ' But keep the divider that was created in design-time
    Next I
     Case 3
    Menu3_items(0).Visible = True ' Make 'parent' menu item visible.
    For I = 1 To Menu3_items.UBound 'Remove items that were added in runtime
        Unload Menu3_items(I) ' But keep the divider that was created in design-time
    Next I
     Case 4
    Menu4_items(0).Visible = True ' Make 'parent' menu item visible.
    For I = 1 To Menu4_items.UBound 'Remove items that were added in runtime
        Unload Menu4_items(I) ' But keep the divider that was created in design-time
    Next I
     Case 5
    Menu5_items(0).Visible = True ' Make 'parent' menu item visible.
    For I = 1 To Menu5_items.UBound 'Remove items that were added in runtime
        Unload Menu5_items(I) ' But keep the divider that was created in design-time
    Next I
     Case 6
    Menu6_items(0).Visible = True ' Make 'parent' menu item visible.
    For I = 1 To Menu6_items.UBound 'Remove items that were added in runtime
        Unload Menu6_items(I) ' But keep the divider that was created in design-time
    Next I
     Case 7
    Menu7_items(0).Visible = True ' Make 'parent' menu item visible.
    For I = 1 To Menu7_items.UBound 'Remove items that were added in runtime
        Unload Menu7_items(I) ' But keep the divider that was created in design-time
    Next I
     Case 8
    Menu8_items(0).Visible = True ' Make 'parent' menu item visible.
    For I = 1 To Menu8_items.UBound 'Remove items that were added in runtime
        Unload Menu8_items(I) ' But keep the divider that was created in design-time
    Next I
     Case 9
    Menu9_items(0).Visible = True ' Make 'parent' menu item visible.
    For I = 1 To Menu9_items.UBound 'Remove items that were added in runtime
        Unload Menu9_items(I) ' But keep the divider that was created in design-time
    Next I
     Case 10
    Menu10_items(0).Visible = True ' Make 'parent' menu item visible.
    For I = 1 To Menu10_items.UBound 'Remove items that were added in runtime
        Unload Menu10_items(I) ' But keep the divider that was created in design-time
    Next I
     Case 11
    Menu11_items(0).Visible = True ' Make 'parent' menu item visible.
    For I = 1 To Menu11_items.UBound 'Remove items that were added in runtime
        Unload Menu11_items(I) ' But keep the divider that was created in design-time
    Next I
     Case 12
    Menu12_items(0).Visible = True ' Make 'parent' menu item visible.
    For I = 1 To Menu12_items.UBound 'Remove items that were added in runtime
        Unload Menu12_items(I) ' But keep the divider that was created in design-time
    Next I
     Case 13
    Menu13_items(0).Visible = True ' Make 'parent' menu item visible.
    For I = 1 To Menu13_items.UBound 'Remove items that were added in runtime
        Unload Menu13_items(I) ' But keep the divider that was created in design-time
    Next I
     Case 14
    Menu14_items(0).Visible = True ' Make 'parent' menu item visible.
    For I = 1 To Menu14_items.UBound 'Remove items that were added in runtime
        Unload Menu14_items(I) ' But keep the divider that was created in design-time
    Next I
     Case 15
    Menu15_items(0).Visible = True ' Make 'parent' menu item visible.
    For I = 1 To Menu15_items.UBound 'Remove items that were added in runtime
        Unload Menu15_items(I) ' But keep the divider that was created in design-time
    Next I
    
    
    End Select
    
    
End Sub



Private Sub FillMenu(M As Integer)
Dim I As Integer
Dim txtline As String
   
Call EmptyMenu(M)  'Clean before adding items
 
 Select Case M
     Case 1
       Open (App.Path + "\menu1.txt") For Input As #1
     Case 2
       Open (App.Path + "\menu2.txt") For Input As #1
     Case 3
       Open (App.Path + "\menu3.txt") For Input As #1
     Case 4
       Open (App.Path + "\menu4.txt") For Input As #1
     Case 5
       Open (App.Path + "\menu5.txt") For Input As #1
     Case 6
       Open (App.Path + "\menu6.txt") For Input As #1
     Case 7
       Open (App.Path + "\menu7.txt") For Input As #1
     Case 8
       Open (App.Path + "\menu8.txt") For Input As #1
     Case 9
       Open (App.Path + "\menu9.txt") For Input As #1
     Case 10
       Open (App.Path + "\menu10.txt") For Input As #1
     Case 11
       Open (App.Path + "\menu11.txt") For Input As #1
     Case 12
       Open (App.Path + "\menu12.txt") For Input As #1
     Case 13
       Open (App.Path + "\menu13.txt") For Input As #1
     Case 14
       Open (App.Path + "\menu14.txt") For Input As #1
     Case 15
       Open (App.Path + "\menu15.txt") For Input As #1
  
       
End Select


Line Input #1, txtline 'blank line
Select Case M

Case 1: mnufile.Caption = txtline: Combo1.AddItem App.Path + "\menu1.txt" ': Combo1.Text = App.Path + "\menu1.txt"
Case 2: mnu_Menu2.Caption = txtline: Combo1.AddItem App.Path + "\menu2.txt"
Case 3: menu3.Caption = txtline: Combo1.AddItem App.Path + "\menu3.txt"
Case 4: MENU4.Caption = txtline: Combo1.AddItem App.Path + "\menu4.txt"
Case 5: MENU5.Caption = txtline: Combo1.AddItem App.Path + "\menu5.txt"
Case 6: MENU6.Caption = txtline: Combo1.AddItem App.Path + "\menu6.txt"
Case 7: MENU7.Caption = txtline: Combo1.AddItem App.Path + "\menu7.txt"
Case 8: MENU8.Caption = txtline: Combo1.AddItem App.Path + "\menu8.txt"
Case 9: MENU9.Caption = txtline: Combo1.AddItem App.Path + "\menu9.txt"
Case 10: MENU10.Caption = txtline: Combo1.AddItem App.Path + "\menu10.txt"
Case 11: MENU11.Caption = txtline: Combo1.AddItem App.Path + "\menu11.txt"
Case 12: MENU12.Caption = txtline: Combo1.AddItem App.Path + "\menu12.txt"
Case 13: MENU13.Caption = txtline: Combo1.AddItem App.Path + "\menu13.txt"
Case 14: MENU14.Caption = txtline: Combo1.AddItem App.Path + "\menu14.txt"
Case 15: MENU15.Caption = txtline: Combo1.AddItem App.Path + "\menu15.txt"
End Select


Line Input #1, txtline 'blank line
Line Input #1, txtline 'blank line
Line Input #1, txtline '*****

done = False

Do  'read text file and add menu items

Line Input #1, txtline

If txtline = "" Then
     done = True
     GoTo done
 Else
I = I + 1
 Select Case M
     Case 1
     menu1_caption(I) = Mid(txtline, 1, 30)
     menu1_command(I) = Mid(txtline, 30, 100)
     
     Case 2
     menu2_caption(I) = Mid(txtline, 1, 30)
     menu2_command(I) = Mid(txtline, 30, 100)
     Case 3
     menu3_caption(I) = Mid(txtline, 1, 30)
     menu3_command(I) = Mid(txtline, 30, 100)
     Case 4
     menu4_caption(I) = Mid(txtline, 1, 30)
     menu4_command(I) = Mid(txtline, 30, 100)
     Case 5
     menu5_caption(I) = Mid(txtline, 1, 30)
     menu5_command(I) = Mid(txtline, 30, 100)
     Case 6
     menu6_caption(I) = Mid(txtline, 1, 30)
     menu6_command(I) = Mid(txtline, 30, 100)
     Case 7
     menu7_caption(I) = Mid(txtline, 1, 30)
     menu7_command(I) = Mid(txtline, 30, 100)
     Case 8
     menu8_caption(I) = Mid(txtline, 1, 30)
     menu8_command(I) = Mid(txtline, 30, 100)
     Case 9
     menu9_caption(I) = Mid(txtline, 1, 30)
     menu9_command(I) = Mid(txtline, 30, 100)
     Case 10
     menu10_caption(I) = Mid(txtline, 1, 30)
     menu10_command(I) = Mid(txtline, 30, 100)
     Case 11
     menu11_caption(I) = Mid(txtline, 1, 30)
     menu11_command(I) = Mid(txtline, 30, 100)
     Case 12
     menu12_caption(I) = Mid(txtline, 1, 30)
     menu12_command(I) = Mid(txtline, 30, 100)
     Case 13
     menu13_caption(I) = Mid(txtline, 1, 30)
     menu13_command(I) = Mid(txtline, 30, 100)
     Case 14
     menu14_caption(I) = Mid(txtline, 1, 30)
     menu14_command(I) = Mid(txtline, 30, 100)
     Case 15
     menu15_caption(I) = Mid(txtline, 1, 30)
     menu15_command(I) = Mid(txtline, 30, 100)
     
     
    End Select
     
End If

Loop Until (EOF(1)) Or done = True
done:
max_items = I

Close #1

    For I = 1 To max_items ' Add new items to the menu
    Select Case M
        Case 1
        Load Menu1_items(I) ' Load new menu item
        Menu1_items(I).Caption = menu1_caption(I) ' Set captions for menu items
    Case 2
        Load Menu2_items(I) ' Load new menu item
        Menu2_items(I).Caption = menu2_caption(I) ' Set captions for menu items
    Case 3
        Load Menu3_items(I) ' Load new menu item
        Menu3_items(I).Caption = menu3_caption(I) ' Set captions for menu items
    Case 4
        Load Menu4_items(I) ' Load new menu item
        Menu4_items(I).Caption = menu4_caption(I) ' Set captions for menu items
    Case 5
        Load Menu5_items(I) ' Load new menu item
        Menu5_items(I).Caption = menu5_caption(I) ' Set captions for menu items
    Case 6
        Load Menu6_items(I) ' Load new menu item
        Menu6_items(I).Caption = menu6_caption(I) ' Set captions for menu items
    Case 7
        Load Menu7_items(I) ' Load new menu item
        Menu7_items(I).Caption = menu7_caption(I) ' Set captions for menu items
    Case 8
        Load Menu8_items(I) ' Load new menu item
        Menu8_items(I).Caption = menu8_caption(I) ' Set captions for menu items
    Case 9
        Load Menu9_items(I) ' Load new menu item
        Menu9_items(I).Caption = menu9_caption(I) ' Set captions for menu items
    Case 10
        Load Menu10_items(I) ' Load new menu item
        Menu10_items(I).Caption = menu10_caption(I) ' Set captions for menu items
    Case 11
        Load Menu11_items(I) ' Load new menu item
        Menu11_items(I).Caption = menu11_caption(I) ' Set captions for menu items
    Case 12
        Load Menu12_items(I) ' Load new menu item
        Menu12_items(I).Caption = menu12_caption(I) ' Set captions for menu items
    Case 13
        Load Menu13_items(I) ' Load new menu item
        Menu13_items(I).Caption = menu13_caption(I) ' Set captions for menu items
    Case 14
        Load Menu14_items(I) ' Load new menu item
        Menu14_items(I).Caption = menu14_caption(I) ' Set captions for menu items
    Case 15
        Load Menu15_items(I) ' Load new menu item
        Menu15_items(I).Caption = menu15_caption(I) ' Set captions for menu items
    
    End Select
    
    
    
    Next I
    Select Case M
       Case 1:    Menu1_items(0).Visible = False ' This is the divider - make it invisible
       Case 2:    Menu2_items(0).Visible = False ' This is the divider - make it invisible
       Case 3:    Menu3_items(0).Visible = False ' This is the divider - make it invisible
       Case 4:    Menu4_items(0).Visible = False ' This is the divider - make it invisible
       Case 5:    Menu5_items(0).Visible = False ' This is the divider - make it invisible
       Case 6:    Menu6_items(0).Visible = False ' This is the divider - make it invisible
       Case 7:    Menu7_items(0).Visible = False ' This is the divider - make it invisible
       Case 8:    Menu8_items(0).Visible = False ' This is the divider - make it invisible
       Case 9:    Menu9_items(0).Visible = False ' This is the divider - make it invisible
       Case 10:   Menu10_items(0).Visible = False ' This is the divider - make it invisible
       Case 11:   Menu11_items(0).Visible = False ' This is the divider - make it invisible
       Case 12:   Menu12_items(0).Visible = False ' This is the divider - make it invisible
       Case 13:   Menu13_items(0).Visible = False ' This is the divider - make it invisible
       Case 14:   Menu14_items(0).Visible = False ' This is the divider - make it invisible
       Case 15:   Menu15_items(0).Visible = False ' This is the divider - make it invisible
       
End Select

End Sub












Private Sub blank4_Click()
    Me.Height = 1
    Timer1.Enabled = True
End Sub

Private Sub Combo1_Click()
pick = UCase(Combo1.Text)
If InStr(1, pick, "MENU1.txt") Then Text3.Text = mnufile.Caption: Command3.Caption = mnufile.Caption
If InStr(1, pick, "MENU2.txt") Then Text3.Text = mnu_Menu2.Caption: Command4.Caption = mnu_Menu2.Caption
If InStr(1, pick, "MENU3.txt") Then Text3.Text = menu3.Caption: Command5.Caption = menu3.Caption
If InStr(1, pick, "MENU4.txt") Then Text3.Text = MENU4.Caption: Command6.Caption = MENU4.Caption
If InStr(1, pick, "MENU5.txt") Then Text3.Text = MENU5.Caption: Command7.Caption = MENU5.Caption
If InStr(1, pick, "MENU6.txt") Then Text3.Text = MENU6.Caption: Command8.Caption = MENU6.Caption
If InStr(1, pick, "MENU7.txt") Then Text3.Text = MENU7.Caption: Command9.Caption = MENU7.Caption
If InStr(1, pick, "MENU8.txt") Then Text3.Text = MENU8.Caption: Command10.Caption = MENU8.Caption
If InStr(1, pick, "MENU9.txt") Then Text3.Text = MENU9.Caption: Command11.Caption = MENU9.Caption
If InStr(1, pick, "MENU10.txt") Then Text3.Text = MENU10.Caption: Command12.Caption = MENU10.Caption
If InStr(1, pick, "MENU11.txt") Then Text3.Text = MENU11.Caption: Command13.Caption = MENU11.Caption
If InStr(1, pick, "MENU12.txt") Then Text3.Text = MENU12.Caption: Command14.Caption = MENU12.Caption
If InStr(1, pick, "MENU13.txt") Then Text3.Text = MENU13.Caption: Command15.Caption = MENU13.Caption
If InStr(1, pick, "MENU14.txt") Then Text3.Text = MENU14.Caption: Command16.Caption = MENU14.Caption
If InStr(1, pick, "MENU15.txt") Then Text3.Text = MENU15.Caption: Command17.Caption = MENU15.Caption

End Sub






Private Sub Command1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
End If

End Sub





Private Function Filenm(strx As String) As String
Dim sps As Integer
'sl As Integer,
'sl = Len(strx)
For sps = Len(strx) To 1 Step -1
If Mid(strx, sps, 1) = "\" Then
Filenm = Mid$(strx, sps + 1)
Exit For
End If
Next
End Function






Private Sub Command10_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
    Combo1.Text = App.Path + "\menu8.txt"
    End If
End Sub

Private Sub Command11_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
    Combo1.Text = App.Path + "\menu9.txt"
    End If
End Sub

Private Sub Command12_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
    Combo1.Text = App.Path + "\menu10.txt"
    End If
End Sub

Private Sub Command13_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
    Combo1.Text = App.Path + "\menu11.txt"
    End If
End Sub

Private Sub Command14_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
    Combo1.Text = App.Path + "\menu12.txt"
    End If
End Sub

Private Sub Command15_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
    Combo1.Text = App.Path + "\menu13.txt"
    End If
End Sub

Private Sub Command16_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
    Combo1.Text = App.Path + "\menu14.txt"
    End If
End Sub



Private Sub Command17_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
    Combo1.Text = App.Path + "\menu15.txt"
    End If
End Sub





Private Sub Command18_Click() 'back button
DDM_Click
End Sub




Private Sub Command2_Click() ' save new item button
 Open Combo1.Text For Append As #1
 Text1.Text = Text1.Text + "                                  "
 myline = Text1.Text + Text2.Text
Print #1, myline
Close #1
Load_Menus
Beep
End Sub



Private Sub Command3_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
    Combo1.Text = App.Path + "\menu1.txt"
    End If
End Sub











Private Sub Command4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
    Combo1.Text = App.Path + "\menu2.txt"
    End If
End Sub

Private Sub Command5_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
    Combo1.Text = App.Path + "\menu3.txt"
    End If
End Sub

Private Sub Command6_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
    Combo1.Text = App.Path + "\menu4.txt"
    End If
End Sub

Private Sub Command7_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
    Combo1.Text = App.Path + "\menu5.txt"
    End If
End Sub

Private Sub Command8_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
    Combo1.Text = App.Path + "\menu6.txt"
    End If
End Sub

Private Sub Command9_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    Text2.Text = FL
    Text1.Text = Filenm(FL)
    Next FN
    Combo1.Text = App.Path + "\menu7.txt"
    End If
End Sub

Sub set_button_names()
 Command3.Caption = mnufile.Caption
Command4.Caption = mnu_Menu2.Caption
 Command5.Caption = menu3.Caption
 Command6.Caption = MENU4.Caption
 Command7.Caption = MENU5.Caption
 Command8.Caption = MENU6.Caption
 Command9.Caption = MENU7.Caption
 Command10.Caption = MENU8.Caption
 Command11.Caption = MENU9.Caption
 Command12.Caption = MENU10.Caption
 Command13.Caption = MENU11.Caption
 Command14.Caption = MENU12.Caption
 Command15.Caption = MENU13.Caption
 Command16.Caption = MENU14.Caption
 Command17.Caption = MENU15.Caption

End Sub



Private Sub DDM_Click() ' drag & drop down mode
set_button_names



If Me.Height > 1000 Then
    Me.Height = 1
    Timer1.Enabled = True
    Timer3.Enabled = False 'turn off mouse position display
Else
Timer3.Enabled = True 'turn on mouse position display
    Me.Height = 4000
    Timer1.Enabled = False
    Timer2.Enabled = False
End If
Combo1.Text = "Pick a menu"
End Sub





Private Sub Form_Load()
stayontop_Click  ' make menubar stay on top
Me.Top = 0
Me.Left = 0
Me.Width = Screen.Width
Me.Height = 1 '375
Call check_data_files
Load_Menus
End Sub


Sub check_data_files()
 chk = Dir(App.Path + "\menu1.txt")
  If chk = "" Then
   Make_data_File App.Path + "\menu1.txt", "Menu 1"
   MsgBox "Since this appears to be the first time MenuBar has ran, I have set up default menus for you. Use the drag and drop menu item to add items to menu system. Use 'Edit this menu' to edit everything, including Menu Names"
  End If
 
 chk = Dir(App.Path + "\menu2.txt"): If chk = "" Then Make_data_File App.Path + "\menu2.txt", "Menu 2"
 chk = Dir(App.Path + "\menu3.txt"): If chk = "" Then Make_data_File App.Path + "\menu3.txt", "Menu 3"
 chk = Dir(App.Path + "\menu4.txt"): If chk = "" Then Make_data_File App.Path + "\menu4.txt", "Menu 4"
 chk = Dir(App.Path + "\menu5.txt"): If chk = "" Then Make_data_File App.Path + "\menu5.txt", "Menu 5"
 chk = Dir(App.Path + "\menu6.txt"): If chk = "" Then Make_data_File App.Path + "\menu6.txt", "Menu 6"
 chk = Dir(App.Path + "\menu7.txt"): If chk = "" Then Make_data_File App.Path + "\menu7.txt", "Menu 7"
 chk = Dir(App.Path + "\menu8.txt"): If chk = "" Then Make_data_File App.Path + "\menu8.txt", "Menu 8"
 chk = Dir(App.Path + "\menu9.txt"): If chk = "" Then Make_data_File App.Path + "\menu9.txt", "Menu 9"
 chk = Dir(App.Path + "\menu10.txt"): If chk = "" Then Make_data_File App.Path + "\menu10.txt", "Menu 10"
 chk = Dir(App.Path + "\menu11.txt"): If chk = "" Then Make_data_File App.Path + "\menu11.txt", "Menu 11"
 chk = Dir(App.Path + "\menu12.txt"): If chk = "" Then Make_data_File App.Path + "\menu12.txt", "Menu 12"
 chk = Dir(App.Path + "\menu13.txt"): If chk = "" Then Make_data_File App.Path + "\menu13.txt", "Menu 13"
 chk = Dir(App.Path + "\menu14.txt"): If chk = "" Then Make_data_File App.Path + "\menu14.txt", "Menu 14"
 chk = Dir(App.Path + "\menu15.txt"): If chk = "" Then Make_data_File App.Path + "\menu15.txt", "Menu 15"
       
End Sub




Sub Make_data_File(a As String, b As String)
Open a For Output As #3
Print #3, b
Print #3, ""
Print #3, ""
Print #3, "*****"
Print #3, "Edit This Menu               " + App.Path + "\menu5.txt"
Close #3
End Sub




Private Sub Load_Menus()
Dim I As Integer
On Error Resume Next

For I = 1 To 15
    FillMenu I
Next I


End Sub












Private Sub mnu_view_log_Click()
runIT App.Path + "\log.txt"
End Sub

Private Sub reload_menus_Click()
Load_Menus
End Sub




Private Sub scott_item1_Click(Index As Integer)
End
End Sub



Private Sub stayontop_Click()

If is_on_top = True Then
    is_on_top = False
  '  stayontop.Checked = False
    Call FormOnTop(Me.hwnd, False)
Else
    is_on_top = True
   ' stayontop.Checked = True
    Call FormOnTop(Me.hwnd, True)
End If

End Sub



Sub runIT(a As String)
On Error Resume Next
cmd = UCase(a)


If InStr(1, cmd, "*") > 0 Then
    a = Right$(a, Len(a) - 1) 'remove the " * " out of string
    logit "1..Running: " + a
    a = Shell(a, vbNormalFocus) 'run programs with parms
Else
    logit "2..Running: " + a
    StartDoc2 a, ""  'run programs without parms & ab-normal file extendtions ex: .lnk
End If


End Sub





Private Sub Menu1_items_Click(Index As Integer)
runIT menu1_command(Index)
End Sub


Private Sub Menu2_items_Click(Index As Integer)
runIT menu2_command(Index)
End Sub





Private Sub Menu3_items_Click(Index As Integer)
runIT menu3_command(Index)
End Sub


Private Sub Menu4_items_Click(Index As Integer)
runIT menu4_command(Index)
End Sub




Private Sub Menu5_items_Click(Index As Integer)
runIT menu5_command(Index)
End Sub


Private Sub Menu6_items_Click(Index As Integer)
runIT menu6_command(Index)
End Sub



Private Sub Menu7_items_Click(Index As Integer)
runIT menu7_command(Index)
End Sub


Private Sub Menu8_items_Click(Index As Integer)
runIT menu8_command(Index)
End Sub



Private Sub Menu9_items_Click(Index As Integer)
runIT menu9_command(Index)
End Sub


Private Sub Menu10_items_Click(Index As Integer)
runIT menu10_command(Index)
End Sub



Private Sub Menu11_items_Click(Index As Integer)
runIT menu11_command(Index)
End Sub


Private Sub Menu12_items_Click(Index As Integer)
runIT menu12_command(Index)
End Sub




Private Sub Menu13_items_Click(Index As Integer)
runIT menu13_command(Index)
End Sub


Private Sub Menu14_items_Click(Index As Integer)
runIT menu14_command(Index)
End Sub




Private Sub Menu15_items_Click(Index As Integer)
runIT menu15_command(Index)
End Sub






Private Sub Timer1_Timer()
Call Module1.GetCursorPos(PositionNow)
'Label2 = "Position: " & PositionNow.x & ", " & PositionNow.y
blank4.Caption = "Position: " & PositionNow.x & ", " & PositionNow.y

If PositionNow.y < 2 Then
    Timer2.Enabled = False 'this is to reset timer 2 so the
    Timer2.Enabled = True  'menubar will stay up when cursor is at top
    Me.Height = 375
End If
End Sub









Private Sub Timer2_Timer()

Static lHwnd As Long
    Dim lCurHwnd As Long
    Dim sText As String * 255

    lCurHwnd = GetForegroundWindow
    lHwnd = lCurHwnd

   ' Debug.Print Str(lHwnd) + "    " + Str(Me.hwnd)

Me.Height = 1
Timer2.Enabled = False

End Sub

Private Sub Timer3_Timer()
Call Module1.GetCursorPos(PositionNow)
'Label2 = "Position: " & PositionNow.x & ", " & PositionNow.y
blank4.Caption = "Position: " & PositionNow.x & ", " & PositionNow.y
End Sub

Private Sub Timer4_Timer(Index As Integer)
If Me.Width <> Screen.Width Then Me.Width = Screen.Width
End Sub

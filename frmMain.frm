VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Font Browser"
   ClientHeight    =   3285
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Fonts"
      Height          =   2775
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   6375
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh Font List"
         Height          =   375
         Left            =   2820
         TabIndex        =   3
         Top             =   300
         Width           =   1635
      End
      Begin VB.ListBox lstFonts 
         Height          =   2010
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2475
      End
      Begin ComctlLib.ProgressBar pbStatus 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   2340
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   556
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblSample 
         Alignment       =   2  'Center
         Caption         =   "Sample text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2700
         TabIndex        =   4
         Top             =   1020
         Width           =   3615
      End
   End
   Begin VB.Label Label1 
      Caption         =   "This is a sample font browsing program. Select a font from below or the above menu. Please rate my code!"
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6435
   End
   Begin VB.Menu mnuFonts 
      Caption         =   "Fonts"
      Begin VB.Menu mnuFontList 
         Caption         =   "Font List"
         Begin VB.Menu mnuFontName 
            Caption         =   ""
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ChangeFont(strFontName As String)
'This sub will change the 'Font' for the sample text
'Look in the lstFonts_Click() even and the menu's font list
'To see how this sub works
'
'The sub also updates the menu and the listbox for a tick or
'selection
'
lblSample.Font = strFontName 'Change font
'
'Select the item in the list
For i = 0 To lstFonts.ListCount - 1
If lstFonts.List(i) = strFontName Then 'If the listitem is the
'selected font, then select that item in the list
lstFonts.ListIndex = i 'Select it via the index
GoTo 2 'Goto that 'Section' in the code
End If '                               |
Next i '                               |
2: '<-----------------------------------

For i = 0 To mnuFontName.Count - 1 'Cycle thru all items in menu
If mnuFontName(i).Caption = strFontName Then 'If its our font
mnuFontName(i).Checked = True 'Place a tick next to the font name
Else
mnuFontName(i).Checked = False 'Uncheck the rest
End If
Next i
End Sub
Sub GetFonts()
On Error Resume Next
lstFonts.Clear 'Reset list
pbStatus.Value = 0 'Reset progressbar
pbStatus.Max = Screen.FontCount 'Maximum number of fonts found in printer memory... (Set Progressbar.MAX to that number)
mnuFontName(0).Caption = Screen.Fonts(0) 'Add first entry (first array) in the font list to the menu.
For i = 1 To Screen.FontCount - 1 'Cycle thru every font
    lstFonts.AddItem Screen.Fonts(i) 'Add the font name
    pbStatus.Value = pbStatus.Value + 1 'Add 1 to the progress
    Load mnuFontName(i) 'Load that menu item
    mnuFontName(0).Caption = Screen.Fonts(i) 'Set its caption to the font name
Next i
Frame1.Caption = Screen.FontCount & " fonts found." 'Display how many fonts have been found
End Sub

Private Sub cmdRefresh_Click()
MousePointer = 11 'Change mousepointer to Hourglass (Working)
GetFonts 'Run this sub (See the DEC's in this form)
MousePointer = 0 'Set back to default mouse pointer
End Sub

Private Sub Form_Load()
MsgBox "Font Browser by Plasma - If you modify/use/improve on this code, please contact me and credit." & _
      vbCrLf & vbCrLf & _
      "Contact: andrewarmstrong@hotmail.com | ICQ# 14344635" & vbCrLf & vbCrLf & _
      "Please rate my code! PLEASE!", vbInformation, "Info/About"
End Sub

Private Sub lstFonts_Click()
ChangeFont lstFonts.List(lstFonts.ListIndex) 'Change the font
'to the selected font in the list
End Sub

Private Sub mnuFontName_Click(Index As Integer)
ChangeFont mnuFontName(Index).Caption 'Change the font
'to the selected font in the list
End Sub

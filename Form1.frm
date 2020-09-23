VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "&Show some  some windows"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Hide"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      ToolTipText     =   "Hide the window!"
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Show"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      ToolTipText     =   "See How it looks like : )"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox txtHwnd 
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0000"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show ALL Windows"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   5520
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Handle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   4920
      Width           =   1530
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This function is used to fetch the caption of titlebar of given window!
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
' This functions is being used to find ALL the child windows of the Desktop window(ie the Applications that are running!)
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
' Variables which are used by the program to "return" the values fetched from callig the above api functions
Dim thehwnd As Long, MyStr As String, junk As Long
' This API is being used by the program to show and unhide the selected window/hwnd
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
' This constant hides window
Private Const SW_HIDE = 0
' This constant shows a window
Private Const SW_SHOW = 5
' This function is used by the program "MOVE" the position of the window that the user/you have JUST asked the program to show!(unhide)
' actually some windows are there Running on ur PCs which are beyond or before the no. of Screen coordinates supported by your Screen!,, so that's why this API is being used to bring that selected window to 0,0 position!
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
' This structure holds the info of selected window's dimentions, and position
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
' This structure is used to get the coordinates of the seletected/specified window! Actually this function is being used
' by the program because our MoveWindow function also required the "dimensions" of the window to be assigned before moving it!, So we want to MOVE the window WITHOUT distorting/disturbing its dimensions,, so we have to find out its dimenstions first :)
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
' You should know till now what's this :D
Dim theRect As RECT
' This function is being used by the program to Redrawn the Window which has been moved! Actually it sends a "Wm_Paint" message to the selected window!
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Sub Command1_Click()
'Create a buffer
MyStr = String(250, Chr$(0))
'reinitialize to get the list re-filled and refreshed!
thehwnd = 0
'reinitialize to get the list re-filled and refreshed!
List1.Clear
' Find ALL the (parent) windows in the "Desktop" window of the OS,, till ALL the child windows have been found
Do
' Find ALL the (parent) windows in the "Desktop" window of the OS,, till ALL the child windows have been found!!! PLEASE Note that ONLY parent windows/Windows of MAIN Applications would be found by the function by this method!! So,, this is a good method to find ONLY Applications, n NOT their children
    thehwnd = FindWindowEx(0, thehwnd, vbNullString, vbNullString)
' Get the caption of the titlebar of the selected window/hwnd!
    GetWindowText thehwnd, MyStr, 250
' Check to see if this hwnd has some caption or not,, if NOT,, this might NOT be an application and can Simply be a part of OSs interior working windows,, like the "Start" button and the Icons of desktop!
' You can REMOVE this coditional(ie the If Statement) to see MORE secrets of Windows,, and to catch more windows
    If Left(MyStr, 1) <> Chr(0) Or Check1.Value = Checked Then
        'Add the details of the found Application in the Listbox for the user to view
        List1.AddItem thehwnd & "  " & MyStr
        'Save the Hwnd of this application in the ItemData for futureuse(ie hiding and showing and ALL that!!)
        List1.ItemData(List1.ListCount - 1) = thehwnd
    End If
Loop Until thehwnd = 0
Me.Caption = "Hey!!,,, L@@k at the stuff that is working on your PC!!!!ALL in disguise!!Check them ALL :O !!"
End Sub

Private Sub Command2_Click()
'First of ALL show the seleted window/application
junk = ShowWindow(Val(txtHwnd), SW_SHOW)
'Get the dimensions, and position of the window
junk = GetWindowRect(Val(txtHwnd), theRect)
'Move the window to Top 0 and Left 0 incase and in MANY Cases,, some windows are hidden and are Alsmost OUT of the screen coordinates supported by ur monitor
junk = MoveWindow(Val(txtHwnd), 0, 0, theRect.Right - theRect.Left, theRect.Bottom - theRect.Top, True)
'Hit a Update to the GUI of that Window!
junk = UpdateWindow(Val(txtHwnd))
End Sub

Private Sub Command3_Click()
'Hide that window
junk = ShowWindow(Val(txtHwnd), SW_HIDE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Vote me!:)
MsgBox vbCrLf & vbCrLf & "Plz Don't forget to vote and comment!!:D:D" & vbCrLf & vbCrLf & "Thanks" & vbCrLf & "-Tushar", vbExclamation
'Open internet explorer with one line of code:)
Shell "start http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=41426&lngWId=1"
End Sub

Private Sub List1_Click()
'Show the handle of EACH window that has been selected by the user
txtHwnd = List1.ItemData(List1.ListIndex)
End Sub


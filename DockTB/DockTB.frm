VERSION 5.00
Begin VB.Form frmDockTB 
   BackColor       =   &H80000016&
   Caption         =   "Docking Toolbars Demo"
   ClientHeight    =   5625
   ClientLeft      =   2700
   ClientTop       =   2070
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   7680
   Begin VB.PictureBox picFrame 
      Height          =   4335
      Left            =   840
      ScaleHeight     =   4275
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin VB.CommandButton cmdButton 
         Caption         =   "Dock the toolbars to my left or right and I resize automatically!"
         Height          =   615
         Left            =   1800
         TabIndex        =   1
         Top             =   1800
         Width           =   1815
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmDockTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------------------------- '
' Filename: DockTB.frm
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Demonstrates the docking toolbars
' ----------------------------------------------------------------- '

Option Explicit

' Create two toolbars
Public Toolbar1 As New CToolbar
Public Toolbar2 As New CToolbar

' Attach the MyFrame's events, so we know when
' a toolbar button is clicked...
Public WithEvents MyFrame As CFrame
Attribute MyFrame.VB_VarHelpID = -1

'--------------------------------------
' toolbar identifiers
Private Const IDW_TOOLBAR1 = &HE800&
Private Const IDW_TOOLBAR2 = &HE804&
'--------------------------------------

'--------------------------------------
' Toolbar1 button identifiers
Private Const ID_NEW = 81
Private Const ID_OPEN = 82
Private Const ID_SAVE = 83

Private Const ID_CUT = 84
Private Const ID_COPY = 85
Private Const ID_PASTE = 86

Private Const ID_PRINT = 87
Private Const ID_HELP = 88

' Toolbar2 button identifiers
Private Const ID_ERASER = 91
Private Const ID_PENCIL = 92
Private Const ID_SELECT = 93

Private Const ID_BRUSH = 94
Private Const ID_SPRAY = 95
Private Const ID_FILL = 96
'--------------------------------------


Private Function SetupToolbars()
    ' Attach the form to the MyFrame object
    MyFrame.hWnd = Me.hWnd
    ' Add our new frame in the WindowList
    Call AddWindowInList(MyFrame)
    ' Initialize all the classes, libraries etc.
    Call Initialize

    ' Our toolbar's style
    Const style = CBRS_GRIPPER Or CBRS_BORDER_3D Or WS_CHILD Or WS_VISIBLE Or _
                CBRS_SIZE_DYNAMIC Or CBRS_TOOLTIPS Or CBRS_FLYBY Or _
                &H800& Or CBRS_FLOATING Or CBRS_TOP

    ' Create the toolbars
    Call Toolbar1.Create(Me.hWnd, style, IDW_TOOLBAR1)
    Call Toolbar2.Create(Me.hWnd, style, IDW_TOOLBAR2)

    ' Please don't ask me why this is required...
    ' Though this is not always required, but it is if you need flat toolbars
    Call SetWindowLong(Toolbar1.m_hWnd, GWL_STYLE, &H5400004E Or CBRS_SIZE_DYNAMIC Or CBRS_TOOLTIPS Or TBSTYLE_FLAT)
    Call SetWindowLong(Toolbar2.m_hWnd, GWL_STYLE, &H5400004E Or CBRS_SIZE_DYNAMIC Or CBRS_TOOLTIPS) 'Or TBSTYLE_FLAT)

    ' Add the buttons to the toolbars
    
    ' Send the size of the TBBUTTON structure to the toolbar,
    ' so it knows which version of the structure we're using...
    Dim tbBtn As TBBUTTON
    Call SendMessage(Toolbar1.m_hWnd, TB_BUTTONSTRUCTSIZE, Len(tbBtn), 0)
    Call SendMessage(Toolbar2.m_hWnd, TB_BUTTONSTRUCTSIZE, Len(tbBtn), 0)

    ' Load the toolbar bitmaps from the resource using the resource ID
    Call Toolbar1.LoadBitmap("101")
    Call Toolbar2.LoadBitmap("102")

    ' Specify the button identifiers for Toolbar1
    Dim Buttons1(9) As Long, Buttons2(6) As Long
    Buttons1(0) = ID_NEW
    Buttons1(1) = ID_OPEN
    Buttons1(2) = ID_SAVE
    Buttons1(3) = 0             ' separator
    Buttons1(4) = ID_CUT
    Buttons1(5) = ID_COPY
    Buttons1(6) = ID_PASTE
    Buttons1(7) = 0             ' separator
    Buttons1(8) = ID_PRINT
    Buttons1(9) = ID_HELP

    ' Specify the button identifiers for Toolbar2
    Buttons2(0) = ID_ERASER
    Buttons2(1) = ID_PENCIL
    Buttons2(2) = ID_SELECT
    Buttons2(3) = 0             ' separator
    Buttons2(4) = ID_BRUSH
    Buttons2(5) = ID_SPRAY
    Buttons2(6) = ID_FILL

    ' Set the toolbar buttons
    Call Toolbar1.SetButtons(Buttons1, 10)
    Call Toolbar2.SetButtons(Buttons2, 7)

    ' Enable toolbar docking
    Call Toolbar1.EnableDocking(CBRS_ALIGN_ANY)
    Call Toolbar2.EnableDocking(CBRS_ALIGN_ANY)

    ' Name the toolbar
    ' This text identifies the toolbar when it is floating...
    Call SetWindowText(Toolbar1.m_hWnd, "Main Toolbar")
    Call SetWindowText(Toolbar2.m_hWnd, "Secondary Toolbar")

    ' Finally subclass MyFrame
    Call MyFrame.Subclass
    ' and enable docking on the frame...
    MyFrame.EnableDocking (CBRS_ALIGN_ANY)

    ' Dock the new toolbars
    Dim r As RECT
    Call MyFrame.DockControlBar1(Toolbar1, AFX_IDW_DOCKBAR_TOP, r)
    Call MyFrame.DockControlBar1(Toolbar2, AFX_IDW_DOCKBAR_BOTTOM, r)
End Function


Private Sub Form_Load()
    ' Create a new CFrame object for MyFrame...
    Set MyFrame = New CFrame

    ' Call the function SetupToolbars defined above,
    ' that sets up all the toolbars...
    ' This is neater than writing all the code in Form_Load()
    Call SetupToolbars

    ' Resize the frame with MainFrame's active area
    Call SetWindowLong(picFrame.hWnd, GWL_ID, AFX_IDW_PANE_FIRST)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Uninitialization, and cleanup
    ' Unregister all the classes that we registered...
    Call UnInitialize
End Sub


' MyFrame's OnCommand event is used to catch toolbar's button clicks
' Also catch menu events here...
Private Sub MyFrame_OnCommand(ID As Long)
    If GetHiWord(ID) = 0 Then
        Dim s As String
        Select Case GetLoWord(ID)

            'Toolbar1's buttons
            Case ID_NEW
                s = "New"
            Case ID_OPEN
                s = "Open"
            Case ID_SAVE
                s = "Save"
            Case ID_CUT
                s = "Cut"
            Case ID_COPY
                s = "Copy"
            Case ID_PASTE
                s = "Paste"
            Case ID_PRINT
                s = "Print"
            Case ID_HELP, 2             '(mnuHelpAbout = 2)
                s = "About"
                MsgBox "Docking Toolbars" & vbCrLf & "Converted from Visual C++ MFC to Visual Basic" & vbCrLf & "by Shaurya Malhotra (shauryamal@gmail.com)", vbInformation, "Docking Toolbars Demo"

            'Toolbar2's buttons
            Case ID_ERASER
                s = "Eraser"
            Case ID_PENCIL
                s = "Pencil"
            Case ID_SELECT
                s = "Select"
            Case ID_BRUSH
                s = "Brush"
            Case ID_SPRAY
                s = "Spray"
            Case ID_FILL
                s = "Fill"

        End Select
        Me.Caption = "You clicked " & s & "!"
    End If
End Sub


' Our frame's resize event
Private Sub picFrame_Resize()
    cmdButton.Left = 50
    cmdButton.Width = picFrame.ScaleWidth - 50
End Sub


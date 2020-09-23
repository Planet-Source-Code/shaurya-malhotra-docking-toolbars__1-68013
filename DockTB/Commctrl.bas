Attribute VB_Name = "Commctrl"
' ----------------------------------------------------------------- '
' Filename: Commctrl.bas
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Common control constants (adapted from MFC headers)
' ----------------------------------------------------------------- '

Option Explicit

#Const WIN32_ = True
#Const WIN32_IE_ = &H401

'------------------ COMMON CONTROL STYLES ------------------
Public Const CCS_TOP As Long = &H1
Public Const CCS_NOMOVEY As Long = &H2
Public Const CCS_BOTTOM As Long = &H3
Public Const CCS_NORESIZE As Long = &H4
Public Const CCS_NOPARENTALIGN As Long = &H8
Public Const CCS_ADJUSTABLE As Long = &H20
Public Const CCS_NODIVIDER As Long = &H40
#If (WIN32_IE_ >= &H300) Then
Public Const CCS_VERT As Long = &H80
Public Const CCS_LEFT As Long = (CCS_VERT Or CCS_TOP)
Public Const CCS_RIGHT As Long = (CCS_VERT Or CCS_BOTTOM)
Public Const CCS_NOMOVEX As Long = (CCS_VERT Or CCS_NOMOVEY)
#End If
'-----------------------------------------------------------

'--------------------- TOOLBAR CONTROL ---------------------
Const TBN_FIRST = 700
Const TBN_LAST = 720

#If Win32 Then
Public Const TOOLBARCLASSNAMEW As String = "ToolbarWindow32"
Public Const TOOLBARCLASSNAMEA As String = "ToolbarWindow32"

#If UNICODE Then
Public Const TOOLBARCLASSNAME As String = TOOLBARCLASSNAMEW
#Else
Public Const TOOLBARCLASSNAME As String = TOOLBARCLASSNAMEA
#End If

#Else
Public Const TOOLBARCLASSNAME As String = "ToolbarWindow"
#End If

Public Type TBBUTTON
    iBitmap As Long
    idCommand As Long
    fsState As Byte
    fsStyle As Byte
#If WIN32_ Then
    bReserved(1 To 2) As Byte
#End If
    dwData As Long
    iString As Long
End Type

Public Type COLORMAP
    from As COLORREF
    to As COLORREF
End Type

Public Const CMB_MASKED = &H2

Public Const TBSTATE_CHECKED = &H1
Public Const TBSTATE_PRESSED = &H2
Public Const TBSTATE_ENABLED = &H4
Public Const TBSTATE_HIDDEN = &H8
Public Const TBSTATE_INDETERMINATE = &H10
Public Const TBSTATE_WRAP = &H20
#If (WIN32_IE_ >= &H300) Then
Public Const TBSTATE_ELLIPSES = &H40
#End If
#If (WIN32_IE_ >= &H400) Then
Public Const TBSTATE_MARKED = &H80
#End If

Public Const TBSTYLE_BUTTON = &H0
Public Const TBSTYLE_SEP = &H1
Public Const TBSTYLE_CHECK = &H2
Public Const TBSTYLE_GROUP = &H4
Public Const TBSTYLE_CHECKGROUP = (TBSTYLE_GROUP Or TBSTYLE_CHECK)
#If (WIN32_IE_ >= &H300) Then
Public Const TBSTYLE_DROPDOWN = &H8
#End If
#If (WIN32_IE_ >= &H400) Then
Public Const TBSTYLE_AUTOSIZE = &H10            ' automatically calculate the cx of the button
Public Const TBSTYLE_NOPREFIX = &H20            ' if this button should not have accel prefix
#End If

Public Const TBSTYLE_TOOLTIPS = &H100
Public Const TBSTYLE_WRAPABLE = &H200
Public Const TBSTYLE_ALTDRAG = &H400
#If (WIN32_IE_ >= &H300) Then
Public Const TBSTYLE_FLAT = &H800
Public Const TBSTYLE_LIST = &H1000
Public Const TBSTYLE_CUSTOMERASE = &H2000
#End If
#If (WIN32_IE_ >= &H400) Then
Public Const TBSTYLE_REGISTERDROP = &H4000
Public Const TBSTYLE_TRANSPARENT = &H8000
Public Const TBSTYLE_EX_DRAWDDARROWS = &H1
#End If

#If (WIN32_IE_ >= &H400) Then

' Toolbar custom draw return flags
Public Const TBCDRF_NOEDGES = &H10000                   ' Don't draw button edges
Public Const TBCDRF_HILITEHOTTRACK = &H20000            ' Use color of the button bk when hottracked
Public Const TBCDRF_NOOFFSET = &H40000                  ' Don't offset button if pressed
Public Const TBCDRF_NOMARK = &H80000                    ' Don't draw default highlight of image/text for TBSTATE_MARKED
Public Const TBCDRF_NOETCHEDEFFECT = &H100000           ' Don't draw etched effect for disabled items

#End If

Public Const TB_ENABLEBUTTON = (WM_USER + 1)
Public Const TB_CHECKBUTTON = (WM_USER + 2)
Public Const TB_PRESSBUTTON = (WM_USER + 3)
Public Const TB_HIDEBUTTON = (WM_USER + 4)
Public Const TB_INDETERMINATE = (WM_USER + 5)
#If (WIN32_IE_ >= &H400) Then
Public Const TB_MARKBUTTON = (WM_USER + 6)
#End If
Public Const TB_ISBUTTONENABLED = (WM_USER + 9)
Public Const TB_ISBUTTONCHECKED = (WM_USER + 10)
Public Const TB_ISBUTTONPRESSED = (WM_USER + 11)
Public Const TB_ISBUTTONHIDDEN = (WM_USER + 12)
Public Const TB_ISBUTTONINDETERMINATE = (WM_USER + 13)
#If (WIN32_IE_ >= &H400) Then
Public Const TB_ISBUTTONHIGHLIGHTED = (WM_USER + 14)
#End If
Public Const TB_SETSTATE = (WM_USER + 17)
Public Const TB_GETSTATE = (WM_USER + 18)
Public Const TB_ADDBITMAP = (WM_USER + 19)

#If WIN32_ Then
Public Type tagTBADDBITMAP
        hInst As Long
        nID As Long
End Type

Public Const HINST_COMMCTRL = -1
Public Const IDB_STD_SMALL_COLOR = 0
Public Const IDB_STD_LARGE_COLOR = 1
Public Const IDB_VIEW_SMALL_COLOR = 4
Public Const IDB_VIEW_LARGE_COLOR = 5
#If (WIN32_IE_ >= &H300) Then
Public Const IDB_HIST_SMALL_COLOR = 8
Public Const IDB_HIST_LARGE_COLOR = 9
#End If

' icon indexes for standard bitmap
Public Const STD_CUT = 0
Public Const STD_COPY = 1
Public Const STD_PASTE = 2
Public Const STD_UNDO = 3
Public Const STD_REDOW = 4
Public Const STD_DELETE = 5
Public Const STD_FILENEW = 6
Public Const STD_FILEOPEN = 7
Public Const STD_FILESAVE = 8
Public Const STD_PRINTPRE = 9
Public Const STD_PROPERTIES = 10
Public Const STD_HELP = 11
Public Const STD_FIND = 12
Public Const STD_REPLACE = 13
Public Const STD_PRINT = 14

' icon indexes for standard view bitmap
Public Const VIEW_LARGEICONS = 0
Public Const VIEW_SMALLICONS = 1
Public Const VIEW_LIST = 2
Public Const VIEW_DETAILS = 3
Public Const VIEW_SORTNAME = 4
Public Const VIEW_SORTSIZE = 5
Public Const VIEW_SORTDATE = 6
Public Const VIEW_SORTTYPE = 7
Public Const VIEW_PARENTFOLDER = 8
Public Const VIEW_NETCONNECT = 9
Public Const VIEW_NETDISCONNECT = 10
Public Const VIEW_NEWFOLDER = 11
#If (WIN32_IE_ >= &H400) Then
Public Const VIEW_VIEWMENU = 12
#End If

#If (WIN32_IE_ >= &H300) Then
Public Const HIST_BACK = 0
Public Const HIST_FORWARD = 1
Public Const HIST_FAVORITES = 2
Public Const HIST_ADDTOFAVORITES = 3
Public Const HIST_VIEWTREE = 4
#End If

#End If

#If (WIN32_IE_ >= &H400) Then
Public Const TB_ADDBUTTONSA = (WM_USER + 20)
Public Const TB_INSERTBUTTONA = (WM_USER + 21)
#Else
Public Const TB_ADDBUTTONS = (WM_USER + 20)
Public Const TB_INSERTBUTTON = (WM_USER + 21)
#End If

Public Const TB_DELETEBUTTON = (WM_USER + 22)
Public Const TB_GETBUTTON = (WM_USER + 23)
Public Const TB_BUTTONCOUNT = (WM_USER + 24)
Public Const TB_COMMANDTOINDEX = (WM_USER + 25)

#If WIN32_ Then

Public Type tagTBSAVEPARAMSA
    hkr As HKEY
    pszSubKey As LPCSTR
    pszValueName As LPCSTR
End Type

Public Type tagTBSAVEPARAMSW
    hkr As HKEY
    pszSubKey As LPCWSTR
    pszValueName As LPCWSTR
End Type

#If UNICODE Then
Public Const TBSAVEPARAMS = TBSAVEPARAMSW
Public Const LPTBSAVEPARAMS = LPTBSAVEPARAMSW
#Else
''Public Const TBSAVEPARAMS = TBSAVEPARAMSA
''Public Const LPTBSAVEPARAMS = LPTBSAVEPARAMSA
#End If

#End If     ' WIN32_

Public Const TB_SAVERESTOREA = (WM_USER + 26)
Public Const TB_SAVERESTOREW = (WM_USER + 76)
Public Const TB_CUSTOMIZE = (WM_USER + 27)
Public Const TB_ADDSTRINGA = (WM_USER + 28)
Public Const TB_ADDSTRINGW = (WM_USER + 77)
Public Const TB_GETITEMRECT = (WM_USER + 29)
Public Const TB_BUTTONSTRUCTSIZE = (WM_USER + 30)
Public Const TB_SETBUTTONSIZE = (WM_USER + 31)
Public Const TB_SETBITMAPSIZE = (WM_USER + 32)
Public Const TB_AUTOSIZE = (WM_USER + 33)
Public Const TB_GETTOOLTIPS = (WM_USER + 35)
Public Const TB_SETTOOLTIPS = (WM_USER + 36)
Public Const TB_SETPARENT = (WM_USER + 37)
Public Const TB_SETROWS = (WM_USER + 39)
Public Const TB_GETROWS = (WM_USER + 40)
Public Const TB_SETCMDID = (WM_USER + 42)
Public Const TB_CHANGEBITMAP = (WM_USER + 43)
Public Const TB_GETBITMAP = (WM_USER + 44)
Public Const TB_GETBUTTONTEXTA = (WM_USER + 45)
Public Const TB_GETBUTTONTEXTW = (WM_USER + 75)
Public Const TB_REPLACEBITMAP = (WM_USER + 46)

#If (WIN32_IE_ >= &H300) Then
Public Const TB_SETINDENT = (WM_USER + 47)
Public Const TB_SETIMAGELIST = (WM_USER + 48)
Public Const TB_GETIMAGELIST = (WM_USER + 49)
Public Const TB_LOADIMAGES = (WM_USER + 50)
Public Const TB_GETRECT = (WM_USER + 51)                ' wParam is the Cmd instead of index
Public Const TB_SETHOTIMAGELIST = (WM_USER + 52)
Public Const TB_GETHOTIMAGELIST = (WM_USER + 53)
Public Const TB_SETDISABLEDIMAGELIST = (WM_USER + 54)
Public Const TB_GETDISABLEDIMAGELIST = (WM_USER + 55)
Public Const TB_SETSTYLE = (WM_USER + 56)
Public Const TB_GETSTYLE = (WM_USER + 57)
Public Const TB_GETBUTTONSIZE = (WM_USER + 58)
Public Const TB_SETBUTTONWIDTH = (WM_USER + 59)
Public Const TB_SETMAXTEXTROWS = (WM_USER + 60)
Public Const TB_GETTEXTROWS = (WM_USER + 61)
#End If     ' WIN32_IE_ >= &h0300

#If UNICODE Then
Public Const TB_GETBUTTONTEXT = TB_GETBUTTONTEXTW
Public Const TB_SAVERESTORE = TB_SAVERESTOREW
Public Const TB_ADDSTRING = TB_ADDSTRINGW
#Else
Public Const TB_GETBUTTONTEXT = TB_GETBUTTONTEXTA
Public Const TB_SAVERESTORE = TB_SAVERESTOREA
Public Const TB_ADDSTRING = TB_ADDSTRINGA
#End If
#If (WIN32_IE_ >= &H400) Then
Public Const TB_GETOBJECT = (WM_USER + 62)              ' wParam == IID, lParam void **ppv
Public Const TB_GETHOTITEM = (WM_USER + 71)
Public Const TB_SETHOTITEM = (WM_USER + 72)             ' wParam == iHotItem
Public Const TB_SETANCHORHIGHLIGHT = (WM_USER + 73)     ' wParam == TRUE/FALSE
Public Const TB_GETANCHORHIGHLIGHT = (WM_USER + 74)
Public Const TB_MAPACCELERATORA = (WM_USER + 78)        ' wParam == ch, lParam int * pidBtn

Public Type TBINSERTMARK
    iButton As Long
    dwFlags As Long
End Type

Public Const TBIMHT_AFTER = &H1                         ' TRUE = insert After iButton, otherwise before
Public Const TBIMHT_BACKGROUND = &H2                    ' TRUE iff missed buttons completely

Public Const TB_GETINSERTMARK = (WM_USER + 79)          ' lParam == LPTBINSERTMARK
Public Const TB_SETINSERTMARK = (WM_USER + 80)          ' lParam == LPTBINSERTMARK
Public Const TB_INSERTMARKHITTEST = (WM_USER + 81)      ' wParam == LPPOINT lParam == LPTBINSERTMARK
Public Const TB_MOVEBUTTON = (WM_USER + 82)
Public Const TB_GETMAXSIZE = (WM_USER + 83)             ' lParam == LPSIZE
Public Const TB_SETEXTENDEDSTYLE = (WM_USER + 84)       ' For TBSTYLE_EX_*
Public Const TB_GETEXTENDEDSTYLE = (WM_USER + 85)       ' For TBSTYLE_EX_*
Public Const TB_GETPADDING = (WM_USER + 86)
Public Const TB_SETPADDING = (WM_USER + 87)
Public Const TB_SETINSERTMARKCOLOR = (WM_USER + 88)
Public Const TB_GETINSERTMARKCOLOR = (WM_USER + 89)

'Public Const TB_SETCOLORSCHEME = CCM_SETCOLORSCHEME    ' lParam is color scheme
'Public Const TB_GETCOLORSCHEME = CCM_GETCOLORSCHEME    ' fills in COLORSCHEME pointed to by lParam

'Public Const TB_SETUNICODEFORMAT = CCM_SETUNICODEFORMAT
'Public Const TB_GETUNICODEFORMAT = CCM_GETUNICODEFORMAT

Public Const TB_MAPACCELERATORW = (WM_USER + 90)        ' wParam == ch, lParam int * pidBtn
#If UNICODE Then
Public Const TB_MAPACCELERATOR = TB_MAPACCELERATORW
#Else
Public Const TB_MAPACCELERATOR = TB_MAPACCELERATORA
#End If
#End If     ' WIN32_IE_ >= &h0400

Public Type TBREPLACEBITMAP
    hInstOld As Long
    nIDOld As Long
    hInstNew As Long
    nIDNew As Long
    nButtons As Long
End Type

#If WIN32_ Then

Public Const TBBF_LARGE = &H1

Public Const TB_GETBITMAPFLAGS = (WM_USER + 41)

#If (WIN32_IE_ >= &H400) Then
Public Const TBIF_IMAGE = &H1
Public Const TBIF_TEXT = &H2
Public Const TBIF_STATE = &H4
Public Const TBIF_STYLE = &H8
Public Const TBIF_LPARAM = &H10
Public Const TBIF_COMMAND = &H20
Public Const TBIF_SIZE = &H40

Public Type TBBUTTONINFOA
    cbSize As Long
    dwMask As Long
    idCommand As Long
    iImage As Long
    fsState As Byte
    fsStyle As Byte
    cx As Integer
    lParam As Long
'    pszText As LPSTR
    cchText As Long
End Type

Public Type TBBUTTONINFOW
    cbSize As Long
    dwMask As Long
    idCommand As Long
    iImage As Long
    fsState As Byte
    fsStyle As Byte
    cx As Integer
    lParam As Long
'    pszText As LPWSTR
    cchText As Long
End Type

#If UNICODE Then
Public Const TBBUTTONINFO = TBBUTTONINFOW
Public Const LPTBBUTTONINFO = LPTBBUTTONINFOW
#Else
'Public Const TBBUTTONINFO = TBBUTTONINFOA
'Public Const LPTBBUTTONINFO = LPTBBUTTONINFOA
#End If

' BUTTONINFO APIs do NOT support the string pool.
Public Const TB_GETBUTTONINFOW = (WM_USER + 63)
Public Const TB_SETBUTTONINFOW = (WM_USER + 64)
Public Const TB_GETBUTTONINFOA = (WM_USER + 65)
Public Const TB_SETBUTTONINFOA = (WM_USER + 66)
#If UNICODE Then
Public Const TB_GETBUTTONINFO = TB_GETBUTTONINFOW
Public Const TB_SETBUTTONINFO = TB_SETBUTTONINFOW
#Else
Public Const TB_GETBUTTONINFO = TB_GETBUTTONINFOA
Public Const TB_SETBUTTONINFO = TB_SETBUTTONINFOA
#End If

Public Const TB_INSERTBUTTONW = (WM_USER + 67)
Public Const TB_ADDBUTTONSW = (WM_USER + 68)

Public Const TB_HITTEST = (WM_USER + 69)

' New post Win95/NT4 for InsertButton and AddButton.  if iString member
' is a pointer to a string, it will be handled as a string like listview
' (although LPSTR_TEXTCALLBACK is not supported).
#If UNICODE Then
Public Const TB_INSERTBUTTON = TB_INSERTBUTTONW
Public Const TB_ADDBUTTONS = TB_ADDBUTTONSW
#Else
Public Const TB_INSERTBUTTON = TB_INSERTBUTTONA
Public Const TB_ADDBUTTONS = TB_ADDBUTTONSA
#End If

Public Const TB_SETDRAWTEXTFLAGS = (WM_USER + 70)       ' wParam == mask lParam == bit values

#End If        ' WIN32_IE_ >= &h0400

Public Const TBN_GETBUTTONINFOA = (TBN_FIRST - 0)
Public Const TBN_GETBUTTONINFOW = (TBN_FIRST - 20)
Public Const TBN_BEGINDRAG = (TBN_FIRST - 1)
Public Const TBN_ENDDRAG = (TBN_FIRST - 2)
Public Const TBN_BEGINADJUST = (TBN_FIRST - 3)
Public Const TBN_ENDADJUST = (TBN_FIRST - 4)
Public Const TBN_RESET = (TBN_FIRST - 5)
Public Const TBN_QUERYINSERT = (TBN_FIRST - 6)
Public Const TBN_QUERYDELETE = (TBN_FIRST - 7)
Public Const TBN_TOOLBARCHANGE = (TBN_FIRST - 8)
Public Const TBN_CUSTHELP = (TBN_FIRST - 9)
#If (WIN32_IE_ >= &H300) Then
Public Const TBN_DROPDOWN = (TBN_FIRST - 10)
#End If
#If (WIN32_IE_ >= &H400) Then
Public Const TBN_GETOBJECT = (TBN_FIRST - 12)

' Structure for TBN_HOTITEMCHANGE notification
Public Type tagNMTBHOTITEM
'    hdr As NMHDR
    idOld As Long
    idNew As Long
    dwFlags As Long            ' HICF_*
End Type


' Hot item change flags
Public Const HICF_OTHER = &H0
Public Const HICF_MOUSE = &H1                           ' Triggered by mouse
Public Const HICF_ARROWKEYS = &H2                       ' Triggered by arrow keys
Public Const HICF_ACCELERATOR = &H4                     ' Triggered by accelerator
Public Const HICF_DUPACCEL = &H8                        ' This accelerator is not unique
Public Const HICF_ENTERING = &H10                       ' idOld is invalid
Public Const HICF_LEAVING = &H20                        ' idNew is invalid
Public Const HICF_RESELECT = &H40                       ' hot item reselected

Public Const TBN_HOTITEMCHANGE = (TBN_FIRST - 13)
Public Const TBN_DRAGOUT = (TBN_FIRST - 14)             ' this is sent when the user clicks down on a button then drags off the button
Public Const TBN_DELETINGBUTTON = (TBN_FIRST - 15)      ' uses TBNOTIFY
Public Const TBN_GETDISPINFOA = (TBN_FIRST - 16)        ' This is sent when the  toolbar needs  some display information
Public Const TBN_GETDISPINFOW = (TBN_FIRST - 17)        ' This is sent when the  toolbar needs  some display information
Public Const TBN_GETINFOTIPA = (TBN_FIRST - 18)
Public Const TBN_GETINFOTIPW = (TBN_FIRST - 19)


Public Type tagNMTBGETINFOTIPA
'    hdr As NMHDR
'    pszText As LPSTR
    cchTextMax As Long
    iItem As Long
    lParam As Long
End Type

Public Type tagNMTBGETINFOTIPW
'    hdr As NMHDR
'    pszText As LPWSTR
    cchTextMax As Long
    iItem As Long
    lParam As Long
End Type

#If UNICODE Then
Public Const TBN_GETINFOTIP = TBN_GETINFOTIPW
Public Const NMTBGETINFOTIP = NMTBGETINFOTIPW
Public Const LPNMTBGETINFOTIP = LPNMTBGETINFOTIPW
#Else
Public Const TBN_GETINFOTIP = TBN_GETINFOTIPA
'Public Const NMTBGETINFOTIP = NMTBGETINFOTIPA
'Public Const LPNMTBGETINFOTIP = LPNMTBGETINFOTIPA
#End If

Public Const TBNF_IMAGE = &H1
Public Const TBNF_TEXT = &H2
Public Const TBNF_DI_SETITEM = &H10000000

Public Type NMTBDISPINFOA
'    hdr As NMHDR
    dwMask As Long          ' [in] Specifies the values requested .[out] Client ask the data to be set for future use
    idCommand As Long       ' [in] id of button we're requesting info for
    lParam As Long          ' [in] lParam of button
    iImage As Long          ' [out] image index
'    pszText As LPSTR       ' [out] new text for item
    cchText As Long         ' [in] size of buffer pointed to by pszText
End Type

Public Type NMTBDISPINFOW
'    hdr As NMHDR
    dwMask As Long          ' [in] Specifies the values requested .[out] Client ask the data to be set for future use
    idCommand As Long       ' [in] id of button we're requesting info for
    lParam As Long          ' [in] lParam of button
    iImage As Long          ' [out] image index
'    pszText As LPWSTR      ' [out] new text for item
    cchText As Long         ' [in] size of buffer pointed to by pszText
End Type


#If UNICODE Then
Public Const TBN_GETDISPINFO = TBN_GETDISPINFOW
Public Const NMTBDISPINFO = NMTBDISPINFOW
Public Const LPNMTBDISPINFO = LPNMTBDISPINFOW
#Else
Public Const TBN_GETDISPINFO = TBN_GETDISPINFOA
'Public Const NMTBDISPINFO = NMTBDISPINFOA
'Public Const LPNMTBDISPINFO = LPNMTBDISPINFOA
#End If

' Return codes for TBN_DROPDOWN
Public Const TBDDRET_DEFAULT = 0
Public Const TBDDRET_NODEFAULT = 1
Public Const TBDDRET_TREATPRESSED = 2                   ' Treat as a standard press button

#End If


#If UNICODE Then
Public Const TBN_GETBUTTONINFO = TBN_GETBUTTONINFOW
#Else
Public Const TBN_GETBUTTONINFO = TBN_GETBUTTONINFOA
#End If

#If (WIN32_IE_ >= &H300) Then
'Public Const TBNOTIFYA = NMTOOLBARA
'Public Const TBNOTIFYW = NMTOOLBARW
'Public Const LPTBNOTIFYA = LPNMTOOLBARA
'Public Const LPTBNOTIFYW = LPNMTOOLBARW
#Else
'Public Const tagNMTOOLBARA = tagTBNOTIFYA
'Public Const NMTOOLBARA = TBNOTIFYA
'Public Const LPNMTOOLBARA = LPTBNOTIFYA
'Public Const tagNMTOOLBARW = tagTBNOTIFYW
'Public Const NMTOOLBARW = TBNOTIFYW
'Public Const LPNMTOOLBARW = LPTBNOTIFYW
#End If

'Public Const TBNOTIFY = NMTOOLBAR
'Public Const LPTBNOTIFY = LPNMTOOLBAR

#If (WIN32_IE_ >= &H300) Then
Public Type tagNMTOOLBARA
'    hdr As NMHDR
    iItem As Long
    TBBUTTON As TBBUTTON
    cchText As Long
'    pszText As LPSTR
End Type

#End If


#If (WIN32_IE_ >= &H300) Then
Public Type tagNMTOOLBARW
'    hdr As NMHDR
    iItem As Long
    TBBUTTON As TBBUTTON
    cchText As Long
'    pszText As LPWSTR
End Type

#End If


#If UNICODE Then
Public Const NMTOOLBAR = NMTOOLBARW
Public Const LPNMTOOLBAR = LPNMTOOLBARW
#Else
'Public Const NMTOOLBAR = NMTOOLBARA
'Public Const LPNMTOOLBAR = LPNMTOOLBARA
#End If

#End If
'-----------------------------------------------------------

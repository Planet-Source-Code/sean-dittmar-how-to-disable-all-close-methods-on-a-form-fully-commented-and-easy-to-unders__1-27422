VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "exit"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is code by Sean Dittmar. Do what you want with it,
'but give credit!

Dim hSysMenu As Long      ' This is the system menu's handle.

Dim zMENU As MENUITEMINFO ' This is a structure that is used
                          ' for the modification of the
                          ' system menu.

Private Sub cmdExit_Click()
End
End Sub

'---------------------IMPORTANT!----------------------
'Hey, a good way to make this tutorial sink in deeper:
'Start a new standard exe program and follow along!
'-----------------------------------------------------

Private Sub Form_Load()
Show

'Step 1: You need to get the system menu's handle.
'=================================================
hSysMenu = GetSystemMenu(Me.hwnd, False)
'                           |       |
'(me.hwnd)------------------+       |
'GetSystemMenu needs the form's     |
'hWnd.                              |
'                                   |
'(False)----------------------------+
'If true, GetSystemMenu returns the system
'menu back to its default state. If false,
'hSysMenu receives the system menu handle.

'Step 2: Populate the MENUITEMINFO structure.
'============================================
With zMENU
    .cbSize = Len(zMENU)
    ' cbSize is given the size of the
    ' MENUITEMINFO structure
    
    .dwTypeData = String(80, 0)
    ' dwTypeData needs to be stretched out to fit
    ' the maximum characters it can hold, 80.
    
    .cch = Len(.dwTypeData)
    ' cch is given the size of dwTypeData
    ' which will be 80.

    .fMask = MENU_STATE
    ' fMask needs this constant.
    
    .wID = SC_CLOSE
    ' This is the option we want because we
    ' wanna disable it.
End With

'Step 3: Get info on a particular system menu.
'========================================
retval = GetMenuItemInfo(hSysMenu, zMENU.wID, False, zMENU)
'                           |         |         |       |
'(hSysMenu)-----------------+         |         |       |
'GetMenuItemInfo needs the system     |         |       |
'menu handle.                         |         |       |
'                                     |         |       |
'(zMENU.wID)--------------------------+         |       |
'This is the item to get info about.            |       |
'                                               |       |
'(False)----------------------------------------+       |
'If true, the prior parameter, zMENU.wID, will be       |
'looked at as a menu position. If false, it will        |
'be look at like a menu id.                             |
'                                                       |
'(zMENU)------------------------------------------------+
'GetMenuItemInfo needs the MENUITEMINFO structure that
'you populated in step 2.

'Step 4: Specify that all "close" methods be grayed out.
'=========================================
Dim lngOldId As Long

With zMENU
    
    lngOldId = .wID         'You need the old wID.
    .wID = xSC_CLOSE        'Change the wID to "no close"
    .fState = MFS_GRAYED    'Make the close methods gray
    .fMask = MENU_ID        'Specifys that the value in
                            'wID is a id and not a state
End With

'Step 5: Make it so there is no way to close the form!
'=====================================================

retval = SetMenuItemInfo(hSysMenu, lngOldId, False, zMENU)

'Tada! The x in the top-right is disabled. But wait! When
'you right-click the title bar, there's the option to close
'the form. Here's how to disable that:

zMENU.fMask = MENU_STATE

retval = SetMenuItemInfo(hSysMenu, zMENU.wID, False, zMENU)

'Thanks for downloading!
End Sub

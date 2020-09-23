Attribute VB_Name = "Module1"
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpmii As MENUITEMINFO) As Long
Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Public Type MENUITEMINFO
    cbSize        As Long
    fMask         As Long
    fType         As Long
    fState        As Long
    wID           As Long
    hSubMenu      As Long
    hbmpChecked   As Long
    hbmpUnchecked As Long
    dwItemData    As Long
    dwTypeData    As String
    cch           As Long
End Type

'Menu item constants.
Public Const SC_CLOSE       As Long = &HF060&
Public Const xSC_CLOSE   As Long = -10
'SetMenuItemInfo fMask constants.
Public Const MENU_STATE     As Long = &H1&
Public Const MENU_ID        As Long = &H2&

'SetMenuItemInfo fState constants.
Public Const MFS_GRAYED     As Long = &H3&
Public Const MFS_CHECKED    As Long = &H8&

'SendMessage constants.
Public Const WM_NCACTIVATE  As Long = &H86


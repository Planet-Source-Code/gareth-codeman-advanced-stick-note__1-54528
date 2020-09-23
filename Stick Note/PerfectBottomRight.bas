Attribute VB_Name = "PerfectBottomRight"



Public Declare Function SHAppBarMessage Lib "shell32.dll" _
    (ByVal dwMessage As Long, pData As APPBARDATA) As Long
    Public Const ABM_GETTASKBARPOS = &H5
    Public Const ABE_TOP = 1
    Public Const ABE_RIGHT = 2
    Public Const ABE_LEFT = 0
    Public Const ABE_BOTTOM = 3


Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type


Public Type APPBARDATA
    cbSize As Long
    hWnd As Long
    uCallBackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long
    End Type





Public Function SendBottomRight(frmCurrent As Form) As Boolean

    On Error GoTo ErrHandler
    SendBottomRight = True
    Dim typAppBarData As APPBARDATA
    Dim intTwipsPerPixel As Integer
    typAppBarData.hWnd = frmCurrent.hWnd
    typAppBarData.cbSize = Len(typAppBarData)
    SHAppBarMessage ABM_GETTASKBARPOS, typAppBarData


    If typAppBarData.uEdge = ABE_BOTTOM Then

        frmCurrent.Top = Screen.Height - frmCurrent.Height - ((Int(Screen.Height / Screen.TwipsPerPixelY) - typAppBarData.rc.Top) * Screen.TwipsPerPixelY)
        frmCurrent.Left = Screen.Width - frmCurrent.Width
    End If


    If typAppBarData.uEdge = ABE_RIGHT Then

        frmCurrent.Top = Screen.Height - frmCurrent.Height
        frmCurrent.Left = Screen.Width - frmCurrent.Width - ((Int(Screen.Width / Screen.TwipsPerPixelX) - typAppBarData.rc.Left) * Screen.TwipsPerPixelX)
    End If


    If typAppBarData.uEdge = ABE_TOP Or typAppBarData.uEdge = ABE_LEFT Then

        frmCurrent.Top = Screen.Height - frmCurrent.Height
        frmCurrent.Left = Screen.Width - frmCurrent.Width
    End If
    Exit Function
ErrHandler:
    SendBottomRight = False
End Function







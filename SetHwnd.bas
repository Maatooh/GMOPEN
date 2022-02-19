Attribute VB_Name = "GetHwnd"
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
        ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
        Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public GHWND As Long

Function GetHwnd(THwnd)
Dim cWnd As Long, pWnd As Long
Dim Title As String * 255
Do
    cWnd = FindWindowEx(pWnd, cWnd, vbNullString, vbNullString)
    If cWnd Then
         GetWindowText cWnd, Title, Len(Title)
         If Not InStr(Title, THwnd) = 0 Then
         GetHwnd = cWnd
         End If
    Else
         Exit Do
    End If
Loop
End Function

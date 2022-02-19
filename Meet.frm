VERSION 5.00
Begin VB.Form Meet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meetbot"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3435
   Icon            =   "Meet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   61
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   229
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer SetKey 
      Interval        =   10
      Left            =   2640
      Top             =   240
   End
   Begin VB.Timer AcceptC 
      Interval        =   3000
      Left            =   2160
      Top             =   240
   End
   Begin VB.CommandButton Ini 
      Caption         =   "Test"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Meet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim Web As New WebDriver
Dim Mkey As New MaatoohVkey
Dim GHWND As Long
Dim Search As Long
Dim OP As Boolean
'ChromeDriver https://inter2match.com/mtbot/chromedriver.exe
'Update https://inter2match.com/mtbot/meet.exe

Private Sub RegisterLibrary()

Shell "cmd.exe /c cd " & App.Path & "&&" & _
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase MaatoohBrws.dll /tlb:MaatoohBrws64.tlb" & "&&" & _
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe /codebase MaatoohBrws.dll /tlb:MaatoohBrws32.tlb"
End Sub

Private Function TimeFx(SetTime As Long)
Dim TM As Long
TM = Timer
Do Until Timer - TM > SetTime
DoEvents
Loop
End Function

Private Sub AcceptC_Timer()
On Error GoTo Err
'--DeleteOld
If Not Dir(App.Path & "\OldMeetBot.exe", vbArchive) = "" Then
Kill (App.Path & "\OldMeetBot.exe")
End If
'---------
GHWND = GetHwnd.GetHwnd("Google Chrome")
If Not GHWND = 0 Then
Web.ExecuteScript ("document.querySelector('#yDmH0d > div.llhEMd.iWO5td > div > div.g3VIld.T9cDKf.vDc8Ic.J9Nfi.iWO5td > div.XfpsVe.J9fJmf > div.CE5PDe').lastChild.click()")
End If
If GHWND = 0 Then
Web.Quit
'---Exit
If OP = True Then
End
End If
'---
End If
Err:
End Sub

Private Sub Form_Load()
'--Close Chrome
TimeFx (6)
'---
Me.Hide
Shell "taskkill /f /im explorer.exe"
OP = False
Ini_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Ini_Click()
On Error GoTo Err
Go:
Call RegisterLibrary
Call Web.Quit
Call Web.SetProfile("C:\ChromeP")
Call Web.Start("Chrome")
Call Web.Get("https://meet.google.com/zko-fffk-rjj")
'---GetHwnd
Do Until Not GHWND = 0
GHWND = GetHwnd.GetHwnd("Google Chrome")
DoEvents
Loop
'---------
Call Mkey.VKeyD(vbKeyF11)
Call Mkey.VKeyU(vbKeyF11)
Call TimeFx(4)
Call Web.ExecuteScript("document.querySelector('#yDmH0d > c-wiz > div > div > div:nth-child(9) > div.crqnQb > div > div > div.vgJExf > div > div > div.d7iDfe.NONs6c > div > div.Sla0Yd > div > div.XCoPyb').firstChild.click()")
OP = True
Exit Sub
Err:
Call SetDriver
GoTo Go
End Sub

Private Sub SetKey_Timer()
On Error GoTo Err
If Not GHWND = 0 Then
'---General
If Not GetAsyncKeyState(vbKeyA) = 0 Then
Web.ExecuteScript ("document.querySelector('#ow3 > div.T4LgNb > div > div:nth-child(9) > div.crqnQb > div.DAQYgc.xPh1xb.P9KVBf > div.rceXCe > div > div.AVk6L.gwmyUe > div > div > span').firstChild.click()")
TimeFx (1)
Exit Sub
End If
If Not GetAsyncKeyState(vbKeyS) = 0 Then
Web.ExecuteScript ("document.querySelector('#ow3 > div.T4LgNb > div > div:nth-child(9) > div.crqnQb > div.DAQYgc.xPh1xb.P9KVBf > div.rceXCe > div > div.CrGlle.dq8L2c > div > span').firstChild.click()")
TimeFx (1)
Exit Sub
End If
If Not GetAsyncKeyState(vbKeyF10) = 0 Then
Web.Quit
Shell "explorer.exe"
TimeFx (1)
End
Exit Sub
End If
If Not GetAsyncKeyState(vbKeyF9) = 0 Then
Call SetUpdate
Exit Sub
End If
'---
End If
Err:
End Sub

Private Sub SetDriver()
Shell "taskkill /f /im ChromeDriver.exe"
TimeFx (2)
If Not Dir(App.Path & "\ChromeDriver.exe", vbArchive) = "" Then
Kill (App.Path & "\ChromeDriver.exe")
End If
Call URLDownloadToFile(0, "https://inter2match.com/mtbot/chromedriver.exe", App.Path & "\ChromeDriver.exe", 0, 0)
End Sub

Private Sub SetUpdate()
If Not Dir(App.Path & "\MeetBot.exe", vbArchive) = "" Then
Name App.Path & "\MeetBot.exe" As App.Path & "\OldMeetBot.exe"
End If
TimeFx (2)
Call URLDownloadToFile(0, "https://inter2match.com/mtbot/Meetbot.exe", App.Path & "\MeetBot.exe", 0, 0)
Do Until Not Dir(App.Path & "\MeetBot.exe", vbArchive) = ""
DoEvents
Loop
MsgBox "Se ha actualizado.", vbInformation, "MeetBot By Maatooh"
Shell App.Path & "\MeetBot.exe"
Web.Quit
End
End Sub

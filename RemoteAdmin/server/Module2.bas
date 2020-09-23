Attribute VB_Name = "Module1"
    '#########################################'
    '   Programmed By Inderpal Singh          '
    '   Email: inderpal0@hotmail.com          '
    '   Date: Dec 19, 2001                    '
    '   Homepage: http://connect.to/lanserver '
    '#########################################'

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
  ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function Get_Desktop(ByVal theFile As String) As Boolean
    Dim lString As String
        DoEvents
        DoEvents
        Call keybd_event(vbKeySnapshot, 1, 0, 0)
        DoEvents
        DoEvents
         'To get the Active Window
        SavePicture Clipboard.GetData(vbCFBitmap), theFile
        Get_Desktop = True
        Exit Function
Trap:
'Error handling
End Function

Public Sub SendFile(FileName As String, WinServer As Winsock)
    Dim FreeF As Integer
    Dim LenFile As Long
    Dim nCnt As Long
    Dim LocData As String
    Dim LoopTimes As Long
    Dim i As Long

    FreeF = FreeFile
    Open FileName For Binary As #99
        nCnt = 1
        LenFile = LOF(99)
        WinServer.SendData "FILESIZE" & LenFile
        DoEvents
        Sleep (400)
        Do Until nCnt >= (LenFile)
            LocData = Space$(1024) 'Set size of chunks
            Get #99, nCnt, LocData 'Get data from the file nCnt is from where to start the get
            If nCnt + 1024 > LenFile Then
                WinServer.SendData Mid$(LocData, 1, (LenFile - nCnt))
            Else
                WinServer.SendData LocData 'Send the chunk
            End If
            nCnt = nCnt + 1024
        Loop
    Close #99
End Sub




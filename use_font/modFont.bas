Attribute VB_Name = "modFont"
'BEFORE USING THIS MODULE PLEASE READ
'readme.txt

Option Explicit
Private Declare Function GetTempPath Lib "Kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function AddFontResource Lib "GDI32" Alias "AddFontResourceA" (ByVal FontFileName As String) As Long
Private Declare Function RemoveFontResource Lib "GDI32" Alias "RemoveFontResourceA" (ByVal FontFileName As String) As Long
Private Declare Function CreateScalableFontResource Lib "GDI32" Alias "CreateScalableFontResourceA" _
   (ByVal fHidden As Long, ByVal lpszResourceFile As String, _
   ByVal lpszFontFile As String, ByVal lpszCurrentPath As String) As Long

Private Function GetTempPathName() As String 'get Windows\Temp folder
    Dim sBuffer As String, lRet As Long
    sBuffer = String$(255, vbNullChar)
    lRet = GetTempPath(255, sBuffer)
    If lRet > 0 Then
    sBuffer = Left$(sBuffer, lRet)
    End If
    GetTempPathName = sBuffer
End Function

Private Function CheckFile(FileName As String) As Boolean 'check if a file exists
On Error GoTo ErrH
    CheckFile = False
    If Dir(FileName) <> "" Then
    If (GetAttr(FileName) And vbDirectory) = 0 Then
    CheckFile = True
    'file Exists
    Kill FileName 'file exists - kill it
    Else '
    'file dosen't exists
    'MsgBox "File doesn't exist!", vbCritical
    Exit Function
    End If
    Else
    'MsgBox "File doesn't exist!", vbCritical
    Exit Function
    End If
ErrH:
End Function

Public Function GetFontName(FileNameTTF As String) As String
    Dim hFile As Integer, Buffer As String, FontName As String, TempName As String, iPos As Integer
    TempName = GetTempPathName & "tempfntinfo.tmp" 'create a temporary fontInfo file in Windows\Temp directory
    CheckFile (TempName) 'check to see if temporary fontInfo exists
    If CreateScalableFontResource(0, TempName, FileNameTTF, vbNullString) Then
    hFile = FreeFile
    Open TempName For Binary Access Read As hFile 'open temp fontInfo file
        Buffer = Space(LOF(hFile))
        Get hFile, , Buffer
        iPos = InStr(Buffer, "FONTRES:") + 8 'font name is after "FONTRES:" (ex. "FONTRES:TheFontFace")
        FontName = Mid(Buffer, iPos, InStr(iPos, Buffer, vbNullChar) - iPos) 'font name
    Close hFile 'close temp fontInfo file
    Kill TempName 'delete temp fontInfo file
    End If
    GetFontName = FontName
End Function

Public Function UseFont(FontFileName As String)
'usage: UseFont("C:\fonts\Font.ttf")
AddFontResource (FontFileName) 'make the system beleve that the font is installed
UseFont = GetFontName(FontFileName)
End Function

Public Function RemoveFont(FontFileName As String)
'REMEMBER to remove the font(s) that you used otherwise the font(s)
'will temporary remain in your system and you will not be able to
'move or delete this file(s) until you restart the computer.
'
'Still if you don't remove the font, programs such as Word will
'recognize it in the font list. After restart the font will
'dissapear from the list.
'
'Remove only the font(s) that you added !
'
'usage: RemoveFont("C:\fonts\Font.ttf")
RemoveFontResource (FontFileName)
End Function


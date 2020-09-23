Attribute VB_Name = "Unlocker"
Public cbuff As String
Public datatype As Integer
Public dataok As Boolean
Public showcon As Boolean
Public strFileName As String

'CLIPBOARD FORMATS FOR VB 6.0
'List taken from MSDN
'vbCFBitmap = Bitmap
'vbCFDIB =  Dib
'vbCFEMetafile =  EnhancedMetafile
'vbCFFiles = FileDrop
'vbCFLink =  No equivalent. For more information, see Dynamic Data Exchange for Visual Basic 6.0 Users.
'vbCFMetafile =  MetafilePict
'vbCFPalette =  Palette
'vbCFRTF =  Rtf
'vbCFText =  Text
'=============================================================================
Public Sub Main()
On Error Resume Next
Con.Initialize
Con.Title = "Clipboard Saver...."
datatype = Clipboard.GetData

Dim a_strArgs() As String
Dim blnDebug As Boolean
Dim i As Integer
Dim fileok As Boolean

dataok = False
showcon = False

If Clipboard.GetFormat(vbCFText) Then
   dataok = True
   cbuff = Clipboard.GetText
End If

   a_strArgs = Split(Command$, " ")
  
    If UBound(a_strArgs) = -1 Or UBound(a_strArgs) > 2 Then
        If Con.Compiled Then
            'MsgBox "Invalid arguments specified."
            Con.ForeColor = conRedHi
            p1 = Con.WriteLine("You did not specify correct arguments.", True, , conAlignCentered)
            Con.ForeColor = conWhite
            dataok = False
            Exit Sub
        Else
            disperr ("Invalid arguments specified.")
        End If
    End If
   
   For i = LBound(a_strArgs) To UBound(a_strArgs)
      
      Select Case LCase(a_strArgs(i))
      
      Case "-f", "/f"
      ' filename specified
         dataok = False
         'showcon = False
         fileok = True
         
         If i = UBound(a_strArgs) Then
            disperr "Filename not specified."
            dataok = False
         Else
            i = i + 1
         End If
         
         If Left(a_strArgs(i), 1) = "-" Or Left(a_strArgs(i), 1) = "/" Then
            'MsgBox "Invalid filename."
            disperr ("Invalid Filename")
         Else
            strFileName = a_strArgs(i)
            dataok = True
            fileok = True
         End If
                        
      Case "-h", "/h"
        showhelp
      
      Case Else
         dataok = False
         disperr ("Invalid argument : " & a_strArgs(i))
      End Select
      
   Next i

    'If Dir(strFilename) = "" Then
    '    dataok = False
    '    disperr ("This filename " & strFilename & " does not exist")
    'Exit Sub
    'End If

If fileok = True Then
  Call filecloser(srtfilename)
Else
    disperr ("Invalid arguments / filename")
    Exit Sub
End If

End Sub


Function disperr(errorstring As String)
Con.ForeColor = conRedHi
'Con.BackColor = conYellowHi
Con.WriteLine (errorstring)
Con.WriteLine (vbCrLf)
Con.BackColor = conBlack
Con.ForeColor = conWhite
Con.ReadLine ("Press <Enter> or <Return> to Continue....")
End Function


Function savefile(filen As String)
Dim fl As Long
fl = FreeFile
'MsgBox filen
Con.WriteLine ("Creating file .... " & filen)
Open filen For Output As #fl
Print #fl, cbuff
Close #fl

End Function

Function showcontents(Buffer As String)
Dim alltext() As String
Dim curpg As Integer
Dim howpg As Integer

alltext = Split(Buffer, vbCr)
k = 0
curpg = 0
howpg = Round((UBound(alltext) / 23), 0)

For j = 0 To UBound(alltext)
   
    If k = 23 Then '    Or j < UBound(alltext) Then
        Con.WriteLine
        curpg = curpg + 1
        Con.ForeColor = conWhiteHi
        p1 = Con.ReadLine("Press <Enter> or <Return> Key to Continue...." & "Page : " & Trim(Str(curpg)) & " of " & Trim(Str(howpg)), 1)
        Con.ForeColor = conWhite
        k = 0
    Else
        Con.WriteLine (alltext(j))
        k = k + 1
    End If
Next

Con.WaitingInput


End Function


Function showhelp()
Con.WriteLine ("")
Con.WriteLine ("Clipboard Text Saver" & vbCrLf & "(c) Shuvro Basu, 2009-2010" & vbCrLf & "===================================================================" & vbCrLf & "Usage : ClipSav /[-]s /[-]f <filename.ext> /[-]h" & vbCrLf & "If you use the -f option, then the Filename has to be supplied to save the contents to the file." & vbCrLf & "-s option displays the text on the screen 23 lines at a time" & vbCrLf & "-h displays this help text. Note: You can use both / or - in before the parameters." & vbCrLf & "For help / support / bugs contact shuvrobasu@gmail.com" & vbCrLf)
Con.WriteLine ("This program was developed using Visual Basic 6.0")
Con.WriteLine ("####################################################################")
Con.WriteLine (vbCrLf)

End Function


Function filecloser(fl As String)
If Not UnLockFile(fl) Then
    disperr ("Unable to open the lock!")
Else
    disperr ("File was unlocked or not locked at all!!")
End If
End Function



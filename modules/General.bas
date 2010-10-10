Attribute VB_Name = "General"
Option Explicit






Function GetApplicationFullPath() As String
    GetApplicationFullPath = App.path & "\" & App.EXEName & ".exe"
End Function



Function GetFileTitle(path As String) As String
      'Returns name of the file: C:\file.ext -> file.ext
          Dim temp() As String
          Dim output As String

10        If path = "" Then
20            GetFileTitle = ""
30            Exit Function
40        End If

50        temp = Split(path, "\")

60        output = temp(UBound(temp))
70        If output = "" Then
80            output = temp(UBound(temp))
90        End If

100       GetFileTitle = output
End Function

Function GetDrive(path As String) As String
          Dim str() As String
10        str = Split(path, ":")
          
20        If Len(str(0)) > 1 Then
30            GetDrive = ""
40        Else
50            GetDrive = str(0)
60        End If
End Function


Function GetPathTo(filepath As String) As String
      'Returns full path to filepath: C:\folder\file.ext -> C:\folder\
10    If filepath = "" Then
20      GetPathTo = ""
30    Else
40        GetPathTo = Mid$(filepath, 1, Len(filepath) - Len(GetFileTitle(filepath)))
50    End If
End Function

Function GetRelativePath(filepath As String, relativeto As String) As String
      'Returns filepath relative to relativeto
        
10        If GetDrive(filepath) <> GetDrive(relativeto) Then
              'if different drive, return full path
20            GetRelativePath = filepath
30            Exit Function
40        End If
          
      'C:\Program Files\Continuum\Server\mylvz.lvz
      'relative to
      'C:\Program Files\Continuum\DCME

      '..\Server\mylvz.lvz

      'C:\Program Files\Continuum\DCME\mylvz.lvz
      'relative to
      'C:\Program Files\Continuum\DCME\

      '..\Server\mylvz.lvz

50        While InStrRev(relativeto, "\") = Len(relativeto) And Len(relativeto) <> 0
60            relativeto = Left$(relativeto, Len(relativeto) - 1)
70        Wend
80        relativeto = relativeto & "\"
          
90        relativeto = LCase$(relativeto)
100       filepath = LCase$(filepath)
          
          Dim build As String
          
          Dim relativeparts() As String
          
110       Do
              Dim pos As Long
120           pos = InStr(filepath, relativeto)
              
              ' And InStrRev(filepath, "\") >= Len(relativeto)
130           If pos = 1 Then
                  'append filename
140               build = build & Mid$(filepath, Len(relativeto) + 1, Len(filepath) - Len(relativeto))
150               Exit Do
160           Else
170               build = build & "..\"
                  
180               Do
190                   relativeparts = Split(relativeto, "\")
                      
                      
200                   If UBound(relativeparts) = 0 Or InStr(relativeparts(UBound(relativeparts)), ":") Then
                          'we reached the root...
210                       GetRelativePath = ""
220                       Exit Function
230                   End If
                      
240                   If Len(relativeparts(UBound(relativeparts))) = 0 Then
250                       relativeto = Left$(relativeto, Len(relativeto) - 1)
260                   Else
270                       relativeto = Left$(relativeto, Len(relativeto) - Len(relativeparts(UBound(relativeparts))))
280                   End If
                      
290               Loop While Len(relativeparts(UBound(relativeparts))) = 0 And Len(relativeto) > 0
                  
300           End If
              
310       Loop While Len(relativeto) > 0
          
320       GetRelativePath = build
End Function


Function GetExtension(filetitle As String) As String
      'Returns extension of file: C:\file.ext -> ext
          Dim temp() As String
          Dim output As String
10        temp = Split(filetitle, ".")
20        If UBound(temp) < 0 Then
30            GetExtension = ""
40            Exit Function
50        End If

60        output = temp(UBound(temp))

70        GetExtension = LCase(output)
End Function

Function GetFileNameWithoutExtension(filetitle As String) As String
      'Returns filename of file with no extension: C:\file.ext -> file
          Dim temp() As String
          Dim output As String
10        temp = Split(filetitle, ".")
20        If UBound(temp) <= 0 Then
30            GetFileNameWithoutExtension = GetFileTitle(filetitle)
40            Exit Function
50        End If

60        output = temp(UBound(temp) - 1)


70        GetFileNameWithoutExtension = GetFileTitle(output)
End Function





Sub toggleLockToolTextBox(txt As TextBox, Optional lck As Boolean = False)
10        If txt.locked And Not lck Then
20            txt.locked = False
30            txt.BorderStyle = vbFixedSingle
40            txt.BackColor = vbWhite
50            txt.Alignment = vbCenter
60            Call ShowCaret(txt.hWnd)
70            txt.selstart = 0
80            txt.sellength = Len(txt.Text)
90        Else
100           txt.locked = True
110           txt.BorderStyle = 0
120           txt.BackColor = &H8000000F
130           txt.Alignment = vbRightJustify
140           Call HideCaret(txt.hWnd)
150       End If

End Sub


Sub removeDisallowedCharacters(ByRef txtbox As TextBox, lowerBound As Single, upperBound As Single, Optional dec As Boolean = False)
10        If lowerBound > upperBound Then
20            txtbox.Text = lowerBound
30            Exit Sub
40        End If
          
50        If (Not IsNumeric(txtbox.Text) And (txtbox.Text <> "-" Or lowerBound >= 0)) _
              Or InStr(txtbox.Text, "e") > 0 Or InStr(txtbox.Text, "E") > 0 _
              Or (Not dec And (InStr(txtbox.Text, ".") > 0 Or InStr(txtbox.Text, ",") > 0)) _
              Or (lowerBound < 0 And InStr(2, txtbox.Text, "-") > 1) _
              Or (lowerBound >= 0 And InStr(txtbox.Text, "-") > 0) Then
              
              Dim oldselstart As Integer
60            oldselstart = txtbox.selstart - 1    'char  typed so always one more
70            If oldselstart < 0 Then oldselstart = 0

              'remove all characters aside from nrs
              Dim i As Integer
              Dim finalresult As String
80            For i = 1 To Len(txtbox.Text)
90                If (Asc(Mid$(txtbox.Text, i, 1)) < Asc("0") Or _
                      Asc(Mid$(txtbox.Text, i, 1)) > Asc("9")) Then
                      Dim result As String
100                   If i - 1 >= 1 Then result = Mid$(txtbox.Text, 1, i - 1)
110                   If i + 1 <= Len(txtbox.Text) Then result = result + Mid$(txtbox.Text, i + 1, Len(txtbox.Text) - (i))
120                   finalresult = result
130               End If
140           Next
150           txtbox.Text = finalresult
160           If oldselstart > Len(txtbox.Text) Then
170               txtbox.selstart = Len(txtbox.Text)
180           Else
190               txtbox.selstart = oldselstart
200           End If
210       End If

220       If val(txtbox.Text) < lowerBound Then
230           txtbox.Text = lowerBound
240       End If

250       If val(txtbox.Text) > upperBound Then
260           txtbox.Text = upperBound
270       End If

End Sub





Function GetLongFilename(ByRef sShortName As String) As String
          Dim sTemp As String
          Dim sLongName As String
          Dim lSlashPos As Long
        If sShortName = "" Then
            GetLongFilename = ""
            Exit Function
        End If
        
        If Left$(sShortName, 2) = "\\" Then
            GetLongFilename = sShortName
        Else
10        If Right$(sShortName, 1) <> "\" Then sShortName = sShortName & "\"
20        lSlashPos = InStr(4, sShortName, "\")    ' start past x:\

30        Do While lSlashPos
40            sTemp = Dir$(Left$(sShortName, lSlashPos - 1), vbNormal Or vbHidden Or vbSystem Or vbDirectory)
50            If LenB(sTemp) = 0 Then Exit Function
60            sLongName = sLongName & "\" & sTemp
70            lSlashPos = InStr(lSlashPos + 1, sShortName, "\")
80        Loop

90        GetLongFilename = Left$(sShortName, 2) & sLongName

        End If
End Function







Function dist(x1 As Integer, y1 As Integer, x2 As Single, y2 As Single) As Single
10        dist = Abs(Math.Sqr(CLng((x2 - x1)) * (x2 - x1) + CLng((y2 - y1)) * (y2 - y1)))
End Function




Function HaveComponent(compname As String) As Boolean
10        compname = LCase(compname)
20        If FileExists(App.path & "\" & compname) Then
30            HaveComponent = True
40        ElseIf FileExists(SysDir & "\" & compname) Then
50            HaveComponent = True
60        ElseIf FileExists(WinDir & "\" & compname) Then
70            HaveComponent = True
80        Else
90            HaveComponent = False
100       End If
End Function




Function DirExists(DirName As String) As Boolean
10        On Error GoTo ErrorHandler
          ' test the directory attribute
        If DirName = "" Then
            DirExists = False
        Else
20        DirExists = GetAttr(DirName) And vbDirectory
        End If
        Exit Function
ErrorHandler:
        DirExists = False
          ' if an error occurs, this function returns False
End Function


Function FileExists(path As String, Optional fileattribute As VbFileAttribute = vbNormal) As Boolean
10        If path = "" Then
20          FileExists = False
30          Exit Function
40        End If
          
50        On Error GoTo errh
60        FileExists = (LCase(Dir$(path, fileattribute)) = LCase$(GetFileTitle(path)))
70        Exit Function
errh:
80        FileExists = False
90        Exit Function
End Function

Function DeleteFile(path As String, Optional fileattribute As VbFileAttribute = vbNormal) As Boolean
10        On Local Error GoTo DeleteFile_Error
20        If FileExists(path, fileattribute) Then
30            Kill path
40        Else
50            DeleteFile = False
60        End If
          
          'If file is still there, we failed
70        DeleteFile = Not FileExists(path, fileattribute)
          
80        Exit Function
DeleteFile_Error:
'90        HandleError Err, "DeleteFile(" & path & "," & fileattribute & ")", False, False
100       DeleteFile = False
End Function




Function DeleteDirectory(FullPath As String) As Boolean
   
'******************************************
'PURPOSE: DELETES A FOLDER, INCLUDING ALL SUB-
'         DIRECTORIES, FILES, REGARDLESS OF THEIR
'         ATTRIBUTES

'PARAMETER: FullPath = FullPath of Folder to Delete

'RETURNS:   True is successful, false otherwise

'REQUIRES:  'VB6
            'Reference to Microsoft Scripting Runtime
            'Caution in use for obvious reasons

'EXAMPLE:   'KillFolder("D:\MyOldFiles")

'******************************************
    On Error Resume Next
    
    Dim oFso
    Set oFso = CreateObject("Scripting.FileSystemObject")
    
    'deletefolder method does not like the "\"
    'at end of fullpath
    
    If Right(FullPath, 1) = "\" Then FullPath = _
        Left(FullPath, Len(FullPath) - 1)
    
    If oFso.FolderExists(FullPath) Then
        
        'Setting the 2nd parameter to true
        'forces deletion of read-only files
        oFso.DeleteFolder FullPath, True
        
        DeleteDirectory = Err.Number = 0 And _
          oFso.FolderExists(FullPath) = False
    End If
    
    Set oFso = Nothing
End Function





Function IsLcase(char As Integer) As Boolean
10        IsLcase = (char >= Asc("a") And char <= Asc("z"))
End Function

Function IsUcase(char As Integer) As Boolean
10        IsUcase = (char >= Asc("A") And char <= Asc("Z"))
End Function

Function IsNumber(char As Integer) As Boolean
10        IsNumber = (char >= Asc("0") And char <= Asc("9"))
End Function

Function decMod(val As Double, modval As Double) As Double
          Dim b As Integer
10        b = val \ modval
20        decMod = val - b * modval
End Function


'rounds number away from 0
'if number is positive, rounds up, else, rounds down
Function RoundAway(X As Single) As Integer
10        RoundAway = Sgn(-X) * Int(-Abs(X))
End Function

Function RoundAwayLong(X As Double) As Long
10        RoundAwayLong = Sgn(-X) * Int(-Abs(X))
End Function

Function Atn2(X As Double, Y As Double) As Double
10        If X = 0 Then
20            Atn2 = Sgn(Y) * PI / 2
30        ElseIf X > 0 Then
40            Atn2 = Atn(Y / X)
50        Else
60            Atn2 = Atn(Y / X) + PI
70        End If
End Function

Function intMinimum(val1 As Integer, val2 As Integer) As Integer
      '10        If val1 <= val2 Then
10        intMinimum = IIf(val1 <= val2, val1, val2)
      '30        Else
      '40            intMinimum = val2
      '50        End If
End Function

Function intMaximum(val1 As Integer, val2 As Integer) As Integer
      '10        If val1 >= val2 Then
10            intMaximum = IIf(val1 >= val2, val1, val2)
      '30        Else
      '40            intMaximum = val2
      '50        End If
End Function

Function longMinimum(val1 As Long, val2 As Long) As Long
      '10        If val1 <= val2 Then
10            longMinimum = IIf(val1 <= val2, val1, val2)
      '30        Else
      '40            longMinimum = val2
      '50        End If
End Function

Function longMaximum(val1 As Long, val2 As Long) As Long
      '10        If val1 >= val2 Then
10            longMaximum = IIf(val1 >= val2, val1, val2)
      '30        Else
      '40            longMaximum = val2
      '50        End If
End Function

Function doubleMinimum(val1 As Double, val2 As Double) As Double
10        If val1 <= val2 Then
20            doubleMinimum = val1
30        Else
40            doubleMinimum = val2
50        End If
End Function

Function doubleMaximum(val1 As Double, val2 As Double) As Double
10        If val1 >= val2 Then
20            doubleMaximum = val1
30        Else
40            doubleMaximum = val2
50        End If
End Function

Function RenameFile(Source As String, Destination As String) As Boolean
10        On Local Error GoTo RenameFile_Error
          
          Dim oFso
20        Set oFso = CreateObject("Scripting.FileSystemObject")
          
      '30        DoEvents
30        If FileExists(Source) Then
          
40            If FileExists(Destination, vbNormal + vbHidden) Then
50                Call SetAttr(Destination, vbNormal)
60                Kill (Destination)
70            End If
      '90            DoEvents

80            oFso.MoveFile Source, Destination
90        End If
      '120       DoEvents
          
100       Set oFso = Nothing
110       RenameFile = True
120       Exit Function
RenameFile_Error:
130       RenameFile = False
'140       HandleError Err, "RenameFile (" & Source & "," & Destination & ")", False,  False
End Function

Sub CreateDir(NewFolder As String)
10        On Error Resume Next
        
20      If DirExists(NewFolder) Then Exit Sub
        
          Dim path() As String
30        path = Split(NewFolder, "\")
          Dim progPath As String
40        progPath = path(0) & "\"
          
          Dim i As Integer
50        For i = 1 To UBound(path)
60            If path(i) <> "" Then
70                progPath = progPath & path(i) & "\"
80                Call CreateSingleDir(progPath)
90            End If
100       Next

End Sub

Sub CreateSingleDir(NewFolder As String)
10        On Error Resume Next
20        MkDir NewFolder
End Sub


' Evaluate the expression.
'source: http://vb-helper.com/howto_evaluate_expressions.html
'Function EvaluateExpression(ByVal expression As String) As Double
'      Const PREC_NONE = 11
'      Const PREC_UNARY = 10   ' Not actually used.
'      Const PREC_POWER = 9
'      Const PREC_TIMES = 8
'      Const PREC_DIV = 7
'      Const PREC_INT_DIV = 6
'      Const PREC_MOD = 5
'      Const PREC_PLUS = 4
'
'      Dim expr As String
'      Dim is_unary As Boolean
'      Dim next_unary As Boolean
'      Dim parens As Integer
'      Dim pos As Integer
'      Dim expr_len As Integer
'      Dim ch As String
'      Dim lexpr As String
'      Dim rexpr As String
'      Dim value As String
'      Dim status As Long
'      Dim best_pos As Integer
'      Dim best_prec As Integer
'
'          ' Remove all spaces.
'10        expr = replace$(expression, " ", "")
'20        expr_len = Len(expr)
'30        If expr_len = 0 Then
'40            EvaluateExpression = 0
'50            Exit Function
'60        End If
'
'          ' If we find + or - now, it is a unary operator.
'70        is_unary = True
'
'          ' So far we have nothing.
'80        best_prec = PREC_NONE
'
'          ' Find the operator with the lowest precedence.
'          ' Look for places where there are no open
'          ' parentheses.
'90        For pos = 1 To expr_len
'              ' Examine the next character.
'100           ch = Mid$(expr, pos, 1)
'
'              ' Assume we will not find an operator. In
'              ' that case, the next operator will not
'              ' be unary.
'110           next_unary = False
'
'120           If ch = " " Then
'                  ' Just skip spaces. We keep them here
'                  ' to make the error messages easier to
'130           ElseIf ch = "(" Then
'                  ' Increase the open parentheses count.
'140               parens = parens + 1
'
'                  ' A + or - after "(" is unary.
'150               next_unary = True
'160           ElseIf ch = ")" Then
'                  ' Decrease the open parentheses count.
'170               parens = parens - 1
'
'                  ' An operator after ")" is not unary.
'180               next_unary = False
'
'                  ' If parens < 0, too many ')'s.
'190               If parens < 0 Then
'200                   Err.Raise vbObjectError + 1001, _
'                          "EvaluateExpression", _
'                          "Too many )s in '" & _
'                          expression & "'"
'210               End If
'220           ElseIf parens = 0 Then
'                  ' See if this is an operator.
'230               If ch = "^" Or ch = "*" Or _
'                     ch = "/" Or ch = "\" Or _
'                     ch = "%" Or ch = "+" Or _
'                     ch = "-" _
'                  Then
'                      ' An operator after an operator
'                      ' is unary.
'240                   next_unary = True
'
'                      ' See if this operator has higher
'                      ' precedence than the current one.
'250                   Select Case ch
'                          Case "^"
'260                           If best_prec >= PREC_POWER Then
'270                               best_prec = PREC_POWER
'280                               best_pos = pos
'290                           End If
'
'300                       Case "*", "/"
'310                           If best_prec >= PREC_TIMES Then
'320                               best_prec = PREC_TIMES
'330                               best_pos = pos
'340                           End If
'
'350                       Case "\"
'360                           If best_prec >= PREC_INT_DIV Then
'370                               best_prec = PREC_INT_DIV
'380                               best_pos = pos
'390                           End If
'
'400                       Case "%"
'410                           If best_prec >= PREC_MOD Then
'420                               best_prec = PREC_MOD
'430                               best_pos = pos
'440                           End If
'
'450                       Case "+", "-"
'                              ' Ignore unary operators
'                              ' for now.
'460                           If (Not is_unary) And _
'                                  best_prec >= PREC_PLUS _
'                              Then
'470                               best_prec = PREC_PLUS
'480                               best_pos = pos
'490                           End If
'500                   End Select
'510               End If
'520           End If
'530           is_unary = next_unary
'540       Next pos
'
'          ' If the parentheses count is not zero,
'          ' there's a ')' missing.
'550       If parens <> 0 Then
'560           Err.Raise vbObjectError + 1002, _
'                  "EvaluateExpression", "Missing ) in '" & _
'                  expression & "'"
'570       End If
'
'          ' Hopefully we have the operator.
'580       If best_prec < PREC_NONE Then
'590           lexpr = Left$(expr, best_pos - 1)
'600           rexpr = Mid$(expr, best_pos + 1)
'610           Select Case Mid$(expr, best_pos, 1)
'                  Case "^"
'620                   EvaluateExpression = _
'                          EvaluateExpression(lexpr) ^ _
'                          EvaluateExpression(rexpr)
'630               Case "*"
'640                   EvaluateExpression = _
'                          EvaluateExpression(lexpr) * _
'                          EvaluateExpression(rexpr)
'650               Case "/"
'660                   EvaluateExpression = _
'                          EvaluateExpression(lexpr) / _
'                          EvaluateExpression(rexpr)
'670               Case "\"
'680                   EvaluateExpression = _
'                          EvaluateExpression(lexpr) \ _
'                          EvaluateExpression(rexpr)
'690               Case "%"
'700                   EvaluateExpression = _
'                          EvaluateExpression(lexpr) Mod _
'                          EvaluateExpression(rexpr)
'710               Case "+"
'720                   EvaluateExpression = _
'                          EvaluateExpression(lexpr) + _
'                          EvaluateExpression(rexpr)
'730               Case "-"
'740                   EvaluateExpression = _
'                          EvaluateExpression(lexpr) - _
'                          EvaluateExpression(rexpr)
'750           End Select
'760           Exit Function
'770       End If
'
'          ' If we do not yet have an operator, there
'          ' are several possibilities:
'          '
'          ' 1. expr is (expr2) for some expr2.
'          ' 2. expr is -expr2 or +expr2 for some expr2.
'          ' 3. expr is Fun(expr2) for a function Fun.
'          ' 4. expr is a primitive.
'          ' 5. It's a literal like "3.14159".
'
'          ' Look for (expr2).
'780       If Left$(expr, 1) = "(" And Right$(expr, 1) = ")" Then
'              ' Remove the parentheses.
'790           EvaluateExpression = EvaluateExpression( _
'                  Mid$(expr, 2, expr_len - 2))
'800           Exit Function
'810       End If
'
'          ' Look for -expr2.
'820       If Left$(expr, 1) = "-" Then
'830           EvaluateExpression = -EvaluateExpression( _
'                  Mid$(expr, 2))
'840           Exit Function
'850       End If
'
'          ' Look for +expr2.
'860       If Left$(expr, 1) = "+" Then
'870           EvaluateExpression = EvaluateExpression( _
'                  Mid$(expr, 2))
'880           Exit Function
'890       End If
'
'          ' Look for Fun(expr2).
'900       If expr_len > 5 And Right$(expr, 1) = ")" Then
'              ' Find the first (.
'910           pos = InStr(expr, "(")
'
'920           If pos > 0 Then
'                  ' See what the function is.
'930               lexpr = LCase$(Left$(expr, pos - 1))
'940               rexpr = Mid$(expr, pos + 1, expr_len - pos - 1)
'950               Select Case lexpr
'                      Case "sin"
'960                       EvaluateExpression = _
'                              Sin(EvaluateExpression(rexpr))
'970                       Exit Function
'980                   Case "cos"
'990                       EvaluateExpression = _
'                              Cos(EvaluateExpression(rexpr))
'1000                      Exit Function
'1010                  Case "tan"
'1020                      EvaluateExpression = _
'                              Tan(EvaluateExpression(rexpr))
'1030                      Exit Function
'1040                  Case "sqr"
'1050                      EvaluateExpression = _
'                              Sqr(EvaluateExpression(rexpr))
'1060                      Exit Function
'1070                  Case "factorial"
'1080                      EvaluateExpression = _
'                              Factorial(EvaluateExpression(rexpr))
'1090                      Exit Function
'                      ' Add other functions (including
'                      ' program-defined functions) here.
'1100              End Select
'1110          End If
'1120      End If
'
'          ' See if it's a primitive.
'1130      On Error Resume Next
'
'          'value = m_Primatives.item(expr)
'1140      status = Err.Number
'1150      On Error GoTo 0
'1160      If status = 0 Then
'              ' We found the primative.
'1170          EvaluateExpression = CDbl(value)
'1180          Exit Function
'1190      End If
'
'          ' It must be a literal like "2.71828".
'1200      On Error Resume Next
'1210      EvaluateExpression = CDbl(expr)
'1220      status = Err.Number
'1230      On Error GoTo 0
'1240      If status <> 0 Then
'1250          Err.Raise status, _
'                  "EvaluateExpression", _
'                  "Error evaluating '" & expression & _
'                  "' as a constant."
'1260      End If
'End Function

' Return the factorial of the expression.
Private Function Factorial(ByVal value As Double) As Double
      Dim result As Double

          ' Make sure the value is an integer.
10        If CLng(value) <> value Then
20            Err.Raise vbObjectError + 1001, _
                  "Factorial", _
                  "Argument must be an integer in Factorial(" & _
                  Format$(value) & ")"
30        End If

40        result = 1
50        Do While value > 1
60            result = result * value
70            value = value - 1
80        Loop
90        Factorial = result
End Function


Function IsTextFile(filepath As String) As Boolean
          Dim ext As String
10        ext = GetExtension(filepath)
20        IsTextFile = (ext = "txt" Or ext = "sss" Or ext = "cfg" Or ext = "ini" Or ext = "set" Or ext = "log")
End Function



Function ExtractFilePaths(files As String) As String()

          'working variables
          Dim cnt As Integer
          Dim tmp As String
          Dim path As String
            
          'dim an array to hold the files selected
          Dim sFileArray() As String

          
          'test the string for a Chr$(0)
          'character. If present, a multiple
          'selection was made.
10         If InStr(files, vbNullChar) Then
              
             'use Split to create an array
             'of the path and files selected
20            sFileArray() = Split(files, vbNullChar)
              
30            For cnt = LBound(sFileArray) To UBound(sFileArray)
40               If cnt = 0 Then
                   'item 0 is always the path
50                  path = sFileArray(0)
60               End If
                 
70            Next
              
                      
80            For cnt = 1 To UBound(sFileArray)
90               sFileArray(cnt) = path & "\" & sFileArray(cnt)
100           Next
           
110        Else
             'no null char, so a single selection was made
120           ReDim sFileArray(1)
130           sFileArray(0) = GetPathTo(files)
140           sFileArray(1) = files
150        End If

160       ExtractFilePaths = sFileArray
             
End Function

Function IsImageType(extension As String) As Boolean
10        IsImageType = (extension = "png" Or _
                         extension = "bmp" Or _
                         extension = "bm2" Or _
                         extension = "gif" Or _
                         extension = "jpg" Or _
                         extension = "jpeg")
End Function






Function ShapesOverlap(Shape1 As shape, Shape2 As shape) As Boolean
10        ShapesOverlap = False
20        If Shape1.Left = Shape2.Left Then
30            If Shape1.Top = Shape2.Top Then
40                If Shape1.width = Shape2.width Then
50                    ShapesOverlap = (Shape1.height = Shape2.height)
60                End If
70            End If
80        End If
End Function

Function VersionToString(version As Long) As String
          'Format:  MMmmrrr
10        VersionToString = version \ 100000 & "." & (version Mod 100000) \ 1000 & "." & version Mod 1000
End Function


Function GetProcessMemory(ByVal app_name As String) As String
          Dim Process As Object
          Dim dMemory As Double
10        For Each Process In GetObject("winmgmts:").ExecQuery("Select WorkingSetSize from Win32_Process Where Name = '" & app_name & "'")
20            dMemory = Process.WorkingSetSize
30        Next
          
40        If dMemory > 0 Then
50            GetProcessMemory = GetKbytes(dMemory)
60        Else
70            GetProcessMemory = "0 Bytes"
80        End If
End Function

Function GetKbytes(ByVal amount) As String
          ' convert to Kbytes
10        amount = amount \ 1024
20        GetKbytes = Format(amount, "###,###,###K")
End Function




Function CheckOverwrite(path As String) As Boolean
          'Checks if a file exists, and asks for confirmation to delete it
          'Returns true if the file is deleted, or does not exist
          '             i.e. if it is safe to create a new file at that path
          'Returns false if the file exists and the user did not want to overwrite it
          
10        If FileExists(path) Then
20            If MessageBox(path & " already exists. Do you want to overwrite it?", vbYesNo + vbQuestion, "Confirm overwrite") = vbYes Then
30                If FileExists(path) Then
40                    Kill path
50                    CheckOverwrite = True
60                End If
70            Else
80                CheckOverwrite = False
90            End If
100       Else
110           CheckOverwrite = True
120       End If
End Function








Function IsIntersecting(rect1 As RECT, rect2 As RECT) As Boolean
          
10        IsIntersecting = (rect1.Left <= rect2.Right And rect1.Right >= rect2.Left And rect1.Top <= rect2.Bottom And rect1.Bottom >= rect2.Top)
          
End Function

Function Intersection(ByRef destRect As RECT, ByRef rect1 As RECT, ByRef rect2 As RECT) As Boolean

          
10        If IsIntersecting(rect1, rect2) Then
20            If rect1.Left < rect2.Left Then destRect.Left = rect2.Left Else destRect.Left = rect1.Left
          
30            If rect1.Right < rect2.Right Then destRect.Right = rect1.Right Else destRect.Right = rect2.Right
          
40            If rect1.Top < rect2.Top Then destRect.Top = rect2.Top Else destRect.Top = rect1.Top
          
50            If rect1.Bottom < rect2.Bottom Then destRect.Bottom = rect1.Bottom Else destRect.Bottom = rect2.Bottom
              
60            Intersection = True
70        Else
80            Intersection = False
90        End If
          
End Function



Function GetUniqueFilename(basepath As String, filename As String) As String
    Dim str1 As String
    Dim strend As String
    Dim i As Long
    
    strend = GetExtension(filename)
    str1 = basepath & GetFileNameWithoutExtension(filename)
    
    'Check if we have _### at the end
    
    i = 1
    While IsNumeric(Right(str1, i))
        i = i + 1
    Wend
    If Left(Right(str1, i + 1), 1) = "_" Then
        str1 = Left(str1, Len(str1) - i - 1)
    End If
    
    i = 1
    
    While FileExists(str1 & "_" & i & "." & strend)
        i = i + 1
    Wend
    
    GetUniqueFilename = str1 & "_" & i & "." & strend
End Function





Function InputBoxValue(prompt As String, title As String, default, minvalue As Long, maxvalue As Long) As Long
    Dim tmpanswer As String
    
inputboxvalue_retry:
    
    tmpanswer = InputBox(prompt, title, default)
    
    If tmpanswer = "" Then
        InputBoxValue = default
    ElseIf IsNumeric(tmpanswer) Then
        Dim tmpret As Long
        tmpret = CLng(tmpanswer)
        
        If tmpret >= minvalue And tmpret <= maxvalue Then
            InputBoxValue = tmpret
        Else
            MessageBox "Value must be between " & minvalue & " and " & maxvalue
            
            GoTo inputboxvalue_retry
        End If
    Else
        GoTo inputboxvalue_retry
    End If
End Function



'Returns a value that is a multiple of 4 ( +1 )
'Used to return the correct position after a chunk, considering the 4-bytes padding
Function Next4bytes(position As Long) As Long
    If (position - 1) Mod 4 <> 0 Then
        Next4bytes = position + 4 - (position - 1) Mod 4

    Else
        Next4bytes = position
    End If
End Function



Function UnsignedAdd(ByVal L1 As Long, ByVal L2 As Long) As Long
    Dim L11 As Long, L12 As Byte, L21 As Long, L22 As Byte, L31 As Long, L32 As Byte
    L11 = L1 And &HFFFFFF
    L12 = (L1 And &H7F000000) \ &H1000000
    If L1 < 0& Then L12 = L12 Or &H80
    L21 = L2 And &HFFFFFF
    L22 = (L2 And &H7F000000) \ &H1000000
    If L2 < 0& Then L22 = L22 Or &H80
    L32 = L12 + L22
    L31 = L11 + L21
    If (L31 And &H1000000) Then L32 = L32 + 1
    UnsignedAdd = (L31 And &HFFFFFF) + (L32 And &H7F) * &H1000000
    If L32 And &H80 Then UnsignedAdd = UnsignedAdd Or &H80000000
End Function


Function UnsignedSubtract(ByVal L1 As Long, ByVal L2 As Long) As Long
    Dim L11 As Long, L12 As Byte, L21 As Long, L22 As Byte, L31 As Long, L32 As Byte
    L11 = L1 And &HFFFFFF
    L12 = (L1 And &H7F000000) \ &H1000000
    If L1 < 0& Then L12 = L12 Or &H80
    L21 = L2 And &HFFFFFF
    L22 = (L2 And &H7F000000) \ &H1000000
    If L2 < 0& Then L22 = L22 Or &H80
    L32 = L12 - L22
    L31 = L11 - L21
    If L31 < 0 Then
        L32 = L32 - 1
        L31 = L31 + &H1000000
    End If
    UnsignedSubtract = L31 + (L32 And &H7F) * &H1000000
    If L32 And &H80 Then UnsignedSubtract = UnsignedSubtract Or &H80000000
End Function
'Function IsIntersecting(rect1 As RECT, rect2 As RECT) As Boolean
'    Dim tmp As RECT
'
'    IsIntersecting = IntersectRect(tmp, rect1, rect2)
'
'End Function

'Function Intersection(rect1 As RECT, rect2 As RECT) As RECT
''    Dim tmp As RECT
'
'    Call IntersectRect(Intersection, rect1, rect2)
'
'End Function



'For the future... sending error reports
'MAPISession1.SignOn
'MAPIMessages1.SessionID = MAPISession1.SessionID
''Compose new message
'MAPIMessages1.Compose
''Address message
'MAPIMessages1.RecipDisplayName = "George Bush"
'MAPIMessages1.RecipAddress = Join Bytes!
'' Resolve recipient name
'MAPIMessages1.AddressResolveUI = True
'MAPIMessages1.ResolveName
''Create the message
'MAPIMessages1.MsgSubject = "I Love ya"
'MAPIMessages1.MsgNoteText = "Hey Bubba"
''Add attachment
'MAPIMessages1.AttachmentPathName = "c:\zxcvzxcv.zip"
''Send the message
'MAPIMessages1.Send False
'MAPISession1.SignOff

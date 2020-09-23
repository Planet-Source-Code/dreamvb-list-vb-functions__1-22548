Attribute VB_Name = "ModPas"
Public VChar(0 To 255) As String

Public CurrentXpos As Integer, CurrentYPos As Integer
Public SkipErr As Integer
Public Auto_Indent As Integer
Public CLF_Colour As Long
Public CLB_Colour As Long
Public Blink_Text As String
Public First_Colour As Long
Public Last_Colour As Long
Public Blink_Text_Enabled As Boolean
Public BlinkTextLeft As Integer
Public BlinkTextTop As Integer
Public BlinkTextFontSize As Integer
Public BlinkTextFontName As String


Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Function BlinkLable(StrString As String, ln As Integer)
Dim X, Y, Z As Integer
Dim StrCol As String

Dim FirstCol, LastCol As String

    X = FindPoint(StrString, "(")
    Y = FindPoint(StrString, ")")
    Z = FindPoint(StrString, "+")
        
    If X = 0 Then
        GetLastError 5, ln
        Exit Function
    Else
        If Y = 0 Then
            GetLastError 6, ln
            Exit Function
        Else
            If Z = 0 Then
                GetLastError 17, ln
                Exit Function
            Else
                FirstCol = "CL" & Trim(UCase(Mid(StrString, X + 1, Z - X - 1)))
                LastCol = "CL" & Trim(UCase(Mid(StrString, Z + 1, Y - Z - 1)))
            End If
        End If
    End If
    X = 0: Y = 0: Z = 0
    
    
    Select Case FirstCol
        Case "CLRED"
            First_Colour = vbRed
        Case "CLBLUE"
            First_Colour = vbBlue
        Case "CLGREEN"
            First_Colour = vbGreen
        Case "CLBLACK"
            First_Colour = vbBlack
        Case "CLYELLOW"
            First_Colour = vbYellow
        Case "CLWHITE"
            First_Colour = vbWhite
        Case "CLDESKTOP"
            First_Colour = vbDesktop
        Case "CLCYAN"
            First_Colour = vbCyan
        Case "CLMAGENTA"
            First_Colour = vbMagenta
        Case Else
            MsgBox FirstCol & " Is not sopported in this verision", vbInformation
            Unload frmWin
            Exit Function
        End Select
    '
    Select Case LastCol
        Case "CLRED"
            Last_Colour = vbRed
        Case "CLBLUE"
            Last_Colour = vbBlue
        Case "CLGREEN"
            Last_Colour = vbGreen
        Case "CLBLACK"
            Last_Colour = vbBlack
        Case "CLYELLOW"
            Last_Colour = vbYellow
        Case "CLWHITE"
            Last_Colour = vbWhite
        Case "CLDESKTOP"
            Last_Colour = vbDesktop
        Case "CLCYAN"
            Last_Colour = vbCyan
        Case "CLMAGENTA"
            Last_Colour = vbMagenta
        Case Else
            MsgBox Last_Colour & " Is not sopported in this verision", vbInformation
             Unload frmWin
            Exit Function
        End Select
        Blink_Text_Enabled = True
        
End Function
Function RestoreOld(window As PictureBox)
    CurrentXpos = 0
    CurrentYPos = 0
    Blink_Text_Enabled = False
    frmWin.lblBlink.Top = 0
    frmWin.lblBlink.Left = 10
    frmWin.lblBlink.Caption = ""
    frmWin.lblBlink.Visible = False
    window.FontSize = 9.75
    window.ForeColor = vbWhite
    window.BackColor = &H8000000F
    
End Function
Function Delay(NumSec As Integer)
Dim Max As Integer
    Max = 100
        If NumSec = 0 Then
            Exit Function
        Else
            Sleep NumSec * Max
        End If
        
End Function
Function SetTextPositionsX(lzString As String, ln As Integer)
Dim StrVal As String
    StrVal = Trim(Left(lzString, Len(lzString) - 1))
    If Len(StrVal) = 0 Then
        GetLastError 8, ln
        Exit Function
    Else
        CurrentXpos = Val(StrVal)
    End If
    StrVal = ""
    
End Function

Function SetTextPositionsY(lzString As String, ln As Integer)
Dim StrVal As String
    StrVal = Trim(Left(lzString, Len(lzString) - 1))
    If Len(StrVal) = 0 Then
        GetLastError 8, ln
        Exit Function
    Else
        CurrentYPos = Val(StrVal)
    End If
    StrVal = ""
    
End Function

Function FindPart(lzStr As String, mPart As String) As Integer
Dim TPos As Integer
    TPos = InStr(lzStr, mPart)
    If TPos Then
        FindPart = 1
    Else
        FindPart = 0
    End If
    
End Function
Function FindPoint(lzStr As String, mPart As String) As Integer
Dim Xpos As Integer
    Xpos = InStr(lzStr, mPart)
    If Xpos > 0 Then
        FindPoint = Xpos
    Else
        FindPoint = 0
    End If
    
End Function

Function ShowMsg(MsgText As String)
Dim MsgStyle As String
    MsgBox MsgText
    
    
End Function
Function IsDigit(ByVal Digit As String) As Boolean
Dim Counter As Integer
    For Counter = 1 To Len(Digit)
        ch = Asc(Mid(Digit, Counter, 1))
        If ch < 48 Then
            IsDigit = False
        ElseIf ch > 57 Then
            IsDigit = False
            Exit Function
        Else
            IsDigit = True
        End If
    Next
    Counter = 0
    
End Function
Function ScreenModes(TMode As Integer, window As PictureBox)
    Select Case TMode
        Case 10
            window.FontSize = 5
        Case 12
            window.FontSize = 16
        Case 13
            window.FontSize = 18
        Case 16
            window.FontSize = 20
            
        Case Else
        MsgBox "Mode " & TMode & " Is Not Sopprted in this verision", vbInformation
    End Select
    
End Function
Function GetLastError(ErrorNum As Integer, LineNum As Integer)
   If SkipErr = False Then
        Exit Function
   Else
    Select Case ErrorNum
        Case 1
            MsgBox "Program with out Procedure not found at Line " & LineNum
        Case 2
            MsgBox "Expected ; not found At Line " & LineNum
        Case 3
            MsgBox "Expected Mode with out value At Line " & LineNum
        Case 4
            MsgBox "Expected Text Colour without = At Line " & LineNum
        Case 5
            MsgBox "Expected ( missing in Statement at Line " & LineNum
        Case 6
            MsgBox "Expected ) missing in statement at line " & LineNum
        Case 7
            MsgBox "Expected , missing in statement at line " & LineNum
        Case 8
            MsgBox "Expected Value missing in function at line " & LineNum
        Case 9
            MsgBox "Invaild data value entered at line " & LineNum
        Case 10
            MsgBox "Procedure without End Sub not found at line " & LineNum
        Case 11
            MsgBox "Expected = not found at line " & LineNum
        Case 12
            MsgBox "Invaild Ellipse Data was entered at line " & LineNum
        Case 13
            MsgBox "Invaild Mesaage Box Style const 1 to 3 are only allowed in this verision", vbInformation
        Case 14
            MsgBox "Var without : not found at line " & LineNum
        Case 15
            MsgBox "Show Expected . found at line " & LineNum
        Case 16
            MsgBox "Const Expected without " & Chr(34) & " at line " & LineNum
        Case 17
            MsgBox "Expected + Missing in Function not found at line " & LineNum
        End Select
        End If
        
End Function
Function SetBKColour(TColour As String, window As PictureBox)
    Select Case UCase(TColour)
        Case "CLRED"
            CLB_Colour = vbRed
        Case "CLBLUE"
            CLB_Colour = vbBlue
        Case "CLGREEN"
            CLB_Colour = vbGreen
        Case "CLBLACK"
            CLB_Colour = vbBlack
        Case "CLYELLOW"
            CLB_Colour = vbYellow
        Case "CLWHITE"
            CLB_Colour = vbWhite
        Case "CLDESKTOP"
            CLB_Colour = vbDesktop
        Case "CLCYAN"
            CLB_Colour = vbCyan
        Case "CLMAGENTA"
            CLB_Colour = vbMagenta
        Case Else
            MsgBox TColour & " Is not sopported in this verision", vbInformation
        End Select
        window.BackColor = CLB_Colour
        
End Function
Function SetTextColour(TColour As String, window As PictureBox)
    Select Case UCase(TColour)
        Case "CLRED"
            CLF_Colour = vbRed
        Case "CLBLUE"
            CLF_Colour = vbBlue
        Case "CLGREEN"
            CLF_Colour = vbGreen
        Case "CLBLACK"
            CLF_Colour = vbBlack
        Case "CLYELLOW"
            CLF_Colour = vbYellow
        Case "CLWHITE"
            CLF_Colour = vbWhite
        Case "CLDESKTOP"
            CLF_Colour = vbDesktop
        Case "CLCYAN"
            CLF_Colour = vbCyan
        Case "CLMAGENTA"
            CLF_Colour = vbMagenta
        Case Else
            MsgBox TColour & " Is not sopported in this verision", vbInformation
        End Select
        window.ForeColor = CLF_Colour
        
End Function
Function GetRightVal(mText As String, ln As Integer) As Integer
Dim X, Y As Integer

    X = FindPoint(mText, "=")
    Y = FindPoint(mText, ";")
    StrVal = Trim(Mid(mText, X + 1, Y - X - 1))
    
    If IsDigit(StrVal) = False Then
        GetLastError 9, ln
        Exit Function
    Else
        GetRightVal = Val(StrVal)
    End If
    
End Function
Function GetStringRight(mText As String) As String
Dim X, Y As Integer
    X = FindPoint(mText, "=")
    Y = FindPoint(mText, ";")
    GetStringRight = Mid(mText, X + 2, Y - X - 3)
    X = 0: Y = 0
    
End Function
Function GetText(mText As String) As String
Dim lPos, Xpos As Integer
Dim StrL As String
Dim X, Y, Z As Integer

    lPos = InStr(mText, "(")
    StrL = Mid(mText, lPos + 1, InStr(lPos + 1, mText, ")") - lPos - 1)
    StrL = Replace(StrL, Chr(34), "")
    For Xpos = 0 To 255
        If InStr(StrL, VChar(Xpos)) Then
            StrL = Replace(StrL, VChar(Xpos), Chr(Xpos))
        End If
    Next
    Z = InStr(StrL, "Space")
    If Z Then
        If Mid(StrL, Z + 5, 1) = "[" Then
            X = InStr(StrL, "[")
            Y = InStr(StrL, "]")
            If Y Then
                Z = Trim(Mid(StrL, X + 1, Y - X - 1))
                StrL = Replace(StrL, "Space[" & Z & "]", Space(Z))
            End If
        End If
    End If
    X = 0: Y = 0: Z = 0
    
    Z = InStr(StrL, "asc")
    If Z Then
        If Mid(StrL, Z + 3, 1) = "[" Then
            X = InStr(StrL, "[")
            Y = InStr(StrL, "]")
            If Y Then
                p = Mid(StrL, X + 1, Y - X - 1)
                StrL = Replace(StrL, "asc[" & p & "]", Asc(p))
            End If
        End If
    End If
    
    GetText = Replace(StrL, "&", "")

End Function
Function PutToScreen(lzStr As String, window As PictureBox)
    If CurrentXpos = 0 Or CurrentYPos = 0 Then
        window.Print lzStr & vbclrf
        Exit Function
    Else
        window.CurrentX = CurrentXpos
        window.CurrentY = CurrentYPos
        window.Print lzStr & vbclrf
    End If
    
End Function
Function RemoveChar(StrString As String, SChar As String) As String
    RemoveChar = Replace(StrString, Chr(9), Chr(32))
    
End Function

Function Plot(lzStr As String, ln As Integer, window As PictureBox)
Dim StrVal1, StrVal2 As String
Dim StrVal As String
Dim k As String
Dim Val1, Val2 As Integer

    k = lzStr
    
    If FindPart(k, "(") = 0 Then
        GetLastError 5, ln
        Exit Function
    ElseIf FindPart(k, ")") = 0 Then
        GetLastError 6, ln
        Exit Function
    ElseIf FindPart(k, ";") = 0 Then
        GetLastError 2, ln
        Exit Function
    Else
        StrVal = GetText(k)
        If FindPart(k, ",") = 0 Then
            GetLastError 7, ln
            Exit Function
        Else
            StrVal = Trim(GetText(k))
            StrVal1 = Trim(Mid(StrVal, FindPoint(StrVal, ",") + 1, Len(StrVal)))
            StrVal2 = Trim(Mid(StrVal, 1, FindPoint(StrVal, ",") - 1))
            
            If IsDigit(StrVal1) = False Then
                GetLastError 9, ln
                Exit Function
            ElseIf IsDigit(StrVal2) = False Then
                Exit Function
                GetLastError 9, ln
            Else
                Val1 = Val(StrVal2)
                Val2 = Val(StrVal1)
                window.PSet (Val1, Val2), CLF_Colour
            End If
    End If
    End If
    k = ""
    StrVal1 = ""
    StrVal2 = ""
    Val1 = 0
    Val2 = 0
    
End Function
Function DrawEllipse(lzStr As String, ln As Integer, window As PictureBox)
Dim EllipseData As Collection
Dim lPos As Integer
Dim StrVal As String, G As String

    Set EllipseData = New Collection
    StrVal = GetText(lzStr)
    If Len(StrVal) = 0 Then
        GetLastError 8, ln
        Exit Function
    Else
        StrVal = Left(StrVal, Len(StrVal)) & ","
        For lPos = 1 To Len(StrVal)
            ch = Mid(StrVal, lPos, 1)
            G = G & ch
            If InStr(G, ",") Then
                mCount = mCount + 1
                
                G = Left(G, Len(G) - 1)
                If IsDigit(G) = False Then
                    GetLastError 9, ln
                    Exit Function
                Else
                   If IsDigit(G) = False Then
                        GetLastError 9, ln
                        Exit Function
                   Else
                        EllipseData.Add G
                G = ""
                End If
                End If
            End If
        Next
    End If
    On Error Resume Next
    If EllipseData.Count = 3 Then
        window.Circle (EllipseData(1), EllipseData(2)), EllipseData(3), CLF_Colour
    Else
        GetLastError 12, ln
        Exit Function
    End If
        StrVal = ""
        G = ""
        mCount = 0
        If Err Then
            GetLastError 9, ln
        End If
        
End Function
Function DrawLine(lzStr As String, ln As Integer, window As PictureBox)
Dim StrVal1, StrVal2 As String
Dim StrVal As String
Dim k As String
Dim Val1, Val2 As Integer

    k = lzStr
    
    If FindPart(k, "(") = 0 Then
        GetLastError 5, ln
        Exit Function
    ElseIf FindPart(k, ")") = 0 Then
        GetLastError 6, ln
        Exit Function
    ElseIf FindPart(k, ";") = 0 Then
        GetLastError 2, ln
        Exit Function
    Else
        StrVal = GetText(k)
        If FindPart(k, ",") = 0 Then
            GetLastError 7, ln
            Exit Function
        Else
            StrVal = Trim(GetText(k))
            StrVal1 = Trim(Mid(StrVal, FindPoint(StrVal, ",") + 1, Len(StrVal)))
            StrVal2 = Trim(Mid(StrVal, 1, FindPoint(StrVal, ",") - 1))
            
            If IsDigit(StrVal1) = False Or IsDigit(StrVal2) = False Then
                GetLastError 9, ln
                Exit Function
            Else
                Val1 = Val(StrVal2)
                Val2 = Val(StrVal1)
                LineTo window.hDC, Val1, Val2
            End If
    End If
    End If
    k = ""
    StrVal1 = ""
    StrVal2 = ""
    Val1 = 0
    Val2 = 0
    
End Function

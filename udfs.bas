Attribute VB_Name = "UDF_Module"
'Function for reversing the content of a cell based on text and delimiter
'Useful for cases like reversing name order eq. John Doe becomes Doe John
Public Function REVERSE(text As String, delimiter As String) As String
Attribute REVERSE.VB_Description = "Function for reversing the content of a cell based on text and delimiter"
Attribute REVERSE.VB_ProcData.VB_Invoke_Func = " \n7"
    
    Dim reversed As String
    Dim textArray() As String
    textArray() = Split(text, delimiter)
    
    'Loop through the exploded array backwards
    For i = UBound(textArray) To LBound(textArray) Step -1
        If i = UBound(textArray) Then
            reversed = textArray(i)
        Else
            reversed = reversed & delimiter & textArray(i)
        End If
    Next i
    
    REVERSE = reversed
    
End Function

'Function for checking the font color of the cell 
Public Function ISFONTCOLOR(target As Range, red As Integer, green As Integer, Blue As Integer)
Attribute ISFONTCOLOR.VB_Description = "Function for checking the font color of the cell. Argument order: targetCell, red, green, blue"
Attribute ISFONTCOLOR.VB_ProcData.VB_Invoke_Func = " \n7"
    If target.Font.color = RGB(red, green, Blue) Then
        ISFONTCOLOR = True
    Else
        ISFONTCOLOR = False
    End If
End Function


'Exploding function for excel. Returns the specified item of the exploded string. 0 based count.
Public Function EXPLODE(text As String, delimiter As String, itemNumber As Integer) As String
    
    Dim output As String
    Dim textArray() As String
    textArray() = Split(text, delimiter)
    
    'Return error if the itemNumber is larger than count of the items in array
    If itemNumber > (LBound(textArray) + UBound(textArray)) Then
        EXPLODE = CVErr(xlErrNA)
    'Also if user tries to enter negative values
    ElseIf itemNumber < 0 Then
        EXPLODE = CVErr(xlErrNA)
    Else
        EXPLODE = textArray(itemNumber)
    End If
    
End Function

'Useful helper function to show contents of the cell as is when concatenating them
'For example concatenating days "1.1.2014 - 1.2.2014" is possible
Public Function ASDISPLAYED(ByVal cell As Range) As String
  ASDISPLAYED = cell.text
End Function

'Helper function for checking validity of email addresses
Public Function ISVALIDEMAIL(email As String) As Boolean
    Dim result As String
    'Email pattern for regex
    Dim pattern As String: pattern = "[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"
    Dim regEx As New RegExp
    
    'Some settings
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = pattern
    End With
    
    If regEx.test(email) Then
        ISVALIDEMAIL = True
    Else
        ISVALIDEMAIL = False
    End If
        
End Function

'Returns the number of items in the list based on delimiter
Public Function LISTLENGTH(text As String, delimiter As String) As Integer

    Dim result As Integer
    Dim textArray() As String
    textArray() = Split(text, delimiter)
    result = LBound(textArray) + UBound(textArray) + 1

    LISTLENGTH = result

End Function

'Converts given rgb value to HEX color
Public Function RGBTOHEX(red As Integer, green As Integer, blue As Integer, includeHash As Boolean) As String
    Dim hexRed As String
    Dim hexGreen As String
    Dim hexBlue As String

    hexRed = Hex(red)
    hexGreen = Hex(green)
    hexBlue = Hex(blue)
    
    If Len(hexRed) = 1 Then
        hexRed = "0" & hexRed
    End If
        
    If Len(hexGreen) = 1 Then
        hexGreen = "0" & hexGreen
    End If
    
    If Len(hexBlue) = 1 Then
        hexBlue = "0" & hexBlue
    End If
    
    If includeHash = False Then
        RGBTOHEX = hexRed & hexGreen & hexBlue
    Else
        RGBTOHEX = "#" & hexRed & hexGreen & hexBlue
    End If
End Function

'Converts HEX color to RGB also accepts html colors with #
Public Function HEXTORGB(hx As String) As String
    If InStr(hx, "#") > 0 Then
        red = Val("&H" & Mid(hx, 2, 2))
        green = Val("&H" & Mid(hx, 4, 2))
        blue = Val("&H" & Mid(hx, 6, 2))
    Else
        red = Val("&H" & Mid(hx, 1, 2))
        green = Val("&H" & Mid(hx, 3, 2))
        blue = Val("&H" & Mid(hx, 5, 2))
    End If
    
    HEXTORGB = red & "," & green & "," & blue
End Function



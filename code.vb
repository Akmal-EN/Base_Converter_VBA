

Function Fullfil0s(init_val As String, mylen As Byte)
    init_val = init_val & ""
    Do While Len(init_val) < mylen
                    init_val = "0" & init_val
    Loop
    Fullfil0s = init_val
End Function

Function Trim0s(init_val As Variant)
    Dim temp_val As String
    temp_val = init_val & ""
    If (temp_val = "") Then
        MsgBox ("Empty Input error!")
        Exit Function
    End If
    Do While (Right(temp_val, 1) = "0" And (Len(temp_val) + 1 - InStrRev(temp_val, ".")) > 1 And (InStr(temp_val, ".") > 0))
        temp_val = Left(temp_val, Len(temp_val) - 1)
    Loop
    If (Left(temp_val, 1) = "-") Then
        Do While (Right(Left(temp_val, 2), 1) = "0" And Left(temp_val, 2) <> "0.")
            temp_val = Right(temp_val, Len(temp_val) - 1)
        Loop
    Else
        Do While (Left(temp_val, 1) = "0" And Left(temp_val, 2) <> "0.")
            temp_val = Right(temp_val, Len(temp_val) - 1)
        Loop
    End If
    If ((Len(temp_val) + 1 - InStrRev(temp_val, ".")) = 1) Then
        temp_val = Left(temp_val, Len(temp_val) - 1)
    End If
    Trim0s = temp_val
End Function


'Base Correspondence of every digit (-1: Unknown letters,0: out of reach, 1: comply)
Function CheckDigits(init_val As String, from_base As Byte)
    Dim ctr As Integer
    ctr = 1
    Do While ctr <= Len(init_val)
        digit = Asc(Left(Right(init_val, ctr), 1))
        If (digit > 64 And digit < 91) Then
            digit = digit - 55
        ElseIf (digit > 47 And digit < 91) Then
            digit = digit - 48
        ElseIf (ctr = Len(init_val) And digit = 45) Then
            digit = 65 - 55
        Else
            CheckDigits = -1
            Exit Function
        End If
        
        'Base correspondence check
        If (digit >= from_base) Then
            CheckDigits = 0
            Exit Function
        'BASE INCORRESPONDENCE ERROR
        End If
        CheckDigits = 1
        ctr = ctr + 1
    Loop
End Function



'Converts a given number in the from_base(2-36) to to_base(2-36)
Function Convert10(init_val As Variant, from_base As Byte, to_base As Byte)
    'Check if the input is empty
    
    'First we convert the input to decimal
    Dim decim As Long
    decim = 0
    

    Dim digit As Integer
    init_val = init_val & ""
    
    Dim neg As Byte
    'If negative - note it
    If (Left(init_val, 1) = "-") Then
        neg = 1
        init_val = Mid(init_val, 2)
    End If

    'ctr - counter
    Dim ctr As Integer
    ctr = 1
    'from right to left, check every letter - convert it to digit
    Do While ctr <= Len(init_val)
        digit = Asc(Left(Right(init_val, ctr), 1))
        If (digit > 64 And digit < 91) Then
            digit = digit - 55
        ElseIf (digit > 47 And digit < 91) Then
            digit = digit - 48
        Else
            MsgBox ("Unknown digit/letter in input")
            Exit Function
        End If
        
        'Check if the digit corresponds to the base
        If (digit >= from_base) Then
            MsgBox ("DIGIT/BASE INCORRESPONDENCE")
            Exit Function
        End If
        
        
        decim = decim + digit * (from_base ^ (ctr - 1))
        ctr = ctr + 1
    Loop
    
    'From decimal convert to  to_base
    
    'If the base we want is already 10 - just give the results
    If (to_base = 10) Then
        If (neg = 1) Then
            Convert10 = "-" & decim
        Else
            Convert10 = decim
        End If

    'Otherwise, divide the decimal result by the base and subtract remainder until it's 0
    Else
        Dim result As Variant
        result = ""
        ctr = 0

        Do While decim <> 0
            digit = decim Mod to_base
            'REMAINDER IS A DIGIT
            If (digit < 9) Then
                    result = digit & result
            Else
                result = Chr(digit + 55) & result
            End If
            decim = decim \ to_base
            ctr = ctr + 1
        Loop
        'Bring back the minus if there was
        If (neg = 1) Then
            Convert10 = "-" & result
        Else
            Convert10 = result
        End If
        
    End If
End Function


Function Complement2(init_val As Variant, Optional from2c As Byte)

    Dim negative As Byte
    negative = 0
    
    Dim index As Byte
    index = 1
    Dim inverted As String
    inverted = ""
    Dim tempstr As String

    'If we need from 2c to 10
    If (from2c = 1) Then
        'PROCESS Conversion FROM 2c to 10
        If (Left(init_val, 1) = "1") Then
        
            init_val = Right(init_val, Len(init_val) - 1)
            If (Convert10(init_val, 2, 10) > 127) Then
                MsgBox ("Range of values converting to 2's complement: -127(10-base) to 127(10-base) ")
                Exit Function
            End If
            init_val = "1" + init_val
            negative = 1
            
        End If
        
        If (negative = 1) Then
            'invert every digit starting from right to left
            Do While index <= Len(init_val)
                tempstr = Left(Right(init_val, index), 1)
                If (tempstr = "0") Then
                    inverted = "1" & inverted
                ElseIf (tempstr = "1") Then
                    inverted = "0" & inverted
                Else
                    MsgBox ("DIGIT/BASE INCORRESPONDENCE")
                    Exit Function
                End If
                index = index + 1
            Loop
            Complement2 = ((-1) * CInt(Convert10(inverted, 2, 10)) - 1)
        Else
            'if the value is not negative - then it's enough leaving it uninverted
            inverted = init_val
            'if it was not negative - then just convert 2base to 10base
            Complement2 = (CInt(Convert10(inverted, 2, 10)))
        End If
        
    Else
    'Now if we want to convert base2 to 2c
        'Note if negative
        If (Left(init_val & "", 2) = "-0" And Len(init_val) <= 9) Then
                        
        ElseIf (Left(init_val & "", 2) = "-1" And Len(init_val) <= 8) Then
            
        ElseIf (Left(init_val & "", 1) = "0" And Len(init_val) <= 8) Then
        
        Else
            MsgBox ("Range of values converting to 2's complement: -127(10-base) to 127(10-base) ")
            Exit Function
        End If
        If (Left(init_val, 1) = "-") Then
        
            init_val = Right(init_val, Len(init_val) - 1)
            If (Len(init_val) > 8) Then
                MsgBox ("Range of values converting to 2's complement: -127(10-base) to 127(10-base) ")
                Exit Function
            End If
            init_val = Fullfil0s(init_val & "", Len(init_val) - 1)
            negative = 1
            
        Else
            init_val = Fullfil0s(Trim0s(init_val), 8)
        End If
        'Invert the base2
        
           If (negative = 1) Then

            Do While index <= Len(init_val)
                tempstr = Left(Right(init_val, index), 1)
                If (tempstr = "0") Then
                    inverted = "1" & inverted
                ElseIf (tempstr = "1") Then
                    inverted = "0" & inverted
                Else
                    MsgBox ("DIGIT/BASE INCORRESPONDENCE")
                    Exit Function
                End If
                index = index + 1
            Loop
            
            'Add 1 to the inverted
            Do While (Left(Right(inverted, index), 1)) = 1
                inverted = Left(inverted, Len(inverted) - index) & "0" & Right(inverted, index - 1)
                index = index + 1
            Loop
            If (index = 1) Then
                inverted = Left(inverted, Len(inverted) - 1) & "1"
            Else
                inverted = Left(inverted, Len(inverted) - index + 1) & "1" & Right(inverted, index - 1)
            End If
            'FulFil 0s on the left
            inverted = Fullfil0s(inverted, 7)
            
            'Negate if necessary
            If (negative = 1) Then
                inverted = "1" & inverted
                If (Right(inverted, 1) = 0) Then
                    inverted = Left(inverted, Len(inverted) - 1) + "1"
                Else
                    inverted = Left(inverted, Len(inverted) - 1) + "0"
                End If
            Else
                inverted = "0" & inverted
                If (Right(inverted, 1) = 0) Then
                    inverted = Left(inverted, Len(inverted) - 1) + "1"
                Else
                    inverted = Left(inverted, Len(inverted) - 1) + "0"
                End If
            End If
        
            Complement2 = inverted
        Else
            Complement2 = init_val
        End If
        
    End If
End Function

Function Fibbo(init_val As Variant, Optional to10 As Byte)

    Dim result As String
    result = ""
    ReDim a(3) As Long
    a(1) = 1
    a(2) = 2
    a(3) = 3
    Dim index As Integer
    index = 1
    negative = 0
    If (Left(init_val, 1) = "-") Then
        negative = 1
        init_val = Right(init_val, Len(init_val) - 1)
    End If
    
    If (to10 = 1) Then
        Dim decim As Long
        decim = 0
        Do While index <= Len(init_val)
            result = Left(Right(init_val, index), 1)
            If (result = "0") Then
            ElseIf (result = "1") Then
                decim = decim + a((index + 2) Mod 3 + 1)
            Else
                MsgBox ("unknown digit")
                Exit Function
            End If
            If (index Mod 3 = 1) Then
                a(1) = a(2) + a(3)
            ElseIf (index Mod 3 = 2) Then
                a(2) = a(1) + a(3)
            Else
                a(3) = a(2) + a(1)
            End If
            index = index + 1
        Loop
        If (negative = 1) Then
            Fibbo = "-" & decim
        Else
            Fibbo = decim
        End If
    
    Else
        Dim temp As Long
        If (VarType(init_val) = 2) Then
        temp = init_val
        Else
        temp = CInt(init_val)
        End If
        Do While (temp > 0)
            a(1) = 1
            a(2) = 2
            a(3) = 3
            index = 1
            'Find the biggest Fibo number under the temp every time
            Do While (a((index Mod 3) + 1) <= temp)
                If (index Mod 3 = 1) Then
                    a(1) = a(2) + a(3)
                ElseIf (index Mod 3 = 2) Then
                    a(2) = a(1) + a(3)
                Else
                    a(3) = a(2) + a(1)
                End If
                index = index + 1
            Loop
            'Until you have all the numbers that sum up to the decimal value
            temp = temp - a(((index + 2) Mod 3) + 1)
            'Get all the indexes into 1
            If (result = "") Then
                result = "1"
            Else
                result = Left(result, Len(result) - index) & "1"
            End If
            'and fill the remaining with 0s
            Do While index <> 1
                result = result & "0"
                index = index - 1
            Loop
        Loop
        If (negative = 1) Then
            Fibbo = "-" & result
        Else
            Fibbo = result
        End If
    
    End If
End Function


Private Sub fromlabel_Click()
MsgBox ("You are supposed to ENTER the BASE you want to CONVERT FROM")
End Sub

Private Sub Label1_Click()
MsgBox ("You are supposed GET the RESULTS at the BOX BELOW")
End Sub

Private Sub Label2_Click()
    MsgBox ("You are supposed to type in the bottom box WHAT YOU WANT TO CONVERT !")
End Sub

Private Sub Label3_Click()
    MsgBox ("Click-toggle the button to NORMALIZE THE RESULT")
End Sub

Private Sub save_btn_Click()
    ActiveCell.Value = InputBox.Value
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = from_box.Value
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = to_box.Value
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = result_box.Value
    ActiveCell.Offset(1, -3).Select
End Sub

Private Sub submit_btn_Click()
    'Check Input
    If (InputBox.Value = "") Then
        MsgBox ("Empty Input!")
    End If
    'Clear result_box
    result_box.Value = ""

    Dim from_base As Byte
    Dim to_base As Byte
    
    If (from_box.Value = "2s complement (8-bit)" Or from_box.Value = "Fibbonaci base") Then
         from_base = 2
    Else
        from_base = CInt(from_box.Value)
    End If

    If (to_box.Value = "2s complement (8-bit)" Or to_box.Value = "Fibbonaci base") Then
        to_base = 2
   Else
       to_base = CInt(to_box.Value)
   End If

   'if the same base - just check the digits
    If (from_box.Value = to_box.Value) Then
        'Check all the digits if they're corresponding
        
        If (CheckDigits(InputBox.Value, from_base) = 1) Then
        result_box.Value = InputBox.Value
        Else
        result_box.Value = "digit/base incorrespondance !"
        End If
        
    Else
        
        Dim init_val As String
        init_val = InputBox.Value
        
        'Fibbonaci base~2s complement (8-bit)
        If (from_box.Value = "2s complement (8-bit)" And to_box.Value = "Fibbonaci base") Then
            result_box.Value = Fibbo(Complement2(init_val, 1))
        
        
        ElseIf (from_box.Value = "Fibbonaci base" And to_box.Value = "2s complement (8-bit)") Then
            result_box.Value = Complement2(Convert10(Fibbo(init_val, 1), 10, 2))
        
        
        '2s complement ~ 10(and other numbers)
        ElseIf (from_box.Value = "2s complement (8-bit)") Then
                result_box.Value = Convert10(Complement2(init_val, 1), 10, to_base)
                
        ElseIf (to_box.Value = "2s complement (8-bit)") Then
            
            If (Left(init_val & "", 1) = "-") Then
                result_box.Value = Complement2(Convert10(init_val, from_base, 2))
            Else
                result_box.Value = Complement2(Fullfil0s(Convert10(init_val, from_base, 2), 8))
            End If
            
        'Fibbonaci ~ 10(other numbers)
        ElseIf (from_box.Value = "Fibbonaci base") Then
            result_box.Value = Convert10(Fibbo(init_val, 1), 10, to_base)
        ElseIf (to_box.Value = "Fibbonaci base") Then
            result_box.Value = Fibbo(Convert10(init_val, from_base, 10))
        Else
                result_box.Value = Convert10(init_val, from_base, to_base)
        End If
        
    End If
    If (result_box.Value <> "") Then
    If (Norm.Value = True) Then
        If (CStr(to_base) = "2s complement (8-bit)") Then
            to_base = "2"
        ElseIf (CStr(to_base) = "Fibbonaci base") Then
            to_base = "Fibo_base"
        End If
            result_box.Value = Normalize(result_box.Value, to_base)
    End If
    End If
    
End Sub


Private Sub to_label_Click()
MsgBox ("You are supposed to ENTER the BASE you want to CONVERT TO")
End Sub

Global globalcounter As Integer
Private Sub UserForm_Initialize()
   
   InputBox = ""
   TextBox1 = ""
   globalcounter = 0
   from_box.List = Array("Fibbonaci base", 2, "2s complement (8-bit)", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36")
   to_box.List = Array("Fibbonaci base", 2, "2s complement (8-bit)", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36")
   With InputBox
        .SetFocus
    End With
    ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count)
    ActiveCell.Value = "Input"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = "From"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = "To"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = "Output"
    ActiveCell.Offset(1, -3).Select
End Sub



Function Normalize(init_val As Variant, base As Byte)
    'string for final result
    Dim temp_str As String
    temp_str = Trim0s(init_val & "")
    'Position of the dot
    Dim pos As Integer
    pos = InStr(temp_str, ".")
    'Check if negative number
    Dim negative As Integer
    negative = InStr(temp_str, "-")
    If ((pos > 2 And negative = 0)) Then
        temp_str = Left(temp_str, pos - 1) & Mid(temp_str, pos + 1)
        'Turn Pos into Exponent
        pos = pos - 2
        temp_str = Left(temp_str, 1) & "." & Mid(temp_str, 2)
    ElseIf (pos > 3 And negative = 1) Then
        pos = pos - 3
        temp_str = Left(temp_str, 2) & "." & Mid(temp_str, 3)
        'Pos = exponent now
    ElseIf (pos = 0) Then
        If ((Len(temp_str) > 1 And negative = 0)) Then
            temp_str = Left(temp_str, 1) & "." & Mid(temp_str, 2)
        ElseIf ((Len(temp_str) > 2 And negative = 1)) Then
            temp_str = Left(temp_str, 2) & "." & Mid(temp_str, 3)
        End If
        pos = Len(Mid(temp_str, InStr(temp_str, ".") + 1))
    End If
        'pos = exponent now
    Normalize = temp_str & " * " & base & " ^ " & pos
End Function





Attribute VB_Name = "modCrypto"
Option Explicit
Public Const sDefaultWHEEL1 As String = "ABCDEFGHIJKLMNOPQRSTVUWXYZ1234567890,qwertyuiopasdfghjk*lzxcvbnm<>"
Public Const sDefaultWHEEL2 As String = "IWEHJKTLZVOPFG1234567890qwerBNMQRYUAS,DXCfghjklzxcvbnmt*yuiopasd<>"


' Encrypts the string:
Function Encrypt_PRO(sINPUT As String, sPASSWORD As String) As String

    Dim sWHEEL1 As String
    Dim sWHEEL2 As String

    sWHEEL1 = sDefaultWHEEL1
    sWHEEL2 = sDefaultWHEEL2


    ' We use password to scramble the wheels:
    ScrambleWheels sWHEEL1, sWHEEL2, sPASSWORD
    
    

    Dim k As Long ' keeps index of the character on the wheel.
    Dim C As String ' to keep single character.
    
    Dim I As Long ' for current character index of source string.
    
    Dim sRESULT As String ' the result.
    sRESULT = ""
    
    For I = 1 To Len(sINPUT)
    
            ' Get character(i)
            C = mid(sINPUT, I, 1)
    
            ' Find character(i) on the first wheel:
            k = InStr(1, sWHEEL1, C, vbBinaryCompare)
            
            If k > 0 Then
                ' Get the character with that index from the second
                ' wheel, and add it to result:
                sRESULT = sRESULT & mid(sWHEEL2, k, 1)
            Else
                ' not found on the wheel, leave as it is:
                sRESULT = sRESULT & C
            End If
    
            ' Rotate first wheel to the left:
            sWHEEL1 = LeftShift(sWHEEL1)
            
            ' Rotate second wheel to the right:
            sWHEEL2 = RightShift(sWHEEL2)
    
    Next I
    
    Encrypt_PRO = sRESULT
    
End Function

' Decrypts the string.
' you may note that the only difference is
' the sWHEEL1 and sWHEEL2 exchange, instead of
' looking for character in sWHEEL1 we look it
' in sWHEEL2:
Function Decrypt_PRO(sINPUT As String, sPASSWORD As String) As String


    Dim sWHEEL1 As String
    Dim sWHEEL2 As String

    sWHEEL1 = sDefaultWHEEL1
    sWHEEL2 = sDefaultWHEEL2

    ' We use password to "de"-scramble the wheels:
    ScrambleWheels sWHEEL1, sWHEEL2, sPASSWORD

    Dim k As Long ' keeps index of the character on the wheel.
    
    Dim I As Long ' for current character index of source string.
    Dim C As String ' to keep single character.
    
    Dim sRESULT As String ' the result.
    sRESULT = ""
    
    For I = 1 To Len(sINPUT)
    
            ' Get character(i)
            C = mid(sINPUT, I, 1)
    
            ' Find character(i) on the second wheel:
            k = InStr(1, sWHEEL2, C, vbBinaryCompare)
            
            If k > 0 Then
                ' Get the character with that index from the first
                ' wheel, and add it to result:
                sRESULT = sRESULT & mid(sWHEEL1, k, 1)
            Else
                ' not found on the wheel, leave as it is:
                sRESULT = sRESULT & C
            End If
    
            ' Rotate first wheel to the left:
            sWHEEL1 = LeftShift(sWHEEL1)
            
            ' Rotate second wheel to the right:
            sWHEEL2 = RightShift(sWHEEL2)
    
    Next I
    
    Decrypt_PRO = sRESULT
    
End Function

' Rotates the wheel (string).
' the first character goes to the end, all
' other characters go one step to the left side.
' For example:
'     "ABCD"
' will be
'     "BCDA"
' after rotation.
Function LeftShift(s As String) As String
    ' tricky way :)
    If Len(s) > 0 Then LeftShift = mid(s, 2, Len(s) - 1) & mid(s, 1, 1)
End Function


' Rotates the wheel (string).
' the last character goes to the beginning, all
' other characters go one step to the right side.
' For example:
'     "ABCD"
' will be
'     "DABC"
' after rotation.
Function RightShift(s As String) As String
    ' tricky way :)
    If Len(s) > 0 Then RightShift = mid(s, Len(s), 1) & mid(s, 1, Len(s) - 1)
End Function


' This sub scrambles the wheels.
' Wheels should be set to the same position
' for both encryption and decryption !
' (and this can be achieved by using the same password :)
' Bigger password = better scramble!
Sub ScrambleWheels(ByRef sW1 As String, ByRef sW2 As String, sPASSWORD As String)

Dim I As Long
Dim k As Long

For I = 1 To Len(sPASSWORD)
    
    For k = 1 To Asc(mid(sPASSWORD, I, 1)) * I
        sW1 = LeftShift(sW1)
        sW2 = RightShift(sW2)
    Next k

Next I

' Who said there are no pointers in VB?

End Sub


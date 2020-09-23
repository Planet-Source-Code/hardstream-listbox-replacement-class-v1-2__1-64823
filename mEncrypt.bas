Attribute VB_Name = "mEncrypt"
Option Explicit

' __________________________________________________________________________
'/Encryption version 1.0                                                    \
'\==========================================================================/
'|Author: HardStream Group                                                  |
'|Date  : 04-01-2006                                                        |
'|Name  : mEncrypt                                                          |
'|Verson: 1.0                                                               |
'/==========================================================================\
'|Version 1.0:                                                              |
'|This is a new encryption module, completely written by HardStream.        |
'|At this moment, the encryption/decryption doesn't work properly (try to   |
'|encrypt/decrypt in a textbox, you'll see it doesn't really work).         |
'|If you know what's wrong, please correct the code and send it to us.      |
'|Our e-mail addresses are:                                                 |
'|1) hardstream@hotmail.com                                                 |
'|2) info@hardstream.nl                                                     |
'|                                                                          |
'|You'll of course get credit for your work.                                |
'|Thank you for your interest.                                              |
'\__________________________________________________________________________/

'Private constants
Private Const AlphaNumeric As String = "A=101|B=103|C=105|D=107|E=109|F=111|G=113|H=115|I=117|J=119|K=121|L=123|M=125|N=202|O=204|P=206|Q=208|R=210|S=212|T=214|U=216|V=218|W=220|X=222|Y=224|Z=226|a=102|b=104|c=106|d=108|e=110|f=112|g=114|h=116|i=118|j=120|k=122|l=124|m=126|n=201|o=203|p=205|q=207|r=209|s=211|t=213|u=215|v=217|w=219|x=221|y=223|z=225|0=301|1=303|2=305|3=307|4=309|5=311|6=313|7=315|8=317|9=319" 'Alphanumeric characters
Private Const SpecialChars As String = "~402 @404 #406 $408 %410 ^412 &414 *416 (418 )420  422 -424 +426 =428 {430 }432 [434 ]436 |438 \440 :442 ;444 ""446 '448 <450 >452 ,454 .456 ?458 /460  462"

'Private variables
Private Result As String

'Get the letter from ID (the actual decryption part)
'You've got to rewrite this if you're going to change the encyption technique
Private Function Letter(ID As String) As String
Dim Base As String, Tmp As String

If ID = 501 Then 'vbCrLf
    Letter = vbCrLf
        Else
    If ID < 400 Then Base = AlphaNumeric Else Base = SpecialChars
    
    Tmp = Split(Base, ID)(0)
    If ID < 400 Then 'Alphanumeric
        Tmp = Right(Tmp, 2)
        Tmp = Left(Tmp, 1)
            Else 'Special
        Tmp = Right(Tmp, 1)
    End If
End If

Letter = Tmp
Tmp = ""
Base = ""
End Function

'Get the ID from letter (the actual encryption part)
'You've got to rewrite this if you're going to change the encyption technique
Private Function ID(Letter As String) As String
Dim Base As String, Tmp As String

If Letter = vbCrLf Then 'vbCrLf
    ID = 501
        Else
    Base = AlphaNumeric & SpecialChars
    
    Tmp = Split(Base, Letter)(0)
    Tmp = Left(Base, Len(Tmp) + 5)
    Tmp = Right(Tmp, 4)
    Tmp = Right(Tmp, 3) 'Make sure there aren't any weird signs (like '=' when identifying an alphanumeric character)
End If

If IsNumeric(Tmp) Then ID = Tmp Else ID = "" 'Check if the ID is really numeric. This is the closest you can get to making sure the ID is correct
Tmp = ""
Base = ""
End Function

'Decrypt a string
Function Decrypt(Expression As String) As String
Dim i As Long
Dim Tmp As String

For i = 1 To Len(Expression) Step 3
    Tmp = Mid(Expression, i, 3)
    Tmp = Letter(Tmp)
    Result = Result & Tmp
Next i

Tmp = ""
Decrypt = Result
Result = ""
End Function

'Encrypt a string
Function Encrypt(Expression As String)
Dim i As Long
Dim Tmp As String

For i = 1 To Len(Expression)
    Tmp = Mid(Expression, i, 1)
    Tmp = ID(Tmp)
    Result = Result & Tmp
Next i

Tmp = ""
Encrypt = Result
Result = ""
End Function

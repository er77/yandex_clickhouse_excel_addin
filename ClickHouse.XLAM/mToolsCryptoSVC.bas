Attribute VB_Name = "mToolsCryptoSVC"
Option Explicit
Option Compare Text

'http://www.codetoad.com/visual_basic_better_xor.asp
Public Function f_XORDecryption(DataIn As String, Optional ByVal vPassword As String = vCurrPasswordLine) As String
  On Error GoTo ErrorHandler
    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim intXOrValue1 As Integer
    Dim intXOrValue2 As Integer
    
    DoEvents
  
    For lonDataPtr = 1 To (Len(DataIn) / 2)
        'The first value to be XOr-ed comes from the data to be encrypted
        intXOrValue1 = Val("&H" & (Mid$(DataIn, (2 * lonDataPtr) - 1, 2)))
        'The second value comes from the code key
        intXOrValue2 = Asc(Mid$(vPassword, ((lonDataPtr Mod Len(vPassword)) + 1), 1))
        
        strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
    Next lonDataPtr
   f_XORDecryption = strDataOut
l_exit:
    Exit Function
ErrorHandler:
Call p_ErrorHandler(0, "XORDecryption")
   
End Function


Public Function f_XOREncryption(DataIn As String, Optional ByVal vPassword As String = vCurrPasswordLine) As String
     On Error GoTo ErrorHandler
    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim temp As Integer
    Dim tempstring As String
    Dim intXOrValue1 As Integer
    Dim intXOrValue2 As Integer
    
    DoEvents
    For lonDataPtr = 1 To Len(DataIn)
        'The first value to be XOr-ed comes from the data to be encrypted
        intXOrValue1 = Asc(Mid$(DataIn, lonDataPtr, 1))
        'The second value comes from the code key
         
        intXOrValue2 = Asc(Mid$(vPassword, ((lonDataPtr Mod Len(vPassword)) + 1), 1))
        
        temp = (intXOrValue1 Xor intXOrValue2)
        tempstring = Hex(temp)
        If Len(tempstring) = 1 Then tempstring = "0" & tempstring
        
        strDataOut = strDataOut + tempstring
    Next lonDataPtr
   f_XOREncryption = strDataOut
l_exit:
    Exit Function
ErrorHandler:
Call p_ErrorHandler(0, "XOREncryption")
     
End Function

Function CRC16HASH(txt) As String
Dim X As Long
Dim mask, i, J, nC, crc As Integer
Dim C As String

crc = &HFFFF
DoEvents
For nC = 1 To Len(txt)
    J = Asc(Mid(txt, nC, 1))
    crc = crc Xor J
    For J = 1 To 8
        mask = 0
        If crc / 2 <> Int(crc / 2) Then mask = &HA001
        crc = Int(crc / 2) And &H7FFF: crc = crc Xor mask
    Next J
Next nC

CRC16HASH = "c" & Hex$(crc)
End Function

 



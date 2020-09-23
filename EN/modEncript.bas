Attribute VB_Name = "modEncript"
Option Explicit

Public Function Encrypt(ByVal data As Collection, ByVal MaxLen As Integer, ByVal SerialN As String) As String
    Dim DataNum As Integer
    Dim i As Integer
    Dim j As Integer
    Dim EncSN As String
    Dim GlobalMot As String
    Dim SNDispatch As String
    Dim Mot(0 To 50) As String
    Dim MotsDispatch(0 To 50) As String
    Dim MotsDispatchXor(0 To 50) As String
    On Error Resume Next
    DataNum = data.Count
    For i = 1 To DataNum
        Mot(i) = data.Item(i)
        If Mot(i) = "" Then Mot(i) = "\Nothing/"
    Next i
     'Small encryption for SerialN
     EncSN = ""
    If SerialN = "" Then SerialN = "0123456789"
    For i = 1 To Len(SerialN)
        EncSN = EncSN + Mid(SerialN, i, 1) + Chr(i)
    Next i
   'Disp SN to MaxLen
    SNDispatch = Chr(Len(EncSN))
    For i = 1 To MaxLen Step Len(EncSN)
        SNDispatch = SNDispatch + EncSN
    Next i
    SNDispatch = SNDispatch + EncSN
    SNDispatch = Left(SNDispatch, MaxLen)
    MotsDispatch(0) = SNDispatch
    For j = 1 To DataNum
        MotsDispatch(j) = Chr(Len(Mot(j)))
        For i = 1 To MaxLen Step Len(Mot(j))
            MotsDispatch(j) = MotsDispatch(j) + Mot(j)
        Next i
        MotsDispatch(j) = MotsDispatch(j) + Mot(j)
        MotsDispatch(j) = Left(MotsDispatch(j), MaxLen)
    Next j

    For j = 0 To DataNum
        For i = 1 To MaxLen
           MotsDispatchXor(j) = MotsDispatchXor(j) + Chr(Asc(Mid(MotsDispatch(j), i, 1)) Xor Asc(Mid(SNDispatch, i, 1)))
        Next i
    Next j
    
    GlobalMot = ""
    For i = 1 To MaxLen
      For j = 0 To DataNum
        GlobalMot = GlobalMot + Mid(MotsDispatchXor(j), i, 1)
      Next j
    Next i
   Encrypt = GlobalMot + Chr(DataNum)
End Function
Public Sub Decrypt(ByVal MotEnc As String, data As Collection, ByVal MaxLen As Integer, ByVal SerialN As String)
    Dim GlobalMot As String
    Dim AscNum As Integer
    Dim DataNum As Integer
    Dim i As Integer
    Dim j As Integer
    Dim EncSN As String
    Dim SNDispatch As String
    Dim TestSN As String
    Dim Mot(0 To 50) As String
    Dim MotsDispatch(0 To 50) As String
    Dim MotsDispatchXor(0 To 50) As String
    Dim buf As String
    Set data = New Collection
    On Error Resume Next
    DataNum = Asc(Right(MotEnc, 1))
    MotEnc = Left(MotEnc, Len(MotEnc) - 1)
    'Small encryption for SerialN
     EncSN = ""
    If SerialN = "" Then SerialN = "0123456789"
    For i = 1 To Len(SerialN)
        EncSN = EncSN + Mid(SerialN, i, 1) + Chr(i)
    Next i
   'Disp SN to MaxLen
    SNDispatch = Chr(Len(EncSN))
    For i = 1 To MaxLen Step Len(EncSN)
        SNDispatch = SNDispatch + EncSN
    Next i
    SNDispatch = SNDispatch + EncSN
    SNDispatch = Left(SNDispatch, MaxLen)
    
    '------------------
    GlobalMot = MotEnc
     For i = 1 To Len(GlobalMot) Step (DataNum + 1)
        For j = 0 To DataNum
            MotsDispatchXor(j) = MotsDispatchXor(j) + Mid(GlobalMot, i + j, 1)
        Next j
    Next i
    
    For j = 0 To DataNum
    MotsDispatch(j) = ""
        For i = 1 To MaxLen
           MotsDispatch(j) = MotsDispatch(j) + Chr(Asc(Mid(MotsDispatchXor(j), i, 1)) Xor Asc(Mid(SNDispatch, i, 1)))
        Next i
    Next j
    If MotsDispatch(0) <> SNDispatch Then GoTo EncErr
    For i = 1 To DataNum
        AscNum = Asc(Left(MotsDispatch(i), 1))
        buf = Mid(MotsDispatch(i), 2, AscNum)
        If buf = "\Nothing/" Then buf = ""
        data.Add buf
    Next i
    Exit Sub
EncErr:
    For i = 1 To DataNum
        data.Add ""
    Next i
End Sub


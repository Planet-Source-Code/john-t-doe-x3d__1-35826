Attribute VB_Name = "ColorMunip"

Public Sub Hex2RGB(strHexColor As String, R As Byte, G As Byte, B As Byte)
    Dim HexColor As String
    Dim i As Byte
    On Error Resume Next
    ' make sure the string is 6 characters l
    '     ong
    ' (it may have been given in &H###### fo
    '     rmat, we want ######)
    strHexColor = Right((strHexColor), 6)
    ' however, it may also have been given a
    '     s or #***** format, so add 0's in front


    For i = 1 To (6 - Len(strHexColor))
        HexColor = HexColor & "0"
    Next
    HexColor = HexColor & strHexColor
    ' convert each set of 2 characters into
    '     bytes, using vb's cbyte function
    R = CByte("&H" & Right$(HexColor, 2))
    G = CByte("&H" & Mid$(HexColor, 3, 2))
    B = CByte("&H" & Left$(HexColor, 2))
End Sub


Public Function RGB2Hex(R As Byte, G As Byte, B As Byte) As String
    On Error Resume Next
    ' convert to long using vb's rgb functio
    '     n, then use the long2rgb function
    RGB2Hex = Long2Hex(RGB(R, G, B))
End Function


Public Sub Long2RGB(LongColor As Long, R As Byte, G As Byte, B As Byte)
    On Error Resume Next
    ' convert to hex using vb's hex function
    '     , then use the hex2rgb function
    Hex2RGB (Hex(LongColor)), R, G, B
End Sub


Public Function RGB2Long(R As Byte, G As Byte, B As Byte) As Long
    On Error Resume Next
    ' use vb's rgb function
    RGB2Long = RGB(R, G, B)
End Function


Public Function Long2Hex(LongColor As Long) As String
    On Error Resume Next
    ' use vb's hex function
    Long2Hex = Hex(LongColor)
End Function


Public Function Hex2Long(strHexColor As String) As Long
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    On Error Resume Next
    ' use the hex2rgb function to get the re
    '     d green and blue bytes
    Hex2RGB strHexColor, R, G, B
    ' convert to long using vb's rgb functio
    '     n
    Hex2Long = RGB(R, G, B)
End Function



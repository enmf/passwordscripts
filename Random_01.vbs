Option Explicit

Sub Main
    'run this to test md5, sha1, sha2/256, sha384, sha2/512 with salt, or sha2/512
    Dim Sin As String, sOut As String, b64 As Boolean
    Dim sH As String, sSecret1 As String
    Dim sH1 As String,  sSecret2 As String

    'insert the text to hash within the sIn quotes
    'and for selected procedures a string for the secret key
    Sin = ""
    sSecret1 = "oilAndVin" 'secret key for StrToSHA512Salt only
    sSecret2 = "test2" 'secret key for StrToSHA512Salt only

    'select as required
    'b64 = False   'output hex
    b64 = True   'output base-64

    'enable any one
    'sH = MD5(sIn, b64)
    'sH = SHA1(sIn, b64)
    'sH = SHA256(sIn, b64)
    'sH = SHA384(sIn, b64)
    sH = StrToSHA512Salt(Sin, sSecret1, b64)
    sH1 = StrToSHA512Salt(Sin, sSecret2, b64)
    'message box and immediate window outputs
    Debug.Print sH & vbNewLine & Len(sH) & " characters in length"
    'MsgBox sH & vbNewLine & Len(sH) & " characters in length"

	Dim loFreeFile As Integer
	loFreeFile = FreeFile

	Open "C:\TEMP\FilePassw" + "225" + ".txt" For Output As #loFreeFile
		Print #loFreeFile, sH
		Print #loFreeFile, sH1
	Close #loFreeFile
    'de-comment this block to place the hash in first cell of sheet1
'    With ThisWorkbook.Worksheets("Sheet1").Cells(1, 1)
'        .Font.Name = "Consolas"
'        .Select: Selection.NumberFormat = "@" 'make cell text
'        .Value = sH
'    End With

End Sub

Public Function MD5(ByVal Sin As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit

    'Test with empty string input:
    'Hex:   d41d8cd98f00...etc
    'Base-64: 1B2M2Y8Asg...etc

    Dim oT As Object, oMD5 As Object
    Dim TextToHash() As Byte
    Dim bytes() As Byte

    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

    TextToHash = oT.GetBytes_4(Sin)
    bytes = oMD5.ComputeHash_2((TextToHash))

    If bB64 = True Then
       MD5 = ConvToBase64String(bytes)
    Else
       MD5 = ConvToHexString(bytes)
    End If

    Set oT = Nothing
    Set oMD5 = Nothing

End Function

Public Function SHA1(Sin As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit

    'Test with empty string input:
    '40 Hex:   da39a3ee5e6...etc
    '28 Base-64:   2jmj7l5rSw0yVb...etc

    Dim oT As Object, oSHA1 As Object
    Dim TextToHash() As Byte
    Dim bytes() As Byte

    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA1 = CreateObject("System.Security.Cryptography.SHA1Managed")

    TextToHash = oT.GetBytes_4(Sin)
    bytes = oSHA1.ComputeHash_2((TextToHash))

    If bB64 = True Then
       SHA1 = ConvToBase64String(bytes)
    Else
       SHA1 = ConvToHexString(bytes)
    End If

    Set oT = Nothing
    Set oSHA1 = Nothing

End Function

Public Function SHA256(Sin As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit

    'Test with empty string input:
    '64 Hex:   e3b0c44298f...etc
    '44 Base-64:   47DEQpj8HBSa+/...etc

    Dim oT As Object, oSHA256 As Object
    Dim TextToHash() As Byte, bytes() As Byte

    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")

    TextToHash = oT.GetBytes_4(Sin)
    bytes = oSHA256.ComputeHash_2((TextToHash))

    If bB64 = True Then
       SHA256 = ConvToBase64String(bytes)
    Else
       SHA256 = ConvToHexString(bytes)
    End If

    Set oT = Nothing
    Set oSHA256 = Nothing

End Function

Public Function SHA384(Sin As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit

    'Test with empty string input:
    '96 Hex:   38b060a751ac...etc
    '64 Base-64:   OLBgp1GsljhM2T...etc

    Dim oT As Object, oSHA384 As Object
    Dim TextToHash() As Byte, bytes() As Byte

    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA384 = CreateObject("System.Security.Cryptography.SHA384Managed")

    TextToHash = oT.GetBytes_4(Sin)
    bytes = oSHA384.ComputeHash_2((TextToHash))

    If bB64 = True Then
       SHA384 = ConvToBase64String(bytes)
    Else
       SHA384 = ConvToHexString(bytes)
    End If

    Set oT = Nothing
    Set oSHA384 = Nothing

End Function

Public Function SHA512(Sin As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit

    'Test with empty string input:
    '128 Hex:   cf83e1357eefb8bd...etc
    '88 Base-64:   z4PhNX7vuL3xVChQ...etc

    Dim oT As Object, oSHA512 As Object
    Dim TextToHash() As Byte, bytes() As Byte

    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA512 = CreateObject("System.Security.Cryptography.SHA512Managed")

    TextToHash = oT.GetBytes_4(Sin)
    bytes = oSHA512.ComputeHash_2((TextToHash))

    If bB64 = True Then
       SHA512 = ConvToBase64String(bytes)
    Else
       SHA512 = ConvToHexString(bytes)
    End If

    Set oT = Nothing
    Set oSHA512 = Nothing

End Function

Function StrToSHA512Salt(ByVal Sin As String, ByVal sSecretKey As String, _
                           Optional ByVal b64 As Boolean = False) As String
    'Returns a sha512 STRING HASH in function name, modified by the parameter sSecretKey.
    'This hash differs from that of SHA512 using the SHA512Managed class.
    'HMAC class inputs are hashed twice;first input and key are mixed before hashing,
    'then the key is mixed with the result and hashed again.

    Dim Asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SecretKey() As Byte
    Dim bytes() As Byte

    'Test results with both strings empty:
    '128 Hex:    b936cee86c9f...etc
    '88 Base-64:   uTbO6Gyfh6pd...etc

    'create text and crypto objects
    Set Asc = CreateObject("System.Text.UTF8Encoding")

    'Any of HMACSHAMD5,HMACSHA1,HMACSHA256,HMACSHA384,or HMACSHA512 can be used
    'for corresponding hashes, albeit not matching those of Managed classes.
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA512")

    'make a byte array of the text to hash
    bytes = Asc.Getbytes_4(Sin)
    'make a byte array of the private key
    SecretKey = Asc.Getbytes_4(sSecretKey)
    'add the private key property to the encryption object
    enc.Key = SecretKey

    'make a byte array of the hash
    bytes = enc.ComputeHash_2((bytes))

    'convert the byte array to string
    If b64 = True Then
       StrToSHA512Salt = ConvToBase64String(bytes)
    Else
       StrToSHA512Salt = ConvToHexString(bytes)
    End If

    'release object variables
    Set Asc = Nothing
    Set enc = Nothing

End Function

Private Function ConvToBase64String(vIn As Variant) As Variant

    Dim oD As Object

    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")

    Set oD = Nothing

End Function

Private Function ConvToHexString(vIn As Variant) As Variant

    Dim oD As Object

    Set oD = CreateObject("MSXML2.DOMDocument")

      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")

    Set oD = Nothing

End Function







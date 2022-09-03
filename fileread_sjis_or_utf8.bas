Option Explicit

'Reads files in UTF8 and returns text
Function ReadTextFromFileAsUTF8(filename As String) As String
  
  Dim text As String
  text = ReadTextFromFileUseADODBStream(filename, "UTF-8") 'UTF-8 with BOM is also readable.
  ReadTextFromFileAsUTF8 = text

End Function

'Reads files in SJIS and returns text
Function ReadTextFromFileAsSJIS(filename As String) As String
  
  Dim text As String
  text = ReadTextFromFileUseADODBStream(filename, "shift_jis")
  ReadTextFromFileAsSJIS = text

End Function

'Reads a file with the specified character encoding and returns the text.
Function ReadTextFromFileUseADODBStream(filename As String, encoding As String) As String

  Dim text As String
  Dim ado As Variant

  Set ado = CreateObject("ADODB.Stream")
  ado.Charset = encoding
  ado.Open
  ado.LoadFromFile filename
  text = ado.ReadText
  ado.Close
  Set ado = Nothing
  
  ReadTextFromFileUseADODBStream = text

End Function


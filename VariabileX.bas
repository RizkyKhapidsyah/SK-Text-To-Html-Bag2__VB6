Attribute VB_Name = "VariabileX"
Public ContinutFisier As String
Public Salvat As Boolean
Public Titlu As String
Public CulText As String
Public Fond As String
Public LinkX As String
Public VLinkX As String
Public ALinkX As String
Public CuloareTemp As String
' Shell32
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Function CautareAvansata(strTextHtml As String) As String
' caractere speciale plus linkuri
Dim intTtoTHT As Long
Dim intContSec As Long
Dim strTempLI As String
Dim strRestul As String
intTtoTHT = 1

Do
Select Case Mid(strTextHtml, intTtoTHT, 1)

Case Chr(34): strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&quot;" & Mid(strTextHtml, intTtoTHT + 1)
Case Chr(13): strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "<br>" & Mid(strTextHtml, intTtoTHT + 1): intTtoTHT = intTtoTHT + 3
Case "&": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&amp;" & Mid(strTextHtml, intTtoTHT + 1)
Case ">": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&gt;" & Mid(strTextHtml, intTtoTHT + 1)
Case "<": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&lt;" & Mid(strTextHtml, intTtoTHT + 1)
Case Chr(32): strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&nbsp;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¡": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&iexcl;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¢": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&cent;" & Mid(strTextHtml, intTtoTHT + 1)
Case "£": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&pound;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¤": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&curren;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¥": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&yen;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¦": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&brvbar;" & Mid(strTextHtml, intTtoTHT + 1)
Case "§": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&sect;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¨": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&uml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "©": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&copy;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ª": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ordf;" & Mid(strTextHtml, intTtoTHT + 1)
Case "«": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&laquo;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¬": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&not;" & Mid(strTextHtml, intTtoTHT + 1)
Case "­": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&shy;" & Mid(strTextHtml, intTtoTHT + 1)
Case "®": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&reg;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¯": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&macr;" & Mid(strTextHtml, intTtoTHT + 1)
Case "°": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&deg;" & Mid(strTextHtml, intTtoTHT + 1)
Case "²": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&sup2;" & Mid(strTextHtml, intTtoTHT + 1)
Case "±": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&plusmn;" & Mid(strTextHtml, intTtoTHT + 1)
Case "³": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&sup3;" & Mid(strTextHtml, intTtoTHT + 1)
Case "´": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&acute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "µ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&micro;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¶": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&para;" & Mid(strTextHtml, intTtoTHT + 1)
Case "·": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&middot;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¸": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&cedil;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¹": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&sup1;" & Mid(strTextHtml, intTtoTHT + 1)
Case "º": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ordm;" & Mid(strTextHtml, intTtoTHT + 1)
Case "»": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&raquo;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¼": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&frac14;" & Mid(strTextHtml, intTtoTHT + 1)
Case "½": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&frac12;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¾": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&frac34;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¿": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&iquest;" & Mid(strTextHtml, intTtoTHT + 1)
Case "À": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Agrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Á": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Aacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Â": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Acirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ã": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Atilde;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ä": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Auml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Å": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Aring;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Æ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&AElig;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ç": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ccedil;" & Mid(strTextHtml, intTtoTHT + 1)
Case "È": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Egrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "É": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Eacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ê": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ecirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ë": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Euml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ì": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Igrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Í": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Iacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Î": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Icirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ï": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Iuml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ð": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ETH;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ñ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ntilde;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ò": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ograve;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ó": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Oacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ô": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ocirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Õ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Otilde;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ö": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ouml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "×": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&times;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ø": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Oslash;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ù": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ugrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ú": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Uacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Û": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ucirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ü": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Uuml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ý": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Yacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Þ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&THORN;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ß": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&szlig;" & Mid(strTextHtml, intTtoTHT + 1)
Case "à": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&agrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "á": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&aacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "â": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&acirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ã": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&atilde;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ä": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&auml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "å": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&aring;" & Mid(strTextHtml, intTtoTHT + 1)
Case "æ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&aelig;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ç": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ccedil;" & Mid(strTextHtml, intTtoTHT + 1)
Case "è": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&egrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "é": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ecirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ë": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&euml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ì": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&igrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "í": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&iacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "î": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&icirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ï": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&iuml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ð": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&eth;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ñ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ntilde;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ò": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ograve;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ó": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&oacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ô": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ocirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "õ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&otilde;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ö": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ouml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "÷": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&divide;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ø": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&oslash;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ù": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ugrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ú": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&uacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "û": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ucirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ü": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&uuml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ý": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&yacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "þ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&thorn;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ÿ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&yuml;" & Mid(strTextHtml, intTtoTHT + 1)

End Select
' *********************************************
  If LCase(Mid(strTextHtml, intTtoTHT, 3)) = "www" Or LCase(Mid(strTextHtml, intTtoTHT, 7)) = "http://" Or LCase(Mid(strTextHtml, intTtoTHT, 8)) = "https://" Or LCase(Mid(strTextHtml, intTtoTHT, 7)) = "mailto:" Or LCase(Mid(strTextHtml, intTtoTHT, 6)) = "ftp://" Or LCase(Mid(strTextHtml, intTtoTHT, 5)) = "news:" Or LCase(Mid(strTextHtml, intTtoTHT, 9)) = "gopher://" Or LCase(Mid(strTextHtml, intTtoTHT, 8)) = "file:///" Or LCase(Mid(strTextHtml, intTtoTHT, 9)) = "telnet://" Or LCase(Mid(strTextHtml, intTtoTHT, 7)) = "wais://" Then
         For intContSec = intTtoTHT To Len(strTextHtml)
             If Mid(strTextHtml, intContSec, 1) = " " Then
               strRestul = Mid(strTextHtml, intContSec)
               strTempLI = Mid(strTextHtml, intTtoTHT, (intContSec - intTtoTHT))
               strTempLI = "<a href=" & Chr(34) & strTempLI & Chr(34) & ">" & strTempLI & "</a>"
               strTextHtml = Left(strTextHtml, intTtoTHT - 1) & strTempLI & strRestul
               intTtoTHT = intTtoTHT + Len(strTempLI)
               intContSec = 0
               strTempLI = ""
               strRestul = ""
               Exit For
             End If
         Next intContSec
  End If
' ********************************************

  intTtoTHT = intTtoTHT + 1
Loop Until intTtoTHT > Len(strTextHtml)

CautareAvansata = strTextHtml
End Function


Public Function CautareSimpla(strTextHtml As String) As String
' doar caractere speciale
Dim intTtoTHT As Long
intTtoTHT = 1

Do
Select Case Mid(strTextHtml, intTtoTHT, 1)

Case Chr(34): strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&quot;" & Mid(strTextHtml, intTtoTHT + 1)
Case Chr(13): strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "<br>" & Mid(strTextHtml, intTtoTHT + 1): intTtoTHT = intTtoTHT + 3
Case "&": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&amp;" & Mid(strTextHtml, intTtoTHT + 1)
Case ">": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&gt;" & Mid(strTextHtml, intTtoTHT + 1)
Case "<": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&lt;" & Mid(strTextHtml, intTtoTHT + 1)
Case Chr(32): strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&nbsp;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¡": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&iexcl;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¢": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&cent;" & Mid(strTextHtml, intTtoTHT + 1)
Case "£": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&pound;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¤": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&curren;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¥": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&yen;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¦": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&brvbar;" & Mid(strTextHtml, intTtoTHT + 1)
Case "§": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&sect;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¨": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&uml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "©": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&copy;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ª": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ordf;" & Mid(strTextHtml, intTtoTHT + 1)
Case "«": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&laquo;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¬": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&not;" & Mid(strTextHtml, intTtoTHT + 1)
Case "­": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&shy;" & Mid(strTextHtml, intTtoTHT + 1)
Case "®": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&reg;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¯": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&macr;" & Mid(strTextHtml, intTtoTHT + 1)
Case "°": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&deg;" & Mid(strTextHtml, intTtoTHT + 1)
Case "²": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&sup2;" & Mid(strTextHtml, intTtoTHT + 1)
Case "±": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&plusmn;" & Mid(strTextHtml, intTtoTHT + 1)
Case "³": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&sup3;" & Mid(strTextHtml, intTtoTHT + 1)
Case "´": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&acute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "µ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&micro;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¶": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&para;" & Mid(strTextHtml, intTtoTHT + 1)
Case "·": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&middot;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¸": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&cedil;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¹": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&sup1;" & Mid(strTextHtml, intTtoTHT + 1)
Case "º": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ordm;" & Mid(strTextHtml, intTtoTHT + 1)
Case "»": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&raquo;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¼": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&frac14;" & Mid(strTextHtml, intTtoTHT + 1)
Case "½": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&frac12;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¾": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&frac34;" & Mid(strTextHtml, intTtoTHT + 1)
Case "¿": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&iquest;" & Mid(strTextHtml, intTtoTHT + 1)
Case "À": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Agrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Á": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Aacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Â": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Acirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ã": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Atilde;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ä": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Auml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Å": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Aring;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Æ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&AElig;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ç": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ccedil;" & Mid(strTextHtml, intTtoTHT + 1)
Case "È": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Egrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "É": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Eacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ê": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ecirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ë": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Euml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ì": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Igrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Í": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Iacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Î": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Icirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ï": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Iuml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ð": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ETH;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ñ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ntilde;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ò": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ograve;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ó": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Oacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ô": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ocirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Õ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Otilde;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ö": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ouml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "×": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&times;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ø": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Oslash;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ù": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ugrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ú": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Uacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Û": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Ucirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ü": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Uuml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Ý": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&Yacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "Þ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&THORN;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ß": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&szlig;" & Mid(strTextHtml, intTtoTHT + 1)
Case "à": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&agrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "á": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&aacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "â": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&acirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ã": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&atilde;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ä": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&auml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "å": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&aring;" & Mid(strTextHtml, intTtoTHT + 1)
Case "æ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&aelig;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ç": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ccedil;" & Mid(strTextHtml, intTtoTHT + 1)
Case "è": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&egrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "é": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ecirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ë": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&euml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ì": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&igrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "í": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&iacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "î": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&icirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ï": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&iuml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ð": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&eth;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ñ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ntilde;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ò": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ograve;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ó": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&oacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ô": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ocirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "õ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&otilde;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ö": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ouml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "÷": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&divide;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ø": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&oslash;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ù": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ugrave;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ú": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&uacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "û": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&ucirc;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ü": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&uuml;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ý": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&yacute;" & Mid(strTextHtml, intTtoTHT + 1)
Case "þ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&thorn;" & Mid(strTextHtml, intTtoTHT + 1)
Case "ÿ": strTextHtml = Left(strTextHtml, intTtoTHT - 1) & "&yuml;" & Mid(strTextHtml, intTtoTHT + 1)

End Select
  intTtoTHT = intTtoTHT + 1
Loop Until intTtoTHT > Len(strTextHtml)

CautareSimpla = strTextHtml
End Function


Public Sub Colorare()
Dim Z As Long
Dim R As Long
Dim strG As String
strG = vbBlack

For Z = 1 To Len(TextToHtml.RTB1.Text)

  If Mid(TextToHtml.RTB1.Text, Z, 1) = "<" Then
      strG = vbBlue
      Z = Z - 1
      TextToHtml.RTB1.SelStart = Z
      TextToHtml.RTB1.SelLength = 1
      TextToHtml.RTB1.SelColor = strG
      Z = Z + 1
  End If
  If Mid(TextToHtml.RTB1.Text, Z, 2) = "<a" Or Mid(TextToHtml.RTB1.Text, Z, 3) = "</a" Then
      strG = vbRed
      Z = Z - 1
      TextToHtml.RTB1.SelStart = Z
      TextToHtml.RTB1.SelLength = 1
      TextToHtml.RTB1.SelColor = strG
      Z = Z + 1
  End If
  
  If Mid(TextToHtml.RTB1.Text, Z, 1) = ">" Then strG = vbBlack
  
  TextToHtml.RTB1.SelStart = Z
  TextToHtml.RTB1.SelLength = 1
  TextToHtml.RTB1.SelColor = strG
Next Z

End Sub

Public Function PreTag(strTextHtml As String) As String
Dim intTtoTHT As Long
Dim intContSec As Long
Dim strTempLI As String
Dim strRestul As String

For intTtoTHT = 1 To Len(strTextHtml) - 9
  
  If LCase(Mid(strTextHtml, intTtoTHT, 3)) = "www" Or LCase(Mid(strTextHtml, intTtoTHT, 7)) = "http://" Or LCase(Mid(strTextHtml, intTtoTHT, 8)) = "https://" Or LCase(Mid(strTextHtml, intTtoTHT, 7)) = "mailto:" Or LCase(Mid(strTextHtml, intTtoTHT, 6)) = "ftp://" Or LCase(Mid(strTextHtml, intTtoTHT, 5)) = "news:" Or LCase(Mid(strTextHtml, intTtoTHT, 9)) = "gopher://" Or LCase(Mid(strTextHtml, intTtoTHT, 8)) = "file:///" Or LCase(Mid(strTextHtml, intTtoTHT, 9)) = "telnet://" Or LCase(Mid(strTextHtml, intTtoTHT, 7)) = "wais://" Then
         For intContSec = intTtoTHT To Len(strTextHtml)
             If Mid(strTextHtml, intContSec, 1) = " " Then
               strRestul = Mid(strTextHtml, intContSec)
               strTempLI = Mid(strTextHtml, intTtoTHT, (intContSec - intTtoTHT))
               strTempLI = "<a href=" & Chr(34) & strTempLI & Chr(34) & ">" & strTempLI & "</a>"
               strTextHtml = Left(strTextHtml, intTtoTHT - 1) & strTempLI & strRestul
               intTtoTHT = intTtoTHT + Len(strTempLI)
               intContSec = 0
               strTempLI = ""
               strRestul = ""
               Exit For
             End If
         Next intContSec
  End If
  
Next intTtoTHT

PreTag = strTextHtml
End Function

Public Function Verificare(ByVal fisier As String) As Boolean
On Error GoTo ErrorMan1

Open fisier For Input As #1
  Verificare = True
Close #1

Exit Function
ErrorMan1:
  Verificare = False
End Function

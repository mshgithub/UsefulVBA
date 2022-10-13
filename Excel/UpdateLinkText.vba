' This macro copies the ID number that comes after "=" at 
' the end of a specific type of URL and replaces the link
' text with the ID number
Sub btnUpdateLinksText()
  For link = 1 To Sheet4.Hyperlinks.Count
    Url = Sheet4.Hyperlinks(link).Address
    LinkText = Right(Url, (Len(Url) - InStr(Url, "=")))
    Sheet4.Hyperlinks(link).TextToDisplay = LinkText
  Next link
End Sub

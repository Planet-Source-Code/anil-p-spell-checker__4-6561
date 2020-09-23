<div align="center">

## Spell Checker


</div>

### Description

Use Microsoft word spell checker utility
 
### More Info
 
Client machine has word 2000

Browser is I.E.

Activex scripting is enabled


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Anil P](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/anil-p.md)
**Level**          |Advanced
**User Rating**    |4.9 (44 globes from 9 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/anil-p-spell-checker__4-6561/archive/master.zip)





### Source Code

```
<SCRIPT LANGUAGE=vbscript>
<!--
'SpellChecker
' PURPOSE: This function accepts Text data for which spell checking has to be done.
' Return's Spelling corrected data
'
Function SpellChecker(TextValue)
	Dim objWordobject
 	Dim objDocobject
 	Dim strReturnValue
 	'Create a new instance of word Application
 	Set objWordobject = CreateObject("word.Application")
 	objWordobject.WindowState = 2
 	objWordobject.Visible = True
	'Create a new instance of Document
 	Set objDocobject = objWordobject.Documents.Add( , , 1, True)
 	objDocobject.Content=TextValue
 	objDocobject.CheckSpelling
 	'Return spell check completed text data
	strReturnValue = objDocobject.Content
	'Close Word Document
 	objDocobject.Close false
	'Set Document to nothing
 	Set objDocobject = Nothing
	'Quit Word
 	objWordobject.Application.Quit True
	'Set word object to nothing
 	Set objWordobject= Nothing
    SpellChecker=strReturnValue
End Function
-->
</SCRIPT>
```


<div align="center">

## Yet another Table coloring


</div>

### Description

This gives an alternating line color on tables, it also allows a realtime highlighting.
 
### More Info
 
All explained int he code

All explained within the code


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jonathan Barton](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jonathan-barton.md)
**Level**          |Intermediate
**User Rating**    |4.5 (36 globes from 8 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML
**Category**       |[Controls/ Forms/ Dialogs/ Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/controls-forms-dialogs-menus__4-3.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jonathan-barton-yet-another-table-coloring__4-6329/archive/master.zip)





### Source Code

```
<%
	'The const's are for the colors and the Flag is needed
	Const sPrimaryColor = "WHITE"
	Const sSecondaryColor = "YELLOW"
	Dim bColorFlag
	dim lCounter
	'This function does all the work for you.
	Function LineColor()
		bColorFlag = not bColorFlag
		If bColorFlag then
			LineColor = sPrimaryColor
		Else
			LineColor = sSecondaryColor
		End if
	End Function
%>
<SCRIPT language='VBScript'>
	Dim sRowOldColor
	Const sHighlightColor = "LIGHTBLUE"
	Sub Document_onclick()
		Dim sID
		'Need to show the parent of the element that was clicked on
		If Window.Event.srcElement.tagName = "TD" Then
			sID = Window.Event.srcElement.ParentElement.id
			msgbox "ID = " & sID,,"This is the ID of the row."
		End If
	End Sub
	Sub Document_onmouseover()
		If Window.Event.srcElement.tagName = "TD" Then
			sRowOldColor = window.event.srcElement.ParentElement.bgcolor
			window.event.SrcElement.ParentElement.bgcolor=sHighlightColor
		End If
	End Sub
	Sub Document_onmouseout ()
		If window.event.srcElement.tagName = "TD" Then
			window.event.srcElement.parentElement.bgcolor = sRowOldColor
		End If
	End Sub
</SCRIPT>
<HTML>
	<HEAD>
		<TITLE>
			Alternating Line Colors
		</TITLE>
	</HEAD>
	<BODY>
		<CENTER>
			<P>
				<BR>
			</P>
			<TABLE width='80%' cellspacing='3' cellpadding='3' border='1'>
			<%
				for lCounter = 1 to 10
					Response.Write("<TR bgcolor='" & LineColor & "' id='" & lCounter & "'>" & vbCrLf)
					Response.Write("<TD><CENTER>" & lCounter & "</CENTER></TD>" & vbCrLf)
					Response.Write("</TR>" & vbCrLf)
				next
			%>
			</TABLE>
		</CENTER>
	</BODY>
</HTML>
```


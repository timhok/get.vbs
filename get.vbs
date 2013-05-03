Set X=CreateObject("MSXML2.XMLHTTP"):Set D=CreateObject("ADODB.Stream"):Set S=Createobject("Scripting.FileSystemObject")
on error resume next'By Timhok
F=split(Wscript.arguments.Item(0),"/")(ubound(split(Wscript.arguments.Item(0),"/"))):X.open "GET", Wscript.arguments.Item(0), false:X.send()
If X.Status=200 Then
D.Open:D.Type=1:D.Write X.ResponseBody:D.Position=0
If S.Fileexists(F) Then S.DeleteFile F
D.SaveToFile F:D.Close
End if
Set S=Nothing:Set D=Nothing:Set X=Nothing
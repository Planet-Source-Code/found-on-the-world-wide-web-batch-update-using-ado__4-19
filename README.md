<div align="center">

## Batch Update using ADO


</div>

### Description

ADO has a great batch update feature that not many people take advantage of. You can use it to update many records at once without making multiple round trips to the database. Here is how to use it.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-batch-update-using-ado__4-19/archive/master.zip)





### Source Code

```
<HTML>
<HEAD><TITLE>Place Document Title Here</TITLE></HEAD>
<BODY BGColor=ffffff Text=000000>
<%
Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open Application("guestDSN")
rs.ActiveConnection = cn
rs.CursorType = adOpenStatic
rs.LockType = adLockBatchOptimistic
rs.Source = "SELECT * FROM authors"
rs.Open
If (rs("au_fname") = "Paul") or (rs("au_fname") = "Johnson") Then
newval = "Melissa"
Else
newval = "Paul"
End If
If err <> 0 Then
%>
<B>Error opening RecordSet</B>
<% Else %>
<B>Opened Successfully</B><P>
<% End If %>
<H2>Before Batch Update</H2>
<TABLE BORDER=1>
<TR>
<% For i = 0 to rs.Fields.Count - 1 %>
<TD><B><%= rs(i).Name %></B></TD>
<% Next %>
</TR>
<% For j = 1 to 5 %>
<TR>
<% For i = 0 to rs.Fields.Count - 1 %>
<TD><%= rs(i) %></TD>
<% Next %>
</TR>
<%
rs.MoveNext
Next
rs.MoveFirst
%>
</Table>
Move randomly in the table and perform updates to table.<BR>
<%
Randomize
r1 = Int(rnd*3) + 1 ' n Itterations
r2 = Int(rnd*2) + 1 ' n places skipped between updates
For i = 1 to r1
response.write "Itteration: " & i & "<br>"
rs("au_fname") = newval
For j = 1 to r2
rs.MoveNext
response.write "Move Next<br>"
Next
Next
rs.UpdateBatch adAffectAll
rs.Requery
rs.MoveFirst
%>
<% rs.MoveFirst %>
<H2>After Changes</H2>
<TABLE BORDER=1>
<TR>
<% For i = 0 to rs.Fields.Count - 1 %>
<TD><B><%= rs(i).Name %></B></TD>
<% Next %>
</TR>
<% For j = 1 to 5 %>
<TR>
<% For i = 0 to rs.Fields.Count - 1 %>
<TD><%= rs(i) %></TD>
<% Next %>
</TR>
<%
rs.MoveNext
Next
rs.Close
Cn.Close
%>
</TABLE>
```


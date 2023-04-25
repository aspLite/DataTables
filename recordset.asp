<!-- #include file="asplite/asplite.asp"-->
<!-- #include file="datatables.asp"-->
<%

'rather than use a real database, we create a dummy recordset 

dim rs : set rs=server.createobject("adodb.recordset")
rs.CursorLocation=3
rs.CursorType=3
rs.LockType=4

'define some colums
rs.fields.append "col1",203,15
rs.fields.append "col2",203,15
rs.fields.append "col3",203,15
rs.fields.append "col4",203,15

'in a real-life scenario, you would typically create and open a recordset with a select-statement: eg: rs.open "select * from mytable". 
rs.open

'fill the recordset with dummy data
dim i
for i=1 to 5000
	rs.addNew
	rs("col1")="abcdefg" & i
	rs("col2")="hijklmn" & i
	rs("col3")="opqrstu" & i
	rs("col4")="vwxyz" & i
next

rs.update

'prepare the recordset for DataTables
'check datatables.asp: most parameters coming back and forth DataTables are managed over there.
dim datatables : set datatables=new cls_datatables

'DataTables returns the INDEX of the column to be sorted on, not the actual column name
'In most cases, you have to translate that index to a column name, like here:
select case aspl.convertNmbr(datatables.OrderCol)	
	case 1 : rs.sort="col2 " & datatables.OrderDir
	case 2 : rs.sort="col3 " & datatables.OrderDir
	case 3 : rs.sort="col4 " & datatables.OrderDir
	case else : rs.sort="col1 " & datatables.OrderDir 'default sorting on col1
end select

'decide on which colums you want to search. You can search colums that you're not including in the DataTable.
dim strSearch : strSearch = datatables.searchValue
if strSearch<>"" then
	rs.filter="col1 like '*" & strSearch & "*'"
	rs.filter=rs.filter & " or col2 like '*" & strSearch & "*'"
	rs.filter=rs.filter & " or col3 like '*" & strSearch & "*'"
	rs.filter=rs.filter & "	or col4 like '*" & strSearch & "*'"
end if

'finally flush recordset as JSON. this also stops executing the page
datatables.dumpJson(rs)
%>
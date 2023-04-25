<%

class cls_datatables

	private draw,StartRecord,RowsPerPage,JsonAnswer,JsonHeader
	public OrderDir, OrderCol, searchValue

	Private Sub Class_Initialize()	

		'read out the search-box
		searchValue=trim(aspL.getRequest("search[value]"))
		searchValue=aspl.sqli(searchValue) 'protect agains sqlInjection!
	
		'which column has to be sorted on ?
		OrderCol = aspL.convertNmbr(aspL.getRequest("Order[0][column]"))		
		
		'asc or desc?
		OrderDir =  trim(lcase(aspL.getRequest("Order[0][dir]")))
		
		'prevent SQLinjection
		if OrderDir<>"asc" and OrderDir<>"desc" then OrderDir="asc"																																			  
		draw = aspL.convertNmbr(aspL.getRequest("draw"))
		
		'where exactly are we in the table?
		StartRecord = aspL.convertNmbr(aspL.getRequest("start"))
		
		if StartRecord = 0 then 
			StartRecord=1 'Absolutepage cannot be 0
		else
			StartRecord=StartRecord+1
		end if

		'how many rows per page? 
		RowsPerPage = aspL.convertNmbr(aspL.getRequest("length"))
		if RowsPerPage <= 0 then RowsPerPage=10
		if RowsPerPage>999999 then RowsPerPage=10	
		
	end sub
	
	public sub dumpJson(rs)
			
		'total number of results
		dim rTotal : rTotal=rs.recordcount
		
		'we only want a portion of the recordset 
		'starting from the startrecord (AbsolutePosition), and the next x rows (pagesize) only
		'if there are no records returned at all, do not set AbsolutePosition as this raises an error
		'no need for pagesize either in that case
		if rTotal>0 then
			rs.AbsolutePosition=StartRecord
			rs.pagesize=RowsPerPage
		end if

		'prepare JSON return - JSON takes care of the recordset paging! - see aspl.json.recordsetPaging
		aspl.json.recordsetPaging=true
		JsonAnswer=aspl.json.toJson("data", rs, false) 

		'finalizing JSON response - preparing header:
		JsonHeader = "{ ""draw"": "& draw &", "& vbcrlf
		JsonHeader = JsonHeader & """recordsTotal"": " & rTotal & ", "
		JsonHeader = JsonHeader & """recordsFiltered"": " & rTotal & ", "	
				  
		'removing from generated JSON initial bracket { and concatenating all together.
		JsonAnswer=right(JsonAnswer,Len(JsonAnswer)-1)
		JsonAnswer = JsonHeader & JsonAnswer

		set rs=nothing
		
		'writing a response and stop executing page
		aspL.dumpJson JsonAnswer		
	
	end sub
	
end class
%>
<%@ Language=VBScript%>

<script language="JScript" runat="server" src="json2.js"></script>

<%
Option Explicit
Response.Buffer = false
Response.CacheControl = "No-store"
%>
<!-- #Include virtual="adovbs.inc" --> 
'config.asp holds sql server connection strings, etc.
<!-- #Include File="config.asp" -->
<!-- #Include File="JsonObject.class.asp" -->

<%
dim reqJson: reqJson = request("json")
dim myJson, rs, mylist, conn
dim i:i=0
dim message:message=""
dim jsObject, jsOut, items, item, jsitem
dim exists:exists = false
dim sql: sql = ""


if ISNULL(reqJson) or ISEMPTY(reqJson) then
	response.write ""
end if

Set myJson = JSON.parse(reqJson) 

select case LCASE(myJson.type)
		
	case "catalog"
	
		dim Catalog:set Catalog = new CatalogItem
		
		select case LCASE(myJson.op)
		
			case "add"
			case "add-version-to-catalog"
				
				if not Catalog.SaveVersionToCatalog(myJson.data.id, myJson.data.catalogid) then				
					message = Catalog.Message
					exists = false
				else
					message = "Title successfully added."
					exists = true
				end if
											
				json = "{""exists"":""" &exists& ""","	
				json = json & """message"":""" &message& ""","
				json = json & """id"":""" & Catalog.id& """"
				json = json & "}"
				
				response.write json
				
				set Catalog = nothing
			
			case "edit"										
			case "delete"
			case "delete-version-from-catalog"
								
				if Catalog.DeleteVersionFromCatalog(myJson.data.id, myJson.data.catalogid) then
					message = "success"
				end if
											
				json = "{""exists"":""" &Catalog.Exists& ""","	
				json = json & """message"":""" &Catalog.Message& """"
				json = json & "}"
				
				response.write json
				
				set Catalog = nothing
			
			case "get-item"
		
				Catalog.LoadRecordById(myJson.id)

				if Catalog.Exists then
				
					json = "{""exists"":""true"","	
					json = json & """message"":""" & Catalog.Message & ""","											
					json = json & """catalog"":" & Catalog.ToJson()
					json = json & "}"
				
					response.write json
					
				else
				
					json = "{""exists"":""false"","	
					json = json & """message"":""" & Catalog.Message & ""","											
					json = json & "}"
					
					response.write json
					
				end if
				
				set Catalog = nothing
			
			case "get-list"
		
				dim myCatalogList: set myCatalogList = new GenericList
								
				Set conn = Server.CreateObject("ADODB.CONNECTION")
				Set rs = Server.CreateObject("ADODB.recordset")
				rs.CursorLocation = adUseClient		
				
				conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & scDataBase & "; UID=" & scSQLUser & "; PWD=" & scSQLpwd'
				
				
				sql = "SET NUMERIC_ROUNDABORT OFF;SET ARITHABORT ON;"
				sql = sql & "Select cat.id, cat.customerid, cat.applicationcustomertypeid as customertypeid, t.[type] as customerType,cat.ndx as [index],cat.[default],"
				sql = sql & " 'customerName' = case when t.[type] = 'reseller' then (Select company from dbo.Resellers Where resellerid=cat.customerId)"
				sql = sql & " when t.[type] = 'company' then (Select companyName from dbo.Company Where companyId=cat.customerId) else 'NA' end,"
				sql = sql & " cat.catalog, cat.sku, cat.timestamp as createDate,"
				sql = sql & "'description' = case when DATALENGTH(cat.description) > 0 then dbo.fn_EncodeBase64(cat.description) else '' end "
				sql = sql & " From ProductCustomerCatalog cat"
				sql = sql & " join ApplicationCustomerType t on cat.applicationcustomertypeid = t.id"
				sql = sql & " ORDER BY cat.catalog"
				sql = sql & " SET NUMERIC_ROUNDABORT ON;SET ARITHABORT OFF;"
				
				'response.write sql
				
				rs.CursorLocation = adUseClient
				rs.Open sql, conn, adOpenStatic, adLockBatchOptimistic
			
				if err.Number <> 0 then				
					m_message = "Database error"
				elseif NOT (rs.bof and rs.eof) then
					
					do while not rs.eof
															
						dim nitem: set nitem = new CatalogItem
					
						nitem.LoadFromRecord(rs)													
						myCatalogList.AddItem nitem

						rs.MoveNext()
					loop

					'rs.close()
					
				end if
					
				dim catalogJson:catalogJson = myCatalogList.ToJson()	
				
				if LEN(catalogJson) > 0 then
				
					json = "{""exists"":""true"","	
					json = json & """message"":""Success"","											
					json = json & """data"":[" & catalogJson & "]"
					json = json & "}"
				
					response.write json
					
				else
				
					json = "{""exists"":""false"","	
					json = json & """message"":""No data."","											
					json = json & """data"":[]"
					json = json & "}"
					
					response.write json
					
				end if
					
								
				if ISOBJECT(rs) then
					if not (rs Is Nothing) then
						if rs.State = 1 Then
							rs.Close
						end if
					end if
				end if
				
				set myCatalogList = nothing
				set rs = nothing
				
		end select
	
end select 


if ISOBJECT(conn) then
	set conn = nothing
end if

set myJson = nothing


%>

<% 'CatalogItem

	Class CatalogItem
	
		private m_id
		private m_customerId
		private m_customerType
		private m_customerName
		private m_index
		private m_name
		private m_sku
		private m_description
		private m_isdefault
		private m_createDate
		private m_exists
		private m_message
		
		Public Property Get Id()
			Id = m_id
		End Property
		Public Property Let Id(value)
			m_id = value
		End Property
		
		Public Property Get CustomerId()
			CustomerId = m_customerId
		End Property
		Public Property Let CustomerId(value)
			m_customerId = value
		End Property
		
		Public Property Get CustomerType()
			CustomerType = m_customerType
		End Property
		Public Property Let CustomerType(value)
			m_customerType = value
		End Property
		
		Public Property Get CustomerName()
			CustomerName = m_customerName
		End Property
		Public Property Let CustomerName(value)
			m_customerName = value
		End Property
		
		Public Property Get Index()
			Index = m_index
		End Property
		Public Property Let Index(value)
			m_index = value
		End Property
		
		Public Property Get Name()
			Name = m_name
		End Property
		Public Property Let Name(value)
			m_name = value
		End Property
		
		Public Property Get Sku()
			Sku = m_sku
		End Property
		Public Property Let Sku(value)
			m_sku = value
		End Property
		
		Public Property Get Description()
			Description = m_description
		End Property
		Public Property Let Description(value)
			m_description = value
		End Property
		
		Public Property Get Isdefault()
			Isdefault = m_isdefault
		End Property
		Public Property Let Isdefault(value)
			m_isdefault = value
		End Property
						
		Public Property Get CreateDate()
			CreateDate = m_createDate
		End Property
		Public Property Let CreateDate(value)
			m_createDate = value
		End Property
						
		Public Property Get Exists()
			Exists = m_exists
		End Property
		Public Property Let Exists(value)
			m_exists = value
		End Property
		
		Public Property Get Message()
			Message = m_message
		End Property
		Public Property Let Message(value)
			m_message = value
		End Property
		
		
		Private Sub Class_Initialize()
	
			m_id = 0			
			m_customerId = 0
			m_customerName = ""
			m_customerType = ""
			m_index = 0
			m_name = ""
			m_sku = ""
			m_description = ""
			m_isdefault = ""	
			m_createDate = now()
					
			m_exists = false
			m_message = ""
			
		End Sub
		
		Sub Class_Terminate
		
						
		End Sub
		
		
		Function SaveVersionToCatalog(versionId, catalogId)
		
			dim sql, conn, cmd
						
			m_exists = false
		
			Set conn = Server.CreateObject("ADODB.CONNECTION")
			Set cmd = Server.CreateObject("ADODB.COMMAND")
			
			conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & scDataBase & "; UID=" & scSQLUser & "; PWD=" & scSQLpwd'
			
		
			'sql = "set nocount off; "
			sql = ""
			sql = sql & "if NOT EXISTS(Select 1 from dbo.CatalogXref where catalogid = ? and sku = ?"
			sql = sql & " and productTypeId = (Select id from dbo.ProductTypeRef Where [type]='course-version')"
			sql = sql & " )begin"
			sql = sql & " Insert CatalogXref(catalogid, sku, productTypeId, timestamp)"
			sql = sql & " Select ?,?,id,getdate() From ProductTypeRef Where [type]='course-version';"
			sql = sql & " Select CONVERT(int,@@IDENTITY) As [id];"
			sql = sql & " end"
			sql = sql & " else"
			sql = sql & " begin Select CONVERT(int,0) as [id]; end "
			'sql = sql & " set nocount on;" 
			
			'response.write sql
			
			WITH cmd
				.Prepared = true
				.CommandType = adCmdText
				.ActiveConnection = conn
				.CommandText = sql			  
			END WITH
			
			cmd.Parameters.Append cmd.CreateParameter("@catalogid", adInteger, adParamInput, 0, catalogId)					
			cmd.Parameters.Append cmd.CreateParameter("@versionid", adVarChar, adParamInput, 10, versionId)
			cmd.Parameters.Append cmd.CreateParameter("@catalogid2", adInteger, adParamInput, 0, catalogId)					
			cmd.Parameters.Append cmd.CreateParameter("@versionid2", adVarChar, adParamInput, 10, versionId)
										
			dim rs:set rs = cmd.Execute
							
			If err.number > 0 then
				m_message = "Error.\nFailed to add title to catalog."
				m_exists = false
			else
				
				if NOT (rs.bof and rs.eof) then
		
					do until rs.EOF
			
						m_id = INT(rs("id"))
				
					rs.MoveNext 
					loop	
		
				end if
				
				if NOT ISNULL(m_id) and ISNUMERIC(m_id) then
					if m_id > 0 then
						m_exists = true
						m_message = "Title added to catalog"	
					
					else
						m_message = "Title already exists in catalog"
					end if
				end if
				
				
			end if
		
			rs.close
			SET rs = nothing
			
						
			SaveVersionToCatalog = m_exists
			
			conn.Close
			
			set cmd = nothing
			set conn = nothing
		
		End Function
		
		Function DeleteVersionFromCatalog(versionId, catalogId)
		
			dim sql, conn, cmd
			dim varRecordsAffected : varRecordsAffected = 0
			
			m_exists = false
		
			Set conn = Server.CreateObject("ADODB.CONNECTION")
			Set cmd = Server.CreateObject("ADODB.COMMAND")
			
			conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & scDataBase & "; UID=" & scSQLUser & "; PWD=" & scSQLpwd'
			
			WITH cmd
				.Prepared = true
				.CommandType = adCmdText
				.ActiveConnection = conn
						  
			END WITH
							
			sql = "Delete dbo.ProductCatalogXref where catalogId = " & catalogId & " and sku = '" & versionId & "'"
			sql = sql & " and productTypeId = (Select id from dbo.ProductTypeRef Where [type] = 'course-version');"
			sql = sql & " Select CONVERT(int,@@RowCount) as cnt;"
			
			'response.write sql
			
			cmd.CommandText = sql
						
			dim rs:set rs = cmd.Execute
			dim cnt:cnt = 0
							
			If err.number > 0 then
				m_message = "Error.\nFailed to delete title from catalog."
				m_exists = false
			else
				
				if NOT (rs.bof and rs.eof) then
		
					do until rs.EOF
			
						cnt = rs("cnt")
				
					rs.MoveNext 
					loop	
		
				end if
				
				if NOT ISNULL(cnt) and ISNUMERIC(cnt) then
					if cnt > 0 then
						m_exists = true
						m_message = "Title deleted from catalog"	
					
					else
						m_message = "Title not in catalog"
					end if
				end if
				
				
			end if
		
			rs.close
			SET rs = nothing
			
																	
			DeleteVersionFromCatalog = m_exists
			
			conn.Close
			
			set cmd = nothing
			set conn = nothing
		
		End Function
		
		Function LoadFromRecord(record)
	
			m_id = IIF(ISNULL(record("id")),0,record("id")) 
			m_customerId = IIF(ISNULL(record("customerid")),0,record("customerid"))
			m_customerType = IIF(ISNULL(record("customerType")),"",record("customerType"))
			m_customerName = IIF(ISNULL(record("customerName")),"",record("customerName"))
			m_index = IIF(ISNULL(record("index")),"",record("index"))
			m_name = IIF(ISNULL(record("catalog")),"",record("catalog"))
			m_sku = IIF(ISNULL(record("sku")),"",record("sku"))
			m_description = IIF(ISNULL(record("description")),"",record("description"))
			m_isdefault = IIF(ISNULL(record("default")),false,record("default"))			
			m_createDate = IIF(ISNULL(record("createDate")),"",record("createDate"))
					
		End Function
	
		Function LoadRecordById(id)
	
			
			dim sql, conn, rs
			dim varRecordsAffected : varRecordsAffected = 0
			dim siteCategoryId: siteCategoryId = 0
			dim authorId: authorId = 0
			
			
			m_exists = false
		
			Set conn = Server.CreateObject("ADODB.CONNECTION")
			Set rs = Server.CreateObject("ADODB.recordset")
			rs.CursorLocation = adUseClient		
			
			conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & scDataBase & "; UID=" & scSQLUser & "; PWD=" & scSQLpwd'
			
			sql = "SET NUMERIC_ROUNDABORT OFF;SET ARITHABORT ON;"
			sql = sql & "Select cat.id, cat.customerid, cat.applicationcustomertypeid as customertypeid, t.[type] as customerType,cat.ndx as [index],cat.[default],"
			sql = sql & " 'customerName' = case when t.[type] = 'reseller' then (Select company from dbo.Resellers Where resellerid=cat.customerId)"
			sql = sql & " when t.[type] = 'company' then (Select companyName from dbo.Company Where companyId=cat.customerId) else 'NA' end,"
			sql = sql & " cat.catalog, cat.sku, cat.timestamp as createDate,"
			sql = sql & "'description' = case when DATALENGTH(cat.description) > 0 then dbo.fn_EncodeBase64(cat.description) else '' end "
			sql = sql & " From ProductCustomerCatalog cat"
			sql = sql & " join ApplicationCustomerType t on cat.applicationcustomertypeid = t.id"
			sql = sql & " cat.id = " & id
			sql = sql & " Order By cat.catalog DESC;"
			sql = sql & " SET NUMERIC_ROUNDABORT ON;SET ARITHABORT OFF;"
					
			'response.write sql
			
			rs.Open sql, conn, adOpenStatic, adLockBatchOptimistic
			
			if err.Number <> 0 then				
				m_message = "Database error"
			elseif NOT (rs.bof and rs.eof) then
			
				m_exists = true
				
				do while not rs.eof
							
					LoadFromRecord(rs)
					
					rs.MoveNext()
				loop
				
			end if
				
			rs.close()
				
					
			LoadRecordById = m_exists
		
		End Function
	
		Sub LoadFromObject(data)
		
			m_id = IIF((ISNULL(data.id) OR LEN(data.id) = 0),0,data.id)
			
			
		End Sub
	
		Function ToJson()
		
			dim json:json = "{""exists"":""" &m_exists& ""","	
			json = json & """message"":""" &m_message& ""","
			json = json & """id"":""" &m_id& ""","			
			json = json & """customerId"":""" &m_customerId& ""","
			json = json & """customerType"":""" &m_customerType& ""","					
			json = json & """customerName"":""" &m_customerName& ""","	
			json = json & """index"":""" &m_index& ""","	
			json = json & """name"":""" &m_name& ""","	
			json = json & """sku"":""" &m_sku& ""","	
			json = json & """description"":""" &m_description& ""","	
			json = json & """isdefault"":""" &m_isdefault& ""","			
			json = json & """createdate"":""" &m_createDate& """"
								
			json = json & "}"
			
			ToJson = json
			
		End Function
		
	end Class

%>

<% 'GenericList

	Class GenericList
	
		private m_index
		private keys

		private m_message
		private m_exists
		private m_items

		
		Public Property Get Message()
			Message = m_message
		End Property
		Public Property Let Message(value)
			m_message = value
		End Property
		
		Public Property Get Exists()
			Exists = m_exists
		End Property
		Public Property Let Exists(value)
			m_exists = value
		End Property
		
		Public Property Get Index()
			m_index = Index
		End Property
			
		Private Sub Class_Initialize()
		
			m_index = -1
			m_message = ""
			m_exists = false
			set m_items = Server.CreateObject("Scripting.Dictionary")
			
			keys = Array
		
		End Sub
		
		Sub Class_Terminate
		
			set m_items = nothing
			if ISOBJECT(keys) then
				set keys = nothing
			end if
			
		End Sub
		
		Function Items()
		
			set Items = m_items
		
		End Function
		
		Function GetItem()
		
			if Not ISOBJECT(m_items) then
				set GetItem = nothing
				exit function
			end if
			if Not ISOBJECT(keys) then
				keys = m_items.Keys
			end if
			if ISOBJECT(keys) then
				if keys.count = 0 then
					set GetItem = nothing
					exit function
				end if
			end if
			
			if m_index = -1 then
				m_index = 0
			end if
			
			if m_items.Exists(keys(m_index)) then
			
				dim item
				SET item = m_items.item(keys(m_index)) 
			
				set GetItem = item
				exit function
			else
				set GetItem = nothing			
			end if	
		
		End Function
		
		Sub AddItem(item)
		
			dim i:i = m_items.count +1
			
			m_items.add i,item
			
		End Sub
		
		Function NextItem()
		
			dim m_continue : m_continue = true
			
			if m_index = m_items.Count-1 OR m_items.Count = 0 then
				m_continue = false
				m_index = -1
			else
				m_index = m_index+1
			end if	
			
			NextItem = m_continue
			
		End Function
		
			
			
			
		Function ToJson()
	
			dim json:json = ""
			
			Dim i:i=0
			Dim keys
			
			if ISOBJECT(m_items) then
				ReDim arrItems(m_items.count)
			end if
		
			Dim item								
			keys = m_items.Keys
			for i = 0 To m_items.Count -1	
			
				SET item = m_items.item(keys(i))
				dim ijson:ijson = item.ToJson()
				
				if NOT ISNULL(ijson) AND LEN(ijson) > 0 then
					json = json & item.ToJson() & ","
				end if
				
			next		
			
			if LEN(json) > 0 then
			
				ToJson = Left(json,Len(json)-1)
			else
				ToJson = json
			end if
			
		End Function
		
	
	End Class
%>
<% 'List
class List

	private m_index
	private keys

	private m_exists
	private m_items

	Public Property Get Exists()
		Exists = m_exists
	End Property
	Public Property Let Exists(value)
		m_exists = value
	End Property
	
	Public Property Get Index()
		m_index = Index
	End Property
		
	Private Sub Class_Initialize()
	
		m_index = -1
		m_exists = false
		set m_items = Server.CreateObject("Scripting.Dictionary")
		
		keys = Array
	
	End Sub
	
	Sub Class_Terminate
	
        set m_items = nothing
		if ISOBJECT(keys) then
			set keys = nothing
		end if
		
    End Sub
	
	Function Items()
	
		set Items = m_items
	
	End Function
	
	Function GetItem()
	
		if Not ISOBJECT(m_items) then
			set GetItem = nothing
			exit function
		end if
		if Not ISOBJECT(keys) then
			keys = m_items.Keys
		end if
		if ISOBJECT(keys) then
			if keys.count = 0 then
				set GetItem = nothing
				exit function
			end if
		end if
		
		if m_index = -1 then
			m_index = 0
		end if
		
		if m_items.Exists(keys(m_index)) then
		
			dim item
			SET item = m_items.item(keys(m_index)) 
		
			set GetItem = item
			exit function
		else
			set GetItem = nothing			
		end if	
	
	End Function
	
	Function NextItem()
	
		dim m_continue : m_continue = true
		
		if m_index = m_items.Count-1 OR m_items.Count = 0 then
			m_continue = false
			m_index = -1
		else
			m_index = m_index+1
		end if	
		
		NextItem = m_continue
		
	End Function
	
	Sub AddListItem(item)
	
		dim i:i = m_items.count +1
		
		m_items.add i,item
		
	End Sub
	
	Sub LoadFromRecordSet(records)
	
		m_exists = false
		dim i:i=0
	
		if ISOBJECT(records) then
		
			if not records.BOF then
				records.MoveFirst
			end if
			'if NOT (records.bof and records.eof) then
		
				m_exists = true
			
				do while not records.EOF
				
					dim item: set item = new ListItem
					item.Load records("key"), records("value")
					
					m_items.add i,item
					i = i+1
				
				records.MoveNext 
				loop	
				
			'end if
		end if
		
	End Sub
	
	Sub LoadFromDictionary(dict)
	
		Dim key,value		
		m_exists = false
		
		keys = dict.Keys
		for i = 0 To dict.Count -1									
			
			key = keys(i)
			value = dict.item(keys(i))
			
			m_items.Add key, value
		next
		
		if m_items.count > 0 then
			m_exists = true
		end if
	
	End Sub
	
	Function GetRecordSet(sql)
			
		Const adOpenStatic = 3    
		Const adUseClient = 3
		Const adLockBatchOptimistic = 4 
	
		m_exists = false
	
		dim conn, rs
		
		Set conn = Server.CreateObject("ADODB.CONNECTION")
		set rs = Server.CreateObject("adodb.recordset")	
		rs.CursorLocation = adUseClient		
		
		conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & scDataBase & "; UID=" & scSQLUser & "; PWD=" & scSQLpwd'
						
		'Set rs = conn.Execute(sql)
		rs.Open sql, conn, adOpenStatic, adLockBatchOptimistic
		Set rs.ActiveConnection = Nothing


		if err.Number <> 0 then
			set GetRecordSet = nothing
		elseif NOT (rs.bof and rs.eof) then
		
			m_exists = true
			
			'do while not rs.eof
			
			'	Response.write "rs - key: " & rs("key") & ", value: " & rs("value") & "<br/>"
			'	rs.MoveNext()
			'loop
								
		end if
		
		set GetRecordSet = rs
		
		conn.Close()
		'rs.Close()
		
		set rs = nothing
		set conn = nothing
	
	End Function
	
	Function GetAsArray()
	
		Dim arrItems()	
		Dim i:i=0
		Dim keys
		
		if ISOBJECT(m_items) then
			ReDim arrItems(m_items.count)
		end if
	
		Dim item								
		keys = m_items.Keys
		for i = 0 To m_items.Count -1									
			SET item = m_items.item(keys(i))
			
			arrItems(i) = item
		next
	
		GetAsArray = arrItems
	
	End Function
	
	Function Serialize()
	
		if NOT ISNULL(m_items) then
		
			'dim items: set items = m_items
			dim jsObject: set jsObject = new JSONobject
			
			dim item
			i=0
			do while NextItem()
				
				set item = GetItem()
				dim jsitem:set jsitem = new JSONobject
				jsitem.add "key", item.Key
				jsitem.add "value", item.KeyValue
				
				jsObject.Add i, jsitem
				
				i = i +1
				
			loop
			
			dim jsOut: set jsOut = new JSONobject
			jsOut.Add "data", jsObject
			
			Serialize = jsOut.Serialize()
			
			set jsObject = nothing
			set jsOut = nothing
		
		end if
	
	End Function
	
end class

%>
<% 'ListItem
class ListItem

	private m_key
	private m_value

	Public Property Get Key()
		Key = m_key
	End Property	
	Public Property Let Key(value)
		m_key = value
	End Property
	
	Public Property Get KeyValue()
		KeyValue = m_value
	End Property	
	Public Property Let KeyValue(value)
		m_value = value
	End Property
	
	Private Sub Class_Initialize()
	
		m_key = 0
		m_value = ""
	
	End Sub
	
	Sub Class_Terminate
        
    End Sub

	sub Load(key, value)
		m_key = key
		m_value = value
	end sub
	
	function ToJson()
	
		dim json: json = "{""key"":" +CSTR(m_key)+ ",""value"":""" &m_value& """}"	
		
		ToJson = json
	
	end function
	
end class

%>
<% 'Utilities

function IIf(test,t,f)

	if test then
	
		IIf = t
	else
		IIf = f
	end if

end function
Function IsNullArray(arr)

	ON ERROR RESUME NEXT
	Dim is_null : is_null = UBound(arr)
	if Err.Number = 0 then
		is_null = false
	else
		is_null = true
	end if
	
	IsNullArray = is_null

End Function

Function FieldExists(rs, fieldName) 

    On Error Resume Next
    FieldExists = rs.Fields(fieldName).name <> ""
    If Err <> 0 Then FieldExists = False
    Err.Clear

End Function

%>

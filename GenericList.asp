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

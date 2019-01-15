<% 'ScoPackageItem
	
	Class ScoPackageItem
	
		private m_id
		private m_courseid
		private m_scoobjectid
		private m_title
		private m_path
		private m_pathtype
		private m_scocontainername
		private m_scoentryfile
		private m_launchdata
		private m_masteryscore
		private m_createdate
		
		private m_exists
		private m_message
		
		Public Property Get Id()
			Id = m_id
		End Property
		Public Property Let Id(value)
			m_id = value
		End Property
		
		
		Public Property Get Courseid()
			Courseid = m_courseid
		End Property
		Public Property Let Courseid(value)
			m_courseid = value
		End Property
		
		Public Property Get Scoobjectid()
			Scoobjectid = m_scoobjectid
		End Property
		Public Property Let Scoobjectid(value)
			m_scoobjectid = value
		End Property
		
		Public Property Get Title()
			Title = m_title
		End Property
		Public Property Let Title(value)
			m_title = value
		End Property
		
		Public Property Get Path()
			Path = m_path
		End Property
		Public Property Let Path(value)
			m_path = value
		End Property
		
		Public Property Get Pathtype()
			Pathtype = m_pathtype
		End Property
		Public Property Let Pathtype(value)
			m_pathtype = value
		End Property
		
		Public Property Get ScoContainerName()
			ScoContainerName = m_scocontainername
		End Property
		Public Property Let ScoContainerName(value)
			m_scocontainername = value
		End Property
		
		Public Property Get ScoEntryFile()
			ScoEntryFile = m_scoentryfile
		End Property
		Public Property Let ScoEntryFile(value)
			m_scoentryfile = value
		End Property
		
		Public Property Get LaunchData()
			LaunchData = m_launchdata
		End Property
		Public Property Let LaunchData(value)
			m_launchdata = value
		End Property
		
		Public Property Get MasteryScore()
			MasteryScore = m_masteryscore
		End Property
		Public Property Let MasteryScore(value)
			m_masteryscore = value
		End Property
		
		Public Property Get CreateDate()
			CreateDate = m_createdate
		End Property
		Public Property Let CreateDate(value)
			m_createdate = value
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
			m_courseid = 0
			m_scoobjectid = ""
			m_title = ""
			m_path = "ScoObjects"
			m_pathtype = "relative"
			m_scocontainername = ""
			m_scoentryfile = "index_lms.html"
			m_launchdata = ""
			m_masteryscore = 0
			m_createdate = NOW()
			m_exists= false
			m_message = ""
			
		End Sub
		
		Sub Class_Terminate
		
		End Sub
		
		Function Save(operation)
		
			dim sql, conn, cmd
						
			m_exists = false
		
			Set conn = Server.CreateObject("ADODB.CONNECTION")
			Set cmd = Server.CreateObject("ADODB.COMMAND")
			
			conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & scDataBase & "; UID=" & scSQLUser & "; PWD=" & scSQLpwd'
			
			if operation = "add" then
			
				if ISNULL(m_title) or LEN(m_title) = 0 then
				
						m_message = "A title must be provided."
				else
				
					sql = "set nocount off; "
					sql = sql & " if NOT EXISTS(select 1 from dbo.Scorm_Sco where scoContainerName = ?)"
					sql = sql & " begin"
					sql = sql & " Insert Scorm_Sco(ScoObjectId, title, path, pathtype, scoContainerName, scoEntryFile, mastery_score, player_height_pxl, player_width_pxl, createDate)"
					sql = sql & " Values(?,?,'ScoObjects','relative',?,?,?,635,850,getdate()); Select convert(int,@@IDENTITY) As [id];"
					sql = sql & " end"
					sql = sql & " else"
					sql = sql & " begin"
					sql = sql & " Select convert(int,0) as [id];"
					sql = sql & " end"					
					sql = sql & " set nocount on;" 
					
					WITH cmd
						.Prepared = true
						.CommandType = adCmdText
						.ActiveConnection = conn
						.CommandText = sql			  
					END WITH
					
					if LEN(m_scoEntryFile) = 0 then
						m_scoEntryFile = "index_lms.html"
					end if
					
					cmd.Parameters.Append cmd.CreateParameter("@scocontainername", adVarChar, adParamInput, 255, m_scoContainerName)
					cmd.Parameters.Append cmd.CreateParameter("@scoobjectid", adVarChar, adParamInput, 50, m_scoObjectId)					
					cmd.Parameters.Append cmd.CreateParameter("@title", adVarChar, adParamInput, 255, m_title)
					cmd.Parameters.Append cmd.CreateParameter("@scocontainername", adVarChar, adParamInput, 255, m_scoContainerName)
					cmd.Parameters.Append cmd.CreateParameter("@scoEntryFile", adVarChar, adParamInput, 255, m_scoEntryFile)
					cmd.Parameters.Append cmd.CreateParameter("@masteryscore", adVarChar, adParamInput, 10, m_masteryScore)
														
					dim rs:set rs = cmd.Execute
									
					If err.number > 0 then
						m_message = "Error.\nFailed to add sco package."
						m_exists = false
					else
						m_exists = true		

						if NOT (rs.bof and rs.eof) then
				
							do until rs.EOF
					
								m_id = INT(rs("id"))
						
							rs.MoveNext 
							loop	
				
						end if
												
					end if
				
					rs.close
					SET rs = nothing
					SET cmd = nothing
				
					if m_id = 0 then
											
						m_message = "Package already exists as described by the 'Sco Container Name'."
						m_exists = false
										
					elseif NOT ISNULL(m_id) and m_id > 0 then
												
						dim xrefId:xrefId = 0
																		
						sql = "set nocount off; "
						sql = sql & " if NOT EXISTS(select 1 from dbo.Scorm_Sco_CourseXref where sco_objectId = ?)"
						sql = sql & " begin"
						sql = sql & " Insert dbo.Scorm_Sco_CourseXref(sco_objectId, courseId, sco_title, createDate)"
						sql = sql & " Values(?,?,?,getdate()); "
						sql = sql & " Select convert(int,@@IDENTITY) As [id];"
						sql = sql & " end"
						sql = sql & " else"
						sql = sql & " begin"
						sql = sql & " Select convert(int,0) as [id]"
						sql = sql & " from dbo.Scorm_Sco_CourseXref x"
						sql = sql & " where x.sco_objectId = ?;"
						sql = sql & " end"					
						sql = sql & " set nocount on;" 
					
						'response.write sql
						Set cmd = Server.CreateObject("ADODB.COMMAND")
						
						WITH cmd
							.Prepared = true
							.CommandType = adCmdText
							.ActiveConnection = conn
							.CommandText = sql			  
						END WITH
					
						cmd.Parameters.Append cmd.CreateParameter("@scoobjectid", adVarChar, adParamInput, 255, m_scoObjectId)
						cmd.Parameters.Append cmd.CreateParameter("@scoobjectid2", adVarChar, adParamInput, 255, m_scoObjectId)	
						cmd.Parameters.Append cmd.CreateParameter("@courseid", adInteger, adParamInput, 0, m_courseId)					
						cmd.Parameters.Append cmd.CreateParameter("@title", adVarChar, adParamInput, 255, m_title)
						cmd.Parameters.Append cmd.CreateParameter("@scoobjectid3", adVarChar, adParamInput, 255, m_scoObjectId)
						
						set rs = cmd.Execute
					
						If err.number > 0 then
							m_message = "Error.\nFailed to associate sco package to course."
							m_exists = false
						else
							m_exists = true		

							if NOT (rs.bof and rs.eof) then
					
								do until rs.EOF
						
									xrefId = INT(rs("id"))
							
								rs.MoveNext 
								loop	
					
							end if
													
						end if
					
						rs.close
						SET rs = nothing
						
						if xrefId = 0 then
						
							if LEN(title) > 0 then
								m_message = "Package already assigned to course " & title & "."
							else
								m_message = "Package already assigned to a course."
							end if
							
							m_exists = false
						else
							m_exists = true
						end if
					
					end if
					
					
					
				end if
				
			else 'Operation edit
			
				if m_id = 0 then
				
					m_message = "An id must be provided to update the record."
					m_exists = false
				else
				
					sql = "set nocount off; "
					sql = sql & " Update Scorm_Sco set title = ?, scoEntryFile = ?, mastery_score = ?"
					sql = sql & " where Id = ?;"
					sql = sql & " set nocount on;" 
			
					WITH cmd
						.Prepared = true
						.CommandType = adCmdText
						.ActiveConnection = conn
						.CommandText = sql			  
					END WITH
								
					cmd.Parameters.Append cmd.CreateParameter("@title", adVarChar, adParamInput, 255, m_title)
					cmd.Parameters.Append cmd.CreateParameter("@scoEntryFile", adVarChar, adParamInput, 255, m_scoEntryFile)
					cmd.Parameters.Append cmd.CreateParameter("@masteryscore", adVarChar, adParamInput, 10, m_masteryScore)
					cmd.Parameters.Append cmd.CreateParameter("@id", adInteger, adParamInput, 0, m_id)
							
					cmd.Execute()
					set cmd = nothing
					
					If err.number > 0 then
						m_message = "Error.\nFailed to update sco package."
						m_exists = false
					else
						m_exists = true			
					end if
					
					sql = "set nocount off; "
					sql = sql & " Update Scorm_Sco_CourseXref set sco_title = ?"
					sql = sql & " Where Sco_ObjectId = (Select scoObjectId From dbo.Scorm_Sco Where id = ?);"
					sql = sql & " set nocount on;" 
			
					Set cmd = Server.CreateObject("ADODB.COMMAND")
			
					WITH cmd
						.Prepared = true
						.CommandType = adCmdText
						.ActiveConnection = conn
						.CommandText = sql			  
					END WITH
								
					cmd.Parameters.Append cmd.CreateParameter("@title", adVarChar, adParamInput, 255, m_title)
					cmd.Parameters.Append cmd.CreateParameter("@id", adInteger, adParamInput, 0, m_id)
							
					cmd.Execute()
							
					
					If err.number > 0 then
						m_message = "Error.\nFailed to update package."
						m_exists = false
					else
						m_exists = true			
					end if
					
				end if
						
			end if				
					
			
			
						
			Save = m_exists
			
			conn.Close
			
			set cmd = nothing
			set conn = nothing
		
		End Function
		
		Function Delete(id)
		
			dim sql, conn, cmd
			
			m_exists = false
		
			Set conn = Server.CreateObject("ADODB.CONNECTION")
			Set cmd = Server.CreateObject("ADODB.COMMAND")
			
			conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & scDataBase & "; UID=" & scSQLUser & "; PWD=" & scSQLpwd'
			
			WITH cmd
				.Prepared = true
				.CommandType = adCmdText
				.ActiveConnection = conn
						  
			END WITH
							
			sql = "Delete dbo.Scorm_Sco_CourseXref Where Sco_ObjectId = "
			sql = sql & " (select scoObjectId From Scorm_sco where id = " & id & ");"
					
			cmd.CommandText = sql
			cmd.Execute()
			
			If err.number > 0 then
				m_message = "Error.\nFailed to delete sco package to course reference."					
			else
				m_exists = true	
				
				
				sql = "Delete dbo.Scorm_sco Where id = " & id
				
				WITH cmd
					.Prepared = true
					.CommandType = adCmdText
					.ActiveConnection = conn
					.CommandText = sql			  
				END WITH
				
				cmd.Execute()
				
				If err.number > 0 then
					m_message = "Error.\nFailed to delete sco package."
											
				end if
					
							
			end if
														
			Delete = m_exists
			
			conn.Close
			
			set cmd = nothing
			set conn = nothing
		
		End Function
		
		Function LoadRecordById(id)
				
			dim sql, conn, rs
			
			
			m_exists = false
		
			Set conn = Server.CreateObject("ADODB.CONNECTION")
			Set rs = Server.CreateObject("ADODB.recordset")
			rs.CursorLocation = adUseClient		
			
			conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & scDataBase & "; UID=" & scSQLUser & "; PWD=" & scSQLpwd'
			
			sql = " Select s.id,x.courseid,s.scoObjectId,s.title,s.path,s.pathtype,s.scoContainerName,s.scoEntryFile,s.launch_data,s.mastery_score,"
			sql = sql & "s.player_height_pxl,s.player_width_pxl,s.createDate"
			sql = sql & " From Scorm_sco s"
			sql = sql & " join Scorm_Sco_CourseXref x on s.scoObjectId=x.sco_objectid"
			sql = sql & " Where x.courseId = " & id &";"
					
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
		
			conn.Close
			
			set rs = nothing
			set conn = nothing
			
		End Function
		
		Function LoadFromRecord(record)
	
			m_id = IIF(ISNULL(record("id")),0,record("id")) 
			m_scoobjectid = IIF(ISNULL(record("scoobjectid")),"",record("scoobjectid"))
			m_courseid = IIF(ISNULL(record("courseid")),0,record("courseid"))
			m_title = IIF(ISNULL(record("title")),"",record("title"))
			m_path = IIF(ISNULL(record("path")),"",record("path"))
			m_pathtype = IIF(ISNULL(record("pathtype")),0,record("pathtype"))
			m_scocontainername = IIF(ISNULL(record("scocontainername")),"",record("scocontainername"))
			m_scoentryfile = IIF(ISNULL(record("scoentryfile")),"",record("scoentryfile"))
			m_launchdata = IIF(ISNULL(record("launch_data")),0,record("launch_data"))
			m_masteryscore = IIF(ISNULL(record("mastery_score")),0,record("mastery_score"))
			m_createDate = IIF(ISNULL(record("createDate")),"",record("createDate"))
				
		End Function
		
		Sub LoadFromObject(data)
		
			m_id = IIF((ISNULL(data.id) OR LEN(data.id) = 0),0,data.id)
			m_courseid = IIF((ISNULL(data.courseid) OR LEN(data.courseid) = 0),0,data.courseid)
			m_title = IIF((ISNULL(data.title) OR LEN(data.title) = 0), "", Base64Decode(data.title))
			m_scocontainername = IIF((ISNULL(data.scocontainername) OR LEN(data.scocontainername) = 0), "", Base64Decode(data.scocontainername))
			m_scoentryfile = IIF((ISNULL(data.scoentryfile) OR LEN(data.scoentryfile) = 0),"",Base64Decode(data.scoentryfile))
			m_masteryscore = IIF((ISNULL(data.mastery_score) OR LEN(data.mastery_score) = 0),"",data.mastery_score)
			
			if LEN(m_title) > 255 then
				m_title = MID(m_title,1,255)
			end if
			if LEN(m_scocontainername) > 255 then
				m_scocontainername = MID(m_scocontainername,1,255)
			end if
			if LEN(m_scoentryfile) > 255 then
				m_scoentryfile = MID(m_scoentryfile,1,255)
			end if
			if LEN(m_masteryscore) > 10 then
				m_masteryscore = MID(m_masteryscore,1,10)
			end if
			
			m_scoObjectId =  m_scoContainerName
			
		End Sub
	
		Function ToJson()
		
			dim json:json = "{""exists"":""" &m_exists& ""","	
			json = json & """message"":""" &m_message& ""","
			json = json & """id"":""" &m_id& ""","			
			json = json & """scoobjectid"":""" &m_scoobjectid& ""","
			json = json & """courseid"":""" &m_courseid& ""","					
			json = json & """title"":""" &m_title& ""","
			json = json & """path"":""" &m_path& ""","
			json = json & """pathtype"":""" &m_pathtype& ""","
			json = json & """scocontainername"":""" &m_scocontainername& ""","
			json = json & """scoentryfile"":""" &m_scoentryfile& ""","
			json = json & """mastery_score"":""" &m_masteryscore& ""","
			json = json & """createdate"":""" &m_createDate& """"
					
			json = json & "}"
			
			ToJson = json
					
		End Function
		
		
	End Class

%>

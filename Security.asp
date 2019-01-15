<%'Security

'Uses Microsofts CapiCom.dll registered component, version 2.1.0.2

const CAPICOM_HASH_ALGORITHM_SHA_256 = 4
const CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS = 5



Class Authenticator

	dim sqlSecurityAccount	
	dim sqlSecurityAccountPassword
	dim varDatabase
	

	Private Sub Class_Initialize()
	
		sqlSecurityAccount = scSQLUser
		sqlSecurityAccountPassword = scSQLpwd
		varDatabase = scDatabase
		
	End Sub
	
	Sub Class_Terminate
        
    End Sub

	
	Sub SetDatabaseContext(dbType)
		
		Select Case dbType
		
			Case "G"
				sqlSecurityAccount = scSQLUser
				sqlSecurityAccountPassword = scSQLPwd
				varDatabase = scDatabase
			Case "SC"
				sqlSecurityAccount = legacySQLUser
				sqlSecurityAccountPassword = legacySQLPwd
				varDatabase = legacyDatabase
			Case "T"
				sqlSecurityAccount = trialSQLUser
				sqlSecurityAccountPassword = trialSQLPWD
				varDatabase = trialDatabase
		End Select
	
	End Sub
	
	Function EncryptMessage(clearText, ByRef hash)
		
		hash = HashString(clearText)		
		
		EncryptMessage = DoEncryptMessage(clearText,CAPICOM_HASH_ALGORITHM_SHA_256,CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS,hash)
	
	End Function
	
	Function DecryptMessage(codedText, hash)
	
		if len(codedText) > 0 and len(hash) > 0 then
			DecryptMessage = DoDecryptMessage(codedText, hash)
		else
			DecryptMessage = ""
		end if
		
	End Function
	
	Function HashString(clearText)
	
		Dim sSalt:sSalt=""
		dim sValue:sValue=""
		dim sHashLength:sHashLength=0
		sSalt = generateRandomString(51)
								
		dim HashData
		set HashData = CreateObject("CAPICOM.HashedData")
				
		With HashData
		  .Algorithm = CAPICOM_HASH_ALGORITHM_SHA_256 ' CAPICOM_HASH_ALGORITHM_SHA256  
		  .Hash clearText & sSalt		  
		  sValue = .Value
		End With
				
		HashString = sValue
	
	End Function
	
	Function DoEncryptMessage(clearText, Algorithm, KeyLength, hash) 

		Dim EncryptedData

		if NOT ISOBJECT(Session("CAPICOMEncryptedData")) then
			set Session("CAPICOMEncryptedData") = CreateObject("CAPICOM.EncryptedData") 
		End If
		
		Set EncryptedData = Session("CAPICOMEncryptedData") 'CreateObject("CAPICOM.EncryptedData") 
		With EncryptedData
			.Algorithm.Name = Algorithm 
			.Algorithm.KeyLength = KeyLength 
			.SetSecret hash 
			.Content = clearText 
		
		DoEncryptMessage =  .Encrypt 
		
		End WIth
		
		Set EncryptedData = Nothing 
		
	End Function 


	Function DoDecryptMessage(encrypted, hash )
	 
		dim decryptedContent : decryptedContent = ""
	 
		'On Error Resume Next
		
		if NOT ISOBJECT(Session("CAPICOMEncryptedData")) then
			set Session("CAPICOMEncryptedData") = CreateObject("CAPICOM.EncryptedData") 
		End If
	 			 
		Dim EncryptedData 
		Set EncryptedData = Session("CAPICOMEncryptedData") 'CreateObject("CAPICOM.EncryptedData") 
		EncryptedData.Algorithm.Name = CAPICOM_HASH_ALGORITHM_SHA_256 'Algorithm 
		EncryptedData.Algorithm.KeyLength = CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS 'KeyLength 
		EncryptedData.SetSecret hash 
		EncryptedData.Decrypt encrypted 
		decryptedContent = EncryptedData.Content
		
		' Error Handler
		If Err.Number <> 0 Then
		   ' Error Occurred / Trap it
		   On Error Goto 0 ' But don't let other errors hide!
			Set EncryptedData = Nothing 
			DoDecryptMessage = ""
		End If
		
		DoDecryptMessage = decryptedContent
		
		Set EncryptedData = Nothing 
		
		
		On Error Goto 0 ' Reset error handling.

		
	End Function

	Function generateRandomString(stringLength)

	    'Declare variables
		Dim sDefaultChars
		Dim iCounter
		Dim sMyRandomString
		Dim iPickedChar
		Dim iDefaultCharactersLength
		Dim iStringLength

		'Initialize variables
		sDefaultChars="abcdefghijklmnopqrstuvxyzABCDEFGHIJKLMNOPQRSTUVXYZ0123456789!@#$^&*"
		iStringLength=stringLength
		iDefaultCharactersLength = Len(sDefaultChars)

        'initialize the random number generator
		Randomize

		'Loop for the number of characters for string
		For iCounter = 1 To iStringLength

		'Next pick a number from 1 to length of character set
		iPickedChar = Int((iDefaultCharactersLength * Rnd) + 1)

		'Next pick a character from the character set using the random number iPickedChar and Mid function
		sMyRandomString = sMyRandomString & Mid(sDefaultChars,iPickedChar,1)

		Next
		generateRandomString = sMyRandomString

	End Function

	Function ValidatePassword(oldPassword, newPassword, newPasswordAgain, ByRef message)
	
		dim varOldPassword : varOldPassword = oldPassword
		dim varNewPassword : varNewPassword = newPassword
		dim varNewPasswordVerify : varNewPasswordVerify = newPasswordAgain
		
		dim isValid : isValid = false
		dim passwordHasUpperCaseChar : passwordHasUpperCaseChar = false
		dim passwordHasMinimumLength : passwordHasMinimumLength = false
		dim passwordHasSpecialChar : passwordHasSpecialChar = false
		dim passwordVerifyMatch : passwordVerifyMatch = false
		dim passwordHasNumericChar : passwordHasNumericChar = false
		dim oldPasswordIsEmpty : oldPasswordIsEmpty = false
		dim newPasswordIsUnique : newPasswordIsUnique = false
		
		if LEN(TRIM(varOldPassword)) > 0 then
			oldPasswordIsEmpty = true
		else		
			message = "Password Update Failed!<br/><br/>You must provide your original password."		
		end if
		
		if StrComp(varNewPassword, varOldPassword, 0) <> 0 then
			newPasswordIsUnique = true
		else		
			if LEN(message) = 0 then
				message = "Password Update Failed!<br/><br/>Your new password must differ from your current password."
			end if
		end if
		
		if StrComp(varNewPassword, varNewPasswordVerify, 0) = 0 then
			passwordVerifyMatch = true
		else		
			if LEN(message) = 0 then
				message = "Password Update Failed!<br/><br/>New passwords do not match."
			end if
		end if
		
		if LEN(varNewPassword) > 11 then
			passwordHasMinimumLength = true
		else
			if LEN(message) = 0 then
				message = "Password Update Failed!<br/><br/>Your new password must be at least 12 characters in length."
			end if
			
		end if
		
		Dim regEx
		set regEx = New RegExp
		
		With regEx
			.Pattern = "[A-Z]"
			.Global = true
			.IgnoreCase = false
		End With
		
		if regEx.Test(Replace(varNewPassword," ","")) then
			passwordHasUpperCaseChar = true
		else 
			if LEN(message) = 0 then
				message = "Password Update Failed!<br/><br/>Your new password must have at least one upper case letter."
			end if
		end if	
			
		regEx.Pattern = "[!@#$^*]"
		if regEx.Test(Replace(varNewPassword," ","")) then	
			passwordHasSpecialChar = true
		else 
			if LEN(message) = 0 then
				message = "Password Update Failed!<br/><br/>Your new password must have at least one special character."		
			end if
		end if
		
		regEx.Pattern = "[0-9]"
		if regEx.Test(Replace(varNewPassword," ","")) then	
			passwordHasNumericChar = true
		else 
			if LEN(message) = 0 then
				message = "Password Update Failed!<br/><br/>Your new password must have at least one numeric character."				
			end if
		end if
		
		if passwordVerifyMatch = false _
			OR passwordHasMinimumLength = false _
				OR passwordHasUpperCaseChar = false _
					OR passwordHasSpecialChar = false _
						OR passwordHasNumericChar = false _
							OR oldPasswordIsEmpty = false then
					
			ValidatePassword = false
		else
			ValidatePassword = true
		end if		
	
	End Function
	
	public Function GetUserRecord(userid, password, ByRef exists, ByRef isPasswordAuthenticated, ByRef isEncrypted, ByRef message)
	
		dim sql, conn, cmd, rs, varMessage
		dim varMemberId : varMemberId = 0
		set objSecurityAccount = new SecurityAccount
		exists = false
		isEncrypted = false		
		isPasswordAuthenticated = false
		
		dim varDbHash, varEncryptedDBPassword, varExists, varDecryptedText, varRecordsAffected
			
		varEncryptedDBPassword = objSecurityAccount.GetCredentials(userid, varExists, varDbHash, varMessage)
				
		if false = varExists then
		
			message = "Either the User Name or Password is invalid<BR> or you have not registered yet."
		else
			if LEN(varDbHash) > 0 then 'Verify encrypted password
			
				isEncrypted = true
				varDecryptedText = DecryptMessage(varEncryptedDBPassword, varDbHash)
								
				if StrComp(password, varDecryptedText, 1) = 0 then		
					isPasswordAuthenticated = true
					
					message = "User authenticated."
				else
					message = "User credentials failed to authenticate."
				end if
			else
				'Verify unencrypted password - user does not have an account security record
				isPasswordAuthenticated = IsCredentialAuthenticated(userid, password, message)
				
			end if
			
		
			Set conn = Server.CreateObject("ADODB.CONNECTION")
			Set cmd = Server.CreateObject("ADODB.Command")
			Set rs = Server.CreateObject("ADODB.recordset")
				
						
			conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & varDatabase & "; UID=" & sqlSecurityAccount & "; PWD=" & sqlSecurityAccountPassword'
			
			sql = "set nocount off; select memberid "
			sql = sql & " from dbo.members"
			sql = sql & " Where userid = ?" 			
			sql = sql & " set nocount on;"
			
			With cmd
				.Prepared = true
				.ActiveConnection = conn
				.CommandText = sql
				.CommandType = adCmdText
			end With		
					
			
			cmd.Parameters.Append cmd.CreateParameter("@userId", adVarChar, adParamInput, 125, userid)		
			
			set rs = cmd.Execute(varRecordsAffected)
			If err.number > 0 or varRecordsAffected = 0 then
				message = "User account failed to load.\n" & err.Description 
			else
				exists = true
			end if
			if NOT (rs.bof and rs.eof) then
			
				'Exists = true
			
				do until rs.EOF
				
					varMemberId = rs("memberid")
				
				rs.MoveNext 
				loop	
				
			end if
			
			conn.Close()
			
			set rs = nothing
			set cmd = nothing
			set conn = nothing
		
		end if
	
		GetUserRecord = varMemberId
	
	End Function

	Function IsCredentialAuthenticated(userid, password, ByRef message)
	
		dim sql, conn, cmd, rs, exists, varRecordsAffected, varMemberId
		varMemberId = 0
		exists = false
	
		Set conn = Server.CreateObject("ADODB.CONNECTION")
		Set cmd = Server.CreateObject("ADODB.Command")
		Set rs = Server.CreateObject("ADODB.recordset")
		
			conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & varDatabase & "; UID=" & sqlSecurityAccount & "; PWD=" & sqlSecurityAccountPassword'
			
			sql = "set nocount off; select memberid "
			sql = sql & " from dbo.members"
			sql = sql & " Where userid = ?" 
			sql = sql & " and password = ?;" 
			sql = sql & " set nocount on;"
			
			With cmd
				.Prepared = true
				.ActiveConnection = conn
				.CommandText = sql
				.CommandType = adCmdText
			end With		
					
			
			cmd.Parameters.Append cmd.CreateParameter("@userId", adVarChar, adParamInput, 125, userid)		
			cmd.Parameters.Append cmd.CreateParameter("@password", adVarChar, adParamInput, 500, password)

			set rs = cmd.Execute(varRecordsAffected)
			
			if err.number > 0 then 'or varRecordsAffected = 0 then
				message = "User account failed to load.\n" & err.Description & "(" & err.number & ")"
			'else
				'exists = true
			end if
			
			if NOT (rs.bof and rs.eof) then
						
				do until rs.EOF
				
					varMemberId = rs("memberid")
					if varMemberId > 0 then					
						exists = true
					end if
				
				rs.MoveNext 
				loop	
				
			end if
			
						
			conn.Close()
			
			set rs = nothing
			set cmd = nothing
			set conn = nothing
	
			IsCredentialAuthenticated = exists
	
	End Function
	
end class

Class SecurityAccount

	dim sqlSecurityAccount
	dim sqlSecurityAccountPassword
	dim varDatabase
	

	private m_memberid
	private m_ForcePasswordReset
	private m_LastForcePasswordResetDate
	private m_Salt
	private m_SaltedHash
	private m_AccountIsLocked
	private m_AccountLockDate
	private m_NumberOfLoginAttempts
	private m_createDate
	private m_exists
	private m_password
	private m_issaltencrypted
	
	Public Property Get Exists()
		Exists = m_exists
	End Property
	Public Property Let Exists(value)
		m_exists = value
	End Property
	Public Property Get MemberId()
		MemberId = m_memberid
	End Property
	Public Property Let MemberId(value)
		m_memberid = value
	End Property
	
	Public Property Get ForcePasswordReset()
		ForcePasswordReset = m_ForcePasswordReset
	End Property
	Public Property Let ForcePasswordReset(value)
		m_ForcePasswordReset = value
	End Property
	
	Public Property Get LastForcePasswordResetDate()
		LastForcePasswordResetDate = m_LastForcePasswordResetDate
	End Property
	Public Property Let LastForcePasswordResetDate(value)
		m_LastForcePasswordResetDate = value
	End Property
	
	Public Property Get Salt()
		Salt = m_Salt
	End Property
	Public Property Let Salt(value)
		m_Salt = value
	End Property
	
	Public Property Get SaltedHash()
		SaltedHash = m_SaltedHash
	End Property
	Public Property Let SaltedHash(value)
		m_SaltedHash = value
	End Property
	
	Public Property Get AccountIsLocked()
		AccountIsLocked = m_AccountIsLocked
	End Property
	Public Property Let AccountIsLocked(value)
		m_AccountIsLocked = value
	End Property
	
	Public Property Get AccountLockDate()
		AccountLockDate = m_AccountLockDate
	End Property
	Public Property Let AccountLockDate(value)
		m_AccountLockDate = value
	End Property
	
	Public Property Get NumberOfLoginAttempts()
		NumberOfLoginAttempts = m_NumberOfLoginAttempts
	End Property
	Public Property Let NumberOfLoginAttempts(value)
		m_NumberOfLoginAttempts = value
	End Property
	
	Public Property Get CreateDate()
		CreateDate = m_createDate
	End Property
	Public Property Let CreateDate(value)
		m_createDate = value
	End Property
		
	Public Property Get Password()
		Password = m_password
	End Property
	Public Property Let Password(value)
		m_password = value
	End Property
	
	Public Property Get IsSaltEncrypted()
		IsSaltEncrypted = m_issaltencrypted
	End Property
	Public Property Let IsSaltEncrypted(value)
		m_issaltencrypted = value
	End Property
	
	
	
	Private Sub Class_Initialize()
			
		sqlSecurityAccount = scSQLUser
		sqlSecurityAccountPassword = scSQLpwd
		varDatabase = scDatabase
	
		m_memberid = 0
		m_ForcePasswordReset = false		
		m_Salt = ""
		m_SaltedHash = ""
		m_AccountIsLocked = false 		
		m_NumberOfLoginAttempts = 0
		m_exists = false
		
	End Sub
	
	Sub Class_Terminate
        
    End Sub

	
	Sub SetDatabaseContext(dbType)
		'response.write "dbtype: " & dbType & "<br/>"
		Select Case dbType
		
			Case "G"
				sqlSecurityAccount = scSQLUser
				sqlSecurityAccountPassword = scSQLPwd
				varDatabase = scDatabase
			Case "SC"
				sqlSecurityAccount = legacySQLUser
				sqlSecurityAccountPassword = legacySQLPwd
				varDatabase = legacyDatabase
			Case "T"
				sqlSecurityAccount = trialSQLUser
				sqlSecurityAccountPassword = trialSQLPWD
				varDatabase = trialDatabase
		End Select
	
	End Sub
	
	
	public Sub SetForcePasswordReset(memberid)
	
		dim sql, conn, rs
		
		Set conn = Server.CreateObject("ADODB.CONNECTION")			
		conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & varDatabase & "; UID=" & sqlSecurityAccount & "; PWD=" & sqlSecurityAccountPassword'
		
		sql = "Update dbo.AccountSecurity Set ForcePasswordReset = 1, LastForcePasswordResetDate = GETDATE() "
		sql = sql & " Where memberId = " & memberid
	
	
		Set rs = conn.Execute(sql)
				
		conn.Close()		
		set conn = nothing
	
	End Sub
	
	public Sub ResetForcePasswordReset()
	
		dim sql, conn, rs
		
		Set conn = Server.CreateObject("ADODB.CONNECTION")			
		conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & varDatabase & "; UID=" & sqlSecurityAccount & "; PWD=" & sqlSecurityAccountPassword'
		
		sql = "Update dbo.AccountSecurity Set ForcePasswordReset = 0, LastForcePasswordResetDate = GETDATE() "
		sql = sql & " Where memberId = " & m_memberid
	
	
		Set rs = conn.Execute(sql)
				
		conn.Close()		
		set conn = nothing
	
	End Sub
	
	public Function CreateSecurityAccount(memberid, ByRef message)
	
		dim sql, conn, cmd, exists
		exists = false
		
		Set conn = Server.CreateObject("ADODB.CONNECTION")
		Set cmd = Server.CreateObject("ADODB.COMMAND")
		
		conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & varDatabase & "; UID=" & sqlSecurityAccount & "; PWD=" & sqlSecurityAccountPassword'
		
		sql = "set nocount off; "
		sql = sql & " if NOT EXISTS(Select 1 from AccountSecurity Where memberid = ? ) begin "
		sql = sql & " INSERT INTO AccountSecurity(MemberId, ForcePasswordReset, LastForcePasswordResetDate) Values(?,?,?); "
		sql = sql & " End set nocount on;"
			
		dim varRecordsAffected : varRecordsAffected = 0
		WITH cmd
			'.Prepared = true
			.CommandType = adCmdText
			.ActiveConnection = conn
			.CommandText = sql			  
		END WITH	
			
		cmd.Parameters.Append cmd.CreateParameter("@p1", adInteger, adParamInput, , memberid)	
		cmd.Parameters.Append cmd.CreateParameter("@p2", adInteger, adParamInput, , memberid)		
		cmd.Parameters.Append cmd.CreateParameter("@p3", adBoolean, adParamInput, , 1)
        cmd.Parameters.Append cmd.CreateParameter("@p3", adDBTimeStamp, adParamInput, , Now())
					
		cmd.Execute varRecordsAffected
		
		If err.number > 0 or varRecordsAffected = 0 then
			message = "Security account failed to create."
		else
			exists = true
		end if
		
		conn.Close()		
			
		set cmd = nothing	
		set conn = nothing
	
		CreateSecurityAccount = exists
	
	end Function
	
	public Function GetCredentials(userid, ByRef exists, ByRef hash, ByRef message)
	
		dim sql, conn, cmd, rs, varPassword
		exists = false
		
		Set conn = Server.CreateObject("ADODB.CONNECTION")
		Set cmd = Server.CreateObject("ADODB.COMMAND")
		Set rs = Server.CreateObject("ADODB.Recordset")
		
		conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & varDatabase & "; UID=" & sqlSecurityAccount & "; PWD=" & sqlSecurityAccountPassword'
				
		
		sql = "select m.memberid, m.password, s.ForcePasswordReset, s.LastForcePasswordResetDate, s.Salt, "
		sql = sql & " Case When s.Salt = 'E' Then CONVERT(nvarchar(128), DecryptByPassphrase(m.userid, s.SaltedHashBin, 1, CONVERT(varbinary, m.memberid))) Else s.SaltedHash End as SaltedHash, "
		sql = sql & " s.AccountIsLocked, s.AccountLockDate, s.NumberOfLoginAttempts, s.createDate "
		sql = sql & " from dbo.members m"
		sql = sql & " left join dbo.AccountSecurity s on m.memberid=s.memberid"
		sql = sql & " Where m.userid = ?"
	
				
				
		WITH cmd
			.Prepared = true
			.CommandType = adCmdText
			.ActiveConnection = conn
			.CommandText = sql			  
		END WITH	
		
		cmd.Parameters.Append cmd.CreateParameter("@p1", adVarChar, adParamInput, 125, userid)
				
		set rs = cmd.Execute() 
				
		if err.number > 0 then
		
			message = err.description
		else
			if NOT (rs.bof and rs.eof) then
			
				do until rs.EOF
					
					exists = true
					m_memberId = rs("memberid")
					varPassword = rs("password")
					hash = rs("saltedhash")
				
					rs.MoveNext 
				loop	
			
			end if
		end if
			
		conn.Close()
		
		set rs = nothing
		set cmd = nothing	
		set conn = nothing
	
		GetCredentials = varPassword
		
	end Function
	
	public Function UpdateCredentials(password, hash, ByRef message)
	
		dim sql, conn, cmd, cmd2, cmd3, cmd4, rs, exists, userid
		exists = false
		
		Set conn = Server.CreateObject("ADODB.CONNECTION")
		Set cmd = Server.CreateObject("ADODB.COMMAND")
		Set cmd2 = Server.CreateObject("ADODB.COMMAND")
		Set cmd3 = Server.CreateObject("ADODB.COMMAND")
		Set cmd4 = Server.CreateObject("ADODB.COMMAND")
		
		conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & varDatabase & "; UID=" & sqlSecurityAccount & "; PWD=" & sqlSecurityAccountPassword'
		
		sql = "set nocount off; Update dbo.Members Set Password = ?" 
		sql = sql & " Where memberId = ?; set nocount on;" 
			
		dim varRecordsAffected : varRecordsAffected = 0
		WITH cmd
			.Prepared = true
			.CommandType = adCmdText
			.ActiveConnection = conn
			.CommandText = sql			  
		END WITH
  				
		
		cmd.Parameters.Append cmd.CreateParameter("@password", adVarChar, adParamInput, 500, password)		
		cmd.Parameters.Append cmd.CreateParameter("@memberId", adInteger, adParamInput, 0, m_memberid)		
				
		cmd.Execute varRecordsAffected ', , adExecuteNoRecords 
		If err.number > 0 or varRecordsAffected = 0 then
			message = "Error.\nPassword failed to update."
		end if
			
		set cmd = nothing	
		
		
		
		sql = "Select userid from dbo.Members where memberid = ?"
		
		
		WITH cmd3
			.Prepared = true
			.CommandType = adCmdText
			.ActiveConnection = conn
			.CommandText = sql			  
		END WITH	
		
		cmd3.Parameters.Append cmd3.CreateParameter("@p1", adInteger, adParamInput, 0, m_memberid)
				
		set rs = cmd3.Execute() 
				
		if err.number > 0 then
		
			message = err.description
		else
			if NOT (rs.bof and rs.eof) then
			
				do until rs.EOF
					
					exists = true
					userid = rs("userid")
					
					rs.MoveNext 
				loop	
			
			end if
		end if
		
		set cmd3 = nothing
		
		
				
		if  NOT ISNULL(varRecordsAffected) AND varRecordsAffected > 0 then
					
			dim varEncryptedHash:varEncryptedHash=""
			
			sql = "Set NOCOUNT ON; Declare @encryptedData varbinary(256) "
			sql = sql & " DECLARE @hash nvarchar(128) "
			sql = sql & " Declare @memberId int "
			sql = sql & " set @memberId = " & m_memberid
			sql = sql & " set @hash = '" & hash & "'"
			sql = sql & " select  @encryptedData=EncryptByPassPhrase(userid, @hash, 1, CONVERT(varbinary, memberid)) "
			sql = sql & " from members "
			sql = sql & " WHERE memberid = " & m_memberId
			sql = sql & " Update AccountSecurity Set SaltedHashBin = @encryptedData, Salt='E' where memberid  = " & m_memberid

		
			WITH cmd4
				.Prepared = true
				.CommandType = adCmdText
				.ActiveConnection = conn
				.CommandText = sql			  
			END WITH	
							
			varRecordsAffected = 0
			cmd4.Execute varRecordsAffected 
			If err.number > 0 or varRecordsAffected = 0 and LEN(message) = 0 then
				message = "Error.\nHash failed to update." & " " & err.Description
			else
				exists = true
			end if
			exists = true
			
			set cmd4 = nothing			
									
						
		end if
			
		conn.Close()		
		
		set cmd = nothing
		set cmd2 = nothing
		set cmd3 = nothing
		set cmd4 = nothing
		set conn = nothing
		
		UpdateCredentials = exists
	
	End Function
	
	public Sub LoadRecord(memberId)
	
		dim sql, conn, rs
		
		Set conn = Server.CreateObject("ADODB.CONNECTION")
		set rs = Server.CreateObject("adodb.recordset")
			
		
		conn.Open "DRIVER={SQL Server}; Server=" & scSQLServer & "; Database=" & varDatabase & "; UID=" & sqlSecurityAccount & "; PWD=" & sqlSecurityAccountPassword'
		
		
		sql = "select m.memberid, m.password, s.ForcePasswordReset, s.LastForcePasswordResetDate, s.Salt, "
		sql = sql & " Case When s.Salt = 'E' Then CONVERT(nvarchar(128), DecryptByPassphrase(m.userid, s.SaltedHashBin, 1, CONVERT(varbinary, m.memberid))) Else s.SaltedHash End as SaltedHash, "
		sql = sql & " s.AccountIsLocked, s.AccountLockDate, s.NumberOfLoginAttempts, s.createDate "
		sql = sql & " from dbo.members m"
		sql = sql & " join dbo.AccountSecurity s on m.memberid=s.memberid"
		sql = sql & " Where m.memberId = " & memberId
	
		Set rs = conn.Execute(sql)
		if NOT (rs.bof and rs.eof) then
		
			Exists = true
		
			do until rs.EOF
			
				LoadFromRecord(rs)			
			
			rs.MoveNext 
			loop	
			
		end if
		
		conn.Close()
		
		set rs = nothing
		set conn = nothing
	
	End Sub
	
	Public Sub LoadFromRecord(record)
			
		MemberId = record("memberid")
		ForcePasswordReset = record("ForcePasswordReset")
		LastForcePasswordResetDate = record("LastForcePasswordResetDate")
		Salt = record("Salt")
		SaltedHash = record("SaltedHash")
		AccountIsLocked = record("AccountIsLocked")
		AccountLockDate = record("AccountLockDate")
		NumberOfLoginAttempts = record("NumberOfLoginAttempts")
		CreateDate = record("createDate")
		Password = record("password")
		
		if NOT ISNULL(record("Salt")) AND record("Salt") = "E" then
			IsSaltEncrypted = true
		else
			IsSaltEncrypted = false
		end if
		
	End Sub

End class


%>

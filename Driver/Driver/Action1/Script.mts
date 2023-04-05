mrowcount=datatable.GetSheet("Action1").GetRowCount
For i = 1 To mrowcount Step 1
	datatable.SetCurrentRow(i)
	If datatable("ModuleExe","Action1")="Y" Then
		Modid=datatable("ModuleID","Action1")
		'msgbox Modid
		
		tcrowcount=datatable.GetSheet("Action2").GetRowCount
		For j = 1 To tcrowcount Step 1
		
			datatable.SetCurrentRow(j)
			If Modid=datatable("ModuleID","Action2") and datatable("TestCaseExe","Action2")="Y" Then
			testcaseid=datatable("TestCaseId","Action2")
			'msgbox testcaseid
			
		tsrowcount=datatable.AddSheet("Action3").GetRowCount
		For k = 1 To tsrowcount Step 1
			datatable.SetCurrentRow(k)
			If testcaseid=datatable("TestCaseID","Action3")   Then
				Keyword=datatable("Keyword","Action3")
				'msgbox Keyword
				
				Select Case Keyword
					
			  					
			
				 Case "lh" 
				Call Launch()
				case "ln"
				Call Login("john","hp")
				Case "ce"	
				call Close()
				Case "nr"
				Call NewOrder()
				
				Case "lnd"
				
				drowcount=datatable.GetSheet("Action4").GetRowCount
					For l = 1 To drocount Step 1
					datatable.SetCurrentRow(l)
					agentname=datatble("username","Action4")	
					password=datatable("password","Action4")
					Call Launch
					Call Login(agentname,password)	
					Call Close
					
					Next
		
				
				
				
				
				End Select
				
				
				
			End If
			
		Next

			End If
			
		Next
	End If
Next

		

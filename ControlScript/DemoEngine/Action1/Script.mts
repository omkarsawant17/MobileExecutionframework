Dim arr_comname()

LoadFunctionLibrary "C:\MobileExecution\Libraries\Keyword.qfl"
Call fnc_DefineConfig()
LoadFunctionLibrary g_str_lib_ScriptFunction

    
    
Set client = DotNetFactory.CreateInstance("experitestClient.Client", "C:\\Program Files\\Experitest\\SeeTest\\clients\\C#\\imageClient.dll")
client.Connect "127.0.0.1",8889
Client.SetSpeed("Fast")
'
str_devicenames=Client.GetConnectedDevices
arr_devicenames=split(str_devicenames,chr(10))
For int_i = 0 To ubound(arr_devicenames) Step 1
	If trim(arr_devicenames(int_i))<>"" Then	
		client.SetDevice arr_devicenames(int_i)	
'		Client.openDevice()
		ReDim arr_comname(0)
		int_arr_com=0
		int_incr=0
		
					intTc_id="42"
'				    intTc_id=QCUtil.CurrentTest.ID
					    DbQuery = "SELECT * from [TCDetails$] where TestCaseNo ='" & intTc_id &"'"
					    Set DbModrec= Fnc_DbConnection(g_str_Controller,DbQuery)
					   
					    
					    For str_step = 0 To DbModrec.Fields.Count-1  Step 1
						    str_colname= DbModrec.Fields(str_step).Name 
						    str_compname= DbModrec.Fields(str_step).Value
						    If isnull (str_compname)=False and str_step>2 Then
						    str_tcid= DbModrec.Fields("TestCaseNo").Value
						    	DbQuery = "SELECT * from [" & str_compname & "$]"
						    	Set DbSteprec= Fnc_DbConnection( g_str_Components,DbQuery)
						    	If ubound(Filter(arr_comname,str_compname))>=0 Then
						    		int_incr=int_incr+1
						    	
						    	End If
						    	ReDim Preserve arr_comname(int_arr_com)
						    	arr_comname(int_arr_com)=str_compname
						    	int_arr_com=int_arr_com+1
								
					    			For comitr = 0 To DbSteprec.RecordCount-1  Step 1
					    				str_step_desc=DbSteprec.Fields("Step Description")
					    				str_step_objName=DbSteprec.Fields("ObjectName")
					    				str_step_actiocode=DbSteprec.Fields("ActionCode")
					    				str_step_data=DbSteprec.Fields(str_tcid)
					    				If ISnull(str_step_data) Then
					    					str_step_data=""
					    				End If
					    				If instr(1,str_step_data,"$$")>0 Then
					    					arr_step_data=split(str_step_data,"$$")
					    					str_step_data=arr_step_data(int_incr)
					    				End If
'					    				
					    				If trim(lcase(str_step_data))<>"na" Then
					    					
					    				
							    				Select Case ucase(str_step_actiocode)
							    					Case "INVOKE"
							    						client.applicationClearData str_step_data
							    						client.Launch str_step_data, true, true
							    					Case "CLICK"
							    						Call fnc_Click(str_step_objName)
							    					Case "LONG CLICK"
							    					
							    					Case "SET"
							    						If Trim(Left(str_step_data, 1)) = "#" And Trim(Right(str_step_data, 1)) = "#" Then
			                                                 str_step_data = Trim(Replace(str_step_data, "#", ""))
			                                                 str_step_data = RetrieveData(str_step_data)
			                                             End if 
							    						 Call fnc_Inputtext(str_step_objName,str_step_data)
							    					Case "SWIPE"
							    					
							    					Case "VERIFY"	
														arr_verifydata=split(str_step_data,"|")	
														If Trim(Left(arr_verifydata(0), 1)) = "#" And Trim(Right(arr_verifydata(0), 1)) = "#" Then
			                                                 arr_verifydata(0) = Trim(Replace(arr_verifydata(0), "#", ""))
			                                                 arr_verifydata(0) = RetrieveData(arr_verifydata(0))
			                                             End if 														
							    						Call fnc_VerifyElement(str_step_objName,arr_verifydata(0),arr_verifydata(1),arr_verifydata(2))
							    					Case "STORE"
							    						str_fetchval=fnc_FetchValue(str_step_objName,"value")
'							    						arr_storedata=split(str_step_data,"|")	
							    						Call Fnc_StoreData(str_step_data, str_fetchval)
							    					Case "WAIT"
							    						WAIT str_step_data
							    					Case "CLOSE"
							    						client.ApplicationClose str_step_data
							    						
							    				End Select
					    				End If
					    				DbSteprec.MoveNext
					    			Next
						    End If
					    Next
	End If
Next


Sub Report()
Dim logLine, outFile, status, errorMessage
logLine = client.GetResultValue ("logLine")
outFile = client.GetResultValue("outFile")
status = client.GetResultValue("status")
If StrComp (status, "True") = 0 then
	Reporter.ReportEvent micPass, logLine, "", outFile
Else 
	errorMessage = client.GetResultValue("errorMessage")
	Reporter.ReportEvent micFail, logLine, errorMessage, outFile
End If
End Sub

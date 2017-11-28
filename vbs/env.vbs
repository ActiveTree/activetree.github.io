'--------定义设置系统环境变量的方法---------
Set pSysEnv = CreateObject("WScript.Shell").Environment("System")

Function IsMatch(Str, Patrn)
  Set r = new RegExp
  r.Pattern = Patrn
  IsMatch = r.test(Str)
End Function


Sub SetEnv(pPath, pValue)
    Dim ExistValueOfPath
    If pValue <> "" 
		Then ExistValueOfPath = pSysEnv(pPath)
	 
	If Right(pValue, 1) = "\" 
		Then pValue = Left(pValue, Len(pValue)-1)
	
	If IsMatch(ExistValueOfPath, "\*?" & Replace(pValue, "\", "\\") & "\\?(\b|;)") 
		Then Exit Sub 
 
	If ExistValueOfPath <> "" 
		Then pValue = ";" & pValue
			pSysEnv(pPath) = ExistValueOfPath & pValue 
		Else
			pSysEnv.Remove(pPath)
    End If
End Sub

'--------获取输入参数设置系统环境变量---------
Do
	InputKey = InputBox("请输入系统变量名")
	If InputKey = VbEmpty Then
		MsgBox "已取消！" 
		Wscript.Quit
	Else
		If InputKey <> "" Then InputValue = Inputbox("请输入系统变量值"): Exit Do
	End If
Loop


If InputValue = VbEmpty 
	Then MsgBox "已取消！" 
    Wscript.Quit
	Else
		SetEnv InputKey,  InputValue 
End If

MsgBox "系统变量设置成功！"
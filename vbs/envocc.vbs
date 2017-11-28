'--------定义设置系统环境变量的方法---------
Set pSysEnv = CreateObject("WScript.Shell").Environment("user")'("System")

Function IsMatch(Str, Patrn)
  Set r = new RegExp
  r.Pattern = Patrn
  IsMatch = r.test(Str)
End Function


Sub SetEnv(pPath, pValue)
    Dim ExistValueOfPath
	
    If pValue <> "" Then 
		ExistValueOfPath = pSysEnv(pPath)
	 
		If Right(pValue, 1) = "\" Then pValue = Left(pValue, Len(pValue)-1)

		If IsMatch(ExistValueOfPath, "\*?" & Replace(pValue, "\", "\\") & "\\?(\b|;)") Then Exit Sub 

		If ExistValueOfPath <> "" Then pValue = ";" & pValue
		
		pSysEnv(pPath) = ExistValueOfPath & pValue 
	Else
		pSysEnv.Remove(pPath)
    End If
End Sub

name = "InputKey"

InputValue = InputBox("Input anything to set, Empty for delete")

'MsgBox InputValue

If InputValue = "" Then
	value = ""
	SetEnv name,  value
	MsgBox "Delete  "& name & " " & "  ok"
Else
	value = "cc"
	SetEnv name,  value
	MsgBox "Set  "& name & " " & value & "  ok"
End If






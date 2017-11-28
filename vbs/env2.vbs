Dim currentpath
currentpath=createobject("Scripting.FileSystemObject").GetFolder(".").Path
MsgBox currentpath

Dim ocrpath
ocrpath=currentpath&"\Tesseract-OCR"

SetTheEnv(ocrpath)

Function SetTheEnv(ocrpath)
	Dim pSysEnv
	Set pSysEnv = CreateObject("WScript.Shell").Environment("System")  
	pSysEnv("TESSDATA_PREFIX")=ocrpath&"\"
	
	Dim ExistValue
	ExistValue=pSysEnv("path")
	Dim target,s,exist,appendvalue
	exist=False
	appendvalue=ocrpath&";"&ocrpath&"\training"
	ExistValue=ExistValue&";"&appendvalue
	ExistValue=reduce(ExistValue,False,";")
	target=split(ExistValue,";")
	ExistValue=""
	For Each s In target
		If s<>"" Then
		ExistValue=ExistValue&s&";"
		End If
	Next
	pSysEnv("path")=ExistValue
	'ExistValue=pSysEnv("wdir")
	'MsgBox("WDIR="&ExistValue)
	MsgBox "环境设置成功!"
End Function
Function reduce(srcstr,casesentive,sp)
	Dim objDict,x,y
	srcarr=split(Trim(srcstr),sp)
	Set objDict=createobject("Scripting.Dictionary")
	For Each x In srcArr
		If Not casesentive Then 
			y=LCase(x)
		Else
			y=x
		End If
		If Not objDict.Exists(y) Then 
			If y<>lcase(driverLetter)&"\" Then
				objDict.Add y,x
			End If
		End If
	Next
	reduce=Join(objDict.Items,sp)
	Set objDict=Nothing
End Function
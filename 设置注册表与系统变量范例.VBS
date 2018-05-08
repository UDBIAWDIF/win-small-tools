Function ReplaceText(srcstr, patrn, replStr)		' 本函数对路径进行格式化，把形如"C:\XXX"格式化成"C:\\XXX"
	Dim regEx										' 建立变量
	Set regEx = New RegExp							' 建立正则表达式
	regEx.Pattern = patrn							' 设置模式
	regEx.IgnoreCase = True							' 设置是否区分大小写
	regEx.Global = True								' 设置全局适用
	ReplaceText = regEx.Replace(srcstr, replStr)	' 作替换
End Function

Function GetPath
	'' Retrieve path to the script file
	Dim path
	path = WScript.ScriptFullName  '' Script file name
	' GetPath = Left(path, InstrRev(path, "\"))
	GetPath = Left(path, InstrRev(path, "\") - 1)
End Function

Dim WSHShell
Set WSHShell = WScript.CreateObject("WScript.Shell")
WSHShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Android SDK Tools\Path", GetPath(), "REG_SZ"
Set WSHShell = Nothing
WScript.echo "已设置好注册表!"

Dim pSysEnv
Set pSysEnv = CreateObject("WScript.Shell").Environment("System")
pSysEnv("ANDROID_HOME") = GetPath()
Set pSysEnv = Nothing

WScript.Quit  '' Terminate script.

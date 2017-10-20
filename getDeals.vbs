'initialData
'------------------------------------------------------------------
Dim weeksFromNow: weeksFromNow = 0
Dim pathToFile: pathToFile="D:\Downloads\getDealsWeekend.txt"
Dim onStartofWeekend: onStartofWeekend="check for the next weekend already"
Dim origen : origen = "bcn"
Dim urlbase: urlbase="https://www.skyscanner.es/transporte/vuelos-desde/"
'-------


Launch("chrome")


'------------------------------------------------------------------
' Sub to launch browsers with the designated url 
'------------------------------------------------------------------
Sub launch(browser)
	Dim sh
	Set sh = CreateObject("WScript.Shell")
	Select Case browser
		Case "chrome"
			sh.Run "chrome -url " + getUrl()
		Case "ie"
			sh.Run "iexplore.exe " + getUrl()
		case "firefox"
			Dim ff : ff="""C:\Program Files (x86)\Mozilla Firefox\firefox.exe"""
			sh.Run ff & " " & getUrl()
		
		case else
			WScript.echo "Browser not found"
	End Select
End Sub

'------------------------------------------------------------------
' Function for obtaining the url for a search
'------------------------------------------------------------------
Function getUrl()
	Dim Otoday : Otoday = now()
	Ostart=checkNextWeekendDate()
	Stoday=sprintf("{0:yyMMdd}", Array(Otoday))
	Sstart=sprintf("{0:yyMMdd}", Array(Ostart))

	
	If Stoday > Sstart Or (Stoday = Sstart And onStartofWeekend = "check for the next weekend already") Then
		While Stoday > Sstart Or (Stoday = Sstart And onStartofWeekend = "check for the next weekend already")
			Ostart=updateWeekend(Ostart+7)
			Sstart=sprintf("{0:yyMMdd}", Array(Ostart))
			WScript.Sleep(300)
		Wend
		'WScript.echo "updating file that says when is next weekend."
	End If

    ''Control for how many weeks from now
	Ostart = Ostart + (7*weeksFromNow)

	
	Oend = Ostart + 2

	Sstart=sprintf("{0:yyMMdd}", Array(Ostart))
	Send=sprintf("{0:yyMMdd}", Array(Oend))


	'WScript.echo "Weeks from now wanted: " & weeksFromNow &  vbNewLine  & "Today: " & Otoday & vbNewLine & "Start: " & Ostart & vbNewLine & "End:   " & Oend 


	'' Form url
	getUrl = urlbase + origen+"/"+ Sstart +"/"+ Send
	'WScript.echo getUrl
	
End Function


'------------------------------------------------------------------
' Function for formatting dates -
'              sprintf("{0:yyMMdd}", Array(dt)) -> YYMMDD
'------------------------------------------------------------------
Function sprintf(sFmt, aData)
	Dim g_oSB : Set g_oSB = CreateObject("System.Text.StringBuilder")
   g_oSB.AppendFormat_4 sFmt, (aData)
   sprintf = g_oSB.ToString()
   g_oSB.Length = 0
End Function


'-------------------------------------------------------------------
'Read last Weekend date from file
'-------------------------------------------------------------------
Function checkNextWeekendDate()
	Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(pathToFile,1)
	checkNextWeekendDate = CDate(objFileToRead.ReadAll())
	objFileToRead.Close
	Set objFileToRead = Nothing
	
End Function
'-------------------------------------------------------------------
'Write new Weekend date to file
'-------------------------------------------------------------------
Function updateWeekend(s)
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(pathToFile,2,true)
	objFileToWrite.WriteLine(s)
	objFileToWrite.Close
	Set objFileToWrite = Nothing
	updateWeekend=s
End Function






'initialData
'------------------------------------------------------------------
Dim weeksFromNow: weeksFromNow = 3
Dim pathToFile: pathToFile="D:\Downloads\getDealsWeekend.txt"
Dim onStartofWeekend: onStartofWeekend="check for the next weekend already"

'------------------------------------------------------------------
' Function for formatting dates -
'              sprintf("{0:yyMMdd}", Array(dt)) -> YYMMDD
'------------------------------------------------------------------
Dim g_oSB : Set g_oSB = CreateObject("System.Text.StringBuilder")

Function sprintf(sFmt, aData)
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


'-------------------------------------------------------------------
' Algorithm to select correct weekend
'-------------------------------------------------------------------
Dim Otoday : Otoday = now()
Ostart=checkNextWeekendDate()
Stoday=sprintf("{0:yyMMdd}", Array(Otoday))
Sstart=sprintf("{0:yyMMdd}", Array(Ostart))

'WScript.echo "about to update because: " & vbNewLine & Otoday & vbNewLine & Ostart & vbNewLine & "--------" & vbNewLine & Stoday & vbNewLine & Sstart & vbNewLine & "---------" & onStartofWeekend

If Stoday > Sstart Or (Stoday = Sstart And onStartofWeekend = "check for the next weekend already") Then
	Ostart=updateWeekend(Ostart+7)
	'WScript.echo "updating when is next weekend expected"
End If

Ostart = Ostart + (7*weeksFromNow)
Oend = Ostart + 2

Sstart=sprintf("{0:yyMMdd}", Array(Ostart))
Send=sprintf("{0:yyMMdd}", Array(Oend))


'WScript.echo "Weeks from now wanted: " & weeksFromNow &  vbNewLine  & "Today: " & Otoday & vbNewLine & "Start: " & Ostart & vbNewLine & "End:   " & Oend 

WScript.echo "Checking for weekend " + Sstart + "-" + Send




' If Stoday < Sstart
' '	WScript.echo Stoday + " is earlier than " + Sstart + vbNewLine + "The weekend we are considering is correct"
' Elseif Stoday = Sstart Then
	' If onstarts = "check for the next weekend already" Then
		' 'WScript.echo Stoday + " is weekend already, check for the next weekend"
		' Ostart=updateWeekend(Ostart+7)
	' Else		
	' '	WScript.echo Stoday + " is weekend already, but im feeling like travelling today "
	' End If
' Else
' '	WScript.echo Stoday + " is later than " + Sstart + vbNewLine + "Updating weekend value and checking for it"
	' Ostart=updateWeekend(Ostart+7)
	' Sstart=sprintf("{0:yyMMdd}", Array(Ostart))
' '	WScript.echo "Next weekend will be " + Sstart
' End If

'Send=sprintf("{0:yyMMdd}", Array(Ostart+2))

'WScript.echo "Checking for weekend " + Sstart + "-" + Send

	

'-------------------------------------------------------------------
' Form url and launch
'-------------------------------------------------------------------
Dim origen : origen = "bcn"
Dim url: url="https://www.skyscanner.es/transporte/vuelos-desde/" + origen+"/"+ Sstart +"/"+ Send
'WScript.echo url

Dim browobj
Set browobj = CreateObject("WScript.Shell")
browobj.Run "chrome -url " + url


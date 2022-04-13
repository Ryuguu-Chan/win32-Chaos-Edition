' ====================================
' Win32Chaos v1.0
' MADE BY OGAN ÖZKUL AKA RYUGUU CHAN
' FROM 10.04.2022 (00:17 AM) TO 11.04.2022 (9:03 PM)
' ====================================

' THIS WHOLE SCRIPT WAS 100% MADE USING THE notepad.exe PROGRAM :)

' the shell object
set AppShell = CreateObject("WScript.Shell")

' Getting the currently logged username
userName = AppShell.ExpandEnvironmentStrings("%USERNAME%")

' the areas which files/directories can be created
CreationLocations = Array("C:\Users\" & userName & "\Desktop", "C:\Users\" & userName & "\Documents", "C:\Users\" & userName & "\Videos", "C:\Users\" & userName & "\Music", "C:\Users\" & userName & "\Downloads", "C:\Users\" & userName & "\Pictures")

' all the keypresses
ChaoKeys = Array("{CAPSLOCK}", "{ENTER}", "{LEFT}", "{RIGHT}", "{UP}", "{DOWN}", "A", "B", "C", "D")

' the prefix that will be included in every file/folder created from it
FolderFilePrefix = "(win32Chaos) "

' for the VM check
VMcheck = false

' the filesystem object
' this will be used creating folders and files
set FileSystem = CreateObject("Scripting.FileSystemObject")

' the mathematic section
' ----------------------

' counting the amouns of digits that's contained in the number n
function GetAmountOfDigits(n)

	' the to be returned value
	R = 0

	' putting n on WorkN
	' since it apparently affects the real one
	WorkN = n

	' repeating it until n is lower than 1
	do while WorkN >= 1

		' placing the most significant digit after the decimal point
		WorkN = (WorkN / 10)

		' interating R to 1
		R = (R + 1)
	loop

	' returning the results
	GetAmountOfDigits = R
end function


' the Do the code Method
sub ExecuteCode(codeN)
	
	' only execute the code if it contains 3 digits (for now)
	if GetAmountOfDigits(codeN) = 3 then

		' getting the least significant digit (the 100s)
		CodeCat = int(codeN / 100)

		' getting the semi significant digit (the 10s)
		CodeAct = int(codeN / 10) - (codeCat * 10)

		' getting the least signficant digit (the 1s)
		CodeAux = (codeN - (codeCat * 100)) - (codeAct * 10)


		' ANALYZING THE CODE
		' ------------------	
		'
		' if CodeCat is 1 -> then it's a GUI type of code
		if CodeCat = 1 then

			' if CodeAct is 2 -> then it's a CMD type of code
			if CodeAct = 2 then
				if CodeAux > 0 then
					for CMDshowLoop = 0 to CodeAux
						AppShell.run("cmd.exe")
					next
				else
					AppShell.run("cmd.exe")
				end if

			' if CodeAct is 3 -> then it's a winver type of code
			elseif CodeAct = 3 then
				if CodeAux > 0 then
					for CMDshowLoop = 0 to CodeAux
						AppShell.run("winver.exe")
					next
				else
					AppShell.run("winver.exe")
				end if

			' if CodeAct is 4 -> then it's a explorer.exe type of code
			elseif CodeAct = 4 then
				if CodeAux > 0 then
					for ExplorerShowLoop = 0 to CodeAux
						AppShell.run("explorer.exe")
					next
				else
					AppShell.run("winver.exe")
				end if
			end if

		' if CodeCat os 2 -> then it's folder/files stuff
		elseif CodeCat = 2 then

			' if the CodeAct is 1 -> Directory Creation Time!
			' but only if CodeAux is less than the amount of locations on CreationLocation
			if CodeAct = 1 then
				if CodeAux < ubound(CreationLocations) then
					FileSystem.CreateFolder(CreationLocations(CodeAux) & "\" & FolderFilePrefix & Int((999999 * Rnd) + 111111))
				end if

			' if CodeAct is 2 -> .txt file creation
			elseif CodeAct = 2 then
				if CodeAux < ubound(CreationLocations) then
					FileSystem.CreateTextFile(CreationLocations(CodeAux) & "\" & FolderFilePrefix & Int((999999 * Rnd) + 111111))
				end if
			end if

		' if CodeCat is 3 -> it's IO stuff
		elseif CodeCat = 3 then

			' if CodeAct is 1 -> it will be about the Keyboard
			if CodeAct = 1 then
				if CodeAux < ubound(ChaoKeys) then
					AppShell.SendKeys(ChaoKeys(CodeAux))
				end if

			' if CodeAct is 2 -> it will be about sound
			elseif CodeAct = 2 then

				' if CodeAux is lower than 5 -> lower the volume x times
				if CodeAux < 5 then
					for LowVolumeIndex = 0 to CodeAux
						AppShell.SendKeys(chr(&hAE))
					next

				' if CodeAux is higher than 5 -> increase the volume x times
				elseif CodeAux > 5 then
					for HighVolumeIndex = 0 to CodeAux
						AppShell.SendKeys(chr(&hAF))
					next

				' otherwise -> just mute it
				else
					AppShell.SendKeys(chr(&hAD))
				end if
			end if
		end if
	end if
end sub

' VM detection
' only run it inside a VM!
set WMI = GetObject("winmgmts:{impersonationLevel=impersonate,authenticationLevel=pktPrivacy}!\\.\root\cimv2")
set colItems = WMI.ExecQuery("Select * from Win32_ComputerSystem")

' iterating through the results
for Each objItem in colItems

	' placing the string inside Model
  	Model = objItem.Model

	' if Model is VIRUTAL (if converted to uppercase) -> all good
    	if InStr(UCase(Model), "VIRTUAL") then
      		VMcheck = true
    	end if
next


' displaying the warning message
WarningMSG = MsgBox("WARNING!"+ vbNewLine +"This script should only be ran inside a Windows Virtual Machine!" + vbNewLine + vbNewLine  + "If you decide to run it on a real machine. Any damages that could happens will be your own fault!" + vbNewLine + "You have been warned" + vbNewLine + vbNewLine + "Continue?", 4, "CAUTIONS!")

' if it doesnt run inside a VM -> close the script!
if VMcheck = false then
	MsgBox("It doesn't run on a VM! Make a Virtual Machine and run the script on it." + vbNewLine + vbNewLine + "Thus the script will be closed!")
else
	' if the user clicked Yes -> start the chaos!
	' otherwise -> go out of there!
	if WarningMSG = 6 then
		' the code
		CODE = 0

		' the chaos intensity level
		' the bigger the value is. The more chaotic thing will become
		INTENSITY = 0

		' 1 sec of wait time (initially)
		' the delay will keep decreasing by INTENSITY
		WAIT_TIME = 1000

		' kick start the RNG
		' it will take the current time as it's seed
		Randomize

		' the infinite loop
		do while true

			' generate the code
			CODE = Int((999 * Rnd) + 100)

			' sleep
			WScript.Sleep(WAIT_TIME - INTENSITY)

			' test
			ExecuteCode(CODE)

			' adding INTENSITY 1 (only if it doesn't make it negative
			if INTENSITY <> WAIT_TIME then
				INTENSITY = INTENSITY + 1
			end if
		
		loop
	else
		MsgBox("The script will be closed then!")
	end if
end if

sub zmain()
	'Verify user would like to run
	messageBox1 = msgbox("Warning: This macro will shut down and delete de-converted processors. Are you sure you would like to continue?", 3, "Master IPL Shut Down & Purge")

	'Run macro if user selects "Yes"
	if messageBox1 = "Yes" then
	
		'Declare variables
		dim rowCount, columnCount, getHowMany, stringText, deleteString, deconvertedProcessors

		'Array for de-converted processors
		deconvertedProcessors = Array("IVR2", "HOU4", "NETS", "AS27", "IVR2", "IVR4", "IVR3", "VISA", "GENP", "0293", "0455", "0445", "0787", "AL14", "0364", "0020", "1058", "0712", "LE9D", "0568", "0165", "L453")
		
		'Process # 1 - Shut down the HCH's for de-converted processors
			'Send cursor home
			SendHostKeys("<HOME>")
			
			'Shut-down command for each processor in the array 
			for each processor in deconvertedProcessors
				deleteString = "SHUTDOWN HCH" + processor + "<ENTER>"
				SendHostKeys(deleteString)
				WaitForHostUpdate(10)
			next
			
		'Process # 2 - Remove all de-converted processors from NetStat log
			'Set rowCount, columnCount, and getHowMany
			rowCount = 3
			columnCount = 2
			getHowMany = 25
			
			'Send cursor home and pause screen
			SendHostKeys("<Home>PAUSE<Enter>")

			'Row loop
			do until rowCount = 21
				'Column loop
				do until columnCount > 54
				stringText = GetString(rowCount, columnCount, getHowMany)
				stringText = Replace(stringText, "?", " ")	
				'Check to see if stringText is a de-converted processor
				for each processor in deconvertedProcessors
					if mid(stringText, 3, 4) = processor then
						deleteString = "DELETE '" + mid(stringText, 1, 8) + mid(stringText, 10, 5) + "',PURGE<ENTER>"
						SendHostKeys(deleteString)
						WaitForHostUpdate(10)
					end if
				next	
				'Check rowCount, columnCount, and stringText conditions
				if rowCount = 20 and columnCount = 54 and stringText <> "                         " then
					SendHostKeys("<HOME><ENTER>")
					WaitForHostUpdate(10)
					rowCount = 3
					columnCount = 2
				else	
				'Add to columnCount
				columnCount = columnCount + 26
				end if
				loop
			'Reset columnCount and add 1 to rowCount
			columnCount	= 2
			rowCount = rowCount + 1
			loop	

			'Send cursor home and change pause to 15
			SendHostKeys("<Home>Pause 15<Enter>")
			SendHostKeys("<Enter>")
			
	end if
end sub
Const ForReading = 1
Const ForWriting = 2
Dim sFolder,oFile,oFile2,str1,strNewFN,strLine,sFile,strNewContents
sFolder = "W:\backup\"
Set oFSO = CreateObject("Scripting.FileSystemObject")
For Each oFile In oFSO.GetFolder(sFolder).Files
  If oFSO.GetExtensionName(oFile.Name) = "csv" Then
    ProcessFiles oFSO, oFile
  End if
Next
Set oFSO = Nothing
Sub ProcessFiles(FSO, File)
	Set oFile2 = FSO.OpenTextFile(File.path, ForReading)
	str1 = ",,,"
	strLine = oFile2.ReadLine
	strNewFN=Replace(strLine,",,,","")
     Do Until oFile2.AtEndOfStream
       strLine = oFile2.ReadLine
	If (InStr(strLine,str1) > 0) Then
     		sFile = sFolder & strNewFN & ".bdy"
		Set File = FSO.CreateTextFile(sFile,2,true)
     		File.Write strNewContents
     		File.Close
     		Set File = Nothing
		strNewFN = Replace(strLine,",,,","") 
		strLine = oFile2.ReadLine
		strNewContents = ""
       End If
	strNewContents = strNewContents & strLine & vbCrLf
     Loop
     		sFile = sFolder & strNewFN & ".bdy"
		Set File = FSO.CreateTextFile(sFile,2,true)
     		File.Write strNewContents
     		File.Close
     		Set File = Nothing
		strNewContents = ""
     oFile2.close
     set oFile2 = Nothing
end sub

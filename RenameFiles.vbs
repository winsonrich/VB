'  RenameFiles.vbs 
'    Renames all files in a specified folder (P1) to a specified naming prefix (P2) 
'    In addition, it appends _(number) to each filename prefix where (number) is a 
'    sequentially increasing number for each file. Filetypes are not modified. 
'          RenameFiles.vbs "D:\backup\" 1-00*.*
'    will rename all files in the C:\images folder to NDL_1.pdf, NDL_2.pdf, etc. 
' 
'   This is very handy for renaming a batch of files obtained from a digital camera! 
' 
Option Explicit 
Dim oCmd, oFolder, oFSO, oFileList, oFile 
Dim sRenamePath, sRenamePrefix, sFileExtension, sConfirmRename 
Dim iFileCount, iFileIndex, fNum
Dim bConfirmEach 
 
Set oCmd = Wscript.Arguments 
 
Select Case (2) 
    Case 2 
      sRenamePath = "D:\backup"
      sRenamePrefix = "NDL"
      bConfirmEach = False 
	
    Case 3 
      sRenamePath = oCmd.item(0) 
      sRenamePrefix = oCmd.item(1) 
      bConfirmEach = oCmd.item(2) 
    Case Else 
      WScript.Echo "RenameFiles.vbs requires 2 parameters:" &_ 
          vbcrlf & "1) Folder Path (or . for current folder)" &_ 
          vbcrlf & "2) File Prefix" &_ 
          vbcrlf & "3) Confirm each file? True/*False* (optional)" 
      WScript.Quit 
End Select 
' 
' Get a list of all files in the specified folder 
' 
Set oFSO = CreateObject("Scripting.FileSystemObject") 
Set oFolder = oFSO.GetFolder(sRenamePath) 
 
Set oFileList = oFolder.Files 
For Each oFile in oFileList 
    iFileCount = iFileCount + 1 
Next 

iFileIndex = 79 - 1 'Input the first NDL Num
For Each oFile in oFileList 
	sFileExtension = Right(oFile.Name,Len(oFile.Name)-InStr(oFile.Name,".")) 
	iFileIndex = iFileIndex + 1
	If ( StrComp( Left(oFile.Name,Len(4)), "1-00" ) <1)Then
            oFile.Move (sRenamePath & "\" & sRenamePrefix & "_" & _ 
               iFileIndex & "." & sFileExtension) 
	End If

Next 
Set oFolder = Nothing 
Set oFSO = Nothing 
WScript.Echo "Done" 
WScript.Quit 
'-------------------------------------------------------- 
Function ZeroFill (sString, iFieldLength) 
    ZeroFill = Right(Ltrim(sString), iFieldLength) 
End Function 

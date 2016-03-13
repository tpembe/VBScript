Const ForReading = 1

Const ForWriting = 2

 

Set objRegEx = New RegExp

objRegEx.Global = True

objRegEx.IgnoreCase = True

objRegEx.Pattern = "D(0[1-9]|[12][0-9]|3[01])[- /.](0[1-9]|1[012])[- /.'](19|20)[0-9]{2}"

 

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFolder = objFSO.GetFolder(".")

Set colFiles = objFolder.Files

 

For Each item in colFiles

                sExtension = objFSO.GetExtensionName(item.Name)

                If sExtension = "qif" Then

                                Set objReadFile = objFSO.OpenTextFile(item.Name, ForReading)

                                strText = objReadFile.ReadAll

                                objReadFile.Close

                                Set colMatches = objRegEx.Execute(strText) 

                                If colMatches.Count > 0 Then

                                                For Each colMatch in colMatches  
																repl = Replace(colMatch.value, "'", "/")
                                                                ss = split(Right(repl,10),"/")

                                                                newDate = "D" & ss(1) & "/" & ss(0) & "/" & ss(2)

                                                                strText = Replace(strText, colMatch.value, newDate)                    

                                                Next

                                End If

                                Set objWriteFile = objFSO.OpenTextFile(item.Name, ForWriting)

                                objWriteFile.write(strText)

                                objWriteFile.Close

                End if

Next

Wscript.Echo "Finished!"
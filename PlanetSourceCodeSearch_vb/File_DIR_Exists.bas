Attribute VB_Name = "modFiles_Functions_DIR_Exists"
'Public Function FSODIRExists(ByVal pFolder As String)
'
'>--> Objective:        Check for file existence
'>--> Dependency:       File System Object
'>--> Compatibility:    Windows 2000 / XP Only
'>-->
'Dim Lfso As New FileSystemObject
'    Folder_Exists = Lfso.FolderExists(pFolder)
'End Function


Function DIRExists(ByVal PathName As String, Optional Directory As Boolean = False) As Boolean
'>--> Objective:        Check for file existence
'>--> Dependency:
'>--> Compatibility:    Windows ALL


    If PathName <> "" Then
        

        If Directory = False Then
            
            DIRExists = (Dir$(PathName) <> "")
        Else
            
            DIRExists = (Dir$(PathName, vbDirectory) <> "")
        End If

        
    End If

End Function

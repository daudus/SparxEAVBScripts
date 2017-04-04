Namespace VBScript

    Module Main
        Sub Main()

        End Sub

        Sub Script()
            !INC Wrappers.Include
            'get the folder from the user
            Dim folder
            folder = New FileSystemFolder
            folder = folder.getUserSelectedFolder("")
            'show messagebox with the name of each subfolder
            Dim subfolders, subfolder
            For Each subfolder In folder.SubFolders
                MsgBox("subfolder name: " & subfolder.Name)
            Next
        End Sub
    End Module
End Namespace

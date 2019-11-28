Sub Convert_PPT_to_imported_images()  
    'PURPOSE: convert
    'SOURCE: https://github.com/Kgotso-Koete/Convert_PPT_to_imported_images
    
    'Get current file's directory path
    'SOURCE: https://stackoverflow.com/questions/12546181/vba-powerpoint-how-to-get-files-current-directory-path-to-a-string-in-vba
    Dim sPath As String
    Dim strFolder As String
    Dim newSaveFile As String
    
    sPath = ActivePresentation.path & "\"
    If Len(sPath) > 0 Then
        MsgBox "Converting " & ActivePresentation.Name & vbNewLine & "to image presentation"
    Else
        MsgBox "File not saved"
    End If
     
    ' Create a new folder to be used to save images and deleted later
    ' SOURCE: https://www.oreilly.com/library/view/vbscript-in-a/0596004885/re69.html
    Dim oFileSys
    Dim oFolder
    strFolder = sPath & "Convert_folder_18926"
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFileSys.CreateFolder(strFolder)
         
    Call SaveSlidesAsImages(sPath)
    Call Deleteslides
    Call BatchImportImages(sPath)
    Call DeleteImages(sPath)
    
    ' Delete the temp folder
    ' SOURCE: https://www.vbsedit.com/html/7991e577-f2ac-4b33-8231-761971d6c0d9.asp
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFolder (strFolder)
    
    'Get name of path of presentation without file extension
    ' SOURCE: https://www.excel-easy.com/vba/string-manipulation.html
    newSaveFile = Left(Application.ActivePresentation.FullName, InStr(Application.ActivePresentation.FullName, ".pptm") - 1) & "_CONVERTED.ppt"
    ActivePresentation.SaveCopyAs FileName:=newSaveFile
    
    'Open new saved presentation and close current one
    Presentations.Open FileName:=newSaveFile
    
End Sub
 

Private Function BatchImportImages(ByVal strDirectory As String)
    'PURPOSE: Import a batch of images as powerpoint slides
    'SOURCE: https://software-solutions-online.com/add-images-folder-powerpoint-presentation-using-vba/
 
    Dim i As Integer
    Dim n As Integer
    Dim arrFilesInFolder As Variant
    Dim FetchLocation As String
    FetchLocation = strDirectory & "Convert_folder_18926"
     
    arrFilesInFolder = GetAllFilesInDirectory(FetchLocation)
    
    For i = LBound(arrFilesInFolder) To UBound(arrFilesInFolder)
        n = UBound(arrFilesInFolder) - i
        Call AddSlideAndImage(arrFilesInFolder(n))
    Next
End Function
Private Function GetAllFilesInDirectory(ByVal strDirectory As String) As Variant
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim arrOutput() As Variant
    Dim i As Integer
    
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Get the folder object
    Set objFolder = objFSO.GetFolder(strDirectory)
    ReDim arrOutput(0)
    i = 1
    'loops through each file in the directory and prints their names and path
    For Each objFile In objFolder.Files
     
        'print file path
        arrOutput(i - 1) = objFile.path
        ReDim Preserve arrOutput(UBound(arrOutput) + 1)
        i = i + 1
    Next objFile
       ReDim Preserve arrOutput(UBound(arrOutput) - 1)
    GetAllFilesInDirectory = arrOutput
End Function

Private Function AddSlideAndImage(ByVal strFile As String)
    Dim objPresentaion As Presentation
    Dim objSlide As Slide
    
    Set objPresentaion = ActivePresentation
    
    Set objSlide = objPresentaion.Slides.Add(1, PpSlideLayout.ppLayoutChart)
    Call objSlide.Shapes.AddPicture(strFile, msoCTrue, msoCTrue, Left:=0, _
    Top:=0, _
    Width:=-1, _
    Height:=-1)
    
  

End Function


Private Function SaveSlidesAsImages(ByVal strDirectory As String)
    'PURPOSE: Save each selected slide as an individual image file
    'SOURCE: https://www.thespreadsheetguru.com/blog/save-your-powerpoint-slides-as-images
    
    Dim FileExtension As String
    Dim SaveLocation As String
    Dim ImageName As String
    Dim SelectedSlides As SlideRange
    Dim sld As Slide
    Dim x As Long
     
    'Inputs
      FileExtension = "png" 'jpg, gif, bmp, emf
      SaveLocation = strDirectory & "Convert_folder_18926\"
      ImageName = "Custom Image"
      
    'Set variable equal to only selected slides in Active Presentation
      On Error GoTo NoSlideSelection
        Set SelectedSlides = ActiveWindow.Selection.SlideRange
      On Error GoTo 0
    
    'Loop through each selected slide
      For x = 1 To SelectedSlides.Count
        
        'Store each slide to a variable
          Set sld = SelectedSlides(x)
          
        'Save Slide as image file
          With ActivePresentation.Slides(sld.SlideIndex)
            .Export SaveLocation & ImageName & _
            sld.SlideIndex & "." & FileExtension, FileExtension
          End With
    
      Next x
    
    Exit Function
    
    'ERROR HANDLERS
NoSlideSelection:
      MsgBox "You do not have any slides in your PowerPoint project.", 16, "No Slides Found"
      Exit Function

End Function

Private Function Deleteslides()
    'PURPOSE: Save each selected slide as an individual image file
    'SOURCE: https://pepitosolis.wordpress.com/2013/04/19/add-or-remove-slides-in-powerpoint-with-vba/
 
    Dim Pre As Presentation
    Set Pre = ActivePresentation
    Dim x As Long
    For x = Pre.Slides.Count To 1 Step -1
        Pre.Slides(x).Delete
    Next x

End Function
  
Private Function DeleteImages(ByVal strDirectory As String)
    'PURPOSE: Save each selected slide as an individual image file
    'SOURCE: http://www.cruto.com/resources/vbscript/vbscript-examples/storage/files/Delete-All-Files-in-a-Folder.asp
  
    Const DeleteReadOnly = True
    Dim DeleteLocation As String
    DeleteLocation = strDirectory & "Convert_folder_18926\*.png"

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.DeleteFile (DeleteLocation), DeleteReadOnly

End Function
 


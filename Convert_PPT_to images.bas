Sub Convert_PPT_to_imported_images()
    'PURPOSE: convert
    'SOURCE: https://github.com/Kgotso-Koete/Convert_PPT_to_imported_images
    
    'Get current file's directory path
    'SOURCE: https://stackoverflow.com/questions/12546181/vba-powerpoint-how-to-get-files-current-directory-path-to-a-string-in-vba
    Dim sPath As String
    Dim strFolder As String
    Dim newSaveFile As String
    
    sPath = ActivePresentation.Path & "\"
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
    
    ' Save slides as images
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
      SaveLocation = sPath & "Convert_folder_18926\"
      ImageName = "Custom Image"
      
    'Set variable equal to only selected slides in Active Presentation
      On Error GoTo NoSlideSelection
        Set SelectedSlides = ActivePresentation.Slides.Range
      On Error GoTo 0
    
    'Loop through each selected slide
      For x = 1 To SelectedSlides.Count
        
        'Store each slide to a variable
          Set sld = SelectedSlides(x)
          
        'Save Slide as image file
          With ActivePresentation.Slides(sld.SlideIndex)
            .Export SaveLocation & sld.SlideIndex & "." & FileExtension, FileExtension
          End With
    
      Next x
      
      ' Open the copy presentation
    ' SOURCE: https://www.excel-easy.com/vba/string-manipulation.html
    newSaveFile = Left(Application.ActivePresentation.FullName, InStr(Application.ActivePresentation.FullName, ".pptm") - 1) & "_CONVERTED.ppt"
    ActivePresentation.SaveCopyAs FileName:=newSaveFile
    
    Dim PreDELETE As Presentation
    Set PreDELETE = Presentations.Open(newSaveFile)
    
    Dim xLen As Long
    For xLen = PreDELETE.Slides.Count To 1 Step -1
        PreDELETE.Slides(xLen).Delete
    Next xLen
    
    With PreDELETE
        .SaveAs newSaveFile
    End With
    
        'Loop through each selected slide
        'Dim objPresentation As Presentation
        Dim objSlide As Slide
    
      For x = 0 To (SelectedSlides.Count - 1)
        'Store each slide to a variable
        Set sld = SelectedSlides(SelectedSlides.Count - x)

    
        Set objSlide = PreDELETE.Slides.Add(1, PpSlideLayout.ppLayoutChart)
        Call objSlide.Shapes.AddPicture(SaveLocation & sld.SlideIndex & "." & FileExtension, msoCTrue, msoCTrue, Left:=0, _
        Top:=0, _
        Width:=-1, _
        Height:=-1)
    
      Next x
      
      With PreDELETE
        .SaveAs newSaveFile
      End With
    
    'ERROR HANDLERS
NoSlideSelection:
    
    ' Delete the temp folder
    ' SOURCE: https://www.vbsedit.com/html/7991e577-f2ac-4b33-8231-761971d6c0d9.asp
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFolder (strFolder)
    
End Sub






Attribute VB_Name = "create_drawing_from_part1"
Dim swApp As SldWorks.SldWorks

Dim swCurModel As SldWorks.ModelDoc2

Dim swDraw As SldWorks.DrawingDoc

Sub main()

    Set swApp = Application.SldWorks

    Set swCurModel = swApp.ActiveDoc

    

    If swCurModel.GetPathName <> "" Then

        Dim drawTemplate As String

        drawTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateDrawing)

        If drawTemplate = "" Then

            MsgBox "Template is not found"

            End

        End If

        

        Set swDraw = swApp.NewDocument(drawTemplate, swDwgPaperSizes_e.swDwgPaperBsize, 0, 0)

        swDraw.Create3rdAngleViews2 swCurModel.GetPathName()

        

    Else

        MsgBox "Please save the model"

    End If

    

End Sub

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
        
        
        Dim swSketchSegment As SldWorks.SketchSegment
        swDraw.ActivateView "Drawing View1"

        Set swSketchSegment = swDraw.CreateCircle2(-0.05, 0, 0, -0.045, 0, 0)
        swDraw.CreateDetailViewAt3 0.4, 0.1, 0, 0, 4#, 1#, "B", 1, 0


    Else

        MsgBox "Please save the model"

    End If
    

End Sub


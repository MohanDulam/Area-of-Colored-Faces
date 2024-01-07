Imports NXOpen
Imports NXOpen.Selection


Module Face_Color
    Dim theSession As Session = NXOpen.Session.GetSession()
    Dim theUI As UI = NXOpen.UI.GetUI()
    Dim lw As ListingWindow = theSession.ListingWindow


    Sub Main()
        ' Variables to store the Selected face and body
        Dim selectedFace As Face = Nothing
        Dim selectedBody As Body = Nothing

        ' calling SelectFace function 
        SelectFace(selectedFace, selectedBody)

        ' check for faces is selected
        If selectedFace IsNot Nothing Then

            ' color code of the selected face
            Dim faceColor As Integer = selectedFace.Color
            ' units of the working part
            Dim areaUnit As Unit = selectedFace.OwningPart.UnitCollection.GetBase("Area")
            ' variable to store the area of same colored faces in part
            Dim dblarea As Double = Nothing

            lw.Open()
            'lw.WriteLine("select Face color code is " & faceColor)

            ' calling the selectedBody function and store the faces in colorFaces array
            Dim colorFaces() As Face = selectedBody.GetFaces()
            ' Declaration of the face list for same colored faces of selected face
            Dim colorFacesList As New List(Of Face)()

            ' looping through the colorFaces to calculate the area of same colored faces of selected face
            For Each face_Color As Face In colorFaces
                ' check for required colored face
                If face_Color.Color = faceColor Then
                    colorFacesList.Add(face_Color) ' add to list
                End If
            Next

            ' variable to store no. of faces have same color of selected face
            Dim faceCount As Integer = colorFacesList.Count

            ' calling function to calculate the area of the same color of selected face
            dblarea = GetFaceArea(colorFacesList.ToArray())

            ' out of the area of same colored faces of selected face
            lw.WriteLine("Number of Faces have same color are " & faceCount.ToString())

            lw.WriteLine("Area of selected Face is " & dblarea.ToString() & " " & areaUnit.Abbreviation)

        End If


    End Sub
    Public Function GetUnloadOption(ByVal dummy As String) As Integer

        'Unloads the image immediately after execution within NX
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately

    End Function
    Public Function GetFaceArea(face As Face) As Double

        ' units of the working part
        Dim areaUnit As Unit = face.OwningPart.UnitCollection.GetBase("Area")
        Dim lengthUnit As Unit = face.OwningPart.UnitCollection.GetBase("Length")

        ' array to pass the fases
        Dim faces() As Face = {face}

        ' Code to get the area of a given faces
        GetFaceArea = face.OwningPart.MeasureManager.NewFaceProperties(areaUnit, lengthUnit, 0.99, faces).Area

    End Function

    Public Function GetFaceArea(face() As Face) As Double

        ' units of the working part
        Dim areaUnit As Unit = face(0).OwningPart.UnitCollection.GetBase("Area")
        Dim lengthUnit As Unit = face(0).OwningPart.UnitCollection.GetBase("Length")

        ' Code to get the area of a given faces
        GetFaceArea = face(0).OwningPart.MeasureManager.NewFaceProperties(areaUnit, lengthUnit, 0.99, face).Area

    End Function
    Public Sub SelectFace(ByRef ObjectFace As Face, ByRef ObjectBody As Body)

        ' Argments to pass the SelectTaggedObject function
        Dim msg As String = "Select the Face"
        Dim title As String = "Select the Colored Face to calculate surface area"
        Dim scope As NXOpen.Selection.SelectionScope = SelectionScope.UseDefault
        Dim highlight As Boolean = False
        Dim type As SelectionType() = {SelectionType.Faces}
        Dim obj As TaggedObject = Nothing
        Dim cursor As Point3d

        ' get the response from the user selection
        Dim response As Response =
            theUI.SelectionManager.SelectTaggedObject(msg, title, scope, highlight, type, obj, cursor)

        ' check for the user selection
        If response <> Response.Cancel And response <> Response.Back Then
            Dim dispObj As Face = obj
            ObjectFace = dispObj
            ObjectBody = ObjectFace.GetBody()
        End If
        
    End Sub

End Module


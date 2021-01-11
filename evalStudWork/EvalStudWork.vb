Imports PdfSharp.Drawing
Imports PdfSharp.Pdf

Public Class EvalStudWork
    Private sPart As Part = Nothing
    Private mPart As Part = Nothing
    Private allRules As New List(Of TableData)
    Private reportStr As New List(Of reportItem)
    Private score As Integer = 0

    'TODO add swpsd to the SolidWorks
    'TODO enable user input to the dll

    Public Sub applyRule(inStudXML As String, inMasterXML As String)
        Try
            readStudXML(inStudXML)
            readMasterXML(inMasterXML)

            allRules = getCheckedMasterRules()
            doRuleEvaluation(allRules)
            doPDF()
        Catch ex As Exception
            MsgBox("Exception in applyRule" & vbCrLf & ex.Message)
        End Try
    End Sub

    Public Function getAllRules() As List(Of TableData)
        Return allRules
    End Function

    Private Function getCheckedMasterRules() As List(Of TableData)
        Try
            Dim tmpAttrib As Attribute
            Dim lstAttribs As List(Of Attribute)
            Dim activeRules As New List(Of TableData)
            Dim allProps As System.Reflection.PropertyInfo() = mPart.GetType().GetProperties()
            For Each p As System.Reflection.PropertyInfo In allProps
                If (p.CanRead) Then
                    If (p.PropertyType.Name.Equals("Attribute")) Then
                        tmpAttrib = p.GetValue(mPart)
                        If (equstr(tmpAttrib.rulecheck, "true")) Then
                            activeRules.Add(New TableData(tmpAttrib.name, tmpAttrib.value, tmpAttrib.rulepoints,
                                                          tmpAttrib.rulecheck, tmpAttrib.ruletype, tmpAttrib.rulevalue,
                                                          tmpAttrib.ruledescription, "part"))
                        End If
                    ElseIf (p.PropertyType.Name.Contains("List")) Then
                        If (equstr(p.Name, "envelope") Or equstr(p.Name, "properties")) Then
                            lstAttribs = p.GetValue(mPart)
                            For Each tmpAttrib In lstAttribs
                                If (equstr(tmpAttrib.rulecheck, "true")) Then
                                    activeRules.Add(New TableData(tmpAttrib.name, tmpAttrib.value, tmpAttrib.rulepoints,
                                                                  tmpAttrib.rulecheck, tmpAttrib.ruletype, tmpAttrib.rulevalue,
                                                                  tmpAttrib.ruledescription, p.Name))
                                End If
                            Next
                        ElseIf (equstr(p.Name, "features")) Then
                            Dim lstFeatures As List(Of Feature)
                            lstFeatures = p.GetValue(mPart)
                            Dim aFeature As Feature
                            For Each aFeature In lstFeatures
                                If (equstr(aFeature.name.rulecheck, "true")) Then
                                    activeRules.Add(New TableData(aFeature.name.name, aFeature.name.value, aFeature.name.rulepoints,
                                                                  aFeature.name.rulecheck, aFeature.name.ruletype, aFeature.name.rulevalue,
                                                                  aFeature.name.ruledescription, "features"))
                                End If
                                If (equstr(aFeature.type.rulecheck, "true")) Then
                                    activeRules.Add(New TableData(aFeature.type.name, aFeature.type.value, aFeature.type.rulepoints,
                                                                  aFeature.type.rulecheck, aFeature.type.ruletype, aFeature.type.rulevalue,
                                                                  aFeature.type.ruledescription, "features"))
                                End If
                                If (equstr(aFeature.base.rulecheck, "true")) Then
                                    activeRules.Add(New TableData(aFeature.base.name, aFeature.base.value, aFeature.base.rulepoints,
                                                                  aFeature.base.rulecheck, aFeature.base.ruletype, aFeature.base.rulevalue,
                                                                  aFeature.base.ruledescription, "features"))
                                End If
                                If (equstr(aFeature.frozen.rulecheck, "true")) Then
                                    activeRules.Add(New TableData(aFeature.frozen.name, aFeature.frozen.value, aFeature.frozen.rulepoints,
                                                                  aFeature.frozen.rulecheck, aFeature.frozen.ruletype, aFeature.frozen.rulevalue,
                                                                  aFeature.frozen.ruledescription, "features"))
                                End If
                                If (equstr(aFeature.suppressed.rulecheck, "true")) Then
                                    activeRules.Add(New TableData(aFeature.suppressed.name, aFeature.suppressed.value, aFeature.suppressed.rulepoints,
                                                                  aFeature.suppressed.rulecheck, aFeature.suppressed.ruletype, aFeature.suppressed.rulevalue,
                                                                  aFeature.suppressed.ruledescription, "features"))
                                End If
                                If (aFeature.attribs.Count > 0) Then
                                    lstAttribs = aFeature.attribs
                                    For Each tmpAttrib In lstAttribs
                                        If (equstr(tmpAttrib.rulecheck, "true")) Then
                                            activeRules.Add(New TableData(tmpAttrib.name, tmpAttrib.value, tmpAttrib.rulepoints,
                                                                          tmpAttrib.rulecheck, tmpAttrib.ruletype, tmpAttrib.rulevalue,
                                                                          tmpAttrib.ruledescription, "attribs"))
                                        End If
                                    Next
                                End If
                                If (aFeature.dimensions.Count > 0) Then
                                    Dim lstDimensions As List(Of Dimension)
                                    lstDimensions = aFeature.dimensions
                                    For Each aDimension As Dimension In lstDimensions
                                        If (equstr(aDimension.name.rulecheck, "true")) Then
                                            activeRules.Add(New TableData(aDimension.name.name, aDimension.name.value, aDimension.name.rulepoints,
                                                                      aDimension.name.rulecheck, aDimension.name.ruletype, aDimension.name.rulevalue,
                                                                      aDimension.name.ruledescription, "dimensions"))
                                        End If
                                        If (equstr(aDimension.type.rulecheck, "true")) Then
                                            activeRules.Add(New TableData(aDimension.type.name, aDimension.type.value, aDimension.type.rulepoints,
                                                                      aDimension.type.rulecheck, aDimension.type.ruletype, aDimension.type.rulevalue,
                                                                      aDimension.type.ruledescription, "dimensions"))
                                        End If
                                        If (equstr(aDimension.value.rulecheck, "true")) Then
                                            activeRules.Add(New TableData(aDimension.value.name, aDimension.value.value, aDimension.value.rulepoints,
                                                                      aDimension.value.rulecheck, aDimension.value.ruletype, aDimension.value.rulevalue,
                                                                      aDimension.value.ruledescription, "dimensions"))
                                        End If
                                    Next
                                End If
                                If (aFeature.sketches.Count > 0) Then
                                    Dim lstSketches As List(Of Sketch)
                                    lstSketches = aFeature.sketches
                                    For Each aSketch As Sketch In lstSketches
                                        If (equstr(aSketch.name.rulecheck, "true")) Then
                                            activeRules.Add(New TableData(aSketch.name.name, aSketch.name.value, aSketch.name.rulepoints,
                                                                      aSketch.name.rulecheck, aSketch.name.ruletype, aSketch.name.rulevalue,
                                                                      aSketch.name.ruledescription, "sketches"))
                                        End If
                                        If (equstr(aSketch.type.rulecheck, "true")) Then
                                            activeRules.Add(New TableData(aSketch.type.name, aSketch.type.value, aSketch.type.rulepoints,
                                                                      aSketch.type.rulecheck, aSketch.type.ruletype, aSketch.type.rulevalue,
                                                                      aSketch.type.ruledescription, "sketches"))
                                        End If
                                        If (equstr(aSketch.status.rulecheck, "true")) Then
                                            activeRules.Add(New TableData(aSketch.status.name, aSketch.status.value, aSketch.status.rulepoints,
                                                                      aSketch.status.rulecheck, aSketch.status.ruletype, aSketch.status.rulevalue,
                                                                      aSketch.status.ruledescription, "sketches"))
                                        End If
                                        If (equstr(aSketch.constrained.rulecheck, "true")) Then
                                            activeRules.Add(New TableData(aSketch.constrained.name, aSketch.constrained.value, aSketch.constrained.rulepoints,
                                                                      aSketch.constrained.rulecheck, aSketch.constrained.ruletype, aSketch.constrained.rulevalue,
                                                                      aSketch.constrained.ruledescription, "sketches"))
                                        End If
                                        If (aSketch.geometries.Count > 0) Then
                                            Dim lstGeometries As List(Of Geometry)
                                            lstGeometries = aSketch.geometries
                                            For Each aGeometry As Geometry In lstGeometries
                                                If (equstr(aGeometry.name.rulecheck, "true")) Then
                                                    activeRules.Add(New TableData(aGeometry.name.name, aGeometry.name.value, aGeometry.name.rulepoints,
                                                                              aGeometry.name.rulecheck, aGeometry.name.ruletype, aGeometry.name.rulevalue,
                                                                              aGeometry.name.ruledescription, "geometries"))
                                                End If

                                                If (aGeometry.attribs.Count > 0) Then
                                                    lstAttribs = aGeometry.attribs
                                                    For Each tmpAttrib In lstAttribs
                                                        If (equstr(tmpAttrib.rulecheck, "true")) Then
                                                            activeRules.Add(New TableData(tmpAttrib.name, tmpAttrib.value, tmpAttrib.rulepoints,
                                                                  tmpAttrib.rulecheck, tmpAttrib.ruletype, tmpAttrib.rulevalue,
                                                                  tmpAttrib.ruledescription, "geometry"))
                                                        End If
                                                    Next
                                                End If

                                                If (aGeometry.relations.Count > 0) Then
                                                    lstAttribs = aGeometry.relations
                                                    For Each tmpAttrib In lstAttribs
                                                        If (equstr(tmpAttrib.rulecheck, "true")) Then
                                                            activeRules.Add(New TableData(tmpAttrib.name, tmpAttrib.value, tmpAttrib.rulepoints,
                                                                                          tmpAttrib.rulecheck, tmpAttrib.ruletype, tmpAttrib.rulevalue,
                                                                                          tmpAttrib.ruledescription, "relations"))
                                                        End If
                                                    Next
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                            Next
                        ElseIf (equstr(p.Name, "equations")) Then
                            Dim lstEquations As List(Of Equation)
                            lstEquations = p.GetValue(mPart)
                            For Each aEquation As Equation In lstEquations
                                If (equstr(aEquation.name.rulecheck, "true")) Then
                                    activeRules.Add(New TableData(aEquation.name.name, aEquation.name.value, aEquation.name.rulepoints,
                                                                  aEquation.name.rulecheck, aEquation.name.ruletype, aEquation.name.rulevalue,
                                                                  aEquation.name.ruledescription, "equations"))
                                End If
                                If (equstr(aEquation.value.rulecheck, "true")) Then
                                    activeRules.Add(New TableData(aEquation.value.name, aEquation.value.value, aEquation.value.rulepoints,
                                                                  aEquation.value.rulecheck, aEquation.value.ruletype, aEquation.value.rulevalue,
                                                                  aEquation.value.ruledescription, "equations"))
                                End If
                                If (equstr(aEquation.status.rulecheck, "true")) Then
                                    activeRules.Add(New TableData(aEquation.status.name, aEquation.status.value, aEquation.status.rulepoints,
                                                                  aEquation.status.rulecheck, aEquation.status.ruletype, aEquation.status.rulevalue,
                                                                  aEquation.status.ruledescription, "equations"))
                                End If
                                If (equstr(aEquation.isglobal.rulecheck, "true")) Then
                                    activeRules.Add(New TableData(aEquation.isglobal.name, aEquation.isglobal.value, aEquation.isglobal.rulepoints,
                                                                  aEquation.isglobal.rulecheck, aEquation.isglobal.ruletype, aEquation.isglobal.rulevalue,
                                                                  aEquation.isglobal.ruledescription, "equations"))
                                End If
                            Next
                        End If
                    End If
                    Debug.Print(">>stop")
                End If
            Next
            Return activeRules
        Catch ex As Exception
            MsgBox("Exception in getCheckedMasterRules" & vbCrLf & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Sub readStudXML(inStudXML As String)
        Try
            Dim serializer As New System.Xml.Serialization.XmlSerializer(GetType(Part))
            Dim file As New System.IO.StreamReader(inStudXML)

            sPart = serializer.Deserialize(file)
            file.Close()
        Catch ex As Exception
            MsgBox("Exception in readInStudXML" & vbCrLf & ex.Message)
        End Try
    End Sub

    Private Sub readMasterXML(inMasterXML As String)
        Try
            Dim serializer As New System.Xml.Serialization.XmlSerializer(GetType(Part))
            Dim file As New System.IO.StreamReader(inMasterXML)

            mPart = DirectCast(serializer.Deserialize(file), Part)
            file.Close()
        Catch ex As Exception
            MsgBox("Exception in readInMasterXML" & vbCrLf & ex.Message)
        End Try
    End Sub

    Private Function equstr(inStrA As String, inStrB As String) As Boolean
        Try
            Dim result As Integer = String.Compare(inStrA, inStrB, True)
            If (result = 0) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox("Exception in equstr" & vbCrLf & ex.Message)
            Return False
        End Try
    End Function

    Private Function chkPoints(aPoints As String) As Integer
        If (aPoints.Equals("") Or aPoints Is Nothing) Then
            Return 0
        Else
            Return CInt(aPoints)
        End If
    End Function

    Private Sub doRuleEvaluation(allRules As List(Of TableData))
        Try
            Dim aPoints As Integer = 0
            For Each aRule As TableData In allRules
                Dim resOk As Boolean
                aPoints = 0
                Dim sValue As String
                sValue = getStudValue(aRule.nameProp(), aRule.sectionProp())

                Select Case aRule.rtypeProp()
                    Case "Exact"
                        resOk = doExact(sValue, aRule.rvalueProp())
                    Case "Discrete"
                        resOk = doDiscrete(sValue, aRule.rvalueProp())
                    Case "Tolerance"
                        resOk = doTolerance(sValue, aRule.rvalueProp())
                    Case "Range"
                        resOk = doRange(sValue, aRule.rvalueProp())
                End Select
                If (resOk) Then
                    aPoints = chkPoints(aRule.rpointsProp())
                    score = score + aPoints
                    resOk = False
                End If
                reportStr.Add(New reportItem(aRule.nameProp(), aRule.valueProp(), aRule.rcheckProp(),
                              CStr(aPoints), aRule.rdescrProp()))
            Next
        Catch ex As Exception
            MsgBox("Exception in doRuleEvaluation" & vbCrLf & ex.Message)
        End Try
    End Sub

    Private Function doExact(value As String, rvalue As String) As Boolean
        Try
            Return String.Equals(value, rvalue)
        Catch ex As Exception
            MsgBox("Exception in doExact" & vbCrLf & ex.Message)
            Return False
        End Try
    End Function

    Private Function doDiscrete(value As String, rvalue As String) As Boolean
        Try
            Dim discVals() As String
            discVals = rvalue.Split(";")
            For Each discVal As String In discVals
                If (String.Equals(discVal, value)) Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox("Exception in doExact" & vbCrLf & ex.Message)
            Return False
        End Try
    End Function

    Private Function doTolerance(value As String, rvalue As String) As Boolean
        Try
            Dim tolVals() As String
            tolVals = rvalue.Split(";")
            Dim intValLB As Integer = 0
            Dim intValHB As Integer = 0
            Dim dblValLB As Double = 0.0
            Dim dblValHB As Double = 0.0
            Dim retVal As Integer = -1
            Dim cmpOK As Boolean = False

            retVal = checkType(value, intValLB, dblValHB)
            If (retVal = -1) Then Return False

            Select Case tolVals.Count
                Case 0
                    Return True
                Case 1
                    If (tolVals(0).Contains("+")) Then
                        tolVals(0) = tolVals(0).Substring(1, tolVals(0).Length)
                        retVal = checkType(tolVals(0), intValLB, dblValLB)
                        If (retVal = -1) Then Return False
                    ElseIf (tolVals(0).Contains("-")) Then
                        tolVals(0) = tolVals(0).Substring(1, tolVals(0).Length)
                        retVal = checkType(tolVals(0), intValHB, dblValHB)
                        If (retVal = -1) Then Return False
                    Else
                        Return False
                    End If
                Case 2
                    retVal = checkType(tolVals(0), intValLB, dblValLB)
                    If (retVal = -1) Then Return False

                    retVal = checkType(tolVals(1), intValHB, dblValHB)
                    If (retVal = -1) Then Return False
            End Select

            Select Case retVal
                Case 2
                    cmpOK = compareVal(CDbl(value), dblValLB, dblValHB)
                Case 1
                    cmpOK = compareVal(CInt(value), intValLB, intValHB)
                Case Else
                    Return False
            End Select

            If (cmpOK) Then Return True

            Return False
        Catch ex As Exception
            MsgBox("Exception in doTolerance" & vbCrLf & ex.Message)
            Return False
        End Try
    End Function

    Private Function doRange(value As String, rvalue As String) As Boolean
        Try
            Dim intValLB As Integer = 0
            Dim intValHB As Integer = 0
            Dim dblValLB As Double = 0.0
            Dim dblValHB As Double = 0.0
            Dim retVal As Integer = -1
            Dim cmpOK As Boolean = False

            retVal = checkType(value, intValLB, dblValHB)
            If (retVal = -1) Then Return False

            Dim rangeVals() As String
            rangeVals = rvalue.Split(";")
            If (rangeVals.Count < 1) Then Return False

            retVal = checkType(rangeVals(0), intValLB, dblValLB)
            If (retVal = -1) Then Return False

            retVal = checkType(rangeVals(1), intValHB, dblValHB)
            If (retVal = -1) Then Return False

            Select Case retVal
                Case 2
                    cmpOK = compareVal2(CDbl(value), IIf(dblValLB < dblValHB, dblValLB, dblValHB),
                                        IIf(dblValLB > dblValHB, dblValLB, dblValHB))
                Case 1
                    cmpOK = compareVal2(CInt(value), IIf(intValLB < intValHB, intValLB, intValHB),
                                        IIf(intValLB > intValHB, intValLB, intValHB))
                Case Else
                    Return False
            End Select

            If (cmpOK) Then Return True

            Return False

        Catch ex As Exception
            MsgBox("Exception in doRange" & vbCrLf & ex.Message)
            Return False
        End Try
    End Function

    Private Function compareVal(value As Double, lowBnd As Double, hghBnd As Double) As Boolean
        Try
            If (value > (value - lowBnd) And value < (value + hghBnd)) Then Return True
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function compareVal(value As Integer, lowBnd As Integer, hghBnd As Integer) As Boolean
        Try
            If (value > (value - lowBnd) And value < (value + hghBnd)) Then Return True
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function compareVal2(value As Double, lowBnd As Double, hghBnd As Double) As Boolean
        Try
            If (value > lowBnd And value < hghBnd) Then Return True
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function compareVal2(value As Integer, lowBnd As Integer, hghBnd As Integer) As Boolean
        Try
            If (value > lowBnd And value < hghBnd) Then Return True
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function checkType(tolValue As String, ByRef intVal As Integer, ByRef dblVal As Double) As Integer
        Try
            If (tolValue.Contains(".") Or tolValue.Contains(",")) Then 'it is an double
                Double.TryParse(tolValue, dblVal)
                Return 2
            Else    'it si an integer
                Integer.TryParse(tolValue, intVal)
                Return 1
            End If
        Catch ex As Exception
            Return -1
        End Try
    End Function

    Private Function getStudValue(name As String, section As String) As String
        Try
            Dim tmpAttrib As Attribute
            Dim lstAttribs As List(Of Attribute)
            Dim activeRules As New List(Of TableData)
            Dim allProps As System.Reflection.PropertyInfo() = sPart.GetType().GetProperties()
            For Each p As System.Reflection.PropertyInfo In allProps
                If (p.CanRead) Then
                    If (p.PropertyType.Name.Equals("Attribute") And equstr("part", section)) Then
                        tmpAttrib = p.GetValue(sPart)
                        If (equstr(tmpAttrib.name, name)) Then
                            Return tmpAttrib.value
                        End If
                    ElseIf (p.PropertyType.Name.Contains("List")) Then
                        If (equstr(p.Name, "envelope") And equstr("envelope", section)) Then
                            lstAttribs = p.GetValue(sPart)
                            For Each tmpAttrib In lstAttribs
                                If (equstr(tmpAttrib.name, name)) Then
                                    Return tmpAttrib.value
                                End If
                            Next
                        ElseIf (equstr(p.Name, "properties") And equstr("properties", section)) Then
                            lstAttribs = p.GetValue(sPart)
                            For Each tmpAttrib In lstAttribs
                                If (equstr(tmpAttrib.name, name)) Then
                                    Return tmpAttrib.value
                                End If
                            Next
                        ElseIf (equstr(p.Name, "features") And (equstr("features", section) Or equstr("sketches", section) _
                                    Or equstr("geometries", section) Or equstr("geometry", section) Or equstr("relations", section))) Then
                            Dim lstFeatures As List(Of Feature)
                            lstFeatures = p.GetValue(sPart)
                            Dim aFeature As Feature
                            For Each aFeature In lstFeatures
                                If (equstr(aFeature.name.name, name)) Then
                                    Return aFeature.name.value
                                End If
                                If (equstr(aFeature.type.name, name)) Then
                                    Return aFeature.type.value
                                End If
                                If (equstr(aFeature.base.name, name)) Then
                                    Return aFeature.base.value
                                End If
                                If (equstr(aFeature.frozen.name, name)) Then
                                    Return aFeature.frozen.value
                                End If
                                If (equstr(aFeature.suppressed.name, name)) Then
                                    Return aFeature.suppressed.value
                                End If
                                If (aFeature.attribs.Count > 0) Then
                                    lstAttribs = aFeature.attribs
                                    For Each tmpAttrib In lstAttribs
                                        If (equstr(tmpAttrib.name, name)) Then
                                            Return tmpAttrib.value
                                        End If
                                    Next
                                End If
                                If (aFeature.dimensions.Count > 0 And equstr("dimensions", section)) Then
                                    Dim lstDimensions As List(Of Dimension)
                                    lstDimensions = aFeature.dimensions
                                    For Each aDimension As Dimension In lstDimensions
                                        If (equstr(aDimension.name.name, name)) Then
                                            Return aDimension.name.value
                                        End If
                                        If (equstr(aDimension.type.name, name)) Then
                                            Return aDimension.type.value
                                        End If
                                        If (equstr(aDimension.value.name, name)) Then
                                            Return aDimension.value.value
                                        End If
                                    Next
                                End If
                                If (aFeature.sketches.Count > 0 And (equstr("sketches", section) _
                                    Or equstr("geometries", section) Or equstr("geometry", section) _
                                    Or equstr("relations", section))) Then
                                    Dim lstSketches As List(Of Sketch)
                                    lstSketches = aFeature.sketches
                                    For Each aSketch As Sketch In lstSketches
                                        If (equstr(aSketch.name.name, name)) Then
                                            Return aSketch.name.value
                                        End If
                                        If (equstr(aSketch.type.name, name)) Then
                                            Return aSketch.type.value
                                        End If
                                        If (equstr(aSketch.status.name, name)) Then
                                            Return aSketch.status.value
                                        End If
                                        If (equstr(aSketch.constrained.name, name)) Then
                                            Return aSketch.constrained.value
                                        End If
                                        If (aSketch.geometries.Count > 0 And (equstr("geometries", section) _
                                            Or equstr("geometry", section) Or equstr("relations", section))) Then
                                            Dim lstGeometries As List(Of Geometry)
                                            lstGeometries = aSketch.geometries
                                            For Each aGeometry As Geometry In lstGeometries
                                                If (equstr(aGeometry.name.name, name)) Then
                                                    Return aGeometry.name.value
                                                End If

                                                If (aGeometry.attribs.Count > 0 And equstr("geometry", section)) Then
                                                    lstAttribs = aGeometry.attribs
                                                    For Each tmpAttrib In lstAttribs
                                                        If (equstr(tmpAttrib.name, name)) Then
                                                            Return tmpAttrib.value
                                                        End If
                                                    Next
                                                End If

                                                If (aGeometry.relations.Count > 0 And equstr("relations", section)) Then
                                                    lstAttribs = aGeometry.relations
                                                    For Each tmpAttrib In lstAttribs
                                                        If (equstr(tmpAttrib.name, name)) Then
                                                            Return tmpAttrib.value
                                                        End If
                                                    Next
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                            Next
                        ElseIf (equstr(p.Name, "equations") And equstr("equations", section)) Then
                            Dim lstEquations As List(Of Equation)
                            lstEquations = p.GetValue(sPart)
                            For Each aEquation As Equation In lstEquations
                                If (equstr(aEquation.name.name, name)) Then
                                    Return aEquation.name.value
                                End If
                                If (equstr(aEquation.value.name, name)) Then
                                    Return aEquation.value.value
                                End If
                                If (equstr(aEquation.status.name, name)) Then
                                    Return aEquation.status.value
                                End If
                                If (equstr(aEquation.isglobal.name, name)) Then
                                    Return aEquation.isglobal.value
                                End If
                            Next
                        End If
                    End If
                End If
            Next
            Return ""
        Catch ex As Exception
            MsgBox("Exception in getStudValue" & vbCrLf & ex.Message)
            Return ""
        End Try
    End Function

    Private Sub doPDF()
        Dim pdf As PdfDocument = New PdfDocument
        pdf.Info.Title = "My First PDF"
        Dim pdfPage As PdfPage = pdf.AddPage
        Dim graph As XGraphics = XGraphics.FromPdfPage(pdfPage)
        Dim font As XFont = New XFont("Verdana", 12, XFontStyle.Bold)
        Dim aRow As String
        Dim xPos As Double = 40.0
        Dim yPos As Double = 50.0
        For Each aRule As reportItem In reportStr
            aRow = aRule.nameProp + "," + aRule.pointsProp + aRule.orgValProp + "," + aRule.masterValProp + "," + aRule.dscrProp
            graph.DrawString(aRow, font, XBrushes.Black,
        New XRect(xPos, yPos, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.TopLeft)
            yPos = yPos + 20
        Next

        'graph.DrawString("This is my first PDF document", font, XBrushes.Black,
        'New XRect(0, 0, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.Center)

        Dim pdfFilename As String = "G:\testing\masterrule\testpage.pdf"
        pdf.Save(pdfFilename)
        Process.Start(pdfFilename)
    End Sub
End Class

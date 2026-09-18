Option Explicit

' CATIA V5 / CATVBA - Compare Module 02 - v0.1.2
' Run M02_Baslat while the successful Compare_M01 result is active.
' Creates unsaved, independent A/B snapshot CATParts and an unsaved preview CATProduct.
' Each accepted source Body is pasted As Result without link into its own destination Body.
' The complete occurrence transform is applied to the independent destination Body.
' Volume, center of gravity and rigid transform quality are validated.
' No source document and no generated document is saved automatically.

Private Const M02_TITLE As String = "CATIA Compare - Modul 02 | v0.1.2"
Private Const M02_M01_PREFIX As String = "Compare_M01_"
Private Const M02_PREFIX As String = "Compare_M02_"
Private Const M02_GROUP_A As String = "A_ORIGINAL"
Private Const M02_GROUP_B As String = "B_REVISED"
Private Const M02_OPEN As Long = 0
Private Const M02_SHOW As Long = 0
Private Const M02_MAX_DEPTH As Long = 100
Private Const M02_MAX_NODES As Long = 100000
Private Const M02_MATRIX_TOL As Double = 0.000001
Private Const M02_COG_TOL_MM As Double = 0.001
Private Const M02_VOL_REL_TOL As Double = 0.000001
Private Const M02_VOL_ABS_TOL_MM3 As Double = 0.001

Private Type M02_Stats
    Nodes As Long
    PartInstances As Long
    CandidateBodies As Long
    CopiedBodies As Long
    SkippedBodies As Long
    VerifiedBodies As Long
    Errors As Long
    Warnings As Long
End Type

Private Type M02_Measure
    VolumeMM3 As Double
    CogX As Double
    CogY As Double
    CogZ As Double
End Type

Private m02App As Object
Private m02Log As Collection
Private m02Running As Boolean
Private m02Stage As String
Private m02RunId As String
Private m02Summary As String
Private m02FirstIssue As String
Private m02ReportPath As String

Public Sub M02_V011_Baslat()
    M02_Baslat
End Sub

Public Sub M02_V012_Baslat()
    M02_Baslat
End Sub

Public Sub M02_Baslat()
    Dim m01Doc As Object, m01Root As Object
    Dim groupA As Object, groupB As Object
    Dim snapADoc As Object, snapBDoc As Object, previewDoc As Object
    Dim statsA As M02_Stats, statsB As M02_Stats
    Dim groupAM(11) As Double, groupBM(11) As Double
    Dim ready As Boolean, viewOK As Boolean, statusText As String
    Dim eNumber As Long, eText As String, eStage As String

    If m02Running Then Exit Sub
    On Error GoTo FatalError
    m02Running = True
    Set m02App = CATIA
    Set m02Log = New Collection
    m02RunId = Format$(Now, "yyyymmdd_hhnnss") & "_" & _
                 Format$(CLng(Timer * 1000#) Mod 100000, "00000")
    m02Summary = ""
    m02FirstIssue = ""
    m02ReportPath = ""
    M02_Log M02_TITLE & " | " & m02RunId

    M02_SetStage "01 - Modul 01 sonucunun okunmasi"
    Set m01Doc = m02App.ActiveDocument
    If TypeName(m01Doc) <> "ProductDocument" Then
        Err.Raise vbObjectError + 2201, M02_TITLE, _
                  "Basarili Compare_M01 CATProduct sonucunu aktif yap."
    End If
    Set m01Root = m01Doc.Product
    If Left$(CStr(m01Root.PartNumber), Len(M02_M01_PREFIX)) <> M02_M01_PREFIX Then
        Err.Raise vbObjectError + 2202, M02_TITLE, _
                  "Aktif Product, Modul 01 sonucu degil."
    End If
    Set groupA = m01Root.Products.Item(M02_GROUP_A)
    Set groupB = m01Root.Products.Item(M02_GROUP_B)
    If groupA.Products.Count = 0 Or groupB.Products.Count = 0 Then
        Err.Raise vbObjectError + 2203, M02_TITLE, _
                  "Modul 01 A/B kaynak gruplari bos."
    End If

    M02_ReadPosition groupA, groupAM
    M02_CheckRigid groupAM, M02_GROUP_A
    M02_ReadPosition groupB, groupBM
    M02_CheckRigid groupBM, M02_GROUP_B

    M02_SetStage "02 - Baglantisiz snapshot CATPart'larinin olusturulmasi"
    Set snapADoc = M02_NewSnapshotPart("M02_A_SNAPSHOT_" & m02RunId)
    Set snapBDoc = M02_NewSnapshotPart("M02_B_SNAPSHOT_" & m02RunId)

    M02_SetStage "03A - A snapshot"
    M02_ScanChildren groupA, groupAM, "A", snapADoc, statsA, 0

    M02_SetStage "03B - B snapshot"
    M02_ScanChildren groupB, groupBM, "B", snapBDoc, statsB, 0

    M02_SetStage "04 - Snapshot kontrol montaji"
    Set previewDoc = M02_CreatePreview(snapADoc, snapBDoc, viewOK)

    ready = (statsA.CopiedBodies > 0 And statsB.CopiedBodies > 0)
    ready = ready And (statsA.CopiedBodies = statsA.VerifiedBodies)
    ready = ready And (statsB.CopiedBodies = statsB.VerifiedBodies)
    ready = ready And (statsA.Errors = 0 And statsB.Errors = 0)
    If Not viewOK Then ready = False

    If ready Then
        If statsA.Warnings + statsB.Warnings = 0 Then
            statusText = "MODUL 02: SNAPSHOT VE YERLESIM KONTROLU TAMAM"
        Else
            statusText = "MODUL 02: SNAPSHOT TAMAM, UYARILARI INCELE"
        End If
    Else
        statusText = "MODUL 02: KONTROL EKSIK - RAPORU INCELE"
    End If

    M02_SetStage "05 - Sonuc raporu"
    m02Summary = statusText & vbCrLf & vbCrLf & _
                 M02_StatsText("A - Orijinal", statsA) & vbCrLf & vbCrLf & _
                 M02_StatsText("B - Revize", statsB) & vbCrLf & vbCrLf & _
                 "Kopyalar: As Result / baglantisiz." & vbCrLf & _
                 "Kontrol: hacim + global agirlik merkezi + rigid donusum." & vbCrLf & _
                 "Belgeler kaydedilmedi."
    M02_Log m02Summary
    M02_FinishReport
    If Not previewDoc Is Nothing Then
        previewDoc.Activate
        M02_TryReframe
    End If
    m02Running = False
    If ready Then
        MsgBox M02_SummaryWithPath(), vbInformation, M02_TITLE
    Else
        MsgBox M02_SummaryWithPath(), vbExclamation, M02_TITLE
    End If
    Exit Sub

FatalError:
    eNumber = Err.Number
    eText = Err.Description
    eStage = m02Stage
    m02Summary = "MODUL 02 DURDU" & vbCrLf & _
                 "Asama: " & eStage & vbCrLf & _
                 "Hata: " & CStr(eNumber) & vbCrLf & eText
    M02_Log m02Summary
    M02_FinishReport
    m02Running = False
    MsgBox M02_SummaryWithPath(), vbExclamation, M02_TITLE
End Sub

Private Function M02_NewSnapshotPart(ByVal partNumber As String) As Object
    Dim doc As Object, part As Object, mainBody As Object
    Set doc = m02App.Documents.Add("Part")
    Set part = doc.Part
    doc.Product.PartNumber = partNumber
    Set mainBody = part.MainBody
    mainBody.Name = "M02_EMPTY_CONTAINER"
    Set M02_NewSnapshotPart = doc
End Function

Private Sub M02_ScanChildren(ByVal parentOccurrence As Object, _
                             ByRef parentWorld() As Double, _
                             ByVal pathText As String, _
                             ByVal destDoc As Object, _
                             ByRef stats As M02_Stats, _
                             ByVal depth As Long)
    Dim children As Object, child As Object
    Dim i As Long
    Set children = parentOccurrence.Products
    For i = 1 To children.Count
        Set child = children.Item(i)
        M02_ScanNode child, parentWorld, pathText & "/" & _
                     M02_Name(child) & "[" & CStr(i) & "]", _
                     destDoc, stats, depth + 1
    Next i
End Sub

Private Sub M02_ScanNode(ByVal occurrence As Object, _
                         ByRef parentWorld() As Double, _
                         ByVal pathText As String, _
                         ByVal destDoc As Object, _
                         ByRef stats As M02_Stats, _
                         ByVal depth As Long)
    Dim localM(11) As Double, worldM(11) As Double
    Dim children As Object, child As Object
    Dim referenceProduct As Object, owner As Object
    Dim i As Long, childCount As Long

    On Error GoTo Failed
    If depth > M02_MAX_DEPTH Or stats.Nodes >= M02_MAX_NODES Then
        M02_Issue stats, pathText, 0, _
                  "Tarama derinlik/dugum sinirina ulasti.", True
        Exit Sub
    End If
    stats.Nodes = stats.Nodes + 1
    M02_ReadPosition occurrence, localM
    M02_Compose parentWorld, localM, worldM
    M02_CheckRigid worldM, pathText

    Set children = occurrence.Products
    childCount = children.Count
    If childCount > 0 Then
        For i = 1 To childCount
            Set child = children.Item(i)
            M02_ScanNode child, worldM, pathText & "/" & _
                         M02_Name(child) & "[" & CStr(i) & "]", _
                         destDoc, stats, depth + 1
        Next i
        Exit Sub
    End If

    Set referenceProduct = occurrence.ReferenceProduct
    If Not referenceProduct Is Nothing Then Set owner = referenceProduct.Parent
    If owner Is Nothing Then
        M02_Issue stats, pathText, 0, "Referans belge yok.", True
    ElseIf TypeName(owner) = "PartDocument" Then
        stats.PartInstances = stats.PartInstances + 1
        M02_CopyPartBodies owner, worldM, pathText, destDoc, stats
    ElseIf occurrence.HasAMasterShapeRepresentation() Then
        M02_Issue stats, pathText, 0, _
                  "Erisilebilir CATPart yok; CGR/V4/alternatif temsil olabilir.", True
    Else
        M02_Issue stats, pathText, 0, "Bos CATPart olmayan yaprak.", False
    End If
    Exit Sub
Failed:
    M02_Issue stats, pathText, Err.Number, Err.Description, True
End Sub

Private Sub M02_CopyPartBodies(ByVal sourceDoc As Object, _
                               ByRef worldM() As Double, _
                               ByVal pathText As String, _
                               ByVal destDoc As Object, _
                               ByRef stats As M02_Stats)
    Dim sourcePart As Object, bodies As Object, body As Object
    Dim i As Long, bodyPath As String
    On Error GoTo Failed
    M02_ActivatePart sourceDoc
    Set sourcePart = sourceDoc.Part
    Set bodies = sourcePart.Bodies
    For i = 1 To bodies.Count
        Set body = bodies.Item(i)
        bodyPath = pathText & "/" & M02_Name(body) & "[" & CStr(i) & "]"
        If CBool(body.InBooleanOperation) Then
            stats.SkippedBodies = stats.SkippedBodies + 1
            M02_Log "ATLANDI Boolean girdisi: " & bodyPath
        ElseIf CBool(sourcePart.IsInactive(body)) Then
            stats.SkippedBodies = stats.SkippedBodies + 1
            M02_Log "ATLANDI pasif Body: " & bodyPath
        ElseIf body.Shapes.Count = 0 Then
            stats.SkippedBodies = stats.SkippedBodies + 1
            M02_Log "ATLANDI bos Body: " & bodyPath
        Else
            stats.CandidateBodies = stats.CandidateBodies + 1
            M02_CopyOneBody sourceDoc, body, worldM, bodyPath, destDoc, stats
        End If
    Next i
    Exit Sub
Failed:
    M02_Issue stats, "Part / " & pathText, Err.Number, Err.Description, True
End Sub

Private Sub M02_CopyOneBody(ByVal sourceDoc As Object, _
                            ByVal sourceBody As Object, _
                            ByRef worldM() As Double, _
                            ByVal bodyPath As String, _
                            ByVal destDoc As Object, _
                            ByRef stats As M02_Stats)
    Dim sourcePart As Object, destPart As Object, destBody As Object
    Dim sourceSelection As Object, destSelection As Object
    Dim sourceMeasure As M02_Measure, destMeasure As M02_Measure
    Dim expectedX As Double, expectedY As Double, expectedZ As Double
    Dim beforeShapes As Long, afterShapes As Long
    Dim stage As String, eNumber As Long, eText As String

    On Error GoTo Failed
    stage = "C01 kaynak olcum"
    M02_ActivatePart sourceDoc
    Set sourcePart = sourceDoc.Part
    M02_MeasureBody sourceDoc, sourcePart, sourceBody, sourceMeasure

    stage = "C02A hedef CATPart aktivasyonu"
    destDoc.Activate
    stage = "C02B hedef Part"
    Set destPart = destDoc.Part
    stage = "C02C yeni Body"
    Set destBody = destPart.Bodies.Add
    stage = "C02D Body adi"
    destBody.Name = M02_SafeName("SNAP_" & CStr(stats.CandidateBodies) & _
                                "_" & M02_Name(sourceBody))
    stage = "C02E Shape sayisi"
    beforeShapes = destBody.Shapes.Count

    stage = "C03 Copy"
    M02_ActivatePart sourceDoc
    Set sourceSelection = sourceDoc.Selection
    sourceSelection.Clear
    sourceSelection.Add sourceBody
    If sourceSelection.Count2 <> 1 Then
        Err.Raise vbObjectError + 2210, M02_TITLE, "Kaynak Body secilemedi."
    End If
    sourceSelection.Copy

    stage = "C04 PasteSpecial AsResult"
    destDoc.Activate
    Set destSelection = destDoc.Selection
    destSelection.Clear
    destSelection.Add destBody
    destSelection.PasteSpecial "CATPrtResultWithOutLink"
    destSelection.Clear
    destPart.Update
    afterShapes = destBody.Shapes.Count
    If afterShapes <= beforeShapes Then
        Err.Raise vbObjectError + 2211, M02_TITLE, _
                  "AsResult hedef Body icinde Shape olusturmadi."
    End If

    stage = "C05 global donusum"
    M02_CheckRigid worldM, bodyPath
    M02_ApplyBodyMove destBody, worldM
    destPart.Update
    stats.CopiedBodies = stats.CopiedBodies + 1

    stage = "C06 hedef olcum"
    M02_MeasureBody destDoc, destPart, destBody, destMeasure
    M02_TransformPoint worldM, sourceMeasure.CogX, sourceMeasure.CogY, _
                       sourceMeasure.CogZ, expectedX, expectedY, expectedZ

    stage = "C07 dogrulama"
    M02_VerifyMeasure sourceMeasure, destMeasure, expectedX, expectedY, _
                      expectedZ, bodyPath
    stats.VerifiedBodies = stats.VerifiedBodies + 1
    M02_Log "SNAPSHOT OK: " & bodyPath & _
            " | hacim(mm3)=" & Format$(destMeasure.VolumeMM3, "0.000000") & _
            " | COG(mm)=" & M02_VectorText(destMeasure.CogX, _
                                            destMeasure.CogY, destMeasure.CogZ)
    Exit Sub
Failed:
    eNumber = Err.Number
    eText = Err.Description
    On Error Resume Next
    If Not sourceSelection Is Nothing Then sourceSelection.Clear
    If Not destSelection Is Nothing Then destSelection.Clear
    On Error GoTo 0
    M02_Issue stats, stage & " / " & bodyPath, eNumber, eText, True
End Sub

Private Sub M02_MeasureBody(ByVal doc As Object, ByVal part As Object, _
                            ByVal body As Object, ByRef result As M02_Measure)
    Dim spa As Object, reference As Object, measurable As Object
    Dim cog(2) As Variant
    Dim stage As String, errorNumber As Long, errorText As String
    On Error GoTo Failed
    stage = "M01 SPAWorkbench"
    Set spa = doc.GetWorkbench("SPAWorkbench")
    If spa Is Nothing Then Err.Raise 91, M02_TITLE, "SPAWorkbench alinamadi."
    stage = "M02 Body referansi"
    Set reference = part.CreateReferenceFromObject(body)
    If reference Is Nothing Then Err.Raise 91, M02_TITLE, "Body referansi alinamadi."
    stage = "M03 Measurable"
    Set measurable = spa.GetMeasurable(reference)
    If measurable Is Nothing Then Err.Raise 91, M02_TITLE, "Measurable alinamadi."
    stage = "M04 Volume"
    result.VolumeMM3 = CDbl(measurable.Volume) * 1000000000#
    If result.VolumeMM3 <= 0# Then
        Err.Raise vbObjectError + 2212, M02_TITLE, "Pozitif kati hacim yok."
    End If
    stage = "M05 GetCOG Variant dizisi"
    measurable.GetCOG cog
    stage = "M06 COG donusumu"
    result.CogX = CDbl(cog(0))
    result.CogY = CDbl(cog(1))
    result.CogZ = CDbl(cog(2))
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Err.Raise errorNumber, M02_TITLE, stage & " | " & errorText
End Sub

Private Sub M02_VerifyMeasure(ByRef sourceM As M02_Measure, _
                              ByRef destM As M02_Measure, _
                              ByVal expectedX As Double, _
                              ByVal expectedY As Double, _
                              ByVal expectedZ As Double, _
                              ByVal pathText As String)
    Dim volumeTol As Double, volumeDelta As Double, cogDelta As Double
    volumeTol = M02_VOL_ABS_TOL_MM3
    If Abs(sourceM.VolumeMM3) * M02_VOL_REL_TOL > volumeTol Then
        volumeTol = Abs(sourceM.VolumeMM3) * M02_VOL_REL_TOL
    End If
    volumeDelta = Abs(destM.VolumeMM3 - sourceM.VolumeMM3)
    If volumeDelta > volumeTol Then
        Err.Raise vbObjectError + 2213, M02_TITLE, _
                  "Hacim uyusmuyor. Delta(mm3)=" & CStr(volumeDelta) & _
                  " tolerans=" & CStr(volumeTol)
    End If
    cogDelta = Sqr((destM.CogX - expectedX) ^ 2 + _
                   (destM.CogY - expectedY) ^ 2 + _
                   (destM.CogZ - expectedZ) ^ 2)
    If cogDelta > M02_COG_TOL_MM Then
        Err.Raise vbObjectError + 2214, M02_TITLE, _
                  "Global agirlik merkezi uyusmuyor. Delta(mm)=" & CStr(cogDelta)
    End If
    M02_Log "DOGRULAMA: " & pathText & _
            " | dV(mm3)=" & Format$(volumeDelta, "0.000000") & _
            " | dCOG(mm)=" & Format$(cogDelta, "0.000000")
End Sub

Private Function M02_CreatePreview(ByVal snapADoc As Object, _
                                   ByVal snapBDoc As Object, _
                                   ByRef viewOK As Boolean) As Object
    Dim doc As Object, root As Object, instanceA As Object, instanceB As Object
    Dim selection As Object
    Dim identity(11) As Double
    On Error GoTo Failed
    Set doc = m02App.Documents.Add("Product")
    Set root = doc.Product
    root.PartNumber = M02_PREFIX & m02RunId
    Set instanceA = root.Products.AddComponent(snapADoc.Product)
    instanceA.Name = "A_SNAPSHOT"
    Set instanceB = root.Products.AddComponent(snapBDoc.Product)
    instanceB.Name = "B_SNAPSHOT"
    M02_SetIdentity identity
    M02_SetOccurrencePosition instanceA, identity
    M02_SetOccurrencePosition instanceB, identity
    Set selection = doc.Selection
    selection.Clear
    selection.Add instanceA
    selection.VisProperties.SetRealColor 225, 70, 70, 1
    selection.VisProperties.SetRealOpacity 160, 1
    selection.VisProperties.SetShow M02_SHOW
    selection.Clear
    selection.Add instanceB
    selection.VisProperties.SetRealColor 55, 200, 95, 1
    selection.VisProperties.SetRealOpacity 160, 1
    selection.VisProperties.SetShow M02_SHOW
    selection.Clear
    viewOK = True
    Set M02_CreatePreview = doc
    Exit Function
Failed:
    viewOK = False
    M02_Log "GORUNUM HATASI: " & CStr(Err.Number) & " | " & Err.Description
    Set M02_CreatePreview = doc
End Function

Private Sub M02_ActivatePart(ByVal partDoc As Object)
    Dim activeDoc As Object, openedDoc As Object, fullPath As String
    If partDoc Is Nothing Then Err.Raise 91, M02_TITLE, "CATPart belgesi yok."
    On Error Resume Next
    partDoc.Activate
    Set activeDoc = m02App.ActiveDocument
    Err.Clear
    On Error GoTo 0
    If activeDoc Is partDoc Then Exit Sub
    fullPath = M02_DocumentPath(partDoc)
    If Len(fullPath) = 0 Then
        Err.Raise vbObjectError + 2215, M02_TITLE, _
                  "Kaynak CATPart penceresi acilamadi; dosya yolu yok."
    End If
    Set openedDoc = m02App.Documents.Open(fullPath)
    If Not openedDoc Is partDoc Then
        Err.Raise vbObjectError + 2216, M02_TITLE, _
                  "Acilan CATPart ayni oturum nesnesi degil."
    End If
    openedDoc.Activate
End Sub

Private Sub M02_ReadPosition(ByVal occurrence As Object, ByRef matrix() As Double)
    Dim raw(11) As Variant, i As Long
    occurrence.Position.GetComponents raw
    For i = 0 To 11
        matrix(i) = CDbl(raw(i))
    Next i
End Sub

Private Sub M02_SetIdentity(ByRef matrix() As Double)
    Dim i As Long
    For i = 0 To 11
        matrix(i) = 0#
    Next i
    matrix(0) = 1#
    matrix(4) = 1#
    matrix(8) = 1#
End Sub

Private Sub M02_ApplyBodyMove(ByVal body As Object, ByRef matrix() As Double)
    Dim raw(11) As Variant, i As Long
    For i = 0 To 11
        raw(i) = CDbl(matrix(i))
    Next i
    body.Move.Apply raw
End Sub

Private Sub M02_SetOccurrencePosition(ByVal occurrence As Object, _
                                      ByRef matrix() As Double)
    Dim raw(11) As Variant, i As Long
    For i = 0 To 11
        raw(i) = CDbl(matrix(i))
    Next i
    occurrence.Position.SetComponents raw
End Sub

Private Sub M02_Compose(ByRef parentM() As Double, ByRef localM() As Double, _
                        ByRef resultM() As Double)
    Dim col As Long, row As Long
    For col = 0 To 2
        For row = 0 To 2
            resultM(col * 3 + row) = _
                parentM(0 * 3 + row) * localM(col * 3 + 0) + _
                parentM(1 * 3 + row) * localM(col * 3 + 1) + _
                parentM(2 * 3 + row) * localM(col * 3 + 2)
        Next row
    Next col
    resultM(9) = parentM(0) * localM(9) + parentM(3) * localM(10) + _
                  parentM(6) * localM(11) + parentM(9)
    resultM(10) = parentM(1) * localM(9) + parentM(4) * localM(10) + _
                   parentM(7) * localM(11) + parentM(10)
    resultM(11) = parentM(2) * localM(9) + parentM(5) * localM(10) + _
                   parentM(8) * localM(11) + parentM(11)
End Sub

Private Sub M02_TransformPoint(ByRef matrix() As Double, _
                               ByVal x As Double, ByVal y As Double, ByVal z As Double, _
                               ByRef outX As Double, ByRef outY As Double, ByRef outZ As Double)
    outX = matrix(0) * x + matrix(3) * y + matrix(6) * z + matrix(9)
    outY = matrix(1) * x + matrix(4) * y + matrix(7) * z + matrix(10)
    outZ = matrix(2) * x + matrix(5) * y + matrix(8) * z + matrix(11)
End Sub

Private Sub M02_CheckRigid(ByRef matrix() As Double, ByVal pathText As String)
    Dim nx As Double, ny As Double, nz As Double
    Dim dxy As Double, dxz As Double, dyz As Double, determinant As Double
    nx = Sqr(matrix(0) ^ 2 + matrix(1) ^ 2 + matrix(2) ^ 2)
    ny = Sqr(matrix(3) ^ 2 + matrix(4) ^ 2 + matrix(5) ^ 2)
    nz = Sqr(matrix(6) ^ 2 + matrix(7) ^ 2 + matrix(8) ^ 2)
    dxy = matrix(0) * matrix(3) + matrix(1) * matrix(4) + matrix(2) * matrix(5)
    dxz = matrix(0) * matrix(6) + matrix(1) * matrix(7) + matrix(2) * matrix(8)
    dyz = matrix(3) * matrix(6) + matrix(4) * matrix(7) + matrix(5) * matrix(8)
    determinant = matrix(0) * (matrix(4) * matrix(8) - matrix(7) * matrix(5)) - _
                  matrix(3) * (matrix(1) * matrix(8) - matrix(7) * matrix(2)) + _
                  matrix(6) * (matrix(1) * matrix(5) - matrix(4) * matrix(2))
    If Abs(nx - 1#) > M02_MATRIX_TOL Or Abs(ny - 1#) > M02_MATRIX_TOL Or _
       Abs(nz - 1#) > M02_MATRIX_TOL Or Abs(dxy) > M02_MATRIX_TOL Or _
       Abs(dxz) > M02_MATRIX_TOL Or Abs(dyz) > M02_MATRIX_TOL Or _
       Abs(determinant - 1#) > M02_MATRIX_TOL Then
        Err.Raise vbObjectError + 2217, M02_TITLE, _
                  "Rigid olmayan/bozuk occurrence donusumu: " & pathText
    End If
End Sub

Private Function M02_StatsText(ByVal label As String, ByRef stats As M02_Stats) As String
    M02_StatsText = label & vbCrLf & _
        "Parca instance: " & CStr(stats.PartInstances) & _
        " | Aday Body: " & CStr(stats.CandidateBodies) & vbCrLf & _
        "Kopyalanan: " & CStr(stats.CopiedBodies) & _
        " | Dogrulanan: " & CStr(stats.VerifiedBodies) & vbCrLf & _
        "Atlanan: " & CStr(stats.SkippedBodies) & _
        " | Hata/Uyari: " & CStr(stats.Errors) & "/" & CStr(stats.Warnings)
End Function

Private Sub M02_Issue(ByRef stats As M02_Stats, ByVal location As String, _
                      ByVal number As Long, ByVal description As String, _
                      ByVal blocking As Boolean)
    Dim prefix As String, message As String
    If blocking Then
        stats.Errors = stats.Errors + 1
        prefix = "HATA"
    Else
        stats.Warnings = stats.Warnings + 1
        prefix = "UYARI"
    End If
    message = prefix & " | " & location & " | " & CStr(number) & " | " & description
    If Len(m02FirstIssue) = 0 Then m02FirstIssue = Left$(message, 240)
    M02_Log message
End Sub

Private Function M02_SummaryWithPath() As String
    Dim text As String
    text = m02Summary
    If Len(m02FirstIssue) > 0 Then text = text & vbCrLf & vbCrLf & m02FirstIssue
    If Len(m02ReportPath) > 0 Then
        text = text & vbCrLf & vbCrLf & "Ayrintili rapor:" & vbCrLf & m02ReportPath
    End If
    M02_SummaryWithPath = text
End Function

Private Sub M02_FinishReport()
    Dim folder As String, target As String
    Dim fso As Object, stream As Object, line As Variant
    On Error GoTo Failed
    If m02Log Is Nothing Then Exit Sub
    folder = Environ$("TEMP")
    If Len(folder) = 0 Then folder = Environ$("TMP")
    If Len(folder) = 0 Then Err.Raise vbObjectError + 2218, M02_TITLE, "TEMP yok."
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    target = folder & "CATIA_Compare_M02_" & m02RunId & ".txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(target, False, True)
    For Each line In m02Log
        stream.WriteLine CStr(line)
    Next line
    stream.Close
    m02ReportPath = target
    Exit Sub
Failed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    m02ReportPath = ""
End Sub

Private Sub M02_SetStage(ByVal text As String)
    m02Stage = text
    M02_Log "ASAMA: " & text
    On Error Resume Next
    m02App.StatusBar = M02_TITLE & " | " & text
    On Error GoTo 0
End Sub

Private Sub M02_Log(ByVal text As String)
    Debug.Print text
    If Not m02Log Is Nothing Then m02Log.Add text
End Sub

Private Sub M02_TryReframe()
    On Error Resume Next
    m02App.ActiveWindow.ActiveViewer.Reframe
    On Error GoTo 0
End Sub

Private Function M02_Name(ByVal obj As Object) As String
    On Error GoTo Failed
    M02_Name = CStr(obj.Name)
    Exit Function
Failed:
    M02_Name = "<name-unavailable>"
End Function

Private Function M02_DocumentPath(ByVal doc As Object) As String
    On Error GoTo Failed
    M02_DocumentPath = CStr(doc.FullName)
    Exit Function
Failed:
    M02_DocumentPath = ""
End Function

Private Function M02_SafeName(ByVal text As String) As String
    Dim bad As Variant
    For Each bad In Array("/", "\", ":", "*", "?", Chr$(34), "<", ">", "|")
        text = Replace(text, CStr(bad), "_")
    Next bad
    M02_SafeName = Left$(text, 70)
End Function

Private Function M02_VectorText(ByVal x As Double, ByVal y As Double, _
                                ByVal z As Double) As String
    M02_VectorText = Format$(x, "0.000000") & ";" & _
                     Format$(y, "0.000000") & ";" & _
                     Format$(z, "0.000000")
End Function

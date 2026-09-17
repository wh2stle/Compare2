Attribute VB_Name = "Compare_Modul01"
Option Explicit

' CATIA V5 / CATVBA - Module 01 - v1.0
' Import this file with File > Import File in the CATIA VBA editor.
' Run CATMain (or M01_Baslat). No UserForm or Windows API declaration needed.
' ASCII source + Windows CRLF: safe for the VBA editor's ANSI importer.
'
' PURPOSE
' Select two saved CATPart/CATProduct paths; reuse open documents if present.
' Build an unsaved inspection CATProduct and inspect every part occurrence.
' Confirm accessible, active, updated Part Design solids using a Reference.
' The report distinguishes verified solids from unreadable/unverified bodies.
'
' SCOPE
' Native Body objects in Parts, Bodies, GS and OGS containers are visited.
' Boolean operands are not counted separately from the resulting body.
' Empty and inactive bodies are reported and skipped. Hidden bodies are read.
' Standalone GSD geometry, CGR, V4 and alternate representations are not solids
' validated by this module. Their presence is reported for later review.
' Counts are per assembly occurrence: a bolt used twice counts twice.
' No bounding box, flattening, cubes, Boolean operation or difference test.
'
' SOURCE HANDLING
' No Save, SaveAs, geometry edit or Update call is made on a source document.
' In-memory edits in an already open document are used and reported.
' Design Mode may load source representations; normal CATIA alerts stay on.
' Only NEW wrapper products receive color/opacity/show changes.
' Only NEW top-level occurrences receive an identity placement.
' Descendant assembly placements and source visibility are left as loaded.
' The new assembly references the sources; it is not an isolated snapshot.
' Document origins coincide; physical alignment must be checked visually.
'
' ENTRY POINTS
' CATMain / M01_Baslat : run the complete module
' M01_SadeceA         : show A, hide B in the active M01 result
' M01_SadeceB         : show B, hide A in the active M01 result
' M01_IkisiniGoster   : show both
' M01_Rapor          : show the last summary and local report path
'
' API references (Dassault CAA Automation documentation, mirrored):
' https://catiadesign.org/_doc/V5Automation/generated/interfaces/ProductStructureInterfaces/interface_Products_41084.htm
' https://catiadesign.org/_doc/V5Automation/generated/interfaces/ProductStructureInterfaces/enum_CatWorkModeType_45790.htm
' https://catiadesign.org/_doc/V5Automation/generated/interfaces/MecModInterfaces/interface_Part_16738.htm
' https://catiadesign.org/_doc/V5Automation/generated/interfaces/MecModInterfaces/interface_OrderedGeometricalSet_37698.htm
' https://catiadesign.org/_doc/V5Automation/generated/interfaces/SpaceAnalysisInterfaces/interface_SPAWorkbench_36841.htm
' https://catiadesign.org/_doc/V5Automation/generated/interfaces/SpaceAnalysisInterfaces/interface_Measurable_34815.htm
' https://catiadesign.org/_doc/V5Automation/generated/interfaces/InfInterfaces/interface_VisPropertySet_21585.htm

Private Const M01_TITLE As String = "CATIA Compare - Modul 01"
Private Const M01_PREFIX As String = "Compare_M01_"
Private Const M01_GROUP_A As String = "A_ORIGINAL"
Private Const M01_GROUP_B As String = "B_REVISED"
Private Const M01_DESIGN_MODE As Long = 2
Private Const M01_OPEN As Long = 0
Private Const M01_SHOW As Long = 0
Private Const M01_HIDE As Long = 1
Private Const M01_MAX_DEPTH As Long = 100
Private Const M01_MAX_NODES As Long = 100000

Private Type M01_Stats
    Nodes As Long
    PartInstances As Long
    Bodies As Long
    Solids As Long
    EmptyBodies As Long
    InactiveBodies As Long
    BooleanOperands As Long
    UnverifiedBodies As Long
    NonBodyShapes As Long
    EmptyProducts As Long
    Errors As Long
    Warnings As Long
End Type

Private mApp As Object
Private mLog As Collection
Private mRunning As Boolean
Private mStage As String
Private mRunId As String
Private mLastSummary As String
Private mReportPath As String
Private mFirstIssue As String

Public Sub CATMain()
    M01_Baslat
End Sub

Public Sub M01_Baslat()
    Dim pathA As String, pathB As String
    Dim docA As Object, docB As Object, resultDoc As Object
    Dim resultRoot As Object, groupA As Object, groupB As Object
    Dim instanceA As Object, instanceB As Object
    Dim statsA As M01_Stats, statsB As M01_Stats
    Dim viewOK As Boolean, ready As Boolean
    Dim errorNumber As Long, errorText As String, errorStage As String
    Dim runStatus As String

    If mRunning Then Exit Sub
    On Error GoTo FatalError
    mRunning = True
    Set mApp = CATIA
    Set mLog = New Collection
    mStage = "00 - Baslangic"
    mRunId = Format$(Now, "yyyymmdd_hhnnss") & "_" & _
             Format$(CLng(Timer * 1000#) Mod 100000, "00000")
    mFirstIssue = ""
    mLastSummary = ""
    mReportPath = ""
    LogLine M01_TITLE & " | v1.0 | " & mRunId
    LogLine "Kapsam: native Part Design Body; sayilar instance bazindadir."
    LogLine "Otomatik hizalama yapilmaz; belge orijinleri ortak kullanilir."
    LogLine "Gizli Body'ler okunur; kaynak gorunurlukleri degistirilmez."

    SetStage "01 - Dosya secimi"
    pathA = PickModelPath("A - Ilk/orijinal CATPart veya CATProduct", "")
    If Len(pathA) = 0 Then GoTo Cancelled
    pathB = PickModelPath("B - Revize CATPart veya CATProduct", pathA)
    If Len(pathB) = 0 Then GoTo Cancelled
    LogLine "DOSYA A: " & pathA
    LogLine "DOSYA B: " & pathB

    SetStage "02A - Ilk modelin acilmasi"
    Set docA = OpenOrReuse(pathA)
    SetStage "02B - Revize modelin acilmasi"
    Set docB = OpenOrReuse(pathB)
    If docA Is docB Then
        Err.Raise vbObjectError + 2101, M01_TITLE, _
                  "CATIA iki secim icin ayni belgeyi dondurdu."
    End If

    SetStage "03 - Kontrol montajinin olusturulmasi"
    Set resultDoc = mApp.Documents.Add("Product")
    Set resultRoot = resultDoc.Product
    resultRoot.PartNumber = M01_PREFIX & mRunId
    Set groupA = resultRoot.Products.AddNewProduct(M01_GROUP_A)
    groupA.Name = M01_GROUP_A
    Set groupB = resultRoot.Products.AddNewProduct(M01_GROUP_B)
    groupB.Name = M01_GROUP_B

    SetStage "04A - A referansinin eklenmesi"
    Set instanceA = groupA.Products.AddComponent(docA.Product)
    SetStage "04B - B referansinin eklenmesi"
    Set instanceB = groupB.Products.AddComponent(docB.Product)
    SetStage "04C - Yeni kok instance konumlarinin ayarlanmasi"
    PutAtDocumentOrigin groupA
    PutAtDocumentOrigin groupB
    PutAtDocumentOrigin instanceA
    PutAtDocumentOrigin instanceB
    LogLine "Alt montaj konumlarina yazilmadi; kaynak referanslari korunuyor."

    SetStage "05A - A Design Mode ve geometri erisimi"
    EnsureDesignMode instanceA, statsA, "A"
    ScanNode instanceA, "A/" & ObjectName(instanceA), statsA, 0
    SetStage "05B - B Design Mode ve geometri erisimi"
    EnsureDesignMode instanceB, statsB, "B"
    ScanNode instanceB, "B/" & ObjectName(instanceB), statsB, 0

    SetStage "06 - Kontrol gorunumu"
    resultDoc.Activate
    viewOK = ApplyPreviewStyle(resultDoc, groupA, groupB)
    TryReframe resultDoc

    ready = (statsA.Solids > 0 And statsB.Solids > 0)
    ready = ready And (statsA.Errors = 0 And statsB.Errors = 0)
    ready = ready And (statsA.UnverifiedBodies = 0 And statsB.UnverifiedBodies = 0)
    If Not viewOK Then ready = False

    If ready Then
        If statsA.Warnings + statsB.Warnings = 0 Then
            runStatus = "MODUL 01: ERISIM KONTROLU TAMAM"
        Else
            runStatus = "MODUL 01: ERISIM TAMAM, UYARILARI INCELE"
        End If
    Else
        runStatus = "MODUL 01: KONTROL EKSIK - RAPORU INCELE"
    End If
    SetStage "07 - Sonuc raporu"
    mLastSummary = runStatus & vbCrLf & vbCrLf & _
                   StatsText("A - Orijinal", statsA) & vbCrLf & vbCrLf & _
                   StatsText("B - Revize", statsB) & vbCrLf & vbCrLf & _
                   "A: kirmizi | B: yesil. Orijinler ortak." & vbCrLf & _
                   "Yerlesimi gozle kontrol et; bu bir fark analizi degildir."
    LogLine mLastSummary
    LogLine "Gorunum: M01_SadeceA / M01_SadeceB / M01_IkisiniGoster"
    LogLine "Kaynak belgeler kaydedilmedi; sonuc montaji da kaydedilmedi."
    FinishReport
    mRunning = False
    SetStatus "Modul 01 tamamlandi."
    If ready Then
        MsgBox SummaryWithPath(), vbInformation, M01_TITLE
    Else
        MsgBox SummaryWithPath(), vbExclamation, M01_TITLE
    End If
    Exit Sub

Cancelled:
    LogLine "Kullanici dosya secimini iptal etti."
    mLastSummary = "Modul 01 iptal edildi."
    mRunning = False
    SetStatus "Modul 01 iptal edildi."
    Exit Sub

FatalError:
    errorNumber = Err.Number
    errorText = Err.Description
    errorStage = mStage
    mLastSummary = "MODUL 01 DURDU" & vbCrLf & _
                   "Asama: " & errorStage & vbCrLf & _
                   "Hata: " & CStr(errorNumber) & vbCrLf & errorText
    LogLine mLastSummary
    LogLine "Acik kaynaklar veya varsa kismi sonuc montaji kapatilmadi."
    FinishReport
    mRunning = False
    SetStatus "Modul 01 hata ile durdu."
    MsgBox SummaryWithPath(), vbExclamation, M01_TITLE
End Sub

' ---------- 01. File selection / opening ----------

Private Function PickModelPath(ByVal titleText As String, _
                               ByVal excludedPath As String) As String
    Dim chosenPath As String, suffix As String
    Do
        ' One wildcard works on CATIA versions with differing multi-filter syntax.
        chosenPath = mApp.FileSelectionBox(titleText, "*.CAT*", M01_OPEN)
        If Len(chosenPath) = 0 Then Exit Function
        suffix = LCase$(Mid$(chosenPath, InStrRev(chosenPath, ".") + 1))
        If suffix <> "catpart" And suffix <> "catproduct" Then
            MsgBox "Lutfen CATPart veya CATProduct sec.", vbExclamation, M01_TITLE
        ElseIf Len(excludedPath) > 0 And _
               StrComp(NormalPath(chosenPath), NormalPath(excludedPath), vbTextCompare) = 0 Then
            MsgBox "A ve B icin farkli dosyalar sec.", vbExclamation, M01_TITLE
        Else
            PickModelPath = chosenPath
            Exit Function
        End If
    Loop
End Function

Private Function OpenOrReuse(ByVal fullPath As String) As Object
    Dim documents As Object, candidate As Object
    Dim i As Long, foundPath As String
    Set documents = mApp.Documents
    For i = 1 To documents.Count
        Set candidate = documents.Item(i)
        foundPath = DocumentPath(candidate)
        If Len(foundPath) > 0 Then
            If StrComp(NormalPath(foundPath), NormalPath(fullPath), vbTextCompare) = 0 Then
                CheckModelDocument candidate, fullPath
                LogLine "ACIK BELGE KULLANILDI: " & foundPath
                NoteSavedState candidate
                Set OpenOrReuse = candidate
                Exit Function
            End If
        End If
    Next i
    Set candidate = documents.Open(fullPath)
    CheckModelDocument candidate, fullPath
    foundPath = DocumentPath(candidate)
    If Len(foundPath) = 0 Then
        Err.Raise vbObjectError + 2102, M01_TITLE, "Acilan belgenin yolu okunamadi."
    End If
    If StrComp(NormalPath(foundPath), NormalPath(fullPath), vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 2103, M01_TITLE, _
                  "CATIA beklenen dosyayi acmadi." & vbCrLf & _
                  "Istenen: " & fullPath & vbCrLf & "Donen: " & foundPath
    End If
    LogLine "BELGE ACILDI: " & foundPath
    NoteSavedState candidate
    Set OpenOrReuse = candidate
End Function

Private Sub CheckModelDocument(ByVal modelDoc As Object, ByVal fullPath As String)
    If modelDoc Is Nothing Then
        Err.Raise vbObjectError + 2104, M01_TITLE, "Belge acilamadi: " & fullPath
    End If
    If TypeName(modelDoc) <> "PartDocument" And TypeName(modelDoc) <> "ProductDocument" Then
        Err.Raise vbObjectError + 2105, M01_TITLE, _
                  "Desteklenmeyen belge tipi: " & TypeName(modelDoc)
    End If
End Sub

Private Sub NoteSavedState(ByVal modelDoc As Object)
    On Error GoTo CannotRead
    If Not modelDoc.Saved Then
        LogLine "BILGI: Kaydedilmemis oturum durumu kullaniliyor: " & ObjectName(modelDoc)
    End If
    Exit Sub
CannotRead:
    LogLine "BILGI: Saved durumu okunamadi: " & ObjectName(modelDoc)
End Sub

' ---------- 02. New result occurrences only ----------

Private Sub PutAtDocumentOrigin(ByVal newOccurrence As Object)
    Dim placement(11) As Variant, readback(11) As Variant
    Dim i As Long
    For i = 0 To 11
        placement(i) = 0#
    Next i
    placement(0) = 1#
    placement(4) = 1#
    placement(8) = 1#
    newOccurrence.Position.SetComponents placement
    newOccurrence.Position.GetComponents readback
    For i = 0 To 11
        If Abs(CDbl(readback(i)) - CDbl(placement(i))) > 0.000000001 Then
            Err.Raise vbObjectError + 2106, M01_TITLE, _
                      "Yeni instance orijine yerlestirilemedi: " & ObjectName(newOccurrence)
        End If
    Next i
End Sub

Private Sub EnsureDesignMode(ByVal occurrence As Object, _
                             ByRef stats As M01_Stats, ByVal label As String)
    On Error GoTo Failed
    occurrence.ApplyWorkMode M01_DESIGN_MODE
    LogLine "DESIGN MODE OK: " & label
    Exit Sub
Failed:
    AddIssue stats, "Design Mode / " & label, Err.Number, Err.Description, True
End Sub

' ---------- 03. Recursive assembly scan ----------

Private Sub ScanNode(ByVal occurrence As Object, ByVal treePath As String, _
                     ByRef stats As M01_Stats, ByVal depth As Long)
    Dim children As Object, child As Object, owner As Object
    Dim referenceProduct As Object
    Dim childCount As Long, i As Long
    Dim why As String, ownerError As Long, ownerText As String
    On Error GoTo Failed
    If depth > M01_MAX_DEPTH Or stats.Nodes >= M01_MAX_NODES Then
        AddIssue stats, treePath, 0, "Tarama derinlik/dugum sinirina ulasti; sonuc eksik.", True
        Exit Sub
    End If
    stats.Nodes = stats.Nodes + 1
    mStage = "05 - Instance taramasi: " & treePath
    SetStatus "Modul 01: " & Left$(treePath, 180)
    If stats.Nodes Mod 25 = 0 Then DoEvents
    ReadPlacement occurrence, treePath, stats

    Set children = occurrence.Products
    childCount = children.Count
    If childCount > 0 Then
        LogLine "MONTAJ: " & treePath & " | alt instance=" & CStr(childCount)
        CheckAssemblyRepresentation occurrence, treePath, stats
        For i = 1 To childCount
            Set child = Nothing
            why = ""
            If GetItem(children, i, child, why) Then
                ScanNode child, treePath & "/" & ObjectName(child) & _
                         "[" & CStr(i) & "]", stats, depth + 1
            Else
                AddIssue stats, treePath & "/child[" & CStr(i) & "]", 0, why, True
            End If
        Next i
        Exit Sub
    End If

    ' Resolve the leaf occurrence's PartDocument, not the active document.
    On Error Resume Next
    Err.Clear
    Set referenceProduct = occurrence.ReferenceProduct
    If Not referenceProduct Is Nothing Then Set owner = referenceProduct.Parent
    ownerError = Err.Number
    ownerText = Err.Description
    Err.Clear
    On Error GoTo Failed
    If ownerError <> 0 Or owner Is Nothing Then
        AddIssue stats, treePath, ownerError, _
                 "Referans belgeye erisilemiyor. " & ownerText, True
        Exit Sub
    End If
    If TypeName(owner) = "PartDocument" Then
        stats.PartInstances = stats.PartInstances + 1
        ScanPart owner, treePath, stats
    ElseIf occurrence.HasAMasterShapeRepresentation() Then
        AddIssue stats, treePath, 0, _
                 "Temsil var ama erisilebilir CATPart yok (CGR/V4/eksik link olabilir).", True
    Else
        stats.EmptyProducts = stats.EmptyProducts + 1
        AddIssue stats, treePath, 0, _
                 "CATPart olmayan yaprak Product; bos veya eksik baglantili olabilir.", True
    End If
    Exit Sub
Failed:
    AddIssue stats, treePath, Err.Number, Err.Description, True
End Sub

Private Sub CheckAssemblyRepresentation(ByVal occurrence As Object, _
                                        ByVal treePath As String, ByRef stats As M01_Stats)
    On Error GoTo Failed
    If occurrence.HasAMasterShapeRepresentation() Then
        AddIssue stats, treePath, 0, _
                 "Montajin kendi master temsili var; sadece alt CATPart Body'leri tarandi.", False
    End If
    Exit Sub
Failed:
    AddIssue stats, "Montaj temsili / " & treePath, Err.Number, Err.Description, True
End Sub

Private Sub ReadPlacement(ByVal occurrence As Object, ByVal treePath As String, _
                          ByRef stats As M01_Stats)
    Dim components(11) As Variant, i As Long, text As String
    On Error GoTo Failed
    occurrence.Position.GetComponents components
    For i = 0 To 11
        If i > 0 Then text = text & "; "
        text = text & CStr(CDbl(components(i)))
    Next i
    LogLine "LOCAL_POSITION: " & treePath & " | " & text
    Exit Sub
Failed:
    AddIssue stats, "Konum / " & treePath, Err.Number, Err.Description, True
End Sub

Private Sub ScanPart(ByVal partDoc As Object, ByVal treePath As String, _
                     ByRef stats As M01_Stats)
    Dim part As Object, spa As Object
    Dim solidsBefore As Long, bodiesBefore As Long
    Dim spaError As Long, spaText As String
    On Error GoTo Failed
    Set part = partDoc.Part
    LogLine "PART: " & treePath & " | " & DocumentPath(partDoc)
    NoteSavedState partDoc
    On Error Resume Next
    Err.Clear
    Set spa = partDoc.GetWorkbench("SPAWorkbench")
    spaError = Err.Number
    spaText = Err.Description
    Err.Clear
    On Error GoTo Failed
    If spaError <> 0 Or spa Is Nothing Then
        AddIssue stats, "Olcum erisimi / " & treePath, spaError, _
                 "SPAWorkbench alinamadi. " & spaText, True
    End If
    solidsBefore = stats.Solids
    bodiesBefore = stats.Bodies
    ScanContainer part, "Part", part, spa, treePath, stats, 0
    If stats.Bodies = bodiesBefore Then
        AddIssue stats, treePath, 0, "Bu Part icinde native Body bulunamadi.", False
    ElseIf stats.Solids = solidsBefore Then
        AddIssue stats, treePath, 0, "Bu Part icin kati Body dogrulanamadi.", False
    End If
    Exit Sub
Failed:
    AddIssue stats, "Part erisimi / " & treePath, Err.Number, Err.Description, True
End Sub

' ---------- 04. Native Bodies inside Part / Body / GS / OGS ----------

Private Sub ScanContainer(ByVal container As Object, ByVal kind As String, _
                          ByVal part As Object, ByVal spa As Object, _
                          ByVal treePath As String, ByRef stats As M01_Stats, _
                          ByVal depth As Long)
    Dim looseCount As Long
    On Error GoTo Failed
    If depth > M01_MAX_DEPTH Then
        AddIssue stats, treePath, 0, "Govde agaci derinlik sinirina ulasti.", True
        Exit Sub
    End If

    ' Only request collections documented for this container type.
    If kind <> "Body" Then
        VisitCollection container, "Bodies", "Body", part, spa, treePath, stats, depth
    End If
    If kind = "Part" Or kind = "Body" Or kind = "HybridBody" Then
        VisitCollection container, "HybridBodies", "HybridBody", _
                        part, spa, treePath, stats, depth
    End If
    If kind = "Part" Or kind = "Body" Or kind = "OrderedGeometricalSet" Then
        VisitCollection container, "OrderedGeometricalSets", "OrderedGeometricalSet", _
                        part, spa, treePath, stats, depth
    End If
    If kind = "HybridBody" Or kind = "OrderedGeometricalSet" Then
        looseCount = container.HybridShapes.Count
        If looseCount > 0 Then
            stats.NonBodyShapes = stats.NonBodyShapes + looseCount
            AddIssue stats, treePath, 0, CStr(looseCount) & _
                     " govde disi GSD elemani var; kati sonuc olarak dogrulanmadi.", False
        End If
    End If
    Exit Sub
Failed:
    AddIssue stats, "Geometri agaci / " & treePath, Err.Number, Err.Description, True
End Sub

Private Sub VisitCollection(ByVal container As Object, ByVal propertyName As String, _
                            ByVal childKind As String, ByVal part As Object, _
                            ByVal spa As Object, ByVal treePath As String, _
                            ByRef stats As M01_Stats, ByVal depth As Long)
    Dim items As Object, item As Object
    Dim i As Long, count As Long
    Dim why As String, childPath As String
    On Error GoTo Failed
    Set items = CallByName(container, propertyName, VbGet)
    count = items.Count
    For i = 1 To count
        Set item = Nothing
        why = ""
        If GetItem(items, i, item, why) Then
            childPath = treePath & "/" & ObjectName(item) & "[" & CStr(i) & "]"
            If childKind = "Body" Then InspectBody part, spa, item, childPath, stats
            ScanContainer item, childKind, part, spa, childPath, stats, depth + 1
        Else
            AddIssue stats, treePath & "/" & propertyName & "[" & CStr(i) & "]", _
                     0, why, True
        End If
    Next i
    Exit Sub
Failed:
    AddIssue stats, treePath & "/" & propertyName, Err.Number, Err.Description, True
End Sub

Private Sub InspectBody(ByVal part As Object, ByVal spa As Object, _
                        ByVal body As Object, ByVal treePath As String, _
                        ByRef stats As M01_Stats)
    Dim reference As Object, measurable As Object
    Dim volumeM3 As Double
    On Error GoTo Unverified
    stats.Bodies = stats.Bodies + 1
    mStage = "05 - Body kontrolu: " & treePath
    If body.InBooleanOperation Then
        stats.BooleanOperands = stats.BooleanOperands + 1
        LogLine "ATLANDI (Boolean girdisi): " & treePath
        Exit Sub
    End If
    If part.IsInactive(body) Then
        stats.InactiveBodies = stats.InactiveBodies + 1
        LogLine "ATLANDI (pasif Body): " & treePath
        Exit Sub
    End If
    If body.Shapes.Count = 0 Then
        stats.EmptyBodies = stats.EmptyBodies + 1
        LogLine "ATLANDI (bos Body): " & treePath
        Exit Sub
    End If
    If Not part.IsUpToDate(body) Then
        Err.Raise vbObjectError + 2107, M01_TITLE, _
                  "Body guncel degil; otomatik Update yapilmadi."
    End If
    If spa Is Nothing Then
        Err.Raise vbObjectError + 2108, M01_TITLE, "Olcum arayuzu kullanilamiyor."
    End If
    Set reference = part.CreateReferenceFromObject(body)
    Set measurable = spa.GetMeasurable(reference)
    volumeM3 = CDbl(measurable.Volume)
    If volumeM3 <= 0# Then
        Err.Raise vbObjectError + 2109, M01_TITLE, "Body icin pozitif kati hacim yok."
    End If
    stats.Solids = stats.Solids + 1
    LogLine "KATI OK: " & treePath & " | hacim(mm^3)=" & _
            Format$(volumeM3 * 1000000000#, "0.000000")
    Exit Sub
Unverified:
    stats.UnverifiedBodies = stats.UnverifiedBodies + 1
    AddIssue stats, "Body dogrulanamadi / " & treePath, Err.Number, Err.Description, True
End Sub

' ---------- 05. Colors and visibility on NEW wrappers ----------

Private Function ApplyPreviewStyle(ByVal resultDoc As Object, _
                                    ByVal groupA As Object, ByVal groupB As Object) As Boolean
    Dim selection As Object
    On Error GoTo Failed
    Set selection = resultDoc.Selection
    selection.Clear
    ' Disable a possible color/opacity inheritance at the NEW result root.
    selection.Add resultDoc.Product
    selection.VisProperties.SetRealColor 200, 200, 200, 0
    selection.VisProperties.SetRealOpacity 255, 0
    selection.Clear
    selection.Add groupA
    selection.VisProperties.SetRealColor 225, 70, 70, 1
    selection.VisProperties.SetRealOpacity 160, 1
    selection.VisProperties.SetShow M01_SHOW
    selection.Clear
    selection.Add groupB
    selection.VisProperties.SetRealColor 55, 200, 95, 1
    selection.VisProperties.SetRealOpacity 160, 1
    selection.VisProperties.SetShow M01_SHOW
    selection.Clear
    ApplyPreviewStyle = True
    LogLine "RENK OK: A kirmizi / B yesil. Yalnizca yeni gruplar renklendirildi."
    Exit Function
Failed:
    LogLine "GORUNUM HATASI: " & CStr(Err.Number) & " | " & Err.Description
    If Len(mFirstIssue) = 0 Then mFirstIssue = "Yeni montaja renk uygulanamadi."
    ClearSelectionQuietly resultDoc
End Function

Public Sub M01_SadeceA()
    SetResultVisibility M01_SHOW, M01_HIDE
End Sub

Public Sub M01_SadeceB()
    SetResultVisibility M01_HIDE, M01_SHOW
End Sub

Public Sub M01_IkisiniGoster()
    SetResultVisibility M01_SHOW, M01_SHOW
End Sub

Private Sub SetResultVisibility(ByVal stateA As Long, ByVal stateB As Long)
    Dim doc As Object, root As Object, groupA As Object, groupB As Object
    Dim selection As Object
    Dim errorText As String
    If mRunning Then Exit Sub
    On Error GoTo Failed
    Set doc = CATIA.ActiveDocument
    If TypeName(doc) <> "ProductDocument" Then GoTo WrongDocument
    Set root = doc.Product
    If Left$(root.PartNumber, Len(M01_PREFIX)) <> M01_PREFIX Then GoTo WrongDocument
    Set groupA = root.Products.Item(M01_GROUP_A)
    Set groupB = root.Products.Item(M01_GROUP_B)
    Set selection = doc.Selection
    selection.Clear
    selection.Add groupA
    selection.VisProperties.SetShow stateA
    selection.Clear
    selection.Add groupB
    selection.VisProperties.SetShow stateB
    selection.Clear
    ' Keep the camera unchanged so A/B toggling exposes placement differences.
    Exit Sub
WrongDocument:
    MsgBox "Once Modul 01'in olusturdugu Compare_M01 montajini aktif yap.", _
           vbInformation, M01_TITLE
    Exit Sub
Failed:
    errorText = CStr(Err.Number) & " | " & Err.Description
    ClearSelectionQuietly doc
    MsgBox "Gorunum degistirilemedi: " & errorText, vbExclamation, M01_TITLE
End Sub

Private Sub TryReframe(ByVal doc As Object)
    On Error GoTo Failed
    doc.Activate
    mApp.ActiveWindow.ActiveViewer.Reframe
    Exit Sub
Failed:
    LogLine "UYARI: Reframe yapilamadi: " & CStr(Err.Number) & " | " & Err.Description
End Sub

Private Sub ClearSelectionQuietly(ByVal doc As Object)
    On Error Resume Next
    If Not doc Is Nothing Then doc.Selection.Clear
    On Error GoTo 0
End Sub

' ---------- 06. Diagnostics ----------

Private Sub AddIssue(ByRef stats As M01_Stats, ByVal location As String, _
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
    If Len(mFirstIssue) = 0 Then mFirstIssue = Left$(message, 220)
    LogLine message
End Sub

Private Function StatsText(ByVal label As String, ByRef stats As M01_Stats) As String
    StatsText = label & vbCrLf & _
                "Parca instance: " & CStr(stats.PartInstances) & _
                " | Body: " & CStr(stats.Bodies) & vbCrLf & _
                "Dogrulanan kati: " & CStr(stats.Solids) & _
                " | Dogrulanamayan: " & CStr(stats.UnverifiedBodies) & vbCrLf & _
                "Bos/pasif/Boolean girdisi: " & CStr(stats.EmptyBodies) & "/" & _
                CStr(stats.InactiveBodies) & "/" & CStr(stats.BooleanOperands) & vbCrLf & _
                "Hata: " & CStr(stats.Errors) & " | Uyari: " & CStr(stats.Warnings)
End Function

Public Sub M01_Rapor()
    If mRunning Then Exit Sub
    If Len(mLastSummary) = 0 Then
        MsgBox "Bu VBA oturumunda henuz rapor yok. Once CATMain'i calistir.", _
               vbInformation, M01_TITLE
    Else
        MsgBox SummaryWithPath(), vbInformation, M01_TITLE
    End If
End Sub

Private Function SummaryWithPath() As String
    Dim message As String
    message = mLastSummary
    If Len(mFirstIssue) > 0 Then message = message & vbCrLf & vbCrLf & mFirstIssue
    If Len(mReportPath) > 0 Then
        message = message & vbCrLf & vbCrLf & "Ayrintili rapor:" & vbCrLf & mReportPath
    ElseIf Not mLog Is Nothing Then
        message = message & vbCrLf & vbCrLf & "Ayrintilar: VBA Immediate Window (Ctrl+G)."
    End If
    SummaryWithPath = message
End Function

Private Sub FinishReport()
    Dim folder As String, target As String
    Dim fso As Object, stream As Object, line As Variant
    Dim errorText As String
    On Error GoTo Failed
    If mLog Is Nothing Then Exit Sub
    folder = Environ$("TEMP")
    If Len(folder) = 0 Then folder = Environ$("TMP")
    If Len(folder) = 0 Then
        Err.Raise vbObjectError + 2110, M01_TITLE, "Gecici rapor klasoru bulunamadi."
    End If
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    target = folder & "CATIA_Compare_M01_" & mRunId & ".txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Unicode preserves Turkish names/paths in the local diagnostic report.
    ' No overwrite: a previous diagnostic file must not be replaced.
    Set stream = fso.CreateTextFile(target, False, True)
    For Each line In mLog
        stream.WriteLine CStr(line)
    Next line
    stream.Close
    mReportPath = target
    Exit Sub
Failed:
    errorText = CStr(Err.Number) & " | " & Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    mReportPath = ""
    Debug.Print "Rapor dosyasi yazilamadi: " & errorText
    mLastSummary = mLastSummary & vbCrLf & "Rapor dosyasi yazilamadi: " & errorText
End Sub

Private Sub SetStage(ByVal text As String)
    mStage = text
    LogLine "ASAMA: " & text
    SetStatus M01_TITLE & " | " & text
End Sub

Private Sub SetStatus(ByVal text As String)
    On Error Resume Next
    If Not mApp Is Nothing Then mApp.StatusBar = text
    On Error GoTo 0
End Sub

Private Sub LogLine(ByVal text As String)
    Debug.Print text
    If Not mLog Is Nothing Then mLog.Add text
End Sub

' ---------- 07. Small read-only helpers ----------

Private Function ObjectName(ByVal obj As Object) As String
    On Error GoTo Unavailable
    ObjectName = CStr(obj.Name)
    Exit Function
Unavailable:
    ObjectName = "<name-unavailable>"
End Function

Private Function DocumentPath(ByVal doc As Object) As String
    On Error GoTo Unavailable
    DocumentPath = CStr(doc.FullName)
    Exit Function
Unavailable:
    DocumentPath = ""
End Function

Private Function NormalPath(ByVal fullPath As String) As String
    NormalPath = Replace(Trim$(fullPath), "/", "\")
End Function

Private Function GetItem(ByVal items As Object, ByVal index As Long, _
                         ByRef item As Object, ByRef why As String) As Boolean
    On Error GoTo Failed
    Set item = items.Item(index)
    If item Is Nothing Then
        why = "Koleksiyon bos nesne dondurdu."
    Else
        GetItem = True
    End If
    Exit Function
Failed:
    why = CStr(Err.Number) & " | " & Err.Description
End Function

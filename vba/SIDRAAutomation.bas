Attribute VB_Name = "SIDRAAutomation"
Option Explicit

Private mRunning As Boolean
Private mApi As SIDRASolutions_SI_API.ISIAPI

Private Function SessionAPI() As SIDRASolutions_SI_API.ISIAPI
    ' SIDRA 9 has a process-wide runtime. Reuse it for repeated button clicks.
    If mApi Is Nothing Then Set mApi = New SIDRASolutions_SI_API.SIAPI
    Set SessionAPI = mApi
End Function

Public Sub SIDRA_RunModels()
    Dim result As String
    If mRunning Then Exit Sub
    result = ExecuteSIDRA(False, "")
End Sub

Public Function SIDRA_TestRun(ByVal outputFolder As String) As String
    SIDRA_TestRun = ExecuteSIDRA(True, outputFolder)
End Function

Public Function SIDRA_VBAReady() As String
    SIDRA_VBAReady = "SIDRA 9 VBA ready"
End Function

Private Function SheetNamed(ByVal wanted As String) As Worksheet
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Worksheets
        If Normal(sheet.Name) = Normal(wanted) Then
            Set SheetNamed = sheet
            Exit Function
        End If
    Next sheet
    Err.Raise vbObjectError + 2100, , "Missing worksheet: " & wanted
End Function

Private Function Normal(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Then Err.Raise vbObjectError + 2101, , "An input contains an Excel error."
    Normal = LCase$(Application.Trim(CStr(v)))
End Function

Private Sub Need(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2102, , message
End Sub

Private Function Number(ByVal v As Variant, ByVal label As String) As Double
    Need Not IsError(v), label & ": the cell contains an Excel error."
    Need Not IsEmpty(v), label & ": enter a number, including 0 for zero demand."
    Need IsNumeric(v) And VarType(v) <> vbBoolean And VarType(v) <> vbString, label & ": enter a numeric value."
    Number = CDbl(v)
    Need Number >= 0, label & ": negative values are not allowed."
End Function

Private Function Dict() As Object
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = vbTextCompare
End Function

Private Function LegCode(ByVal label As Variant) As Long
    Select Case Replace(Replace(Normal(label), "-", ""), " ", "")
        Case "south": LegCode = 0
        Case "southeast": LegCode = 1
        Case "east": LegCode = 2
        Case "northeast": LegCode = 3
        Case "north": LegCode = 4
        Case "northwest": LegCode = 5
        Case "west": LegCode = 6
        Case "southwest": LegCode = 7
        Case Else: Err.Raise vbObjectError + 2103, , "Invalid SIDRA origin leg: " & CStr(label)
    End Select
End Function

Private Function DirectionColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal direction As String) As Long
    Dim c As Long
    For c = 4 To 7
        If Normal(ws.Cells(headerRow, c).Value2) = Normal(direction) Then
            DirectionColumn = c
            Exit Function
        End If
    Next c
    Err.Raise vbObjectError + 2104, , "Missing " & direction & " column on Inputs row " & headerRow
End Function

Private Function MappingFor(ByVal ws As Worksheet, ByVal junction As String) As Variant
    Dim r As Long, c As Long, origins(0 To 3) As Variant, used As Object
    Set used = Dict()
    For r = 6 To ws.Cells(ws.Rows.count, 9).End(xlUp).Row
        If Normal(ws.Cells(r, 9).Value2) = Normal(junction) Then
            For c = 0 To 3
                origins(c) = LegCode(ws.Cells(r, c + 10).Value2)
                Need Not used.Exists(CStr(origins(c))), junction & ": duplicate origin legs in the mapping table."
                used.Add CStr(origins(c)), True
            Next c
            MappingFor = origins
            Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 2105, , "Add " & junction & " to the approach mapping table in Inputs columns I:M."
End Function

Private Function ReadConfig() As Object
    Dim wsP As Worksheet, wsS As Worksheet, wsI As Worksheet
    Dim cfg As Object, scenario As Object, junction As Object, volumes As Object
    Dim scenarios As New Collection, peaks As New Collection, junctions As New Collection
    Dim usedCases As Object, usedPeaks As Object, usedJunctions As Object
    Dim r As Long, rr As Long, lastRow As Long, peakStart As Long, declared As Long
    Dim label As String, raw As String, n As Long, fso As Object, path As String
    Dim hv As Double, foundHV As Boolean, directions As Variant, totals(0 To 3) As Variant, d As Long, peak As Variant
    Set wsP = SheetNamed("Paths")
    On Error Resume Next
    Set wsS = SheetNamed("Sceanrios")
    On Error GoTo 0
    If wsS Is Nothing Then Set wsS = SheetNamed("Scenarios")
    Set wsI = SheetNamed("Inputs")
    Set cfg = Dict(): Set fso = CreateObject("Scripting.FileSystemObject")
    Set usedCases = Dict(): Set usedPeaks = Dict(): Set usedJunctions = Dict()
    For r = 1 To wsP.UsedRange.Row + wsP.UsedRange.Rows.count
        label = Normal(wsP.Cells(r, 2).Value2)
        If label = "sidra file directory" Then path = Trim$(CStr(wsP.Cells(r, 3).Value2))
        If label = "year" Then cfg.Add "ReferenceYear", wsP.Cells(r, 3).Value2
    Next r
    Need Len(path) > 0, "Paths: enter the SIDRA baseline file path."
    If Left$(path, 1) = Chr$(34) And Right$(path, 1) = Chr$(34) Then path = Mid$(path, 2, Len(path) - 2)
    If InStr(path, ":") = 0 And Left$(path, 2) <> "\\" Then path = fso.BuildPath(ThisWorkbook.path, path)
    Need fso.FileExists(path), "SIDRA baseline not found: " & path
    Need LCase$(fso.GetExtensionName(path)) = "sip9", "Choose a SIDRA 9 .sip9 baseline."
    cfg.Add "Project", path
    If Not cfg.Exists("ReferenceYear") Then cfg.Add "ReferenceYear", ""
    lastRow = wsS.UsedRange.Row + wsS.UsedRange.Rows.count
    For r = 1 To lastRow
        label = Normal(wsS.Cells(r, 2).Value2)
        If label = "sceanrios" Or label = "scenarios" Then
            declared = CLng(Number(wsS.Cells(r, 3).Value2, "Scenario count"))
            Need declared = Number(wsS.Cells(r, 3).Value2, "Scenario count"), "Scenario count must be a whole number."
        ElseIf Left$(label, 9) = "sceanrio " Or Left$(label, 9) = "scenario " Then
            n = CLng(Val(Mid$(label, 10)))
            Need n > 0 And Not usedCases.Exists(CStr(n)), "Invalid or duplicate scenario number."
            raw = Trim$(CStr(wsS.Cells(r, 3).Value2))
            Need Len(raw) > 0, "Scenario " & n & " has no year/label."
            Set scenario = Dict()
            scenario.Add "Number", n: scenario.Add "Label", raw
            scenarios.Add scenario: usedCases.Add CStr(n), True
        ElseIf label = "peaks" Then
            peakStart = r
        End If
    Next r
    Need declared > 0 And scenarios.count = declared, "Scenario count must match the number of scenario rows."
    Need peakStart > 0, "Scenarios: Peaks section is missing."
    For r = peakStart To lastRow
        raw = Trim$(CStr(wsS.Cells(r, 3).Value2))
        If Len(raw) > 0 Then
            Need Not usedPeaks.Exists(raw), "Duplicate peak: " & raw
            peaks.Add raw: usedPeaks.Add raw, True
        End If
    Next r
    Need peaks.count > 0, "Enter at least one peak."
    directions = Array("NB", "EB", "SB", "WB")
    lastRow = wsI.Cells(wsI.Rows.count, 3).End(xlUp).Row
    For r = 1 To Application.Max(lastRow, 4)
        If Replace(Normal(wsI.Cells(r, 1).Value2), " ", "") = "hv%" Then
            hv = Number(wsI.Cells(r, 2).Value2, "HV %")
            foundHV = True
        End If
        raw = Trim$(CStr(wsI.Cells(r, 2).Value2))
        If Left$(Normal(raw), 9) = "junction " Then
            Need Not usedJunctions.Exists(raw), "Duplicate junction block: " & raw
            Set junction = Dict(): Set volumes = Dict()
            junction.Add "Name", raw: junction.Add "Year", wsI.Cells(r, 1).Value2
            junction.Add "Origins", MappingFor(wsI, raw)
            rr = r + 1
            Do While rr <= lastRow
                If Len(Trim$(CStr(wsI.Cells(rr, 2).Value2))) > 0 Then Exit Do
                label = Trim$(CStr(wsI.Cells(rr, 3).Value2))
                If Len(label) > 0 Then
                    Need Not volumes.Exists(label), raw & ": duplicate peak " & label
                    For d = 0 To 3
                        totals(d) = Number(wsI.Cells(rr, DirectionColumn(wsI, r, CStr(directions(d)))).Value2, raw & " " & label & " " & directions(d))
                    Next d
                    volumes.Add label, Array(totals(0), totals(1), totals(2), totals(3))
                End If
                rr = rr + 1
            Loop
            For Each peak In peaks
                Need volumes.Exists(CStr(peak)), raw & ": missing input row for " & CStr(peak)
            Next peak
            junction.Add "Volumes", volumes
            junctions.Add junction: usedJunctions.Add raw, True
        End If
    Next r
    Need foundHV And hv <= 1, "Enter HV as a percentage, e.g. 2% (stored as 0.02)."
    Need junctions.count > 0, "Inputs has no junction blocks."
    cfg.Add "HV", hv: cfg.Add "Scenarios", scenarios: cfg.Add "Peaks", peaks: cfg.Add "Junctions", junctions
    Set ReadConfig = cfg
End Function

Private Function FindSite(ByVal project As SIDRASolutions_SI_API.ISIAPIProject, ByVal junction As String) As SIDRASolutions_SI_API.ISIAPISite
    Dim folder As SIDRASolutions_SI_API.ISIAPISiteFolder, site As SIDRASolutions_SI_API.ISIAPISite
    Dim f As Long, s As Long, count As Long
    For f = 0 To project.SiteFolders.count - 1
        Set folder = project.SiteFolders.Item_2(f)
        For s = 0 To folder.Sites.count - 1
            Set site = folder.Sites.Item_2(s)
            If Normal(site.Name) = Normal(junction) Then
                count = count + 1
                Set FindSite = site
            End If
        Next s
    Next f
    Need count = 1, "Expected one SIDRA Site named '" & junction & "'; found " & count & "."
End Function

Private Sub NoNetworks(ByVal project As SIDRASolutions_SI_API.ISIAPIProject)
    Dim f As Long
    For f = 0 To project.NetworkFolders.count - 1
        Need project.NetworkFolders.Item_2(f).Networks.count = 0, "Use independent Sites; this VBA version does not process Networks."
    Next f
End Sub

Private Sub AllocateVolumes(ByVal site As SIDRASolutions_SI_API.ISIAPISite, ByVal origins As Variant, ByVal totals As Variant, ByVal hv As Double, ByVal apply As Boolean, ByVal audit As Object, ByVal caseName As String)
    Dim mov As SIDRASolutions_SI_API.ISIAPIMovement_vehicle_od
    Dim lv As SIDRASolutions_SI_API.ISIAPIMovement_vehicle_od_mc, heavy As SIDRASolutions_SI_API.ISIAPIMovement_vehicle_od_mc
    Dim extra As SIDRASolutions_SI_API.ISIAPIMovement_vehicle_od_mc
    Dim d As Long, o As Long, i As Long, mc As Long, k As Long, found As Boolean
    Dim base(0 To 7) As Double, enabled(0 To 7) As Boolean
    Dim baseTotal As Double, share As Double, turnTotal As Double, lightCount As Double, heavyCount As Double, appliedTotal As Double
    Dim directions As Variant
    directions = Array("NB", "EB", "SB", "WB")
    Need site.Unittimeforvolumes = 60, site.Name & ": set Unit Time for Volumes to 60 minutes."
    Need site.MovementClasses.MovementClassExists(1) And site.MovementClasses.MovementClassExists(2), site.Name & ": light/heavy classes are missing."
    Need site.MovementClasses.item(1).Is_included And site.MovementClasses.item(2).Is_included, site.Name & ": include both light and heavy classes."
    For o = 0 To 7
        If site.Legs.LegExists(o) Then
            found = False
            For i = 0 To 3
                If CLng(origins(i)) = o Then found = True
            Next i
            Need found, site.Name & ": origin leg " & o & " is absent from the approach mapping."
        End If
    Next o
    For i = 0 To 3
        o = CLng(origins(i)): baseTotal = 0: appliedTotal = 0
        Need site.Legs.LegExists(o), site.Name & ": mapped origin leg " & o & " does not exist."
        For d = 0 To 7
            base(d) = 0: enabled(d) = False
            If site.MovementVehicleODs.MovementExists(o, d) Then
                Set mov = site.MovementVehicleODs.item(o, d)
                If mov.Exists Then
                    Need mov.MovementVehicleODMCs.MovementClassExists(1) And mov.MovementVehicleODMCs.MovementClassExists(2), site.Name & ": missing light/heavy class for a turn."
                    Set lv = mov.MovementVehicleODMCs.item(1): Set heavy = mov.MovementVehicleODMCs.item(2)
                    Need lv.Exists And heavy.Exists, site.Name & ": enable light and heavy vehicles for movement " & mov.Display_od_id
                    For mc = 3 To 12
                        If mov.MovementVehicleODMCs.MovementClassExists(mc) And site.MovementClasses.MovementClassExists(mc) Then
                            Set extra = mov.MovementVehicleODMCs.item(mc)
                            If extra.Exists And site.MovementClasses.item(mc).Is_included Then
                                Need Abs(extra.Volume) < 0.000001, site.Name & ": nonzero custom-class demand is not supported by these two-class inputs."
                            End If
                        End If
                    Next mc
                    enabled(d) = True: base(d) = CDbl(lv.Volume) + CDbl(heavy.Volume)
                    baseTotal = baseTotal + base(d)
                End If
            End If
        Next d
        Need baseTotal > 0 Or CDbl(totals(i)) = 0, site.Name & " " & directions(i) & ": no baseline turning proportions."
        If apply Then
            For d = 0 To 7
                If enabled(d) Then
                    share = 0
                    If baseTotal > 0 Then share = base(d) / baseTotal
                    turnTotal = CDbl(totals(i)) * share
                    lightCount = turnTotal * (1 - hv): heavyCount = turnTotal * hv
                    Set mov = site.MovementVehicleODs.item(o, d)
                    Set lv = mov.MovementVehicleODMCs.item(1): Set heavy = mov.MovementVehicleODMCs.item(2)
                    lv.Volume = CSng(lightCount): heavy.Volume = CSng(heavyCount)
                    Need Abs(CDbl(lv.Volume) - lightCount) <= Application.Max(0.0001, lightCount * 0.00001), "Light-vehicle volume was not accepted."
                    Need Abs(CDbl(heavy.Volume) - heavyCount) <= Application.Max(0.0001, heavyCount * 0.00001), "Heavy-vehicle volume was not accepted."
                    appliedTotal = appliedTotal + CDbl(lv.Volume) + CDbl(heavy.Volume)
                    audit.WriteLine Csv(caseName) & "," & Csv(directions(i)) & "," & o & "," & d & "," & Csv(mov.Display_od_id) & "," & Csv(totals(i)) & "," & Csv(share) & "," & Csv(lv.Volume) & "," & Csv(heavy.Volume) & "," & Csv(lv.Growth_rate) & "," & Csv(lv.Flow_scale) & "," & Csv(lv.Peak_flow_factor)
                End If
            Next d
            Need Abs(appliedTotal - CDbl(totals(i))) <= Application.Max(0.001, CDbl(totals(i)) * 0.00001), site.Name & ": approach allocation does not reconcile."
        End If
    Next i
End Sub

Private Sub KeepOnlySite(ByVal project As SIDRASolutions_SI_API.ISIAPIProject, ByVal chosen As SIDRASolutions_SI_API.ISIAPISite)
    Dim f As Long, s As Long, folder As SIDRASolutions_SI_API.ISIAPISiteFolder
    Dim other As SIDRASolutions_SI_API.ISIAPISite, chosenId As String
    chosenId = CStr(chosen.Site_id)
    For f = project.SiteFolders.count - 1 To 0 Step -1
        Set folder = project.SiteFolders.Item_2(f)
        For s = folder.Sites.count - 1 To 0 Step -1
            Set other = folder.Sites.Item_2(s)
            If CStr(other.Site_id) <> chosenId Then Need folder.RemoveSite(other), "Could not remove an unused Site from the copy."
        Next s
        If folder.Sites.count = 0 Then Need project.RemoveSiteFolder(folder), "Could not remove an empty folder from the copy."
    Next f
End Sub

Private Function SafeName(ByVal label As String) As String
    Dim item As Variant
    For Each item In Array("<", ">", ":", Chr$(34), "/", "\", "|", "?", "*")
        label = Replace(label, CStr(item), "_")
    Next item
    label = Trim$(label)
    Do While Len(label) > 0
        If Right$(label, 1) <> "." And Right$(label, 1) <> " " Then Exit Do
        label = Left$(label, Len(label) - 1)
    Loop
    Need Len(label) > 0, "An output file/folder name is empty."
    SafeName = label
End Function

Private Function Csv(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    If IsNumeric(v) And VarType(v) <> vbString Then s = Replace(s, Application.International(xlDecimalSeparator), ".")
    Csv = Chr$(34) & Replace(s, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function Diagnostics(ByVal site As SIDRASolutions_SI_API.ISIAPISite) As String
    Dim i As Long
    For i = 0 To site.DiagnosticMsgs.count - 1
        If Len(Diagnostics) > 0 Then Diagnostics = Diagnostics & " | "
        Diagnostics = Diagnostics & site.DiagnosticMsgs.item(i).message
    Next i
End Function

Private Function ProcessCase(ByVal api As SIDRASolutions_SI_API.ISIAPI, ByVal cfg As Object, ByVal scenario As Object, ByVal junction As Object, ByVal peak As String, ByVal root As String, ByVal audit As Object, ByVal log As Object) As Variant
    Dim fso As Object, folder As String, stem As String, model As String, label As String
    Dim site As SIDRASolutions_SI_API.ISIAPISite, vehicle As SIDRASolutions_SI_API.ISIAPIOutputSiteVehicle
    Dim opened As Boolean, status As String, message As String, e As String
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.BuildPath(root, SafeName("Scenario " & scenario("Number") & "_" & scenario("Label")))
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    stem = SafeName(junction("Name") & "_Scenario " & scenario("Number") & "_" & peak)
    model = fso.BuildPath(folder, stem & ".sip9")
    label = junction("Name") & " " & scenario("Label") & " " & peak & " Peak"
    fso.CopyFile cfg("Project"), model, False
    Need api.OpenProject(model), "Cannot open copied model: " & api.LastErrorMessage
    opened = True
    Set site = FindSite(api.project, CStr(junction("Name")))
    KeepOnlySite api.project, site
    site.Name = stem: site.Title = label: site.siteFolder.Name = fso.GetFileName(folder)
    AllocateVolumes site, junction("Origins"), junction("Volumes")(peak), cfg("HV"), True, audit, label
    If site.Process Then
        Set vehicle = site.Outputset.OutputSiteVehicle
        status = "OK"
        If site.HasWarnings Then status = "WARNING"
        message = site.ProcessingWarnings & " " & Diagnostics(site)
        ProcessCase = Array(label, CDbl(vehicle.Deg_satn), CDbl(vehicle.Delay_control_average), CStr(vehicle.Level_of_service), status, model)
    Else
        status = "FAILED": message = site.ProcessingError & " " & Diagnostics(site)
        ProcessCase = Array(label, Empty, Empty, "FAILED", status, model)
    End If
    Need api.SubmitDataChanges, "Could not commit copied model: " & api.LastErrorMessage
    Need api.CloseProject, "Could not close copied model."
    opened = False
    log.WriteLine Csv(label) & "," & Csv(status) & "," & Csv(message) & "," & Csv(model)
    Exit Function
Failed:
    e = Err.Description
    On Error Resume Next
    If opened Then api.CloseProject
    ProcessCase = Array(label, Empty, Empty, "FAILED", "FAILED", model)
    log.WriteLine Csv(label) & ",FAILED," & Csv(e) & "," & Csv(model)
End Function



Private Function ExecuteSIDRA(ByVal quiet As Boolean, ByVal outputOverride As String) As String
    Dim cfg As Object, fso As Object, api As SIDRASolutions_SI_API.ISIAPI
    Dim wsI As Worksheet, wsO As Worksheet, junction As Variant, scenario As Variant, peak As Variant
    Dim site As SIDRASolutions_SI_API.ISIAPISite, root As String, preflight As String
    Dim audit As Object, log As Object, result As Variant, r As Long, i As Long, count As Long, failures As Long, warnings As Long
    Dim originalDate As Date, originalSize As Double, errText As String
    Dim oldEvents As Boolean, oldScreen As Boolean, oldCancel As Long, oldStatus As Variant
    If mRunning Then ExecuteSIDRA = "ERROR: a run is already active.": Exit Function
    mRunning = True
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    oldCancel = Application.EnableCancelKey: oldStatus = Application.StatusBar
    On Error GoTo Failed
    Application.EnableEvents = False: Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Set wsI = SheetNamed("Inputs"): Set wsO = SheetNamed("Outputs")
    wsI.Range("D1").Value2 = "Validating inputs..."
    wsO.Range("B3:G" & Application.Max(3, wsO.UsedRange.Row + wsO.UsedRange.Rows.count)).Hyperlinks.Delete
    wsO.Range("B3:G" & Application.Max(3, wsO.UsedRange.Row + wsO.UsedRange.Rows.count)).ClearContents
    wsI.Calculate
    Set cfg = ReadConfig()
    Need Not ThisWorkbook.ReadOnly, "Save a writable copy of this .xlsm workbook first."
    Need Len(ThisWorkbook.path) > 0, "Save this workbook to a local folder first."
    Set fso = CreateObject("Scripting.FileSystemObject")
    originalDate = fso.GetFile(cfg("Project")).DateLastModified
    originalSize = fso.GetFile(cfg("Project")).Size
    Set api = SessionAPI()
    Need api.IsLicensed, "SIDRA 9 is not licensed. Open SIDRA and check its licence."
    preflight = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName & ".sip9")
    fso.CopyFile cfg("Project"), preflight, False
    Need api.OpenProject(preflight), "Cannot open baseline copy: " & api.LastErrorMessage
    NoNetworks api.project
    For Each junction In cfg("Junctions")
        Set site = FindSite(api.project, CStr(junction("Name")))
        For Each peak In cfg("Peaks")
            AllocateVolumes site, junction("Origins"), junction("Volumes")(CStr(peak)), cfg("HV"), False, Nothing, ""
        Next peak
    Next junction
    Need api.CloseProject, "Could not close preflight copy."
    fso.DeleteFile preflight, True: preflight = ""
    If Len(outputOverride) > 0 Then
        root = outputOverride
    Else
        root = fso.BuildPath(ThisWorkbook.path, "SIDRA_Results_" & Format$(Now, "yyyymmdd_hhnnss"))
    End If
    Need Not fso.FolderExists(root) And Not fso.FileExists(root), "Output folder already exists; use a new folder: " & root
    fso.CreateFolder root
    Set audit = fso.CreateTextFile(fso.BuildPath(root, "Volume_Audit.csv"), False)
    Set log = fso.CreateTextFile(fso.BuildPath(root, "Run_Log.csv"), False)
    audit.WriteLine "Case,Direction,Origin,Destination,MovementID,ApproachTotal,TurnShare,AppliedLV,AppliedHV,GrowthRate,FlowScale,PeakFlowFactor"
    log.WriteLine "Case,Status,Message,Model"
    log.WriteLine Csv("Run information") & ",INFO," & Csv("Paths Year=" & cfg("ReferenceYear") & "; scenario years are labels; no additional growth; peak-specific inputs; HV=" & cfg("HV")) & "," & Csv(cfg("Project"))
    wsO.Range("F2").Value2 = "Status": wsO.Range("G2").Value2 = "Model"
    r = 3
    For Each scenario In cfg("Scenarios")
        For Each junction In cfg("Junctions")
            For Each peak In cfg("Peaks")
                count = count + 1
                Application.StatusBar = "SIDRA: " & junction("Name") & " / Scenario " & scenario("Number") & " / " & peak
                wsI.Range("D1").Value2 = "Running model " & count & "..."
                DoEvents
                result = ProcessCase(api, cfg, scenario, junction, CStr(peak), root, audit, log)
                For i = 0 To 4
                    wsO.Cells(r, i + 2).Value2 = result(i)
                Next i
                wsO.Hyperlinks.Add wsO.Cells(r, 7), CStr(result(5)), , , "Open SIDRA file"
                If result(4) = "FAILED" Then failures = failures + 1
                If result(4) = "WARNING" Then warnings = warnings + 1
                r = r + 1
            Next peak
        Next junction
    Next scenario
    audit.Close: Set audit = Nothing: log.Close: Set log = Nothing
    ' All project files are closed by ProcessCase. Keep the API session alive.
    Set api = Nothing
    Need fso.GetFile(cfg("Project")).DateLastModified = originalDate And fso.GetFile(cfg("Project")).Size = originalSize, "Baseline changed during the run. Close it in other applications and repeat."
    wsO.Range("C3:C" & r - 1).NumberFormat = "0.000"
    wsO.Range("D3:D" & r - 1).NumberFormat = "0.00"
    wsO.Columns("B").ColumnWidth = 42: wsO.Columns("C:D").ColumnWidth = 15
    wsO.Columns("E:F").ColumnWidth = 14: wsO.Columns("G").ColumnWidth = 23
    wsI.Range("D1").Value2 = "Completed: " & count & " models; " & failures & " failed; " & warnings & " warnings."
    wsI.Range("D2").Value2 = "Last run: " & Format$(Now, "dd mmm yyyy hh:nn")
    wsO.Range("B1").Value2 = "SIDRA | ANALYSIS RESULTS" & vbLf & _
        "Run: " & Format$(Now, "dd mmm yyyy hh:nn")
    wsO.PageSetup.PrintArea = "$A$1:$G$" & r - 1
    wsO.Range("B3:G" & r - 1).Font.Size = 11
    wsO.Range("B3:G" & r - 1).VerticalAlignment = xlCenter
    wsO.Range("B3:G" & r - 1).Rows.RowHeight = 23
    wsI.PageSetup.PrintArea = "$A$1:$M$" & Application.Max(13, _
        wsI.Cells(wsI.Rows.count, 3).End(xlUp).Row, wsI.Cells(wsI.Rows.count, 9).End(xlUp).Row)
    wsI.PageSetup.FitToPagesTall = False
    ThisWorkbook.Save
    ThisWorkbook.SaveCopyAs fso.BuildPath(root, "SIDRA_Results.xlsm")
    ExecuteSIDRA = "COMPLETED: " & count & " models; " & failures & " failed; " & warnings & " warnings."
    If Not quiet Then MsgBox ExecuteSIDRA & vbCrLf & vbCrLf & "Outputs sheet updated." & vbCrLf & root, vbInformation, "SIDRA automation"
    GoTo CleanUp
Failed:
    errText = Err.Description
    ExecuteSIDRA = "ERROR: " & errText
    On Error Resume Next
    If Not wsI Is Nothing Then wsI.Range("D1").Value2 = ExecuteSIDRA
    If Not ThisWorkbook.ReadOnly Then ThisWorkbook.Save
    If Not quiet Then MsgBox ExecuteSIDRA, vbExclamation, "SIDRA automation"
CleanUp:
    On Error Resume Next
    If Not audit Is Nothing Then audit.Close
    If Not log Is Nothing Then log.Close
    If Not api Is Nothing Then api.CloseProject
    If Len(preflight) > 0 Then
        If Not fso Is Nothing Then
            If fso.FileExists(preflight) Then fso.DeleteFile preflight, True
        End If
    End If
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    Application.EnableCancelKey = oldCancel: Application.StatusBar = oldStatus
    mRunning = False
End Function

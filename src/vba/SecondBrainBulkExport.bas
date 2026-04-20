Attribute VB_Name = "SecondBrainBulkExport"
Option Compare Database
Option Explicit

'===========================================================================
' SecondBrainBulkExport
' Módulo temporal inyectado por AccessExtension para exportación masiva.
' Exporta toda la estructura de la BD en una sola ejecución COM.
'
' Puntos de entrada:
'   ExportToJsonFile(outputPath, mode)  → JSON para SecondBrain
'   ExportToFiles(outputPath, mode)     → Carpetas para "Export Objects"
'
' mode: "full" | "tables" | "queries" | "forms" | "reports" | "macros" | "modules"
'
' Este módulo se elimina automáticamente después de la ejecución.
'===========================================================================

' ───────────────────────────────────────────────────────────────────────────
' PUNTO DE ENTRADA 1 — JSON para SecondBrain
' ───────────────────────────────────────────────────────────────────────────
Public Sub ExportToJsonFile(ByVal outputPath As String, Optional ByVal mode As String = "full")
    On Error GoTo ErrH

    Dim json As String
    json = BuildJsonExport(mode)
    WriteUTF8File outputPath, json
    Debug.Print "SecondBrainBulkExport: JSON escrito en " & outputPath
    Exit Sub
ErrH:
    Debug.Print "SecondBrainBulkExport.ExportToJsonFile ERROR " & Err.Number & ": " & Err.Description
End Sub

' ───────────────────────────────────────────────────────────────────────────
' PUNTO DE ENTRADA 2 — Carpetas estilo access-analyzer
' ───────────────────────────────────────────────────────────────────────────
Public Sub ExportToFiles(ByVal outputPath As String, Optional ByVal mode As String = "full")
    On Error GoTo ErrH

    CreateExportFolders outputPath
    If mode = "full" Or mode = "tables"   Then ExportTablesFiles outputPath
    If mode = "full" Or mode = "queries"  Then ExportQueriesFiles outputPath
    If mode = "full" Or mode = "forms"    Then ExportFormsFiles outputPath
    If mode = "full" Or mode = "reports"  Then ExportReportsFiles outputPath
    If mode = "full" Or mode = "macros"   Then ExportMacrosFiles outputPath
    If mode = "full" Or mode = "modules"  Then ExportModulesFiles outputPath
    If mode = "full" Then ExportSummaryFile outputPath
    Debug.Print "SecondBrainBulkExport: Carpetas escritas en " & outputPath
    Exit Sub
ErrH:
    Debug.Print "SecondBrainBulkExport.ExportToFiles ERROR " & Err.Number & ": " & Err.Description
End Sub

' ───────────────────────────────────────────────────────────────────────────
' CONSTRUCCIÓN DEL JSON
' ───────────────────────────────────────────────────────────────────────────
Private Function BuildJsonExport(ByVal mode As String) As String
    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sb As String
    sb = "{" & vbCrLf
    sb = sb & "  ""database"": " & JsonStr(CurrentProject.Name) & "," & vbCrLf
    sb = sb & "  ""source_path"": " & JsonStr(CurrentProject.FullName) & "," & vbCrLf
    sb = sb & "  ""generated"": " & JsonStr(Format(Now, "yyyy-mm-ddThh:nn:ss")) & "," & vbCrLf

    ' Tables + columns + indexes + primary_keys
    If mode = "full" Or mode = "tables" Then
        sb = sb & "  ""tables"": " & BuildTablesJson(db) & "," & vbCrLf
        sb = sb & "  ""columns"": " & BuildColumnsJson(db) & "," & vbCrLf
        sb = sb & "  ""indexes"": " & BuildIndexesJson(db) & "," & vbCrLf
        sb = sb & "  ""primary_keys"": " & BuildPrimaryKeysJson(db) & "," & vbCrLf
        sb = sb & "  ""linked_tables"": " & BuildLinkedTablesJson(db) & "," & vbCrLf
    Else
        sb = sb & "  ""tables"": []," & vbCrLf
        sb = sb & "  ""columns"": []," & vbCrLf
        sb = sb & "  ""indexes"": []," & vbCrLf
        sb = sb & "  ""primary_keys"": []," & vbCrLf
        sb = sb & "  ""linked_tables"": []," & vbCrLf
    End If

    ' Relationships
    If mode = "full" Or mode = "tables" Then
        sb = sb & "  ""relationships"": " & BuildRelationshipsJson(db) & "," & vbCrLf
        sb = sb & "  ""foreign_keys"": " & BuildForeignKeysJson(db) & "," & vbCrLf
    Else
        sb = sb & "  ""relationships"": []," & vbCrLf
        sb = sb & "  ""foreign_keys"": []," & vbCrLf
    End If

    ' Queries
    If mode = "full" Or mode = "queries" Then
        sb = sb & "  ""queries"": " & BuildQueriesJson(db) & "," & vbCrLf
    Else
        sb = sb & "  ""queries"": []," & vbCrLf
    End If

    ' Forms
    If mode = "full" Or mode = "forms" Then
        sb = sb & "  ""forms"": " & BuildFormsJson() & "," & vbCrLf
    Else
        sb = sb & "  ""forms"": []," & vbCrLf
    End If

    ' Reports
    If mode = "full" Or mode = "reports" Then
        sb = sb & "  ""reports"": " & BuildReportsJson() & "," & vbCrLf
    Else
        sb = sb & "  ""reports"": []," & vbCrLf
    End If

    ' Macros
    If mode = "full" Or mode = "macros" Then
        sb = sb & "  ""macros"": " & BuildMacrosJson() & "," & vbCrLf
    Else
        sb = sb & "  ""macros"": []," & vbCrLf
    End If

    ' Modules
    If mode = "full" Or mode = "modules" Then
        sb = sb & "  ""modules"": " & BuildModulesJson() & "," & vbCrLf
    Else
        sb = sb & "  ""modules"": []," & vbCrLf
    End If

    ' References (always lightweight)
    sb = sb & "  ""references"": " & BuildReferencesJson() & "," & vbCrLf
    sb = sb & "  ""startup_options"": []," & vbCrLf
    sb = sb & "  ""warnings"": []" & vbCrLf
    sb = sb & "}"

    BuildJsonExport = sb
End Function

' ───────────────────────────────────────────────────────────────────────────
' TABLAS
' ───────────────────────────────────────────────────────────────────────────
Private Function BuildTablesJson(db As DAO.Database) As String
    Dim items As String
    Dim tbl As DAO.TableDef
    For Each tbl In db.TableDefs
        If IsUserTable(tbl) Then
            Dim isLinked As Boolean
            isLinked = (tbl.Attributes And dbAttachedTable) <> 0 Or (tbl.Attributes And dbAttachedODBC) <> 0
            Dim srcTable As String: srcTable = ""
            On Error Resume Next
            If isLinked Then srcTable = tbl.SourceTableName
            On Error GoTo 0
            If Len(items) > 0 Then items = items & "," & vbCrLf
            items = items & "    {""table_name"": " & JsonStr(tbl.Name) & _
                            ", ""table_type"": " & JsonStr(IIf(isLinked, "LINKED", "LOCAL")) & _
                            ", ""source_table"": " & JsonStr(srcTable) & "}"
        End If
    Next tbl
    BuildTablesJson = "[" & vbCrLf & items & vbCrLf & "  ]"
End Function

Private Function BuildColumnsJson(db As DAO.Database) As String
    Dim items As String
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    Dim pos As Integer
    For Each tbl In db.TableDefs
        If IsUserTable(tbl) Then
            pos = 1
            For Each fld In tbl.Fields
                If Len(items) > 0 Then items = items & "," & vbCrLf
                items = items & "    {""table_name"": " & JsonStr(tbl.Name) & _
                                ", ""column_name"": " & JsonStr(fld.Name) & _
                                ", ""data_type"": " & JsonStr(GetFieldType(fld)) & _
                                ", ""is_nullable"": " & JsonStr(IIf(fld.Required, "NO", "YES")) & _
                                ", ""character_maximum_length"": " & IIf(fld.Type = dbText, CStr(fld.Size), "null") & _
                                ", ""ordinal_position"": " & CStr(pos) & "}"
                pos = pos + 1
            Next fld
        End If
    Next tbl
    BuildColumnsJson = "[" & vbCrLf & items & vbCrLf & "  ]"
End Function

Private Function BuildIndexesJson(db As DAO.Database) As String
    Dim items As String
    Dim tbl As DAO.TableDef
    Dim idx As DAO.Index
    Dim idxFld As DAO.Field
    For Each tbl In db.TableDefs
        If IsUserTable(tbl) Then
            For Each idx In tbl.Indexes
                For Each idxFld In idx.Fields
                    If Len(items) > 0 Then items = items & "," & vbCrLf
                    items = items & "    {""table_name"": " & JsonStr(tbl.Name) & _
                                    ", ""index_name"": " & JsonStr(idx.Name) & _
                                    ", ""field_name"": " & JsonStr(idxFld.Name) & _
                                    ", ""is_unique"": " & JsonBool(idx.Unique) & _
                                    ", ""is_primary"": " & JsonBool(idx.Primary) & _
                                    ", ""sort_order"": " & JsonStr(IIf(idxFld.Attributes And dbDescending, "desc", "asc")) & "}"
                Next idxFld
            Next idx
        End If
    Next tbl
    BuildIndexesJson = "[" & vbCrLf & items & vbCrLf & "  ]"
End Function

Private Function BuildPrimaryKeysJson(db As DAO.Database) As String
    Dim items As String
    Dim tbl As DAO.TableDef
    Dim idx As DAO.Index
    Dim idxFld As DAO.Field
    For Each tbl In db.TableDefs
        If IsUserTable(tbl) Then
            For Each idx In tbl.Indexes
                If idx.Primary Then
                    For Each idxFld In idx.Fields
                        If Len(items) > 0 Then items = items & "," & vbCrLf
                        items = items & "    {""table_schema"": ""dbo""" & _
                                        ", ""table_name"": " & JsonStr(tbl.Name) & _
                                        ", ""column_name"": " & JsonStr(idxFld.Name) & "}"
                    Next idxFld
                End If
            Next idx
        End If
    Next tbl
    BuildPrimaryKeysJson = "[" & vbCrLf & items & vbCrLf & "  ]"
End Function

Private Function BuildLinkedTablesJson(db As DAO.Database) As String
    Dim items As String
    Dim tbl As DAO.TableDef
    For Each tbl In db.TableDefs
        Dim isLinked As Boolean
        isLinked = (tbl.Attributes And dbAttachedTable) <> 0 Or (tbl.Attributes And dbAttachedODBC) <> 0
        If isLinked Then
            Dim srcTbl As String: srcTbl = ""
            Dim conn As String: conn = ""
            On Error Resume Next
            srcTbl = tbl.SourceTableName
            conn = tbl.Connect
            On Error GoTo 0
            If Len(items) > 0 Then items = items & "," & vbCrLf
            items = items & "    {""name"": " & JsonStr(tbl.Name) & _
                            ", ""source_table"": " & JsonStr(srcTbl) & _
                            ", ""connect_string"": " & JsonStr(conn) & _
                            ", ""is_odbc"": " & JsonBool((tbl.Attributes And dbAttachedODBC) <> 0) & "}"
        End If
    Next tbl
    BuildLinkedTablesJson = "[" & vbCrLf & items & vbCrLf & "  ]"
End Function

' ───────────────────────────────────────────────────────────────────────────
' RELACIONES Y FOREIGN KEYS
' ───────────────────────────────────────────────────────────────────────────
Private Function BuildRelationshipsJson(db As DAO.Database) As String
    Dim items As String
    Dim rel As DAO.Relation
    For Each rel In db.Relations
        If Not (Left$(rel.Name, 4) = "MSys" Or Left$(rel.Name, 4) = "~sq_") Then
            If Len(items) > 0 Then items = items & "," & vbCrLf
            items = items & "    {""name"": " & JsonStr(rel.Name) & _
                            ", ""table"": " & JsonStr(rel.Table) & _
                            ", ""foreign_table"": " & JsonStr(rel.ForeignTable) & "}"
        End If
    Next rel
    BuildRelationshipsJson = "[" & vbCrLf & items & vbCrLf & "  ]"
End Function

Private Function BuildForeignKeysJson(db As DAO.Database) As String
    Dim items As String
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    For Each rel In db.Relations
        If Not (Left$(rel.Name, 4) = "MSys" Or Left$(rel.Name, 4) = "~sq_") Then
            For Each fld In rel.Fields
                If Len(items) > 0 Then items = items & "," & vbCrLf
                items = items & "    {""table_name"": " & JsonStr(rel.ForeignTable) & _
                                ", ""column_name"": " & JsonStr(fld.ForeignName) & _
                                ", ""referenced_table"": " & JsonStr(rel.Table) & _
                                ", ""referenced_column"": " & JsonStr(fld.Name) & "}"
            Next fld
        End If
    Next rel
    BuildForeignKeysJson = "[" & vbCrLf & items & vbCrLf & "  ]"
End Function

' ───────────────────────────────────────────────────────────────────────────
' QUERIES
' ───────────────────────────────────────────────────────────────────────────
Private Function BuildQueriesJson(db As DAO.Database) As String
    Dim items As String
    Dim qry As DAO.QueryDef
    For Each qry In db.QueryDefs
        If IsUserQuery(qry) Then
            Dim sql As String: sql = ""
            On Error Resume Next
            sql = qry.SQL
            On Error GoTo 0
            If Len(items) > 0 Then items = items & "," & vbCrLf
            items = items & "    {""name"": " & JsonStr(qry.Name) & _
                            ", ""type"": " & JsonStr(InferQueryType(sql)) & _
                            ", ""sql"": " & JsonStr(sql) & "}"
        End If
    Next qry
    BuildQueriesJson = "[" & vbCrLf & items & vbCrLf & "  ]"
End Function

' ───────────────────────────────────────────────────────────────────────────
' FORMULARIOS
' ───────────────────────────────────────────────────────────────────────────
Private Function BuildFormsJson() As String
    Dim items As String
    Dim i As Integer
    For i = 0 To CurrentProject.AllForms.Count - 1
        Dim formName As String
        formName = CurrentProject.AllForms(i).Name
        If Len(items) > 0 Then items = items & "," & vbCrLf
        items = items & "    " & BuildUiObjectJson("Form", formName)
    Next i
    BuildFormsJson = "[" & vbCrLf & items & vbCrLf & "  ]"
End Function

' ───────────────────────────────────────────────────────────────────────────
' INFORMES
' ───────────────────────────────────────────────────────────────────────────
Private Function BuildReportsJson() As String
    Dim items As String
    Dim i As Integer
    For i = 0 To CurrentProject.AllReports.Count - 1
        Dim reportName As String
        reportName = CurrentProject.AllReports(i).Name
        If Len(items) > 0 Then items = items & "," & vbCrLf
        items = items & "    " & BuildUiObjectJson("Report", reportName)
    Next i
    BuildReportsJson = "[" & vbCrLf & items & vbCrLf & "  ]"
End Function

Private Function BuildUiObjectJson(ByVal objType As String, ByVal objName As String) As String
    ' Abre en modo diseño oculto para leer RecordSource/controles directamente.
    ' Evita SaveAsText (escribe+lee archivos temporales por cada objeto) y el
    ' bucle VBE línea-a-línea. Mucho más rápido para bases con 80+ forms.
    Dim recordSource As String: recordSource = ""
    Dim controlsJson As String: controlsJson = "[]"

    On Error Resume Next
    Err.Clear
    If objType = "Form" Then
        DoCmd.OpenForm objName, acDesign, , , , acHidden
        If Err.Number = 0 Then
            recordSource = Forms(objName).RecordSource
            controlsJson = BuildControlsJsonFromControls(Forms(objName).Controls)
            DoCmd.Close acForm, objName, acSaveNo
        End If
    Else
        DoCmd.OpenReport objName, acViewDesign, , , acHidden
        If Err.Number = 0 Then
            recordSource = Reports(objName).RecordSource
            controlsJson = BuildControlsJsonFromControls(Reports(objName).Controls)
            DoCmd.Close acReport, objName, acSaveNo
        End If
    End If
    On Error GoTo 0

    BuildUiObjectJson = "{""name"": " & JsonStr(objName) & _
                        ", ""record_source"": " & JsonStr(recordSource) & _
                        ", ""controls"": " & controlsJson & "}"
End Function

Private Function BuildControlsJsonFromControls(ByVal ctrlCollection As Object) As String
    ' Lee nombre, tipo, ControlSource y Caption de cada control desde la colección
    ' del form/report abierto en diseño — sin archivo temporal.
    Dim items As String
    Dim ctl As Object
    On Error Resume Next
    For Each ctl In ctrlCollection
        Dim ctrlName As String: ctrlName = ""
        Dim ctrlType As String: ctrlType = ""
        Dim ctrlSource As String: ctrlSource = ""
        Dim ctrlCaption As String: ctrlCaption = ""

        ctrlName = ctl.Name
        ctrlType = TypeName(ctl)
        ctrlSource = ctl.ControlSource
        ctrlCaption = ctl.Caption

        If Len(ctrlName) > 0 Then
            If Len(items) > 0 Then items = items & "," & vbCrLf
            items = items & "      {""name"": " & JsonStr(ctrlName) & _
                            ", ""control_type"": " & JsonStr(ctrlType) & _
                            ", ""control_source"": " & JsonStr(ctrlSource) & _
                            ", ""caption"": " & JsonStr(ctrlCaption) & "}"
        End If
    Next ctl
    On Error GoTo 0
    BuildControlsJsonFromControls = "[" & vbCrLf & items & vbCrLf & "    ]"
End Function

' Parsea el valor de una propiedad del archivo SaveAsText (p.ej. RecordSource = "tblXxx")
Private Function ParsePropertyFromSavedText(ByVal filePath As String, ByVal propName As String) As String
    On Error GoTo ErrH
    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Input As #fNum
    Dim line As String
    Dim search As String: search = UCase(propName & " =")
    Do While Not EOF(fNum)
        Line Input #fNum, line
        If UCase(Left$(Trim(line), Len(search))) = search Then
            ' valor entre comillas: RecordSource = "tblXxx"
            Dim val As String
            val = Mid$(line, InStr(line, "=") + 1)
            val = Trim(val)
            If Left$(val, 1) = """" Then val = Mid$(val, 2, Len(val) - 2)
            ParsePropertyFromSavedText = val
            Close #fNum
            Exit Function
        End If
    Loop
    Close #fNum
    Exit Function
ErrH:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
End Function

' Parsea controles básicos del archivo SaveAsText — devuelve JSON array
Private Function ParseControlsFromSavedText(ByVal filePath As String) As String
    On Error GoTo ErrH
    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Input As #fNum

    Dim items As String
    Dim line As String
    Dim inControl As Boolean: inControl = False
    Dim ctrlName As String: ctrlName = ""
    Dim ctrlType As String: ctrlType = ""
    Dim ctrlSource As String: ctrlSource = ""
    Dim ctrlCaption As String: ctrlCaption = ""

    Do While Not EOF(fNum)
        Line Input #fNum, line
        Dim trimLine As String: trimLine = Trim(line)

        ' Inicio de un control: "Begin <Type> <Name>"
        If Left$(trimLine, 6) = "Begin " And InStr(trimLine, " ") > 0 Then
            Dim parts() As String
            parts = Split(trimLine, " ")
            If UBound(parts) >= 2 Then
                ctrlType = parts(1)
                ctrlName = parts(2)
                ctrlSource = ""
                ctrlCaption = ""
                inControl = True
            End If
        ElseIf trimLine = "End" And inControl Then
            ' Fin del control
            If Len(ctrlName) > 0 Then
                If Len(items) > 0 Then items = items & "," & vbCrLf
                items = items & "      {""name"": " & JsonStr(ctrlName) & _
                                ", ""control_type"": " & JsonStr(ctrlType) & _
                                ", ""control_source"": " & JsonStr(ctrlSource) & _
                                ", ""caption"": " & JsonStr(ctrlCaption) & "}"
            End If
            ctrlName = ""
            ctrlType = ""
            inControl = False
        ElseIf inControl Then
            If UCase(Left$(trimLine, 14)) = "CONTROLSOURCE " Or UCase(Left$(trimLine, 14)) = "CONTROLSOURCE=" Then
                ctrlSource = ExtractPropValue(trimLine)
            ElseIf UCase(Left$(trimLine, 8)) = "CAPTION " Or UCase(Left$(trimLine, 8)) = "CAPTION=" Then
                ctrlCaption = ExtractPropValue(trimLine)
            End If
        End If
    Loop

    Close #fNum
    ParseControlsFromSavedText = "[" & vbCrLf & items & vbCrLf & "    ]"
    Exit Function
ErrH:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
    ParseControlsFromSavedText = "[]"
End Function

Private Function ExtractPropValue(ByVal line As String) As String
    Dim pos As Long
    pos = InStr(line, "=")
    If pos = 0 Then pos = InStr(line, " ")
    If pos = 0 Then Exit Function
    Dim val As String
    val = Trim(Mid$(line, pos + 1))
    If Left$(val, 1) = """" Then val = Mid$(val, 2, Len(val) - 2)
    ExtractPropValue = val
End Function

' ───────────────────────────────────────────────────────────────────────────
' MACROS
' ───────────────────────────────────────────────────────────────────────────
Private Function BuildMacrosJson() As String
    Dim items As String
    Dim i As Integer
    For i = 0 To CurrentProject.AllMacros.Count - 1
        Dim macroName As String
        macroName = CurrentProject.AllMacros(i).Name
        If Len(items) > 0 Then items = items & "," & vbCrLf
        items = items & "    {""name"": " & JsonStr(macroName) & "}"
    Next i
    BuildMacrosJson = "[" & vbCrLf & items & vbCrLf & "  ]"
End Function

' ───────────────────────────────────────────────────────────────────────────
' MÓDULOS VBA
' ───────────────────────────────────────────────────────────────────────────
Private Function BuildModulesJson() As String
    Dim items As String
    On Error Resume Next
    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject
    If vbProj Is Nothing Then
        BuildModulesJson = "[]"
        Exit Function
    End If
    Dim k As Integer
    For k = 1 To vbProj.VBComponents.Count
        Dim vbComp As Object
        Set vbComp = vbProj.VBComponents(k)
        ' Solo módulos estándar y de clase (no Form_* ni Report_*)
        If vbComp.Type = 1 Or vbComp.Type = 2 Then
            Dim code As String: code = ""
            If vbComp.CodeModule.CountOfLines > 0 Then
                code = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
            End If
            If Len(items) > 0 Then items = items & "," & vbCrLf
            items = items & "    {""name"": " & JsonStr(vbComp.Name) & _
                            ", ""type"": " & JsonStr(GetComponentTypeVBA(vbComp.Type)) & _
                            ", ""code"": " & JsonStr(code) & "}"
        End If
    Next k
    On Error GoTo 0
    BuildModulesJson = "[" & vbCrLf & items & vbCrLf & "  ]"
End Function

' ───────────────────────────────────────────────────────────────────────────
' REFERENCIAS VBA
' ───────────────────────────────────────────────────────────────────────────
Private Function BuildReferencesJson() As String
    Dim items As String
    On Error Resume Next
    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject
    If Not vbProj Is Nothing Then
        Dim ref As Object
        For Each ref In vbProj.References
            If Not ref.BuiltIn Then
                If Len(items) > 0 Then items = items & "," & vbCrLf
                items = items & "    {""name"": " & JsonStr(ref.Name) & _
                                ", ""path"": " & JsonStr(ref.FullPath) & "}"
            End If
        Next ref
    End If
    On Error GoTo 0
    BuildReferencesJson = "[" & vbCrLf & items & vbCrLf & "  ]"
End Function

' ───────────────────────────────────────────────────────────────────────────
' EXPORTACIÓN A FICHEROS (ExportToFiles)
' ───────────────────────────────────────────────────────────────────────────
Private Sub CreateExportFolders(ByVal basePath As String)
    On Error Resume Next
    MkDir basePath
    MkDir basePath & "\01_Tablas"
    MkDir basePath & "\02_Consultas"
    MkDir basePath & "\03_Formularios"
    MkDir basePath & "\04_Informes"
    MkDir basePath & "\05_Macros"
    MkDir basePath & "\06_Codigo_VBA"
    On Error GoTo 0
End Sub

Private Sub ExportSummaryFile(ByVal basePath As String)
    Dim db As DAO.Database: Set db = CurrentDb
    Dim content As String
    content = "EXPORTACION COMPLETA" & vbCrLf & String(60, "=") & vbCrLf
    content = content & "Aplicacion: " & CurrentProject.Name & vbCrLf
    content = content & "Archivo: " & CurrentProject.FullName & vbCrLf
    content = content & "Exportado: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf
    content = content & "Tablas: " & CountUserTables(db) & vbCrLf
    content = content & "Consultas: " & CountUserQueries(db) & vbCrLf
    content = content & "Formularios: " & CurrentProject.AllForms.Count & vbCrLf
    content = content & "Informes: " & CurrentProject.AllReports.Count & vbCrLf
    content = content & "Macros: " & CurrentProject.AllMacros.Count & vbCrLf
    WriteUTF8File basePath & "\00_RESUMEN.txt", content
End Sub

Private Sub ExportTablesFiles(ByVal basePath As String)
    Dim db As DAO.Database: Set db = CurrentDb
    Dim tbl As DAO.TableDef
    Dim content As String
    For Each tbl In db.TableDefs
        If IsUserTable(tbl) Then
            content = "[TABLA] " & tbl.Name & vbCrLf & String(40, "-") & vbCrLf
            Dim fld As DAO.Field
            For Each fld In tbl.Fields
                content = content & fld.Name & " | " & GetFieldType(fld) & " | " & IIf(fld.Required, "Requerido", "Opcional") & vbCrLf
            Next fld
            WriteUTF8File basePath & "\01_Tablas\" & CleanName(tbl.Name) & ".txt", content
        End If
    Next tbl
End Sub

Private Sub ExportQueriesFiles(ByVal basePath As String)
    Dim db As DAO.Database: Set db = CurrentDb
    Dim qry As DAO.QueryDef
    For Each qry In db.QueryDefs
        If IsUserQuery(qry) Then
            Dim sql As String: sql = ""
            On Error Resume Next: sql = qry.SQL: On Error GoTo 0
            Dim content As String
            content = "-- Consulta: " & qry.Name & vbCrLf & sql
            WriteUTF8File basePath & "\02_Consultas\" & CleanName(qry.Name) & ".sql", content
        End If
    Next qry
End Sub

Private Sub ExportFormsFiles(ByVal basePath As String)
    Dim i As Integer
    For i = 0 To CurrentProject.AllForms.Count - 1
        Dim fName As String: fName = CurrentProject.AllForms(i).Name
        On Error Resume Next
        Application.SaveAsText acForm, fName, basePath & "\03_Formularios\" & CleanName(fName) & ".formulario.txt"
        On Error GoTo 0
        ExportUiVbaCode basePath & "\03_Formularios", "Form_" & fName, fName & "_codigo"
    Next i
End Sub

Private Sub ExportReportsFiles(ByVal basePath As String)
    Dim i As Integer
    For i = 0 To CurrentProject.AllReports.Count - 1
        Dim rName As String: rName = CurrentProject.AllReports(i).Name
        On Error Resume Next
        Application.SaveAsText acReport, rName, basePath & "\04_Informes\" & CleanName(rName) & ".informe.txt"
        On Error GoTo 0
        ExportUiVbaCode basePath & "\04_Informes", "Report_" & rName, rName & "_codigo"
    Next i
End Sub

Private Sub ExportMacrosFiles(ByVal basePath As String)
    Dim i As Integer
    For i = 0 To CurrentProject.AllMacros.Count - 1
        Dim mName As String: mName = CurrentProject.AllMacros(i).Name
        On Error Resume Next
        Application.SaveAsText acMacro, mName, basePath & "\05_Macros\" & CleanName(mName) & ".macro.txt"
        On Error GoTo 0
    Next i
End Sub

Private Sub ExportModulesFiles(ByVal basePath As String)
    On Error Resume Next
    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject
    If vbProj Is Nothing Then Exit Sub
    Dim k As Integer
    For k = 1 To vbProj.VBComponents.Count
        Dim vbComp As Object
        Set vbComp = vbProj.VBComponents(k)
        If vbComp.Type = 1 Or vbComp.Type = 2 Then
            ExportVbaComponent basePath & "\06_Codigo_VBA", vbComp
        End If
    Next k
    On Error GoTo 0
End Sub

Private Sub ExportUiVbaCode(ByVal folder As String, ByVal compName As String, ByVal fileName As String)
    On Error Resume Next
    Dim vbComp As Object
    Set vbComp = Application.VBE.ActiveVBProject.VBComponents(compName)
    If Not vbComp Is Nothing Then
        If vbComp.CodeModule.CountOfLines > 0 Then
            ExportVbaComponent folder, vbComp, fileName
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub ExportVbaComponent(ByVal folder As String, ByVal vbComp As Object, Optional ByVal customName As String = "")
    Dim fileName As String
    fileName = IIf(Len(customName) > 0, CleanName(customName), CleanName(vbComp.Name))
    Dim content As String
    content = "' Modulo: " & vbComp.Name & vbCrLf & "' Tipo: " & GetComponentTypeVBA(vbComp.Type) & vbCrLf & vbCrLf
    Dim i As Long
    For i = 1 To vbComp.CodeModule.CountOfLines
        content = content & vbComp.CodeModule.Lines(i, 1) & vbCrLf
    Next i
    WriteUTF8File folder & "\" & fileName & ".bas", content
End Sub

' ───────────────────────────────────────────────────────────────────────────
' FUNCIONES AUXILIARES COMPARTIDAS
' ───────────────────────────────────────────────────────────────────────────
Private Function IsUserTable(tbl As DAO.TableDef) As Boolean
    If (tbl.Attributes And (dbSystemObject Or dbHiddenObject)) <> 0 Then Exit Function
    IsUserTable = Not (Left$(UCase$(tbl.Name), 4) = "MSYS" Or Left$(UCase$(tbl.Name), 4) = "USYS")
End Function

Private Function IsUserQuery(qry As DAO.QueryDef) As Boolean
    IsUserQuery = Not (Left$(qry.Name, 4) = "~sq_" Or Left$(UCase$(qry.Name), 4) = "MSYS")
End Function

Private Function CountUserTables(db As DAO.Database) As Integer
    Dim tbl As DAO.TableDef
    For Each tbl In db.TableDefs
        If IsUserTable(tbl) Then CountUserTables = CountUserTables + 1
    Next tbl
End Function

Private Function CountUserQueries(db As DAO.Database) As Integer
    Dim qry As DAO.QueryDef
    For Each qry In db.QueryDefs
        If IsUserQuery(qry) Then CountUserQueries = CountUserQueries + 1
    Next qry
End Function

Private Function GetFieldType(f As DAO.Field) As String
    Select Case f.Type
        Case dbBoolean:  GetFieldType = "Yes/No"
        Case dbByte:     GetFieldType = "Byte"
        Case dbInteger:  GetFieldType = "Integer"
        Case dbLong:     GetFieldType = "Long"
        Case dbCurrency: GetFieldType = "Currency"
        Case dbSingle:   GetFieldType = "Single"
        Case dbDouble:   GetFieldType = "Double"
        Case dbDate:     GetFieldType = "Date/Time"
        Case dbText:     GetFieldType = "Text"
        Case dbMemo:     GetFieldType = "Memo"
        Case dbGUID:     GetFieldType = "GUID"
        Case dbBinary:   GetFieldType = "Binary"
        Case Else:       GetFieldType = "Type_" & CStr(f.Type)
    End Select
End Function

Private Function GetComponentTypeVBA(componentType As Integer) As String
    Select Case componentType
        Case 1:   GetComponentTypeVBA = "Module"
        Case 2:   GetComponentTypeVBA = "Class"
        Case 3:   GetComponentTypeVBA = "Form"
        Case 100: GetComponentTypeVBA = "Report"
        Case Else: GetComponentTypeVBA = "Type_" & componentType
    End Select
End Function

Private Function InferQueryType(ByVal sql As String) As String
    Dim s As String: s = UCase(Trim(sql))
    If Left$(s, 6) = "SELECT" Then: InferQueryType = "SELECT": Exit Function
    If Left$(s, 6) = "INSERT" Then: InferQueryType = "INSERT": Exit Function
    If Left$(s, 6) = "UPDATE" Then: InferQueryType = "UPDATE": Exit Function
    If Left$(s, 6) = "DELETE" Then: InferQueryType = "DELETE": Exit Function
    If Left$(s, 12) = "TRANSFORM S" Then: InferQueryType = "CROSSTAB": Exit Function
    InferQueryType = "OTHER"
End Function

Private Function CleanName(ByVal nameIn As String) As String
    Dim r As String: r = nameIn
    r = Replace(r, " ", "_"): r = Replace(r, "/", "_"): r = Replace(r, "\", "_")
    r = Replace(r, ":", "_"): r = Replace(r, "*", "_"): r = Replace(r, "?", "_")
    r = Replace(r, """", "_"): r = Replace(r, "<", "_"): r = Replace(r, ">", "_")
    r = Replace(r, "|", "_")
    CleanName = r
End Function

' ── JSON helpers ────────────────────────────────────────────────────────────
Private Function JsonStr(ByVal s As String) As String
    ' Escapa el string para JSON: \ " y saltos de línea
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbTab, "\t")
    JsonStr = """" & s & """"
End Function

Private Function JsonBool(ByVal b As Boolean) As String
    JsonBool = IIf(b, "true", "false")
End Function

' ── Escritura UTF-8 ─────────────────────────────────────────────────────────
Private Sub WriteUTF8File(ByVal filePath As String, ByVal content As String)
    On Error GoTo Fallback
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText content
        .SaveToFile filePath, 2
        .Close
    End With
    Exit Sub
Fallback:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Dim fNum As Integer: fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, content
    Close #fNum
End Sub

Attribute VB_Name = "modVCSExport"
'---------------------------------------------------------------------------------------
' Module    : modVCSExport
' Purpose   : Standalone exporter that writes the objects of the current Access database
'           : to text source files, using the same folder layout and file formats as the
'           : MSAccess VCS add-in (export format 5.0.0). The add-in does NOT need to be
'           : installed. The companion module modVCSImport rebuilds a database from
'           : these files.
'           :
'           : Usage (Immediate window, inside the database you want to export):
'           :   ExportSource                            ' Export to <database file>.src\
'           :   ExportSource "C:\Work\MyApp.src"        ' Export to a specific folder
'           :   ExportSource , "*"                      ' Also save data for all local tables
'           :   ExportSource , "tblStates;tblColors"    ' Also save data for these tables
'           :   ExportObject acForm, "frmMain"          ' Export a single object
'           :
'           : Tables listed for data export are remembered in vcs-options.json
'           : ("TablesToExportData"), exactly like the add-in.
'           :
'           : Differences from the add-in (all of them import-compatible):
'           :  - Query SQL is written as stored by Access (not reformatted), and the
'           :    query .json omits the Design View layout (DesignLayout).
'           :  - Conditional formatting stays inline in the .form/.report file
'           :    (same as the add-in option DecodeConditionalFormatting = false).
'           :  - Connection strings are kept in source (UseEnvForConnections = Never),
'           :    with passwords removed.
'           :  - Command bars, navigation pane groups, VBE UserForms and ADP projects
'           :    are not exported.
'           :
'           : Uses late binding only. No additional VBA references are required.
'---------------------------------------------------------------------------------------
Option Compare Binary
Option Explicit

' Add-in version and export format emulated by this module
Private Const VCS_ADDIN_VERSION As String = "5.0.1"
Private Const VCS_EXPORT_FORMAT As String = "5.0.0"

' Modules belonging to this tool. These are never exported.
Private Const SKIP_MODULES As String = "|modVCSExport|modVCSImport|"

' Remove UID/PWD values from connection strings when a password is present
Private Const STRIP_CREDENTIALS As Boolean = True

' Access constants that may not exist in older versions
Private Const AC_TABLE_DATA_MACRO As Long = 12
Private Const AC_EXPORT_ALL_PROPERTIES As Long = 32
Private Const AC_EMBED_SCHEMA As Long = 8

' DAO constants (late bound)
Private Const DB_BOOLEAN As Long = 1
Private Const DB_BYTE As Long = 2
Private Const DB_INTEGER As Long = 3
Private Const DB_LONG As Long = 4
Private Const DB_CURRENCY As Long = 5
Private Const DB_SINGLE As Long = 6
Private Const DB_DOUBLE As Long = 7
Private Const DB_DATE As Long = 8
Private Const DB_BINARY As Long = 9
Private Const DB_TEXT As Long = 10
Private Const DB_LONG_BINARY As Long = 11
Private Const DB_MEMO As Long = 12
Private Const DB_GUID As Long = 15
Private Const DB_VAR_BINARY As Long = 17
Private Const DB_NUMERIC As Long = 19
Private Const DB_TIME As Long = 22
Private Const DB_ATTACHMENT As Long = 101
Private Const DB_AUTO_INCR_FIELD As Long = 16
Private Const DB_RELATION_INHERITED As Long = 4
Private Const DB_OPEN_DYNASET As Long = 2
Private Const DB_OPEN_SNAPSHOT As Long = 4
Private Const DB_QSELECT As Long = 0
Private Const DB_QSQL_PASS_THROUGH As Long = 112
Private Const DB_QSPT_BULK As Long = 144

' VBIDE constants (late bound)
Private Const VBEXT_CT_STD_MODULE As Long = 1
Private Const VBEXT_RK_PROJECT As Long = 1
Private Const VBEXT_PP_LOCKED As Long = 1

' ADODB constants (late bound)
Private Const AD_TYPE_BINARY As Long = 1
Private Const AD_TYPE_TEXT As Long = 2
Private Const AD_SAVE_CREATE_OVERWRITE As Long = 2
Private Const AD_READ_ALL As Long = -1

' XSLT used by the add-in to pretty-print XML (2-space indent)
Private Const XSLT_INDENT As String = "<xsl:stylesheet xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"" version=""1.0""><xsl:output method=""xml""/><xsl:template match=""@*""><xsl:copy/></xsl:template><xsl:template match=""*""><xsl:param name=""indent"" select=""''""/><xsl:text>&#xA;</xsl:text><xsl:value-of select=""$indent""/><xsl:copy><xsl:apply-templates select=""@*|*|text()""><xsl:with-param name=""indent"" select=""concat($indent, '  ')""/></xsl:apply-templates></xsl:copy><xsl:if test=""count(../*)&gt;0 and ../*[last()]=. and not(following-sibling::*)""><xsl:text>&#xA;</xsl:text><xsl:value-of select=""substring($indent,3)""/></xsl:if></xsl:template></xsl:stylesheet>"

#If VBA7 Then
    Private Declare PtrSafe Function GetACP Lib "kernel32" () As Long
#Else
    Private Declare Function GetACP Lib "kernel32" () As Long
#End If

' Module state for the current export operation
Private m_FSO As Object             ' Scripting.FileSystemObject
Private m_Dbs As Object             ' DAO.Database (CurrentDb)
Private m_VBProject As Object       ' VBIDE.VBProject of the current database
Private m_Folder As String          ' Export folder, with trailing backslash
Private m_Written As Object         ' Files written during this export (for orphan cleanup)
Private m_TableData As Object       ' TablesToExportData option
Private m_Log As Collection         ' Log lines
Private m_Errors As Long
Private m_Warnings As Long
Private m_AnsiCharset As String     ' Charset used by the VBE for module export
Private m_DefaultPrintJson As Variant


'---------------------------------------------------------------------------------------
' Procedure : ExportSource
' Purpose   : Export all supported objects of the current database to source files.
'           : ExportFolder - Defaults to <database full path>.src\
'           : TableData    - "" = use the TablesToExportData list in vcs-options.json,
'           :                "*" = add every local table to that list,
'           :                "tblA;tblB" = add these tables to that list.
'           :                (Data is saved in the add-in's "Tab Delimited" format.)
'---------------------------------------------------------------------------------------
'
Public Sub ExportSource(Optional ByVal ExportFolder As String, Optional ByVal TableData As String, _
    Optional ByVal ShowMessage As Boolean = True)

    Dim sngStart As Single

    sngStart = Timer
    If Not BeginExport(ExportFolder, TableData) Then Exit Sub

    LogLine "Beginning export of " & CurrentProject.Name
    LogLine "Export folder: " & m_Folder
    LogLine "Export format: " & VCS_EXPORT_FORMAT & " (standalone modVCSExport)"
    LogLine CStr(Now)

    CloseOpenObjects

    ExportVcsOptions
    ExportProject
    ExportVbeProject
    ExportVbeReferences
    ExportProjectProperties
    ExportSavedSpecs
    ExportAllModules
    ExportSharedImages
    ExportThemes
    ExportDbProperties
    ExportImexSpecs
    ExportAllTables
    ExportAllQueries
    ExportAllOfType acForm
    ExportAllOfType acMacro
    ExportAllOfType acReport
    ExportAllTableDataMacros
    ExportAllRelations
    ExportDocuments

    RemoveOrphanedFiles
    FinishExport "Export", sngStart, ShowMessage

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportObject
' Purpose   : Export a single database object (form, report, macro, module, query or
'           : table). For tables, the table data (if listed in TablesToExportData) and
'           : any table data macros are exported as well.
'---------------------------------------------------------------------------------------
'
Public Sub ExportObject(ByVal ObjectType As AcObjectType, ByVal ObjectName As String, _
    Optional ByVal ExportFolder As String)

    Dim sngStart As Single

    sngStart = Timer
    If Not BeginExport(ExportFolder, vbNullString) Then Exit Sub
    LogLine "Exporting " & ObjectName & " to " & m_Folder

    If Not m_FSO.FileExists(m_Folder & "vcs-options.json") Then
        ' Make sure the folder can be used for a build.
        ExportVcsOptions
        ExportProject
        ExportVbeProject
        ExportVbeReferences
    End If

    On Error Resume Next
    If SysCmd(acSysCmdGetObjectState, ObjectType, ObjectName) <> 0 Then
        DoCmd.Close ObjectType, ObjectName, acSavePrompt
    End If
    Err.Clear
    On Error GoTo 0

    Select Case ObjectType
        Case acForm, acReport
            ExportFormOrReport ObjectType, ObjectName
        Case acMacro
            ExportMacro ObjectName
        Case acModule
            If IsSkippedModule(ObjectName) Then
                LogWarning ObjectName & " is part of the export/import tool and is not exported."
            Else
                ExportModule ObjectName
            End If
        Case acQuery
            ExportQuery ObjectName
        Case acTable
            ExportTableDef ObjectName
            If m_TableData.Exists(ObjectName) Then ExportTableData ObjectName
            If HasDataMacro(ObjectName) Then ExportTableDataMacro ObjectName
        Case Else
            LogWarning "Object type " & ObjectType & " is not supported by ExportObject."
    End Select

    FinishExport "Export", sngStart, (m_Errors > 0)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BeginExport
' Purpose   : Initialize module state for an export operation.
'---------------------------------------------------------------------------------------
'
Private Function BeginExport(ByVal strFolder As String, ByVal strTableData As String) As Boolean

    Set m_FSO = CreateObject("Scripting.FileSystemObject")
    Set m_Log = New Collection
    Set m_Written = NewDict
    m_Errors = 0
    m_Warnings = 0
    m_DefaultPrintJson = Empty
    m_AnsiCharset = GetAnsiCharset

    If CurrentProject.ProjectType <> acMDB Then
        MsgBox "This exporter supports Access database files (.accdb/.mdb) only.", vbExclamation
        Exit Function
    End If

    Set m_Dbs = CurrentDb
    Set m_VBProject = GetVBProject
    If m_VBProject Is Nothing Then
        MsgBox "Unable to access the VBA project of the current database.", vbExclamation
        Exit Function
    End If
    If m_VBProject.Protection = VBEXT_PP_LOCKED Then
        MsgBox "The VBA project is locked (password protected or compiled ACCDE/MDE)." & vbCrLf & _
            "Please unlock it (or use the original .accdb) before exporting.", vbExclamation
        Exit Function
    End If

    ' Resolve the export folder
    If Len(strFolder) = 0 Then strFolder = CurrentProject.FullName & ".src"
    If Right$(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
    m_Folder = strFolder
    EnsureFolder m_Folder

    LoadTableDataOptions strTableData
    BeginExport = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : FinishExport
' Purpose   : Write the log file and show a summary.
'---------------------------------------------------------------------------------------
'
Private Sub FinishExport(ByVal strOperation As String, ByVal sngStart As Single, ByVal blnShowMessage As Boolean)

    Dim strLogFile As String
    Dim strSummary As String
    Dim varLine As Variant
    Dim strText As String

    On Error Resume Next
    SysCmd acSysCmdClearStatus
    On Error GoTo 0
    strSummary = "Done. (" & Format$(Timer - sngStart, "0.0") & " seconds)"
    If m_Errors > 0 Or m_Warnings > 0 Then
        strSummary = strSummary & " " & m_Errors & " error(s), " & m_Warnings & " warning(s)."
    End If
    LogLine strSummary

    ' Save log file (the add-in uses a logs subfolder, which is ignored by git)
    On Error Resume Next
    strLogFile = m_Folder & "logs\" & strOperation & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".log"
    For Each varLine In m_Log
        strText = strText & varLine & vbCrLf
    Next varLine
    EnsureFolder m_Folder & "logs"
    WriteTextFile strLogFile, strText
    Err.Clear
    On Error GoTo 0

    If blnShowMessage Then
        MsgBox "Export complete." & vbCrLf & vbCrLf & m_Folder & vbCrLf & vbCrLf & strSummary & _
            IIf(m_Errors + m_Warnings > 0, vbCrLf & vbCrLf & "See the log file for details:" & vbCrLf & strLogFile, ""), _
            IIf(m_Errors > 0, vbExclamation, vbInformation), "Export Source"
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CloseOpenObjects
' Purpose   : Close open objects so that saved designs are exported.
'---------------------------------------------------------------------------------------
'
Private Sub CloseOpenObjects()

    Dim lngCnt As Long
    Dim obj As Object

    On Error Resume Next
    For lngCnt = Forms.Count - 1 To 0 Step -1
        DoCmd.Close acForm, Forms(lngCnt).Name, acSavePrompt
    Next lngCnt
    For lngCnt = Reports.Count - 1 To 0 Step -1
        DoCmd.Close acReport, Reports(lngCnt).Name, acSavePrompt
    Next lngCnt
    For Each obj In CurrentData.AllQueries
        If obj.IsLoaded Then DoCmd.Close acQuery, obj.Name, acSavePrompt
    Next obj
    For Each obj In CurrentData.AllTables
        If obj.IsLoaded Then DoCmd.Close acTable, obj.Name, acSavePrompt
    Next obj
    For Each obj In CurrentProject.AllMacros
        If obj.IsLoaded Then DoCmd.Close acMacro, obj.Name, acSavePrompt
    Next obj
    Err.Clear

End Sub


'=======================================================================================
' Single-file components
'=======================================================================================

'---------------------------------------------------------------------------------------
' Procedure : ExportVcsOptions
' Purpose   : Write vcs-options.json (required by the add-in and by modVCSImport).
'           : Options that describe this exporter's output are set explicitly; any
'           : other options already present in the file are preserved.
'---------------------------------------------------------------------------------------
'
Private Sub ExportVcsOptions()

    Dim dFile As Object
    Dim dInfo As Object
    Dim dOptions As Object
    Dim dExisting As Object
    Dim dPrint As Object
    Dim varKey As Variant
    Dim varName As Variant

    On Error GoTo ErrHandler

    ' Load any existing options
    Set dFile = ReadJsonFile(m_Folder & "vcs-options.json")
    If Not dFile Is Nothing Then Set dExisting = GetDict(dFile, "Options")
    If dExisting Is Nothing Then Set dExisting = NewDict

    Set dPrint = NewDict
    For Each varName In Array("Orientation", "PaperSize", "Duplex", "PrintQuality", "DisplayFrequency", _
        "Collate", "Resolution", "DisplayFlags", "Color", "Copies", "ICMMethod", "DefaultSource", "Scale", _
        "ICMIntent", "FormName", "PaperLength", "DitherType", "MediaType", "PaperWidth", "TTOption")
        dPrint.Add varName, (varName = "Orientation" Or varName = "PaperSize")
    Next varName

    ' Options in the same order as the add-in writes them.
    Set dOptions = NewDict
    AddOption dOptions, dExisting, "ExportFolder", vbNullString
    AddOption dOptions, dExisting, "UseFastSave", True
    dOptions.Add "SplitLayoutFromVBA", True
    dOptions.Add "ExportFormatVersion", VCS_EXPORT_FORMAT
    AddOption dOptions, dExisting, "RunBeforeExport", vbNullString
    AddOption dOptions, dExisting, "RunAfterExport", vbNullString
    dOptions.Add "SanitizeLevel", 2
    dOptions.Add "SanitizeColors", 1
    dOptions.Add "SaveQuerySQL", True
    dOptions.Add "FormatSQL", False
    AddOption dOptions, dExisting, "UseDeterministicQueryExport", True
    dOptions.Add "SaveTableSQL", True
    dOptions.Add "SavePrintVars", True
    dOptions.Add "DecodeConditionalFormatting", False
    dOptions.Add "ExtractThemeFiles", False
    dOptions.Add "UseEnvForConnections", 2
    dOptions.Add "ExportPrintSettings", dPrint
    dOptions.Add "ExportLayoutSvg", False
    AddOption dOptions, dExisting, "LayoutSvgImageEmbed", "Small"
    AddOption dOptions, dExisting, "LayoutSvgScaleMode", "Universal"
    AddOption dOptions, dExisting, "ExportAfterMerge", False
    AddOption dOptions, dExisting, "ForceImportOriginalQuerySQL", False
    AddOption dOptions, dExisting, "UseMergeBuild", False
    AddOption dOptions, dExisting, "RunBeforeBuild", vbNullString
    AddOption dOptions, dExisting, "RunAfterBuild", vbNullString
    AddOption dOptions, dExisting, "RunBeforeMerge", vbNullString
    AddOption dOptions, dExisting, "RunAfterMerge", vbNullString
    dOptions.Add "TablesToExportData", SortDict(m_TableData)
    AddOption dOptions, dExisting, "SchemaExports", NewDict
    AddOption dOptions, dExisting, "ShowDebug", False
    AddOption dOptions, dExisting, "BreakOnError", False
    AddOption dOptions, dExisting, "MaxLogFiles", 20
    dOptions.Add "StripPublishOption", True
    dOptions.Add "PreserveRubberDuckID", False
    dOptions.Add "SaveAllDocumentProperties", False
    dOptions.Add "SaveLinkedTableDef", False
    AddOption dOptions, dExisting, "HashAlgorithm", "SHA256"
    AddOption dOptions, dExisting, "UseShortHash", True
    AddOption dOptions, dExisting, "UseGitIntegration", False
    AddOption dOptions, dExisting, "ShowVCSLegacy", True

    ' Keep any other options that were already in the file.
    For Each varKey In dExisting.Keys
        If Not dOptions.Exists(varKey) Then dOptions.Add varKey, dExisting(varKey)
    Next varKey

    Set dInfo = NewDict
    dInfo.Add "AddinVersion", VCS_ADDIN_VERSION
    #If Win64 Then
        dInfo.Add "AccessVersion", Application.Version & " 64-bit"
    #Else
        dInfo.Add "AccessVersion", Application.Version & " 32-bit"
    #End If

    Set dFile = NewDict
    dFile.Add "Info", dInfo
    dFile.Add "Options", dOptions
    WriteTextFile m_Folder & "vcs-options.json", JsonEncode(dFile)
    Exit Sub

ErrHandler:
    LogError "Error writing vcs-options.json"

End Sub


Private Sub AddOption(dOptions As Object, dExisting As Object, ByVal strName As String, ByVal varDefault As Variant)
    If dExisting.Exists(strName) Then
        dOptions.Add strName, dExisting(strName)
    Else
        dOptions.Add strName, varDefault
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadTableDataOptions
' Purpose   : Load the list of tables to export data for, and apply any additions.
'---------------------------------------------------------------------------------------
'
Private Sub LoadTableDataOptions(ByVal strTableData As String)

    Dim dFile As Object
    Dim dList As Object
    Dim varKey As Variant
    Dim varName As Variant
    Dim strName As String
    Dim tdf As Object

    Set m_TableData = NewDict

    Set dFile = ReadJsonFile(m_Folder & "vcs-options.json")
    If dFile Is Nothing Then
        ' Same defaults as the add-in
        m_TableData.Add "USysRegInfo", TableDataFormat("Tab Delimited")
        m_TableData.Add "USysRibbons", TableDataFormat("Tab Delimited")
    Else
        Set dList = GetDict(GetDict(dFile, "Options"), "TablesToExportData")
        If Not dList Is Nothing Then
            For Each varKey In dList.Keys
                If IsObject(dList(varKey)) Then
                    m_TableData.Add varKey, dList(varKey)
                Else
                    m_TableData.Add varKey, TableDataFormat("Tab Delimited")
                End If
            Next varKey
        End If
    End If

    strTableData = Trim$(strTableData)
    If strTableData = "*" Then
        For Each tdf In m_Dbs.TableDefs
            If IsExportableTable(tdf.Name) And Len(tdf.Connect) = 0 Then
                If Not m_TableData.Exists(tdf.Name) Then m_TableData.Add tdf.Name, TableDataFormat("Tab Delimited")
            End If
        Next tdf
    ElseIf Len(strTableData) > 0 Then
        For Each varName In Split(Replace(strTableData, ",", ";"), ";")
            strName = Trim$(varName)
            If Len(strName) > 0 Then
                If Not m_TableData.Exists(strName) Then m_TableData.Add strName, TableDataFormat("Tab Delimited")
            End If
        Next varName
    End If

End Sub


Private Function TableDataFormat(ByVal strFormat As String) As Object
    Set TableDataFormat = NewDict
    TableDataFormat.Add "Format", strFormat
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExportProject
' Purpose   : project.json (file format is used to create the database on build)
'---------------------------------------------------------------------------------------
'
Private Sub ExportProject()

    Dim dItems As Object
    Dim prj As Object

    On Error GoTo ErrHandler
    Set prj = CurrentProject
    Set dItems = NewDict
    dItems.Add "FileFormat", prj.FileFormat
    dItems.Add "RemovePersonalInformation", CBool(prj.RemovePersonalInformation)
    WriteSingleFile "project.json", "clsDbProject", dItems, "Project"
    Exit Sub

ErrHandler:
    LogError "Error exporting project.json"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportVbeProject
' Purpose   : vbe-project.json
'---------------------------------------------------------------------------------------
'
Private Sub ExportVbeProject()

    Dim dItems As Object

    On Error GoTo ErrHandler
    Set dItems = NewDict
    With m_VBProject
        dItems.Add "Name", .Name
        dItems.Add "Description", .Description
        dItems.Add "FileName", m_FSO.GetFileName(.FileName)
        dItems.Add "HelpFile", ValidHelpFile(.HelpFile)
        dItems.Add "HelpContextId", .HelpContextID
        dItems.Add "ConditionalCompilationArguments", CStr(Nz(Application.GetOption("Conditional Compilation Arguments"), vbNullString))
        dItems.Add "Mode", .Mode
        dItems.Add "Protection", .Protection
        dItems.Add "Type", .Type
    End With
    WriteSingleFile "vbe-project.json", "clsDbVbeProject", dItems, "VBE Project"
    Exit Sub

ErrHandler:
    LogError "Error exporting vbe-project.json"

End Sub


Private Function ValidHelpFile(ByVal strFile As String) As String
    Select Case LCase$(m_FSO.GetExtensionName(strFile))
        Case "hlp", "chm": ValidHelpFile = strFile
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExportVbeReferences
' Purpose   : vbe-references.json (in priority order, built-in references excluded)
'---------------------------------------------------------------------------------------
'
Private Sub ExportVbeReferences()

    Dim dItems As Object
    Dim dRef As Object
    Dim ref As Object
    Dim strName As String
    Dim blnBroken As Boolean

    On Error GoTo ErrHandler
    Set dItems = NewDict

    For Each ref In m_VBProject.References
        If Not ref.BuiltIn Then
            strName = vbNullString
            blnBroken = False
            On Error Resume Next
            blnBroken = ref.IsBroken
            strName = ref.Name
            Err.Clear
            On Error GoTo ErrHandler
            If (blnBroken And ref.Type <> VBEXT_RK_PROJECT) Or Len(strName) = 0 Then strName = ref.Guid
            If Not dItems.Exists(strName) Then
                Set dRef = NewDict
                If ref.Type = VBEXT_RK_PROJECT Then
                    dRef.Add "FullPath", GetRelativePath(ref.FullPath)
                Else
                    If Len(ref.Guid) > 0 Then dRef.Add "GUID", ref.Guid
                    dRef.Add "Version", CStr(ref.Major) & "." & CStr(ref.Minor)
                End If
                dItems.Add strName, dRef
            End If
        End If
    Next ref

    WriteSingleFile "vbe-references.json", "clsDbVbeReference", dItems, "VBE References"
    Exit Sub

ErrHandler:
    LogError "Error exporting vbe-references.json"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportDbProperties
' Purpose   : dbs-properties.json (DAO database properties, sorted by name)
'---------------------------------------------------------------------------------------
'
Private Sub ExportDbProperties()

    Dim dItems As Object
    Dim dProp As Object
    Dim prp As Object
    Dim varValue As Variant
    Dim lngType As Long

    On Error GoTo ErrHandler
    Set dItems = NewDict

    For Each prp In m_Dbs.Properties
        Select Case prp.Name
            Case "Connection", "Last VCS Export", "Last VCS Version"
                ' Skip these properties
            Case Else
                varValue = Empty
                On Error Resume Next
                varValue = prp.Value
                lngType = prp.Type
                If Err.Number <> 0 Then
                    Err.Clear
                    On Error GoTo ErrHandler
                Else
                    On Error GoTo ErrHandler
                    Select Case prp.Name
                        Case "AppIcon", "Name"
                            If Len(Nz(varValue, vbNullString)) > 0 Then varValue = GetRelativePath(CStr(varValue))
                    End Select
                    If lngType = DB_DATE Then
                        If IsDate(varValue) Then varValue = ToIsoUtc(CDate(varValue))
                    End If
                    Set dProp = NewDict
                    dProp.Add "Value", varValue
                    dProp.Add "Type", lngType
                    If Not dItems.Exists(prp.Name) Then dItems.Add prp.Name, dProp
                End If
        End Select
    Next prp

    WriteSingleFile "dbs-properties.json", "clsDbProperty", SortDict(dItems), "Database Properties (DAO)"
    Exit Sub

ErrHandler:
    LogError "Error exporting dbs-properties.json"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportProjectProperties
' Purpose   : proj-properties.json (CurrentProject.Properties)
'---------------------------------------------------------------------------------------
'
Private Sub ExportProjectProperties()

    Dim dItems As Object
    Dim prp As Object
    Dim varValue As Variant

    On Error GoTo ErrHandler
    Set dItems = NewDict

    For Each prp In CurrentProject.Properties
        Select Case prp.Name
            Case "Connection", "Last VCS Export", "Last VCS Version", "VCS Build Path", "VCS Source Path"
                ' Skip these
            Case Else
                varValue = prp.Value
                If prp.Name = "AppIcon" Then
                    If Len(Nz(varValue, vbNullString)) > 0 Then varValue = GetRelativePath(CStr(varValue))
                End If
                If Not dItems.Exists(prp.Name) Then dItems.Add prp.Name, varValue
        End Select
    Next prp

    WriteSingleFile "proj-properties.json", "clsDbProjProperty", SortDict(dItems), "Project Properties (Access)"
    Exit Sub

ErrHandler:
    LogError "Error exporting proj-properties.json"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportDocuments
' Purpose   : documents.json (SummaryInfo / UserDefined properties of the database)
'---------------------------------------------------------------------------------------
'
Private Sub ExportDocuments()

    Dim dItems As Object
    Dim dDocs As Object
    Dim dDoc As Object
    Dim dProp As Object
    Dim doc As Object
    Dim prp As Object
    Dim varValue As Variant

    On Error GoTo ErrHandler
    Set dItems = NewDict
    Set dDocs = NewDict

    For Each doc In m_Dbs.Containers("Databases").Documents
        If doc.Name <> "MSysDB" And Left$(doc.Name, 1) <> "~" Then
            Set dDoc = NewDict
            For Each prp In doc.Properties
                Select Case prp.Name
                    Case "AllPermissions", "Container", "DateCreated", "LastUpdated", "Name", "Owner", _
                         "GUID", "Permissions", "UserName", "KeepLocal", "Replicable"
                        ' Standard DAO properties
                    Case Else
                        On Error Resume Next
                        varValue = Empty
                        varValue = prp.Value
                        If Err.Number = 0 Then
                            Set dProp = NewDict
                            dProp.Add "Type", prp.Type
                            dProp.Add "Value", varValue
                            If Not dDoc.Exists(prp.Name) Then dDoc.Add prp.Name, dProp
                        End If
                        Err.Clear
                        On Error GoTo ErrHandler
                End Select
            Next prp
            If dDoc.Count > 0 Then dDocs.Add doc.Name, SortDict(dDoc)
        End If
    Next doc

    If dDocs.Count > 0 Then dItems.Add "Databases", SortDict(dDocs)
    WriteSingleFile "documents.json", "clsDbDocument", dItems, "Database Documents Properties (DAO)"
    Exit Sub

ErrHandler:
    LogError "Error exporting documents.json"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : WriteSingleFile
' Purpose   : Write a root-level json file, or remove it when there is no content.
'---------------------------------------------------------------------------------------
'
Private Sub WriteSingleFile(ByVal strFileName As String, ByVal strClass As String, dItems As Object, ByVal strDescription As String)
    If dItems.Count > 0 Then
        WriteTextFile m_Folder & strFileName, BuildJsonFile(strClass, dItems, strDescription)
    Else
        DeleteFile m_Folder & strFileName
    End If
End Sub


'=======================================================================================
' VBA modules
'=======================================================================================

Private Sub ExportAllModules()
    Dim obj As Object
    LogLine "Exporting modules..."
    For Each obj In CurrentProject.AllModules
        If Not IsSkippedModule(obj.Name) Then ExportModule obj.Name
    Next obj
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportModule
' Purpose   : Export a standard or class module through the VBE (preserves attributes),
'           : convert to UTF-8 with BOM, and place it in the @Folder subfolder.
'---------------------------------------------------------------------------------------
'
Private Sub ExportModule(ByVal strName As String)

    Dim cmp As Object
    Dim strTemp As String
    Dim strCode As String
    Dim strExt As String
    Dim strFolder As String
    Dim strBase As String

    On Error GoTo ErrHandler

    Status "Exporting module " & strName
    Set cmp = m_VBProject.VBComponents(strName)
    If cmp.Type = VBEXT_CT_STD_MODULE Then strExt = ".bas" Else strExt = ".cls"

    strTemp = GetTempFile
    cmp.Export strTemp
    strCode = SanitizeVBA(ReadTextFile(strTemp, m_AnsiCharset))
    DeleteFile strTemp

    strFolder = m_Folder & "modules\" & GetFolderAnnotation(strCode)
    strBase = strFolder & GetSafeFileName(strName)
    WriteTextFile strBase & strExt, strCode
    WriteObjectJson strBase & ".json", Nothing, vbNullString, vbNullString, "Modules", strName, acModule
    Exit Sub

ErrHandler:
    LogError "Error exporting module " & strName

End Sub


Private Function IsSkippedModule(ByVal strName As String) As Boolean
    IsSkippedModule = (InStr(1, SKIP_MODULES, "|" & strName & "|", vbTextCompare) > 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SanitizeVBA
' Purpose   : Remove trailing spaces from each line and trailing blank lines.
'---------------------------------------------------------------------------------------
'
Private Function SanitizeVBA(ByVal strCode As String) As String

    Dim varLines As Variant
    Dim lngLast As Long
    Dim lngLine As Long
    Dim strOut() As String

    varLines = Split(NormalizeLineEndings(strCode), vbCrLf)
    lngLast = -1
    For lngLine = UBound(varLines) To 0 Step -1
        If Len(Trim$(varLines(lngLine))) > 0 Then
            lngLast = lngLine
            Exit For
        End If
    Next lngLine
    If lngLast < 0 Then Exit Function

    ReDim strOut(0 To lngLast)
    For lngLine = 0 To lngLast
        strOut(lngLine) = RTrim$(varLines(lngLine))
    Next lngLine
    SanitizeVBA = Join(strOut, vbCrLf) & vbCrLf

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFolderAnnotation
' Purpose   : Return the relative subfolder for a Rubberduck '@Folder("A.B") annotation,
'           : such as "A\B\", or an empty string when there is no annotation.
'---------------------------------------------------------------------------------------
'
Private Function GetFolderAnnotation(ByVal strCode As String) As String

    Const TAG As String = "'@FOLDER("

    Dim lngPos As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngEol As Long
    Dim varSegments As Variant
    Dim lngSeg As Long
    Dim strPath As String

    If Left$(strCode, 2) <> vbCrLf Then strCode = vbCrLf & strCode
    lngPos = InStr(1, UCase$(strCode), vbCrLf & TAG, vbBinaryCompare)
    If lngPos = 0 Then Exit Function

    lngEol = InStr(lngPos + 2, strCode, vbCrLf)
    If lngEol = 0 Then lngEol = Len(strCode) + 1
    lngStart = InStr(lngPos, strCode, """")
    If lngStart = 0 Or lngStart > lngEol Then Exit Function
    lngEnd = InStr(lngStart + 1, strCode, """")
    If lngEnd = 0 Or lngEnd > lngEol Then Exit Function
    If lngEnd <= lngStart + 1 Then Exit Function

    varSegments = Split(Mid$(strCode, lngStart + 1, lngEnd - lngStart - 1), ".")
    For lngSeg = 0 To UBound(varSegments)
        If Len(Trim$(varSegments(lngSeg))) > 0 Then
            strPath = strPath & GetSafeFileName(Trim$(varSegments(lngSeg))) & "\"
        End If
    Next lngSeg
    GetFolderAnnotation = strPath

End Function


'=======================================================================================
' Forms, reports and macros
'=======================================================================================

Private Sub ExportAllOfType(ByVal intType As Long)

    Dim colItems As Object
    Dim obj As Object

    Select Case intType
        Case acForm
            LogLine "Exporting forms..."
            Set colItems = CurrentProject.AllForms
        Case acReport
            LogLine "Exporting reports..."
            Set colItems = CurrentProject.AllReports
        Case acMacro
            LogLine "Exporting macros..."
            Set colItems = CurrentProject.AllMacros
    End Select

    For Each obj In colItems
        If intType = acMacro Then
            ExportMacro obj.Name
        Else
            ExportFormOrReport intType, obj.Name
        End If
    Next obj

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportFormOrReport
' Purpose   : SaveAsText, sanitize, split the code-behind into a .cls file, and write
'           : print settings / description / hidden flag to the companion .json file.
'---------------------------------------------------------------------------------------
'
Private Sub ExportFormOrReport(ByVal intType As Long, ByVal strName As String)

    Dim strTemp As String
    Dim strText As String
    Dim strVBA As String
    Dim strSafe As String
    Dim strFolder As String
    Dim strBase As String
    Dim dPrint As Object

    On Error GoTo ErrHandler

    Status "Exporting " & strName
    strTemp = GetTempFile
    Application.SaveAsText intType, strName, strTemp
    strText = ReadSourceFile(strTemp)
    DeleteFile strTemp

    strSafe = GetSafeFileName(strName)
    Set dPrint = NewDict
    strText = SanitizeObject(strText, intType, strSafe, strVBA, dPrint)

    If intType = acForm Then strFolder = "forms\" Else strFolder = "reports\"
    strFolder = m_Folder & strFolder & GetFolderAnnotation(strVBA)
    strBase = strFolder & strSafe

    WriteTextFile strBase & IIf(intType = acForm, ".form", ".report"), strText
    If Len(strVBA) > 0 Then WriteTextFile strBase & ".cls", strVBA

    ' Companion json: print settings (if not the default printer settings) + metadata
    If dPrint.Count > 0 Then
        If IsDefaultPrintSettings(dPrint) Then Set dPrint = Nothing
    Else
        Set dPrint = Nothing
    End If
    If dPrint Is Nothing Then
        WriteObjectJson strBase & ".json", Nothing, vbNullString, vbNullString, _
            IIf(intType = acForm, "Forms", "Reports"), strName, intType
    Else
        WriteObjectJson strBase & ".json", dPrint, IIf(intType = acForm, "clsDbForm", "clsDbReport"), _
            strName & " Print Settings", IIf(intType = acForm, "Forms", "Reports"), strName, intType
    End If
    Exit Sub

ErrHandler:
    LogError "Error exporting " & IIf(intType = acForm, "form ", "report ") & strName

End Sub


Private Sub ExportMacro(ByVal strName As String)

    Dim strTemp As String
    Dim strText As String
    Dim strVBA As String
    Dim strBase As String
    Dim dPrint As Object

    On Error GoTo ErrHandler

    Status "Exporting macro " & strName
    strTemp = GetTempFile
    Application.SaveAsText acMacro, strName, strTemp
    strText = ReadSourceFile(strTemp)
    DeleteFile strTemp

    Set dPrint = NewDict
    strText = SanitizeObject(strText, acMacro, GetSafeFileName(strName), strVBA, dPrint)
    strBase = m_Folder & "macros\" & GetSafeFileName(strName)
    WriteTextFile strBase & ".macro", strText
    WriteObjectJson strBase & ".json", Nothing, vbNullString, vbNullString, "Scripts", strName, acMacro
    Exit Sub

ErrHandler:
    LogError "Error exporting macro " & strName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SanitizeObject
' Purpose   : Apply the add-in's sanitize rules (SanitizeLevel = Standard,
'           : SanitizeColors = Minimal) to SaveAsText output. Printer blocks are
'           : removed and parsed into dPrint. The code-behind (if any) is returned in
'           : strVBA and replaced in the output with a reference to the .cls file.
'---------------------------------------------------------------------------------------
'
Private Function SanitizeObject(ByVal strText As String, ByVal intType As Long, ByVal strSafeName As String, _
    ByRef strVBA As String, dPrint As Object) As String

    Dim varLines As Variant
    Dim strOut() As String
    Dim blnDelete() As Boolean
    Dim lngOut As Long
    Dim lngLine As Long
    Dim lngCnt As Long
    Dim strLine As String
    Dim strTrim As String
    Dim lngSkipIndent As Long
    Dim lngContIndent As Long
    Dim strCapture As String
    Dim dHex As Object
    Dim colStack As Collection
    Dim blnReportRight As Boolean
    Dim blnReportBottom As Boolean
    Dim blnViewportDone As Boolean
    Dim lngPos As Long
    Dim strProp As String
    Dim strValue As String
    Dim strResult() As String

    strVBA = vbNullString
    varLines = Split(NormalizeLineEndings(strText), vbCrLf)
    ReDim strOut(0 To UBound(varLines) + 2)
    ReDim blnDelete(0 To UBound(varLines) + 2)
    lngSkipIndent = -1
    lngContIndent = -1
    Set dHex = NewDict
    Set colStack = New Collection

    For lngLine = 0 To UBound(varLines)
        strLine = varLines(lngLine)
        strTrim = Trim$(strLine)

        ' Inside a block that is being removed
        If lngSkipIndent >= 0 Then
            If strTrim = "End" And IndentOf(strLine) = lngSkipIndent Then
                lngSkipIndent = -1
            ElseIf Len(strCapture) > 0 Then
                dHex(strCapture) = dHex(strCapture) & strTrim
            End If
            GoTo NextLine
        End If

        ' Continuation lines of a removed property
        If lngContIndent >= 0 Then
            If IndentOf(strLine) > lngContIndent And Len(strTrim) > 0 Then GoTo NextLine
            lngContIndent = -1
        End If

        Select Case strTrim
            Case "PrtMip = Begin", "PrtDevMode = Begin", "PrtDevModeW = Begin", _
                 "PrtDevNames = Begin", "PrtDevNamesW = Begin"
                lngSkipIndent = IndentOf(strLine)
                strCapture = Left$(strTrim, InStr(1, strTrim, " ") - 1)
                dHex(strCapture) = vbNullString
                GoTo NextLine
            Case "GUID = Begin", "NameMap = Begin", "dbLongBinary ""DOL"" = Begin", "dbBinary ""GUID"" = Begin"
                lngSkipIndent = IndentOf(strLine)
                strCapture = vbNullString
                GoTo NextLine
            Case "NoSaveCTIWhenDisabled =1", "AllowPivotTableView =0", "AllowPivotChartView =0", _
                 "dbByte ""PublishToWeb"" =""1""", "PublishOption =1"
                GoTo NextLine
            Case "Version =21"
                strLine = Replace(strLine, "Version =21", "Version =20")
            Case "CodeBehindForm"
                ' Split the VBA code into a separate .cls file
                strOut(lngOut) = strLine
                lngOut = lngOut + 1
                If lngLine < UBound(varLines) Then
                    strVBA = SanitizeVBA(JoinLines(varLines, lngLine + 1, UBound(varLines)))
                    strOut(lngOut) = "' See """ & strSafeName & ".cls"""
                    lngOut = lngOut + 1
                End If
                Exit For
        End Select

        If Left$(strTrim, 10) = "Checksum =" Then GoTo NextLine
        If Left$(strTrim, 12) = "ColumnInfo =" Or Left$(strTrim, 10) = "BaseInfo =" Then
            lngContIndent = IndentOf(strLine)
            GoTo NextLine
        End If
        If Left$(strTrim, 15) = "WebImagePadding" Then GoTo NextLine

        If intType = acReport Then
            If Not blnReportRight And Left$(strLine, 11) = "    Right =" Then
                blnReportRight = True
                GoTo NextLine
            End If
            If Not blnReportBottom And Left$(strLine, 12) = "    Bottom =" Then
                blnReportBottom = True
                GoTo NextLine
            End If
        ElseIf intType = acForm Then
            If Not blnViewportDone Then
                If Left$(strLine, 10) = "    Left =" Then
                    strLine = "    Left =1000"
                ElseIf Left$(strLine, 9) = "    Top =" Then
                    strLine = "    Top =1000"
                ElseIf Left$(strLine, 11) = "    Right =" Then
                    strLine = "    Right =50000"
                ElseIf Left$(strLine, 12) = "    Bottom =" Then
                    strLine = "    Bottom =50000"
                    blnViewportDone = True
                End If
            End If
        End If

        ' Track blocks so redundant theme colors can be removed
        If strTrim = "Begin" Or Left$(strTrim, 6) = "Begin " Or Right$(strTrim, 8) = " = Begin" Then
            colStack.Add NewDict
        ElseIf strTrim = "End" Then
            If colStack.Count > 0 Then
                CloseColorBlock colStack(colStack.Count), blnDelete
                colStack.Remove colStack.Count
            End If
        ElseIf colStack.Count > 0 Then
            lngPos = InStr(1, strTrim, "=")
            If lngPos > 1 Then
                strProp = RTrim$(Left$(strTrim, lngPos - 1))
                strValue = Trim$(Mid$(strTrim, lngPos + 1))
                TrackColorProperty colStack(colStack.Count), strProp, strValue, lngOut
            End If
        End If

        strOut(lngOut) = strLine
        lngOut = lngOut + 1
NextLine:
    Next lngLine

    ' Assemble output (skipping removed color lines)
    ReDim strResult(0 To lngOut)
    lngCnt = 0
    For lngLine = 0 To lngOut - 1
        If Not blnDelete(lngLine) Then
            strResult(lngCnt) = strOut(lngLine)
            lngCnt = lngCnt + 1
        End If
    Next lngLine
    If lngCnt > 0 Then
        ReDim Preserve strResult(0 To lngCnt - 1)
        SanitizeObject = Join(strResult, vbCrLf)
    End If

    ' Parse the captured printer blocks
    If intType <> acMacro Then BuildPrintSettings dHex, dPrint

End Function


Private Function JoinLines(varLines As Variant, ByVal lngFirst As Long, ByVal lngLast As Long) As String
    Dim strLines() As String
    Dim lngCnt As Long
    If lngLast < lngFirst Then Exit Function
    ReDim strLines(0 To lngLast - lngFirst)
    For lngCnt = lngFirst To lngLast
        strLines(lngCnt - lngFirst) = varLines(lngCnt)
    Next lngCnt
    JoinLines = Join(strLines, vbCrLf)
End Function


Private Function IndentOf(ByVal strLine As String) As Long
    IndentOf = Len(strLine) - Len(LTrim$(strLine))
End Function


Private Function IsColorBase(ByVal strBase As String) As Boolean
    Select Case strBase
        Case "Back", "AlternateBack", "Border", "Fore", "Gridline", "HoverFore", "Hover", _
             "PressedFore", "Pressed", "DatasheetFore", "DatasheetBack", "DatasheetGridlines"
            IsColorBase = True
    End Select
End Function


Private Sub TrackColorProperty(dBlock As Object, ByVal strProp As String, ByVal strValue As String, ByVal lngIndex As Long)

    Dim strBase As String

    If strProp = "UseTheme" Then
        If strValue = "0" Then dBlock("UseThemeOff") = True
    ElseIf Right$(strProp, 15) = "ThemeColorIndex" Then
        strBase = Left$(strProp, Len(strProp) - 15)
        If IsColorBase(strBase) Then dBlock("T:" & strBase) = Val(strValue)
    ElseIf Right$(strProp, 5) = "Color" Then
        strBase = Left$(strProp, Len(strProp) - 5)
        If IsColorBase(strBase) Then
            If IsNumeric(strValue) Then
                If Val(strValue) >= 0 Then dBlock("C:" & strBase) = lngIndex
            End If
        End If
    End If

End Sub


Private Sub CloseColorBlock(dBlock As Object, blnDelete() As Boolean)

    Dim varKey As Variant
    Dim strBase As String

    If dBlock.Exists("UseThemeOff") Then Exit Sub
    For Each varKey In dBlock.Keys
        If Left$(varKey, 2) = "C:" Then
            strBase = Mid$(varKey, 3)
            If dBlock.Exists("T:" & strBase) Then
                ' Color is derived from the theme, so the literal value is noise.
                If dBlock("T:" & strBase) <> -1 Then blnDelete(dBlock(varKey)) = True
            End If
        End If
    Next varKey

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BuildPrintSettings
' Purpose   : Convert the PrtDevNames/PrtDevMode/PrtMip blocks into the add-in's
'           : "Device", "Printer" and "Margins" dictionaries.
'---------------------------------------------------------------------------------------
'
Private Sub BuildPrintSettings(dHex As Object, dPrint As Object)

    Dim bte() As Byte
    Dim lngSize As Long
    Dim lngBase As Long
    Dim lngFields As Long
    Dim blnWide As Boolean
    Dim dDevice As Object
    Dim dPrinter As Object
    Dim dMargins As Object

    On Error GoTo ErrHandler

    ' Device (only when the object uses a specific printer)
    If dHex.Exists("PrtDevNames") Then
        lngSize = HexToBytes(dHex("PrtDevNames"), bte)
    ElseIf dHex.Exists("PrtDevNamesW") Then
        lngSize = HexToBytes(dHex("PrtDevNamesW"), bte)
        blnWide = True
    End If
    If lngSize >= 8 Then
        If ReadInt16(bte, 0) > 0 And ReadInt16(bte, 6) <> 1 Then
            Set dDevice = NewDict
            dDevice.Add "DriverName", ReadCString(bte, ReadInt16(bte, 0) * IIf(blnWide, 2, 1), blnWide)
            dDevice.Add "DeviceName", ReadCString(bte, ReadInt16(bte, 2) * IIf(blnWide, 2, 1), blnWide)
            dDevice.Add "Port", ReadCString(bte, ReadInt16(bte, 4) * IIf(blnWide, 2, 1), blnWide)
            dDevice.Add "Default", False
            dPrint.Add "Device", dDevice
        End If
    End If

    ' Printer (orientation and paper size)
    lngSize = 0
    lngBase = 32
    If dHex.Exists("PrtDevMode") Then
        lngSize = HexToBytes(dHex("PrtDevMode"), bte)
    ElseIf dHex.Exists("PrtDevModeW") Then
        lngSize = HexToBytes(dHex("PrtDevModeW"), bte)
        lngBase = 64
    End If
    If lngSize >= lngBase + 16 Then
        lngFields = ReadInt32(bte, lngBase + 8)
        If lngFields <> 0 Then
            Set dPrinter = NewDict
            If (lngFields And &H1&) <> 0 Then dPrinter.Add "Orientation", OrientationName(ReadInt16(bte, lngBase + 12))
            If (lngFields And &H2&) <> 0 Then dPrinter.Add "PaperSize", PaperSizeName(ReadInt16(bte, lngBase + 14))
            dPrint.Add "Printer", dPrinter
        End If
    End If

    ' Margins and column layout
    lngSize = 0
    If dHex.Exists("PrtMip") Then lngSize = HexToBytes(dHex("PrtMip"), bte)
    If lngSize >= 48 Then
        If ReadInt32(bte, 44) > 0 Then
            Set dMargins = NewDict
            dMargins.Add "LeftMargin", TwipsToInches(ReadInt32(bte, 0))
            dMargins.Add "TopMargin", TwipsToInches(ReadInt32(bte, 4))
            dMargins.Add "RightMargin", TwipsToInches(ReadInt32(bte, 8))
            dMargins.Add "BotMargin", TwipsToInches(ReadInt32(bte, 12))
            dMargins.Add "DataOnly", (ReadInt32(bte, 16) <> 0)
            dMargins.Add "Width", TwipsToInches(ReadInt32(bte, 20))
            dMargins.Add "Height", TwipsToInches(ReadInt32(bte, 24))
            dMargins.Add "DefaultSize", (ReadInt32(bte, 28) <> 0)
            dMargins.Add "Columns", ReadInt32(bte, 32)
            dMargins.Add "ColumnSpacing", TwipsToInches(ReadInt32(bte, 36))
            dMargins.Add "RowSpacing", TwipsToInches(ReadInt32(bte, 40))
            dMargins.Add "ItemLayout", ColumnLayoutName(ReadInt32(bte, 44))
            dMargins.Add "FastPrint", ReadInt32(bte, 48)
            dMargins.Add "Datasheet", ReadInt32(bte, 52)
            dPrint.Add "Margins", dMargins
        End If
    End If
    Exit Sub

ErrHandler:
    LogWarning "Unable to parse printer settings (" & Err.Description & ")"
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsDefaultPrintSettings
' Purpose   : True if the print settings match the default printer, in which case the
'           : add-in does not write them to source. (FastPrint/Datasheet are ignored.)
'---------------------------------------------------------------------------------------
'
Private Function IsDefaultPrintSettings(dPrint As Object) As Boolean

    Dim dCompare As Object
    Dim dMargins As Object
    Dim varKey As Variant
    Dim varSub As Variant

    If IsEmpty(m_DefaultPrintJson) Then m_DefaultPrintJson = GetDefaultPrintJson

    ' Copy without the reserved settings
    Set dCompare = NewDict
    For Each varKey In dPrint.Keys
        If varKey = "Margins" Then
            Set dMargins = NewDict
            For Each varSub In dPrint(varKey).Keys
                If varSub <> "FastPrint" And varSub <> "Datasheet" Then dMargins.Add varSub, dPrint(varKey)(varSub)
            Next varSub
            dCompare.Add varKey, dMargins
        Else
            dCompare.Add varKey, dPrint(varKey)
        End If
    Next varKey

    IsDefaultPrintSettings = (JsonEncode(dCompare) = m_DefaultPrintJson)

End Function


Private Function GetDefaultPrintJson() As String

    Dim prt As Object
    Dim dDefault As Object
    Dim dPrinter As Object
    Dim dMargins As Object

    On Error GoTo ErrHandler
    Set prt = Application.Printer

    Set dPrinter = NewDict
    dPrinter.Add "Orientation", OrientationName(prt.Orientation)
    dPrinter.Add "PaperSize", PaperSizeName(prt.PaperSize)

    Set dMargins = NewDict
    dMargins.Add "LeftMargin", TwipsToInches(prt.LeftMargin)
    dMargins.Add "TopMargin", TwipsToInches(prt.TopMargin)
    dMargins.Add "RightMargin", TwipsToInches(prt.RightMargin)
    dMargins.Add "BotMargin", TwipsToInches(prt.BottomMargin)
    dMargins.Add "DataOnly", CBool(prt.DataOnly)
    dMargins.Add "Width", TwipsToInches(prt.ItemSizeWidth)
    dMargins.Add "Height", TwipsToInches(prt.ItemSizeHeight)
    dMargins.Add "DefaultSize", CBool(prt.DefaultSize)
    dMargins.Add "Columns", CLng(prt.ItemsAcross)
    dMargins.Add "ColumnSpacing", TwipsToInches(prt.ColumnSpacing)
    dMargins.Add "RowSpacing", TwipsToInches(prt.RowSpacing)
    dMargins.Add "ItemLayout", ColumnLayoutName(prt.ItemLayout)

    Set dDefault = NewDict
    dDefault.Add "Printer", dPrinter
    dDefault.Add "Margins", dMargins
    GetDefaultPrintJson = JsonEncode(dDefault)
    Exit Function

ErrHandler:
    ' No printer installed. Treat all print settings as non-default.
    GetDefaultPrintJson = "(none)"
    Err.Clear

End Function


Private Function TwipsToInches(ByVal lngTwips As Long) As Single
    TwipsToInches = CSng(Round(lngTwips / 1440, 4))
End Function


Private Function OrientationName(ByVal lngValue As Long) As Variant
    Select Case lngValue
        Case acPRORPortrait: OrientationName = "Portrait"
        Case acPRORLandscape: OrientationName = "Landscape"
        Case Else: OrientationName = lngValue
    End Select
End Function


Private Function ColumnLayoutName(ByVal lngValue As Long) As Variant
    Select Case lngValue
        Case acPRHorizontalColumnLayout: ColumnLayoutName = "Horizontal Columns"
        Case acPRVerticalColumnLayout: ColumnLayoutName = "Vertical Columns"
        Case Else: ColumnLayoutName = lngValue
    End Select
End Function


Private Function PaperSizeName(ByVal lngValue As Long) As Variant
    Select Case lngValue
        Case acPRPS10x14: PaperSizeName = "10x14"
        Case acPRPS11x17: PaperSizeName = "11x17"
        Case acPRPSA3: PaperSizeName = "A3"
        Case acPRPSA4: PaperSizeName = "A4"
        Case acPRPSA4Small: PaperSizeName = "A4 Small"
        Case acPRPSA5: PaperSizeName = "A5"
        Case acPRPSB4: PaperSizeName = "B4"
        Case acPRPSB5: PaperSizeName = "B5"
        Case acPRPSCSheet: PaperSizeName = "C Size Sheet"
        Case acPRPSDSheet: PaperSizeName = "D Size Sheet"
        Case acPRPSEnv10: PaperSizeName = "Envelope #10"
        Case acPRPSEnv11: PaperSizeName = "Envelope #11"
        Case acPRPSEnv12: PaperSizeName = "Envelope #12"
        Case acPRPSEnv14: PaperSizeName = "Envelope #14"
        Case acPRPSEnv9: PaperSizeName = "Envelope #9"
        Case acPRPSEnvB4: PaperSizeName = "Envelope B4"
        Case acPRPSEnvB5: PaperSizeName = "Envelope B5"
        Case acPRPSEnvB6: PaperSizeName = "Envelope B6"
        Case acPRPSEnvC3: PaperSizeName = "Envelope C3"
        Case acPRPSEnvC4: PaperSizeName = "Envelope C4"
        Case acPRPSEnvC5: PaperSizeName = "Envelope C5"
        Case acPRPSEnvC6: PaperSizeName = "Envelope C6"
        Case acPRPSEnvC65: PaperSizeName = "Envelope C65"
        Case acPRPSEnvDL: PaperSizeName = "Envelope DL"
        Case acPRPSEnvItaly: PaperSizeName = "Italian Envelope"
        Case acPRPSEnvMonarch: PaperSizeName = "Monarch Envelope"
        Case acPRPSEnvPersonal: PaperSizeName = "Envelope"
        Case acPRPSESheet: PaperSizeName = "E Size Sheet"
        Case acPRPSExecutive: PaperSizeName = "Executive"
        Case acPRPSFanfoldLglGerman: PaperSizeName = "German Legal Fanfold"
        Case acPRPSFanfoldStdGerman: PaperSizeName = "German Standard Fanfold"
        Case acPRPSFanfoldUS: PaperSizeName = "U.S. Standard Fanfold"
        Case acPRPSFolio: PaperSizeName = "Folio"
        Case acPRPSLedger: PaperSizeName = "Ledger"
        Case acPRPSLegal: PaperSizeName = "Legal"
        Case acPRPSLetter: PaperSizeName = "Letter"
        Case acPRPSLetterSmall: PaperSizeName = "Letter Small"
        Case acPRPSNote: PaperSizeName = "Note"
        Case acPRPSQuarto: PaperSizeName = "Quarto"
        Case acPRPSStatement: PaperSizeName = "Statement"
        Case acPRPSTabloid: PaperSizeName = "Tabloid"
        Case acPRPSUser: PaperSizeName = "User-Defined"
        Case Else: PaperSizeName = lngValue
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteObjectJson
' Purpose   : Write the companion .json file for an object. Adds the "Properties"
'           : (Description) and "Hidden" keys. Removes the file if there is no content.
'---------------------------------------------------------------------------------------
'
Private Sub WriteObjectJson(ByVal strFile As String, dItems As Object, ByVal strClass As String, _
    ByVal strDescription As String, ByVal strContainer As String, ByVal strName As String, ByVal intType As Long)

    If dItems Is Nothing Then
        Set dItems = NewDict
        strClass = vbNullString
        strDescription = strName & " Metadata"
    End If
    AddObjectMetadata dItems, strContainer, strName, intType

    If dItems.Count > 0 Then
        WriteTextFile strFile, BuildJsonFile(strClass, dItems, strDescription)
    Else
        DeleteFile strFile
    End If

End Sub


Private Sub AddObjectMetadata(dItems As Object, ByVal strContainer As String, ByVal strName As String, ByVal intType As Long)

    Dim prp As Object
    Dim dProps As Object
    Dim dProp As Object
    Dim varValue As Variant
    Dim blnHidden As Boolean

    On Error Resume Next
    Set prp = m_Dbs.Containers(strContainer).Documents(strName).Properties("Description")
    If Err.Number = 0 And Not prp Is Nothing Then
        varValue = prp.Value
        If Err.Number = 0 Then
            Set dProp = NewDict
            dProp.Add "Type", prp.Type
            dProp.Add "Value", varValue
            Set dProps = NewDict
            dProps.Add "Description", dProp
            dItems.Add "Properties", dProps
        End If
    End If
    Err.Clear

    blnHidden = Application.GetHiddenAttribute(intType, strName)
    If Err.Number = 0 And blnHidden Then dItems.Add "Hidden", True
    Err.Clear

End Sub


'=======================================================================================
' Queries
'=======================================================================================

Private Sub ExportAllQueries()
    Dim obj As Object
    LogLine "Exporting queries..."
    For Each obj In CurrentData.AllQueries
        If Left$(obj.Name, 1) <> "~" Then ExportQuery obj.Name
    Next obj
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportQuery
' Purpose   : Write the query SQL (.sql) and its metadata (.json).
'---------------------------------------------------------------------------------------
'
Private Sub ExportQuery(ByVal strName As String)

    Dim qdf As Object
    Dim dItems As Object
    Dim dProps As Object
    Dim dCols As Object
    Dim lngType As Long
    Dim blnPassThrough As Boolean
    Dim strBase As String
    Dim strSql As String

    On Error GoTo ErrHandler

    Status "Exporting query " & strName
    Set qdf = m_Dbs.QueryDefs(strName)
    lngType = qdf.Type
    blnPassThrough = (lngType = DB_QSQL_PASS_THROUGH Or lngType = DB_QSPT_BULK)
    strSql = qdf.SQL
    If Len(strSql) = 0 Then strSql = ";"

    strBase = m_Folder & "queries\" & GetSafeFileName(strName)
    WriteTextFile strBase & ".sql", strSql

    Set dItems = NewDict
    dItems.Add "QueryType", lngType
    If blnPassThrough Then dItems.Add "Connect", SanitizeConnect(qdf.Connect, "query " & strName)

    Set dProps = GetQueryProperties(qdf, blnPassThrough)
    If dProps.Count > 0 Then dItems.Add "QueryProperties", dProps

    If lngType = DB_QSELECT Then
        Set dCols = GetQueryColumns(qdf)
        If dCols.Count > 0 Then dItems.Add "Columns", dCols
    End If

    AddObjectMetadata dItems, "Tables", strName, acQuery
    WriteTextFile strBase & ".json", BuildJsonFile("clsDbQuery", dItems, strName)
    Exit Sub

ErrHandler:
    LogError "Error exporting query " & strName

End Sub


Private Function GetQueryProperties(qdf As Object, ByVal blnPassThrough As Boolean) As Object

    Dim dProps As Object
    Dim prp As Object
    Dim varName As Variant
    Dim varValue As Variant

    Set dProps = NewDict
    On Error Resume Next
    For Each varName In Array("ODBCTimeout", "MaxRecords", "ReturnsRecords", "RecordsetType", "UseTransaction", _
        "FailOnError", "RecordLocks", "LogMessages", "OrderByOn", "OrderByOnLoad", "FilterOnLoad", "TotalsRow", _
        "Orientation", "DefaultView", "SubdatasheetName", "SubdatasheetExpanded", "SubdatasheetHeight", _
        "LinkChildFields", "LinkMasterFields", "RowHeight")
        Set prp = Nothing
        Set prp = qdf.Properties(CStr(varName))
        If Err.Number = 0 And Not prp Is Nothing Then
            varValue = prp.Value
            If Err.Number = 0 Then
                If Not IsDefaultQueryProperty(CStr(varName), varValue, blnPassThrough) Then
                    dProps.Add CStr(varName), TypedValue(prp.Type, varValue)
                End If
            End If
        End If
        Err.Clear
    Next varName
    Set GetQueryProperties = dProps

End Function


Private Function IsDefaultQueryProperty(ByVal strName As String, ByVal varValue As Variant, ByVal blnPassThrough As Boolean) As Boolean

    If IsNull(varValue) Or IsEmpty(varValue) Then
        IsDefaultQueryProperty = True
        Exit Function
    End If

    Select Case strName
        Case "ODBCTimeout": IsDefaultQueryProperty = (varValue = 60)
        Case "MaxRecords", "RecordsetType", "RecordLocks", "Orientation", "SubdatasheetHeight"
            IsDefaultQueryProperty = (varValue = 0)
        Case "ReturnsRecords": IsDefaultQueryProperty = (Not blnPassThrough) Or (varValue = True)
        Case "UseTransaction", "OrderByOnLoad": IsDefaultQueryProperty = (varValue = True)
        Case "FailOnError", "LogMessages", "OrderByOn", "FilterOnLoad", "TotalsRow", "SubdatasheetExpanded"
            IsDefaultQueryProperty = (varValue = False)
        Case "DefaultView": IsDefaultQueryProperty = (varValue = 2)
        Case "RowHeight": IsDefaultQueryProperty = (varValue = -1 Or varValue = 65535)
        Case "SubdatasheetName", "LinkChildFields", "LinkMasterFields"
            IsDefaultQueryProperty = (Len(CStr(varValue)) = 0)
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetQueryColumns
' Purpose   : Column properties set in the query designer (width, caption, format...).
'---------------------------------------------------------------------------------------
'
Private Function GetQueryColumns(qdf As Object) As Object

    Dim dCols As Object
    Dim dCol As Object
    Dim colFields As Object
    Dim fld As Object
    Dim prp As Object
    Dim lngCnt As Long
    Dim lngFld As Long
    Dim varName As Variant
    Dim varValue As Variant
    Dim blnInherited As Boolean

    Set dCols = NewDict
    On Error Resume Next
    Set colFields = qdf.Fields
    lngCnt = colFields.Count
    If Err.Number <> 0 Then
        ' Source tables may not be available (for example linked tables)
        Err.Clear
        Set GetQueryColumns = dCols
        Exit Function
    End If

    For lngFld = 0 To lngCnt - 1
        Set fld = colFields(lngFld)
        Set dCol = NewDict
        For Each varName In Array("ColumnWidth", "ColumnHidden", "ColumnOrder", "Caption", "Description", _
            "Format", "DecimalPlaces", "InputMask", "TextAlign", "DisplayControl")
            Set prp = Nothing
            Set prp = fld.Properties(CStr(varName))
            If Err.Number = 0 And Not prp Is Nothing Then
                blnInherited = prp.Inherited
                varValue = prp.Value
                If Err.Number = 0 And Not blnInherited And Not IsNull(varValue) Then
                    Select Case varName
                        Case "ColumnWidth": If varValue <> -1 Then dCol.Add varName, varValue
                        Case "ColumnHidden": If varValue <> False Then dCol.Add varName, varValue
                        Case "ColumnOrder", "TextAlign": If varValue <> 0 Then dCol.Add varName, varValue
                        Case "DisplayControl": If varValue <> 109 Then dCol.Add varName, varValue
                        Case "Caption", "Description": If Len(varValue) > 0 Then dCol.Add varName, varValue
                        Case "Format", "InputMask": If Len(varValue) > 0 Then dCol.Add varName, TypedValue("dbText", varValue)
                        Case "DecimalPlaces": If varValue <> 255 Then dCol.Add varName, TypedValue("dbByte", varValue)
                    End Select
                End If
            End If
            Err.Clear
        Next varName
        If dCol.Count > 0 Then
            If Not dCols.Exists(fld.Name) Then dCols.Add fld.Name, dCol
        End If
    Next lngFld
    Err.Clear

    Set GetQueryColumns = SortDict(dCols)

End Function


Private Function TypedValue(ByVal varType As Variant, ByVal varValue As Variant) As Object
    Set TypedValue = NewDict
    TypedValue.Add "Type", varType
    TypedValue.Add "Value", varValue
End Function


'=======================================================================================
' Tables
'=======================================================================================

Private Function IsExportableTable(ByVal strName As String) As Boolean
    IsExportableTable = Not (strName Like "MSys*" Or strName Like "~*")
End Function


Private Sub ExportAllTables()

    Dim tdf As Object
    Dim colNames As Collection
    Dim varName As Variant

    LogLine "Exporting tables..."
    Set colNames = New Collection
    m_Dbs.TableDefs.Refresh
    For Each tdf In m_Dbs.TableDefs
        If IsExportableTable(tdf.Name) Then colNames.Add tdf.Name
    Next tdf

    For Each varName In colNames
        ExportTableDef CStr(varName)
        If m_TableData.Exists(varName) Then ExportTableData CStr(varName)
    Next varName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportTableDef
' Purpose   : Local tables: XML schema (.xml) + informational DDL (.sql).
'           : Linked tables: connection details (.json).
'---------------------------------------------------------------------------------------
'
Private Sub ExportTableDef(ByVal strName As String)

    Dim tdf As Object
    Dim strBase As String
    Dim strTemp As String
    Dim strXml As String
    Dim dItems As Object
    Dim blnHidden As Boolean

    On Error GoTo ErrHandler

    Status "Exporting table " & strName
    Set tdf = m_Dbs.TableDefs(strName)
    strBase = m_Folder & "tbldefs\" & GetSafeFileName(strName)

    If Len(tdf.Connect) > 0 Then
        ExportLinkedTable tdf, strBase
        Exit Sub
    End If

    ' Table structure as XSD schema with all table and field properties
    strTemp = GetTempFile(".xsd")
    Application.ExportXML acExportTable, strName, , strTemp, , , , AC_EXPORT_ALL_PROPERTIES
    strXml = SanitizeXML(ReadTextFile(strTemp, "utf-8"), "table " & strName)
    DeleteFile strTemp
    If Len(strXml) > 0 Then WriteTextFile strBase & ".xml", strXml

    ' Informational SQL (not used on import)
    WriteTextFile strBase & ".sql", GetTableDDL(tdf)

    ' Hidden flag (the description is part of the XML schema)
    Set dItems = NewDict
    On Error Resume Next
    blnHidden = Application.GetHiddenAttribute(acTable, strName)
    Err.Clear
    On Error GoTo ErrHandler
    If blnHidden Then
        dItems.Add "Hidden", True
        WriteTextFile strBase & ".json", BuildJsonFile(vbNullString, dItems, strName & " Metadata")
    Else
        DeleteFile strBase & ".json"
    End If
    Exit Sub

ErrHandler:
    LogError "Error exporting table " & strName

End Sub


Private Sub ExportLinkedTable(tdf As Object, ByVal strBase As String)

    Dim dItems As Object
    Dim strPK As String

    On Error GoTo ErrHandler

    Set dItems = NewDict
    dItems.Add "Name", tdf.Name
    dItems.Add "Connect", SanitizeConnect(GetRelativeConnect(tdf.Connect), "linked table " & tdf.Name)
    dItems.Add "SourceTableName", tdf.SourceTableName
    dItems.Add "Attributes", tdf.Attributes
    strPK = GetPrimaryKeyList(tdf)
    If Len(strPK) > 0 Then dItems.Add "PrimaryKey", strPK
    AddObjectMetadata dItems, "Tables", tdf.Name, acTable

    WriteTextFile strBase & ".json", BuildJsonFile("clsDbTableDef", dItems, "Linked Table")
    DeleteFile strBase & ".xml"
    DeleteFile strBase & ".sql"
    Exit Sub

ErrHandler:
    LogError "Error exporting linked table " & tdf.Name

End Sub


Private Function GetPrimaryKeyList(tdf As Object) As String

    Dim idx As Object
    Dim fld As Object
    Dim strList As String

    On Error GoTo ErrHandler
    For Each idx In tdf.Indexes
        If idx.Primary Then
            For Each fld In idx.Fields
                If Len(strList) > 0 Then strList = strList & ", "
                strList = strList & "[" & fld.Name & "]"
            Next fld
            Exit For
        End If
    Next idx
    GetPrimaryKeyList = strList
    Exit Function

ErrHandler:
    ' Index information may be unavailable when the source is offline.
    Err.Clear

End Function


Private Function GetTableDDL(tdf As Object) As String

    Dim fld As Object
    Dim idx As Object
    Dim strFields As String
    Dim strLine As String
    Dim strPK As String
    Dim strPKName As String

    On Error Resume Next
    For Each fld In tdf.Fields
        strLine = "  [" & fld.Name & "] "
        If (fld.Attributes And DB_AUTO_INCR_FIELD) <> 0 Then
            strLine = strLine & "AUTOINCREMENT"
        Else
            strLine = strLine & GetTypeString(fld.Type)
            If fld.Type = DB_TEXT Or fld.Type = DB_VAR_BINARY Then strLine = strLine & " (" & fld.Size & ")"
        End If
        If Len(strFields) > 0 Then strFields = strFields & "," & vbCrLf
        strFields = strFields & strLine
    Next fld

    For Each idx In tdf.Indexes
        If idx.Primary Then
            strPKName = idx.Name
            strPK = GetPrimaryKeyList(tdf)
            Exit For
        End If
    Next idx
    If Len(strPK) > 0 Then
        strFields = strFields & "," & vbCrLf & "   CONSTRAINT [" & strPKName & "] PRIMARY KEY (" & strPK & ")"
    End If
    Err.Clear

    GetTableDDL = "CREATE TABLE [" & tdf.Name & "] (" & vbCrLf & strFields & vbCrLf & ")"

End Function


Private Function GetTypeString(ByVal lngType As Long) As String
    Select Case lngType
        Case DB_LONG_BINARY: GetTypeString = "LONGBINARY"
        Case DB_BINARY: GetTypeString = "BINARY"
        Case DB_BOOLEAN: GetTypeString = "BIT"
        Case DB_CURRENCY: GetTypeString = "CURRENCY"
        Case DB_DATE, DB_TIME: GetTypeString = "DATETIME"
        Case DB_GUID: GetTypeString = "GUID"
        Case DB_MEMO: GetTypeString = "LONGTEXT"
        Case DB_DOUBLE: GetTypeString = "DOUBLE"
        Case DB_SINGLE: GetTypeString = "SINGLE"
        Case DB_BYTE: GetTypeString = "BYTE"
        Case DB_INTEGER: GetTypeString = "SHORT"
        Case DB_LONG: GetTypeString = "LONG"
        Case DB_NUMERIC: GetTypeString = "NUMERIC"
        Case Else: GetTypeString = "VARCHAR"
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExportTableData
' Purpose   : Save table data as "Tab Delimited" (.txt) or "XML Format" (.xml).
'---------------------------------------------------------------------------------------
'
Private Sub ExportTableData(ByVal strName As String)

    Dim strFormat As String
    Dim strBase As String
    Dim dSetting As Object

    On Error GoTo ErrHandler

    strFormat = "Tab Delimited"
    Set dSetting = GetDict(m_TableData, strName)
    If Not dSetting Is Nothing Then
        If dSetting.Exists("Format") Then strFormat = CStr(dSetting("Format"))
    End If

    strBase = m_Folder & "tables\" & GetSafeFileName(strName)
    Select Case strFormat
        Case "XML Format"
            ExportTableDataXml strName, strBase & ".xml"
            DeleteFile strBase & ".txt"
        Case "Tab Delimited"
            ExportTableDataTxt strName, strBase & ".txt"
            DeleteFile strBase & ".xml"
        Case Else
            ' "No Data" or unknown format
    End Select
    Exit Sub

ErrHandler:
    LogError "Error exporting data for table " & strName

End Sub


Private Function IsBinaryFieldType(ByVal lngType As Long) As Boolean
    Select Case lngType
        Case DB_BINARY, DB_LONG_BINARY, DB_VAR_BINARY
            IsBinaryFieldType = True
        Case Is >= DB_ATTACHMENT
            ' Attachment and multi-value (complex) fields
            IsBinaryFieldType = True
    End Select
End Function


Private Sub ExportTableDataTxt(ByVal strName As String, ByVal strFile As String)

    Dim tdf As Object
    Dim fld As Object
    Dim rst As Object
    Dim strFields As String
    Dim strOrder As String
    Dim strSql As String
    Dim strLines() As String
    Dim strValues() As String
    Dim lngLines As Long
    Dim lngCol As Long
    Dim lngCols As Long
    Dim blnBinary() As Boolean

    On Error GoTo ErrHandler

    Status "Exporting data for " & strName
    Set tdf = m_Dbs.TableDefs(strName)
    lngCols = tdf.Fields.Count
    If lngCols = 0 Then Exit Sub
    ReDim strValues(0 To lngCols - 1)
    ReDim blnBinary(0 To lngCols - 1)

    lngCol = 0
    For Each fld In tdf.Fields
        If Len(strFields) > 0 Then strFields = strFields & ", "
        strFields = strFields & "[" & fld.Name & "]"
        strValues(lngCol) = fld.Name
        blnBinary(lngCol) = IsBinaryFieldType(fld.Type)
        If Not blnBinary(lngCol) Then
            If Len(strOrder) > 0 Then strOrder = strOrder & ", "
            strOrder = strOrder & "[" & fld.Name & "]"
        End If
        lngCol = lngCol + 1
    Next fld

    ' Header row
    ReDim strLines(0 To 1023)
    strLines(0) = Join(strValues, vbTab)
    lngLines = 1

    strSql = "SELECT " & strFields & " FROM [" & strName & "]"
    If Len(strOrder) > 0 Then strSql = strSql & " ORDER BY " & strOrder
    Set rst = m_Dbs.OpenRecordset(strSql, DB_OPEN_SNAPSHOT)
    Do While Not rst.EOF
        For lngCol = 0 To lngCols - 1
            If blnBinary(lngCol) Then
                strValues(lngCol) = "UNSUPPORTED DATA TYPE"
            Else
                strValues(lngCol) = EscapeDataValue(CStr(Nz(rst.Fields(lngCol).Value, vbNullString)))
            End If
        Next lngCol
        If lngLines > UBound(strLines) Then ReDim Preserve strLines(0 To UBound(strLines) * 2 + 1)
        strLines(lngLines) = Join(strValues, vbTab)
        lngLines = lngLines + 1
        rst.MoveNext
    Loop
    rst.Close

    ReDim Preserve strLines(0 To lngLines - 1)
    WriteTextFile strFile, Join(strLines, vbCrLf)
    Exit Sub

ErrHandler:
    LogError "Error exporting data for table " & strName

End Sub


Private Function EscapeDataValue(ByVal strValue As String) As String
    strValue = Replace(strValue, "\", Chr$(26))
    strValue = Replace(strValue, vbCrLf, "\r\n")
    strValue = Replace(strValue, vbCr, "\r")
    strValue = Replace(strValue, vbLf, "\n")
    strValue = Replace(strValue, vbTab, "\t")
    EscapeDataValue = Replace(strValue, Chr$(26), "\\")
End Function


Private Sub ExportTableDataXml(ByVal strName As String, ByVal strFile As String)

    Dim strTemp As String
    Dim strXml As String

    On Error GoTo ErrHandler
    strTemp = GetTempFile(".xml")
    If Len(m_Dbs.TableDefs(strName).Connect) = 0 Then
        Application.ExportXML acExportTable, strName, strTemp, , , , , AC_EMBED_SCHEMA
    Else
        Application.ExportXML acExportTable, strName, strTemp
    End If
    strXml = SanitizeXML(ReadTextFile(strTemp, "utf-8"), "table data " & strName)
    DeleteFile strTemp
    If Len(strXml) > 0 Then WriteTextFile strFile, strXml
    Exit Sub

ErrHandler:
    LogError "Error exporting XML data for table " & strName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SanitizeXML
' Purpose   : Remove volatile content (timestamps, GUIDs, NameMap) and pretty-print the
'           : XML the same way as the add-in.
'---------------------------------------------------------------------------------------
'
Private Function SanitizeXML(ByVal strXml As String, ByVal strLabel As String) As String

    Dim objXml As Object
    Dim objXsl As Object
    Dim objNode As Object
    Dim objRoot As Object
    Dim objList As Object
    Dim lngCnt As Long
    Dim strOut As String

    On Error GoTo ErrHandler

    ' Some exports do not encode ampersands correctly
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "&[A-Za-z]{2,6};"
        If Not .Test(strXml) Then strXml = Replace(strXml, "&", "&amp;")
    End With

    Set objXml = CreateObject("MSXML2.DOMDocument.6.0")
    objXml.async = False
    If Not objXml.LoadXML(strXml) Then
        LogError "Unable to parse the XML for " & strLabel & " (" & objXml.parseError.reason & ")"
        Exit Function
    End If

    ' Table data with embedded schema
    Set objRoot = objXml.SelectSingleNode("/root/dataroot")
    If Not objRoot Is Nothing Then
        RemoveAttribute objRoot, "generated"
        ' Keep the schema only when it is needed to import the data
        If objXml.SelectNodes("//*[(namespace-uri()='http://www.w3.org/2001/XMLSchema' and local-name()='element' and @*[namespace-uri()='urn:schemas-microsoft-com:officedata' and ((local-name()='jetType' and (string()='complex' or string()='oleobject')) or (local-name()='expression'))])]").Length = 0 Then
            objXml.replaceChild objRoot, objXml.SelectSingleNode("/root")
        End If
    End If

    ' Table data without schema
    Set objRoot = objXml.SelectSingleNode("/dataroot")
    If Not objRoot Is Nothing Then RemoveAttribute objRoot, "generated"

    ' Remove noise
    Set objList = objXml.SelectNodes("//*[(namespace-uri()='urn:schemas-microsoft-com:officedata' and local-name()='tableProperty' and (@name='NameMap' or @name='GUID' or @name='PublishToWeb')) or (namespace-uri()='urn:schemas-microsoft-com:officedata' and local-name()='fieldProperty' and @name='GUID')]")
    For lngCnt = objList.Length - 1 To 0 Step -1
        Set objNode = objList.Item(lngCnt)
        objNode.ParentNode.RemoveChild objNode
    Next lngCnt

    ' Pretty-print
    Set objXsl = CreateObject("MSXML2.DOMDocument.6.0")
    objXsl.async = False
    objXsl.LoadXML XSLT_INDENT
    On Error Resume Next
    strOut = objXml.transformNode(objXsl)
    If Err.Number <> 0 Or Len(strOut) = 0 Then
        Err.Clear
        strOut = objXml.XML
    End If
    SanitizeXML = strOut
    Exit Function

ErrHandler:
    LogError "Error sanitizing XML for " & strLabel

End Function


Private Sub RemoveAttribute(objNode As Object, ByVal strName As String)
    On Error Resume Next
    objNode.Attributes.removeNamedItem strName
    Err.Clear
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportAllTableDataMacros
' Purpose   : Table data macros (tdmacros\*.xml)
'---------------------------------------------------------------------------------------
'
Private Sub ExportAllTableDataMacros()

    Dim tdf As Object
    Dim colNames As Collection
    Dim varName As Variant

    Set colNames = New Collection
    For Each tdf In m_Dbs.TableDefs
        If IsExportableTable(tdf.Name) And Len(tdf.Connect) = 0 Then colNames.Add tdf.Name
    Next tdf

    For Each varName In colNames
        If HasDataMacro(CStr(varName)) Then ExportTableDataMacro CStr(varName)
    Next varName

End Sub


Private Function HasDataMacro(ByVal strTable As String) As Boolean
    On Error GoTo ErrHandler
    HasDataMacro = (DCount("[Name]", "MSysObjects", "(LvExtra Is Not Null) And (Type = 1) And ([Name] = """ & _
        Replace(strTable, """", """""") & """)") > 0)
    Exit Function
ErrHandler:
    ' Unable to read MSysObjects; try exporting anyway.
    HasDataMacro = True
    Err.Clear
End Function


Private Sub ExportTableDataMacro(ByVal strTable As String)

    Dim strTemp As String
    Dim strXml As String
    Dim strFile As String

    strFile = m_Folder & "tdmacros\" & GetSafeFileName(strTable) & ".xml"
    strTemp = GetTempFile(".xml")

    On Error Resume Next
    Application.SaveAsText AC_TABLE_DATA_MACRO, strTable, strTemp
    If Err.Number <> 0 Then
        ' Error 2950 means there is no data macro for this table.
        If Err.Number <> 2950 Then LogWarning "Unable to export data macros for " & strTable & " (" & Err.Description & ")"
        Err.Clear
        DeleteFile strTemp
        Exit Sub
    End If
    On Error GoTo ErrHandler

    If m_FSO.FileExists(strTemp) Then
        strXml = SanitizeXML(ReadSourceFile(strTemp), "data macros of " & strTable)
        DeleteFile strTemp
        If Len(strXml) > 0 Then WriteTextFile strFile, strXml
    End If
    Exit Sub

ErrHandler:
    LogError "Error exporting data macros for " & strTable

End Sub


'=======================================================================================
' Relationships
'=======================================================================================

Private Sub ExportAllRelations()

    Dim rel As Object
    Dim fld As Object
    Dim dItems As Object
    Dim dField As Object
    Dim colFields As Collection
    Dim strFile As String

    On Error GoTo ErrHandler
    LogLine "Exporting relationships..."

    For Each rel In m_Dbs.Relations
        If rel.Name <> "MSysNavPaneGroupsMSysNavPaneGroupToObjects" _
            And rel.Name <> "MSysNavPaneGroupCategoriesMSysNavPaneGroups" _
            And (rel.Attributes And DB_RELATION_INHERITED) = 0 Then

            Set colFields = New Collection
            For Each fld In rel.Fields
                Set dField = NewDict
                dField.Add "Name", fld.Name
                dField.Add "ForeignName", fld.ForeignName
                colFields.Add dField
            Next fld

            Set dItems = NewDict
            dItems.Add "Name", rel.Name
            dItems.Add "Attributes", rel.Attributes
            dItems.Add "Table", rel.Table
            dItems.Add "ForeignTable", rel.ForeignTable
            dItems.Add "Fields", colFields

            strFile = rel.Name
            If InStr(1, strFile, "].") > 0 Then strFile = Mid$(strFile, InStr(1, strFile, "].") + 2)
            WriteTextFile m_Folder & "relations\" & GetSafeFileName(strFile) & ".json", _
                BuildJsonFile("clsDbRelation", dItems, "Database relationship")
        End If
    Next rel
    Exit Sub

ErrHandler:
    LogError "Error exporting relationships"

End Sub


'=======================================================================================
' Shared images, themes and specifications
'=======================================================================================

Private Function TableExists(ByVal strName As String) As Boolean
    Dim strTest As String
    On Error Resume Next
    strTest = m_Dbs.TableDefs(strName).Name
    TableExists = (Err.Number = 0)
    Err.Clear
End Function


Private Sub ExportThemes()

    Dim rst As Object
    Dim rstFiles As Object
    Dim strFile As String

    On Error GoTo ErrHandler
    If Not TableExists("MSysResources") Then Exit Sub

    Set rst = m_Dbs.OpenRecordset("SELECT [Name], [Data] FROM MSysResources WHERE [Type]='thmx'", DB_OPEN_DYNASET)
    Do While Not rst.EOF
        strFile = m_Folder & "themes\" & GetSafeFileName(Nz(rst.Fields("Name").Value, vbNullString)) & ".thmx"
        Set rstFiles = rst.Fields("Data").Value
        If Not rstFiles.EOF Then
            EnsureFolder m_FSO.GetParentFolderName(strFile)
            DeleteFile strFile
            rstFiles.Fields("FileData").SaveToFile strFile
            MarkWritten strFile
        End If
        rstFiles.Close
        rst.MoveNext
    Loop
    rst.Close
    Exit Sub

ErrHandler:
    LogError "Error exporting themes"

End Sub


Private Sub ExportSharedImages()

    Dim rst As Object
    Dim rstFiles As Object
    Dim dItems As Object
    Dim dSeen As Object
    Dim strName As String
    Dim strFileName As String
    Dim strBase As String
    Dim strImage As String
    Dim strHash As String

    On Error GoTo ErrHandler
    If Not TableExists("MSysResources") Then Exit Sub

    Set dSeen = NewDict
    Set rst = m_Dbs.OpenRecordset("SELECT * FROM MSysResources WHERE [Type]='img' ORDER BY [Id]", DB_OPEN_DYNASET)
    Do While Not rst.EOF
        strName = Nz(rst.Fields("Name").Value, vbNullString)
        strBase = m_Folder & "images\" & GetSafeFileName(strName)
        If dSeen.Exists(strBase) Then
            LogWarning "Duplicate shared image name skipped: " & strName
        Else
            dSeen.Add strBase, True
            Set rstFiles = rst.Fields("Data").Value
            If Not rstFiles.EOF Then
                strFileName = Nz(rstFiles.Fields("FileName").Value, vbNullString)
                strImage = strBase & "." & m_FSO.GetExtensionName(strFileName)
                EnsureFolder m_FSO.GetParentFolderName(strImage)
                DeleteFile strImage
                rstFiles.Fields("FileData").SaveToFile strImage
                MarkWritten strImage

                Set dItems = NewDict
                dItems.Add "Name", strName
                dItems.Add "FileName", strFileName
                dItems.Add "Extension", Nz(rst.Fields("Extension").Value, vbNullString)
                strHash = GetShortFileHash(strImage)
                If Len(strHash) > 0 Then dItems.Add "ContentHash", strHash
                WriteTextFile strBase & ".json", BuildJsonFile("clsDbSharedImage", dItems, "Shared Image Gallery Item")
            End If
            rstFiles.Close
        End If
        rst.MoveNext
    Loop
    rst.Close
    Exit Sub

ErrHandler:
    LogError "Error exporting shared images"

End Sub


' First 7 characters of the SHA256 hash of a file (used by the add-in to detect
' changes). Returns an empty string if .NET hashing is not available.
Private Function GetShortFileHash(ByVal strFile As String) As String

    Dim objHash As Object
    Dim bteData() As Byte
    Dim bteHash() As Byte
    Dim lngCnt As Long
    Dim strHash As String

    On Error GoTo ErrHandler
    bteData = ReadBinaryFile(strFile)
    Set objHash = CreateObject("System.Security.Cryptography.SHA256Managed")
    bteHash = objHash.ComputeHash_2((bteData))
    For lngCnt = 0 To 3
        strHash = strHash & LCase$(Right$("0" & Hex$(bteHash(lngCnt)), 2))
    Next lngCnt
    GetShortFileHash = Left$(strHash, 7)
    Exit Function

ErrHandler:
    Err.Clear

End Function


Private Sub ExportImexSpecs()

    Dim rst As Object
    Dim rstCols As Object
    Dim fld As Object
    Dim dItems As Object
    Dim dCols As Object
    Dim dCol As Object
    Dim strName As String

    On Error GoTo ErrHandler
    If Not TableExists("MSysIMEXSpecs") Then Exit Sub

    Set rst = m_Dbs.OpenRecordset("SELECT * FROM MSysIMEXSpecs", DB_OPEN_SNAPSHOT)
    Do While Not rst.EOF
        Set dItems = NewDict
        For Each fld In rst.Fields
            If fld.Name <> "SpecID" Then dItems.Add fld.Name, fld.Value
        Next fld

        Set dCols = NewDict
        Set rstCols = m_Dbs.OpenRecordset("SELECT * FROM MSysIMEXColumns WHERE SpecID=" & rst.Fields("SpecID").Value, DB_OPEN_SNAPSHOT)
        Do While Not rstCols.EOF
            Set dCol = NewDict
            For Each fld In rstCols.Fields
                If fld.Name <> "SpecID" And fld.Name <> "FieldName" Then dCol.Add fld.Name, fld.Value
            Next fld
            If Not dCols.Exists(Nz(rstCols.Fields("FieldName").Value, vbNullString)) Then
                dCols.Add Nz(rstCols.Fields("FieldName").Value, vbNullString), dCol
            End If
            rstCols.MoveNext
        Loop
        rstCols.Close
        dItems.Add "Columns", dCols

        strName = Nz(rst.Fields("SpecName").Value, "Spec " & rst.Fields("SpecID").Value)
        WriteTextFile m_Folder & "imexspecs\" & GetSafeFileName(strName) & ".json", _
            BuildJsonFile("clsDbImexSpec", dItems, "Import/Export Specification from MSysIMEXSpecs")
        rst.MoveNext
    Loop
    rst.Close
    Exit Sub

ErrHandler:
    LogError "Error exporting import/export specifications"

End Sub


Private Sub ExportSavedSpecs()

    Dim prj As Object
    Dim colSpecs As Object
    Dim objSpec As Object
    Dim dItems As Object
    Dim strDescription As String
    Dim strXml As String
    Dim lngCnt As Long

    On Error GoTo ErrHandler
    Set prj = CurrentProject
    Set colSpecs = prj.ImportExportSpecifications

    For lngCnt = 0 To colSpecs.Count - 1
        Set objSpec = colSpecs(lngCnt)
        Set dItems = NewDict
        dItems.Add "Name", objSpec.Name
        On Error Resume Next
        strDescription = vbNullString
        strDescription = objSpec.Description
        If Err.Number = 0 Then dItems.Add "Description", strDescription
        Err.Clear
        On Error GoTo ErrHandler
        strXml = objSpec.XML
        Do While InStr(1, strXml, vbCr & vbCr) > 0
            strXml = Replace(strXml, vbCr & vbCr, vbCr)
        Loop
        dItems.Add "XML", strXml
        WriteTextFile m_Folder & "savedspecs\" & GetSafeFileName(objSpec.Name) & ".json", _
            BuildJsonFile("clsDbSavedSpec", dItems, "Saved Import/Export Specification")
    Next lngCnt
    Exit Sub

ErrHandler:
    LogError "Error exporting saved import/export specifications"

End Sub


'=======================================================================================
' Orphaned file cleanup
'=======================================================================================

'---------------------------------------------------------------------------------------
' Procedure : RemoveOrphanedFiles
' Purpose   : Remove source files for objects that no longer exist in the database.
'           : Only files with extensions this exporter manages are removed.
'---------------------------------------------------------------------------------------
'
Private Sub RemoveOrphanedFiles()

    Dim varFolder As Variant
    Dim varFile As Variant

    ' Never remove files when an object could not be exported (its previous files
    ' would look like orphans).
    If m_Errors > 0 Then
        LogWarning "Orphaned files were not removed because errors occurred during the export."
        Exit Sub
    End If

    On Error Resume Next
    For Each varFolder In Array("modules", "forms", "reports", "macros", "queries", "tbldefs", "tables", _
        "relations", "tdmacros", "themes", "images", "imexspecs", "savedspecs")
        If m_FSO.FolderExists(m_Folder & varFolder) Then
            CleanFolder m_FSO.GetFolder(m_Folder & varFolder), CStr(varFolder), True
        End If
    Next varFolder

    ' Root-level files that are only written when there is content
    For Each varFile In Array("proj-properties.json", "documents.json", "hidden-attributes.json")
        If m_FSO.FileExists(m_Folder & varFile) And Not WasWritten(m_Folder & varFile) Then
            DeleteFile m_Folder & varFile
        End If
    Next varFile
    Err.Clear

End Sub


Private Sub CleanFolder(objFolder As Object, ByVal strCategory As String, ByVal blnRoot As Boolean)

    Dim objFile As Object
    Dim objSub As Object
    Dim colFiles As Collection
    Dim colSubs As Collection
    Dim varItem As Variant

    Set colFiles = New Collection
    Set colSubs = New Collection
    For Each objFile In objFolder.Files
        colFiles.Add objFile.Path
    Next objFile
    For Each objSub In objFolder.SubFolders
        colSubs.Add objSub.Path
    Next objSub

    For Each varItem In colFiles
        If IsManagedExtension(strCategory, LCase$(m_FSO.GetExtensionName(varItem))) Then
            If Not WasWritten(CStr(varItem)) Then
                LogLine "  Removing orphaned file: " & Mid$(varItem, Len(m_Folder) + 1)
                DeleteFile CStr(varItem)
            End If
        End If
    Next varItem

    If strCategory <> "themes" Then
        For Each varItem In colSubs
            CleanFolder m_FSO.GetFolder(varItem), strCategory, False
        Next varItem
    End If

    ' Remove empty subfolders
    If Not blnRoot Then
        If objFolder.Files.Count = 0 And objFolder.SubFolders.Count = 0 Then objFolder.Delete True
    End If

End Sub


Private Function IsManagedExtension(ByVal strCategory As String, ByVal strExt As String) As Boolean
    Select Case strCategory
        Case "modules": IsManagedExtension = (strExt = "bas" Or strExt = "cls" Or strExt = "json")
        Case "forms": IsManagedExtension = (strExt = "form" Or strExt = "bas" Or strExt = "cls" Or strExt = "json" Or strExt = "svg")
        Case "reports": IsManagedExtension = (strExt = "report" Or strExt = "bas" Or strExt = "cls" Or strExt = "json" Or strExt = "svg")
        Case "macros": IsManagedExtension = (strExt = "macro" Or strExt = "bas" Or strExt = "json")
        Case "queries": IsManagedExtension = (strExt = "sql" Or strExt = "json" Or strExt = "qdef" Or strExt = "bas")
        Case "tbldefs": IsManagedExtension = (strExt = "xml" Or strExt = "json" Or strExt = "sql")
        Case "tables": IsManagedExtension = (strExt = "txt" Or strExt = "xml")
        Case "relations", "imexspecs", "savedspecs": IsManagedExtension = (strExt = "json")
        Case "tdmacros": IsManagedExtension = (strExt = "xml")
        Case "themes": IsManagedExtension = (strExt = "thmx")
        Case "images": IsManagedExtension = True
    End Select
End Function


'=======================================================================================
' Connection strings and paths
'=======================================================================================

'---------------------------------------------------------------------------------------
' Procedure : GetRelativePath
' Purpose   : Paths inside the database folder are stored as "rel:<relative path>".
'---------------------------------------------------------------------------------------
'
Private Function GetRelativePath(ByVal strPath As String) As String

    Dim strBase As String

    strBase = CurrentProject.Path & "\"
    If StrComp(strPath, CurrentProject.Path, vbTextCompare) = 0 Then
        GetRelativePath = "rel:"
    ElseIf StrComp(Left$(strPath, Len(strBase)), strBase, vbTextCompare) = 0 Then
        GetRelativePath = "rel:" & Mid$(strPath, Len(strBase) + 1)
    Else
        GetRelativePath = strPath
    End If

End Function


Private Function GetRelativeConnect(ByVal strConnect As String) As String

    Dim varParts As Variant
    Dim lngPart As Long
    Dim strPath As String

    varParts = Split(strConnect, ";")
    For lngPart = 0 To UBound(varParts)
        If StrComp(Left$(varParts(lngPart), 9), "DATABASE=", vbTextCompare) = 0 Then
            strPath = Mid$(varParts(lngPart), 10)
            If Right$(strPath, 1) = "\" And Len(strPath) > 3 Then strPath = Left$(strPath, Len(strPath) - 1)
            If Len(strPath) > 0 And Left$(strPath, 4) <> "rel:" Then
                varParts(lngPart) = Left$(varParts(lngPart), 9) & GetRelativePath(strPath)
            End If
        End If
    Next lngPart
    GetRelativeConnect = Join(varParts, ";")

End Function


'---------------------------------------------------------------------------------------
' Procedure : SanitizeConnect
' Purpose   : Remove credentials from a connection string when a password is present.
'---------------------------------------------------------------------------------------
'
Private Function SanitizeConnect(ByVal strConnect As String, ByVal strLabel As String) As String

    Dim varParts As Variant
    Dim lngPart As Long
    Dim strOut As String
    Dim strKey As String
    Dim blnHasPassword As Boolean

    SanitizeConnect = strConnect
    If Not STRIP_CREDENTIALS Or Len(strConnect) = 0 Then Exit Function

    varParts = Split(strConnect, ";")
    For lngPart = 0 To UBound(varParts)
        If StrComp(Left$(Trim$(varParts(lngPart)), 4), "PWD=", vbTextCompare) = 0 Then
            If Len(Mid$(Trim$(varParts(lngPart)), 5)) > 0 Then blnHasPassword = True
        End If
    Next lngPart
    If Not blnHasPassword Then Exit Function

    For lngPart = 0 To UBound(varParts)
        strKey = UCase$(Trim$(varParts(lngPart)))
        If Not (Left$(strKey, 4) = "UID=" Or Left$(strKey, 4) = "PWD=" Or Len(strKey) = 0) Then
            strOut = strOut & varParts(lngPart) & ";"
        End If
    Next lngPart
    If Right$(strConnect, 1) <> ";" And Right$(strOut, 1) = ";" Then strOut = Left$(strOut, Len(strOut) - 1)

    LogWarning "Removed credentials from the connection string of " & strLabel & _
        ". Supply them again after the database is rebuilt."
    SanitizeConnect = strOut

End Function


'=======================================================================================
' Safe file names
'=======================================================================================

Private Function GetSafeFileName(ByVal strName As String) As String
    Dim strSafe As String
    strSafe = Replace(strName, "%", "%25")      ' This must be done first.
    strSafe = Replace(strSafe, "<", "%3C")
    strSafe = Replace(strSafe, ">", "%3E")
    strSafe = Replace(strSafe, ":", "%3A")
    strSafe = Replace(strSafe, """", "%22")
    strSafe = Replace(strSafe, "/", "%2F")
    strSafe = Replace(strSafe, "\", "%5C")
    strSafe = Replace(strSafe, "|", "%7C")
    strSafe = Replace(strSafe, "?", "%3F")
    strSafe = Replace(strSafe, "*", "%2A")
    GetSafeFileName = strSafe
End Function


'=======================================================================================
' File and text helpers
'=======================================================================================

Private Function GetVBProject() As Object

    Dim prj As Object
    Dim strFile As String

    On Error Resume Next
    For Each prj In Application.VBE.VBProjects
        strFile = vbNullString
        strFile = prj.FileName
        If StrComp(strFile, CurrentProject.FullName, vbTextCompare) = 0 Then
            Set GetVBProject = prj
            Exit For
        End If
    Next prj
    If GetVBProject Is Nothing Then Set GetVBProject = Application.VBE.ActiveVBProject
    Err.Clear

End Function


Private Function GetAnsiCharset() As String
    Dim lngCodePage As Long
    lngCodePage = GetACP
    Select Case lngCodePage
        Case 874: GetAnsiCharset = "windows-874"
        Case 932: GetAnsiCharset = "shift_jis"
        Case 936: GetAnsiCharset = "gb2312"
        Case 949: GetAnsiCharset = "ks_c_5601-1987"
        Case 950: GetAnsiCharset = "big5"
        Case 1250 To 1258: GetAnsiCharset = "windows-" & lngCodePage
        Case Else: GetAnsiCharset = "windows-1252"
    End Select
End Function


Private Function GetTempFile(Optional ByVal strExtension As String = ".tmp") As String
    GetTempFile = m_FSO.GetSpecialFolder(2).Path & "\VCS_" & m_FSO.GetBaseName(m_FSO.GetTempName) & strExtension
End Function


Private Sub EnsureFolder(ByVal strFolder As String)
    If Right$(strFolder, 1) = "\" Then strFolder = Left$(strFolder, Len(strFolder) - 1)
    If Len(strFolder) = 0 Then Exit Sub
    If m_FSO.FolderExists(strFolder) Then Exit Sub
    EnsureFolder m_FSO.GetParentFolderName(strFolder)
    m_FSO.CreateFolder strFolder
End Sub


Private Sub DeleteFile(ByVal strFile As String)
    On Error Resume Next
    If m_FSO.FileExists(strFile) Then m_FSO.DeleteFile strFile, True
    Err.Clear
End Sub


Private Sub MarkWritten(ByVal strFile As String)
    m_Written(LCase$(m_FSO.GetAbsolutePathName(strFile))) = True
End Sub


Private Function WasWritten(ByVal strFile As String) As Boolean
    WasWritten = m_Written.Exists(LCase$(m_FSO.GetAbsolutePathName(strFile)))
End Function


Private Function NormalizeLineEndings(ByVal strText As String) As String
    strText = Replace(strText, vbCrLf, vbLf)
    strText = Replace(strText, vbCr, vbLf)
    NormalizeLineEndings = Replace(strText, vbLf, vbCrLf)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadTextFile
' Purpose   : Read a text file with the given charset. Line endings are normalized
'           : to CRLF and any byte order mark is removed.
'---------------------------------------------------------------------------------------
'
Private Function ReadTextFile(ByVal strFile As String, Optional ByVal strCharset As String = "utf-8") As String

    Dim strText As String

    If Not m_FSO.FileExists(strFile) Then Exit Function
    With CreateObject("ADODB.Stream")
        .Type = AD_TYPE_TEXT
        .Charset = strCharset
        .Open
        .LoadFromFile strFile
        strText = .ReadText(AD_READ_ALL)
        .Close
    End With
    If Left$(strText, 1) = ChrW$(&HFEFF) Then strText = Mid$(strText, 2)
    ReadTextFile = NormalizeLineEndings(strText)

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadSourceFile
' Purpose   : Read a file created by SaveAsText. (UTF-16 in .accdb files, ANSI in
'           : Access 2000-2003 .mdb files.)
'---------------------------------------------------------------------------------------
'
Private Function ReadSourceFile(ByVal strFile As String) As String

    Dim bteData() As Byte

    bteData = ReadBinaryFile(strFile, 3)
    If UBound(bteData) >= 1 Then
        If bteData(0) = &HFF And bteData(1) = &HFE Then
            ReadSourceFile = ReadTextFile(strFile, "unicode")
            Exit Function
        End If
    End If
    If UBound(bteData) >= 2 Then
        If bteData(0) = &HEF And bteData(1) = &HBB And bteData(2) = &HBF Then
            ReadSourceFile = ReadTextFile(strFile, "utf-8")
            Exit Function
        End If
    End If
    If Val(m_Dbs.Version) < 5 Then
        ReadSourceFile = ReadTextFile(strFile, m_AnsiCharset)
    Else
        ReadSourceFile = ReadTextFile(strFile, "utf-8")
    End If

End Function


Private Function ReadBinaryFile(ByVal strFile As String, Optional ByVal lngMaxBytes As Long = -1) As Byte()

    Dim bteEmpty() As Byte

    ReDim bteEmpty(0 To 0)
    If Not m_FSO.FileExists(strFile) Then
        ReadBinaryFile = bteEmpty
        Exit Function
    End If
    With CreateObject("ADODB.Stream")
        .Type = AD_TYPE_BINARY
        .Open
        .LoadFromFile strFile
        If .Size = 0 Then
            ReadBinaryFile = bteEmpty
        Else
            ReadBinaryFile = .Read(lngMaxBytes)
        End If
        .Close
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteTextFile
' Purpose   : Write text with CRLF line endings and a trailing line break. UTF-8 output
'           : (the default) always includes a byte order mark, like the add-in.
'---------------------------------------------------------------------------------------
'
Private Sub WriteTextFile(ByVal strFile As String, ByVal strText As String, Optional ByVal strCharset As String = "utf-8")

    strText = NormalizeLineEndings(strText)
    If Right$(strText, 2) <> vbCrLf Then strText = strText & vbCrLf
    EnsureFolder m_FSO.GetParentFolderName(strFile)

    ' If the existing file name differs only by case, replace it so the new case is used.
    If m_FSO.FileExists(strFile) Then
        If StrComp(m_FSO.GetFile(strFile).Name, m_FSO.GetFileName(strFile), vbBinaryCompare) <> 0 Then
            m_FSO.DeleteFile strFile, True
        End If
    End If

    With CreateObject("ADODB.Stream")
        .Type = AD_TYPE_TEXT
        .Charset = strCharset
        .Open
        .WriteText strText
        .SaveToFile strFile, AD_SAVE_CREATE_OVERWRITE
        .Close
    End With
    MarkWritten strFile

End Sub


'---------------------------------------------------------------------------------------
' Binary helpers for printer settings
'---------------------------------------------------------------------------------------
'
Private Function HexToBytes(ByVal strHex As String, bteOut() As Byte) As Long

    Dim lngCnt As Long
    Dim lngSize As Long

    strHex = Replace(strHex, "0x", vbNullString, , , vbTextCompare)
    strHex = Replace(strHex, ",", vbNullString)
    strHex = Replace(strHex, " ", vbNullString)
    lngSize = Len(strHex) \ 2
    If lngSize = 0 Then Exit Function
    ReDim bteOut(0 To lngSize - 1)
    For lngCnt = 0 To lngSize - 1
        bteOut(lngCnt) = CByte("&H" & Mid$(strHex, lngCnt * 2 + 1, 2))
    Next lngCnt
    HexToBytes = lngSize

End Function


Private Function ReadInt16(bte() As Byte, ByVal lngPos As Long) As Long
    If lngPos < 0 Or lngPos + 1 > UBound(bte) Then Exit Function
    ReadInt16 = CLng(bte(lngPos)) + CLng(bte(lngPos + 1)) * 256&
    If ReadInt16 >= 32768 Then ReadInt16 = ReadInt16 - 65536
End Function


Private Function ReadInt32(bte() As Byte, ByVal lngPos As Long) As Long
    Dim dblValue As Double
    If lngPos < 0 Or lngPos + 3 > UBound(bte) Then Exit Function
    dblValue = CDbl(bte(lngPos)) + CDbl(bte(lngPos + 1)) * 256# + _
        CDbl(bte(lngPos + 2)) * 65536# + CDbl(bte(lngPos + 3)) * 16777216#
    If dblValue >= 2147483648# Then dblValue = dblValue - 4294967296#
    ReadInt32 = CLng(dblValue)
End Function


Private Function ReadCString(bte() As Byte, ByVal lngPos As Long, ByVal blnWide As Boolean) As String
    Dim strOut As String
    Dim lngCode As Long
    Do While lngPos <= UBound(bte)
        If blnWide Then
            If lngPos + 1 > UBound(bte) Then Exit Do
            lngCode = CLng(bte(lngPos)) + CLng(bte(lngPos + 1)) * 256&
            lngPos = lngPos + 2
        Else
            lngCode = bte(lngPos)
            lngPos = lngPos + 1
        End If
        If lngCode = 0 Then Exit Do
        If blnWide Then
            strOut = strOut & ChrW$(lngCode)
        Else
            strOut = strOut & Chr$(lngCode)
        End If
    Loop
    ReadCString = strOut
End Function


'---------------------------------------------------------------------------------------
' Procedure : ToIsoUtc
' Purpose   : Convert a local date/time to ISO 8601 UTC (as used in dbs-properties.json)
'---------------------------------------------------------------------------------------
'
Private Function ToIsoUtc(ByVal dteLocal As Date) As String
    Dim dteUtc As Date
    On Error Resume Next
    dteUtc = dteLocal
    With CreateObject("WbemScripting.SWbemDateTime")
        .SetVarDate dteLocal, True
        dteUtc = .GetVarDate(False)
    End With
    Err.Clear
    ToIsoUtc = Format$(dteUtc, "yyyy-mm-dd") & "T" & Format$(dteUtc, "hh:nn:ss") & ".000Z"
End Function


'=======================================================================================
' Logging
'=======================================================================================

Private Sub LogLine(ByVal strText As String)
    Debug.Print strText
    If Not m_Log Is Nothing Then m_Log.Add strText
End Sub


Private Sub Status(ByVal strText As String)
    On Error Resume Next
    SysCmd acSysCmdSetStatus, Left$(strText, 200)
    DoEvents
    Err.Clear
End Sub


Private Sub LogWarning(ByVal strText As String)
    m_Warnings = m_Warnings + 1
    LogLine "  WARNING: " & strText
End Sub


Private Sub LogError(ByVal strText As String)
    m_Errors = m_Errors + 1
    If Err.Number <> 0 Then strText = strText & " (Error " & Err.Number & ": " & Err.Description & ")"
    LogLine "  ERROR: " & strText
    Err.Clear
End Sub


'=======================================================================================
' Dictionaries and JSON
'=======================================================================================

Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
    NewDict.CompareMode = vbTextCompare
End Function


' Return a child dictionary, or Nothing if it does not exist.
Private Function GetDict(dParent As Object, ByVal strKey As String) As Object
    If dParent Is Nothing Then Exit Function
    If Not dParent.Exists(strKey) Then Exit Function
    If IsObject(dParent(strKey)) Then
        If TypeName(dParent(strKey)) = "Dictionary" Then Set GetDict = dParent(strKey)
    End If
End Function


' Return a copy of the dictionary sorted by key (case-insensitive, like the add-in)
Private Function SortDict(dSource As Object) As Object

    Dim varKeys() As Variant
    Dim varKey As Variant
    Dim varTemp As Variant
    Dim lngCnt As Long
    Dim lngPos As Long
    Dim dSorted As Object

    Set dSorted = NewDict
    If dSource.Count > 0 Then
        ReDim varKeys(0 To dSource.Count - 1)
        lngCnt = 0
        For Each varKey In dSource.Keys
            varKeys(lngCnt) = varKey
            lngCnt = lngCnt + 1
        Next varKey
        ' Insertion sort, using the database sort order like the add-in (lists are small)
        For lngCnt = 1 To UBound(varKeys)
            varTemp = varKeys(lngCnt)
            lngPos = lngCnt - 1
            Do While lngPos >= 0
                If StrComp(varKeys(lngPos), varTemp, vbDatabaseCompare) > 0 Then
                    varKeys(lngPos + 1) = varKeys(lngPos)
                    lngPos = lngPos - 1
                Else
                    Exit Do
                End If
            Loop
            varKeys(lngPos + 1) = varTemp
        Next lngCnt
        For lngCnt = 0 To UBound(varKeys)
            dSorted.Add varKeys(lngCnt), dSource(varKeys(lngCnt))
        Next lngCnt
    End If
    Set SortDict = dSorted

End Function


Private Function BuildJsonFile(ByVal strClass As String, dItems As Object, ByVal strDescription As String) As String

    Dim dFile As Object
    Dim dInfo As Object

    Set dInfo = NewDict
    dInfo.Add "Class", strClass
    dInfo.Add "Description", strDescription
    Set dFile = NewDict
    dFile.Add "Info", dInfo
    dFile.Add "Items", dItems
    BuildJsonFile = JsonEncode(dFile)

End Function


'---------------------------------------------------------------------------------------
' Procedure : JsonEncode
' Purpose   : Convert a value to JSON using the add-in's layout (2-space indent, CRLF,
'           : "key": value). Dictionary items with an Empty value are omitted.
'---------------------------------------------------------------------------------------
'
Private Function JsonEncode(ByVal varValue As Variant, Optional ByVal lngLevel As Long = 0) As String

    Dim varKey As Variant
    Dim varItem As Variant
    Dim strOut As String
    Dim lngCnt As Long
    Dim lngIndex As Long

    If IsObject(varValue) Then
        If varValue Is Nothing Then
            JsonEncode = "null"
        ElseIf TypeName(varValue) = "Dictionary" Then
            strOut = "{"
            For Each varKey In varValue.Keys
                If IsObject(varValue(varKey)) Then
                    varItem = Empty
                    Set varItem = varValue(varKey)
                Else
                    varItem = varValue(varKey)
                End If
                If Not IsEmptyItem(varItem) Then
                    If lngCnt > 0 Then strOut = strOut & ","
                    strOut = strOut & vbCrLf & Space$((lngLevel + 1) * 2) & """" & JsonEscape(CStr(varKey)) & """: " & _
                        JsonEncode(varItem, lngLevel + 1)
                    lngCnt = lngCnt + 1
                End If
            Next varKey
            JsonEncode = strOut & vbCrLf & Space$(lngLevel * 2) & "}"
        ElseIf TypeName(varValue) = "Collection" Then
            strOut = "["
            For Each varItem In varValue
                If lngCnt > 0 Then strOut = strOut & ","
                strOut = strOut & vbCrLf & Space$((lngLevel + 1) * 2) & JsonEncode(varItem, lngLevel + 1)
                lngCnt = lngCnt + 1
            Next varItem
            JsonEncode = strOut & vbCrLf & Space$(lngLevel * 2) & "]"
        Else
            JsonEncode = "null"
        End If
    ElseIf IsArray(varValue) Then
        strOut = "["
        For lngIndex = LBound(varValue) To UBound(varValue)
            If lngCnt > 0 Then strOut = strOut & ","
            strOut = strOut & vbCrLf & Space$((lngLevel + 1) * 2) & JsonEncode(varValue(lngIndex), lngLevel + 1)
            lngCnt = lngCnt + 1
        Next lngIndex
        JsonEncode = strOut & vbCrLf & Space$(lngLevel * 2) & "]"
    Else
        Select Case VarType(varValue)
            Case vbNull, vbEmpty
                JsonEncode = "null"
            Case vbBoolean
                JsonEncode = IIf(varValue, "true", "false")
            Case vbString
                JsonEncode = """" & JsonEscape(varValue) & """"
            Case vbDate
                JsonEncode = """" & CStr(varValue) & """"
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, 20
                JsonEncode = Replace(CStr(varValue), ",", ".")
            Case Else
                JsonEncode = """" & JsonEscape(CStr(varValue)) & """"
        End Select
    End If

End Function


Private Function IsEmptyItem(ByVal varItem As Variant) As Boolean
    If Not IsObject(varItem) Then IsEmptyItem = IsEmpty(varItem)
End Function


Private Function JsonEscape(ByVal strText As String) As String

    Dim lngCode As Long

    strText = Replace(strText, "\", "\\")
    strText = Replace(strText, """", "\""")
    strText = Replace(strText, Chr$(8), "\b")
    strText = Replace(strText, Chr$(12), "\f")
    strText = Replace(strText, vbLf, "\n")
    strText = Replace(strText, vbCr, "\r")
    strText = Replace(strText, vbTab, "\t")
    For lngCode = 0 To 31
        If InStr(1, strText, ChrW$(lngCode), vbBinaryCompare) > 0 Then
            strText = Replace(strText, ChrW$(lngCode), "\u" & Right$("000" & Hex$(lngCode), 4))
        End If
    Next lngCode
    For lngCode = 127 To 159
        If InStr(1, strText, ChrW$(lngCode), vbBinaryCompare) > 0 Then
            strText = Replace(strText, ChrW$(lngCode), "\u" & Right$("000" & Hex$(lngCode), 4))
        End If
    Next lngCode
    JsonEscape = strText

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadJsonFile
' Purpose   : Parse a json file into dictionaries (objects) and collections (arrays).
'           : Returns Nothing if the file does not exist or cannot be parsed.
'---------------------------------------------------------------------------------------
'
Private Function ReadJsonFile(ByVal strFile As String) As Object

    Dim strJson As String
    Dim lngPos As Long
    Dim varResult As Variant

    If Not m_FSO.FileExists(strFile) Then Exit Function
    strJson = ReadTextFile(strFile)
    lngPos = 1
    JsonSkipSpace strJson, lngPos
    If Mid$(strJson, lngPos, 1) <> "{" Then Exit Function

    On Error GoTo ErrHandler
    JsonParseValue strJson, lngPos, varResult
    If IsObject(varResult) Then Set ReadJsonFile = varResult
    Exit Function

ErrHandler:
    LogWarning "Unable to parse " & strFile & " (" & Err.Description & ")"
    Err.Clear

End Function


Private Sub JsonSkipSpace(strJson As String, lngPos As Long)
    Do While lngPos <= Len(strJson)
        Select Case Mid$(strJson, lngPos, 1)
            Case " ", vbTab, vbCr, vbLf
                lngPos = lngPos + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub


Private Sub JsonParseValue(strJson As String, lngPos As Long, varOut As Variant)

    JsonSkipSpace strJson, lngPos
    Select Case Mid$(strJson, lngPos, 1)
        Case "{"
            Set varOut = JsonParseObject(strJson, lngPos)
        Case "["
            Set varOut = JsonParseArray(strJson, lngPos)
        Case """"
            varOut = JsonParseString(strJson, lngPos)
        Case "t"
            varOut = True
            lngPos = lngPos + 4
        Case "f"
            varOut = False
            lngPos = lngPos + 5
        Case "n"
            varOut = Null
            lngPos = lngPos + 4
        Case Else
            varOut = JsonParseNumber(strJson, lngPos)
    End Select

End Sub


Private Function JsonParseObject(strJson As String, lngPos As Long) As Object

    Dim dResult As Object
    Dim strKey As String
    Dim varItem As Variant

    Set dResult = NewDict
    lngPos = lngPos + 1
    Do
        JsonSkipSpace strJson, lngPos
        Select Case Mid$(strJson, lngPos, 1)
            Case "}"
                lngPos = lngPos + 1
                Exit Do
            Case ","
                lngPos = lngPos + 1
            Case """"
                strKey = JsonParseString(strJson, lngPos)
                JsonSkipSpace strJson, lngPos
                If Mid$(strJson, lngPos, 1) <> ":" Then Err.Raise vbObjectError + 513, , "Expected ':' at position " & lngPos
                lngPos = lngPos + 1
                varItem = Empty
                JsonParseValue strJson, lngPos, varItem
                If dResult.Exists(strKey) Then dResult.Remove strKey
                dResult.Add strKey, varItem
            Case Else
                Err.Raise vbObjectError + 513, , "Invalid JSON at position " & lngPos
        End Select
    Loop
    Set JsonParseObject = dResult

End Function


Private Function JsonParseArray(strJson As String, lngPos As Long) As Collection

    Dim colResult As Collection
    Dim varItem As Variant

    Set colResult = New Collection
    lngPos = lngPos + 1
    Do
        JsonSkipSpace strJson, lngPos
        Select Case Mid$(strJson, lngPos, 1)
            Case "]"
                lngPos = lngPos + 1
                Exit Do
            Case ","
                lngPos = lngPos + 1
            Case ""
                Err.Raise vbObjectError + 513, , "Unexpected end of JSON"
            Case Else
                varItem = Empty
                JsonParseValue strJson, lngPos, varItem
                colResult.Add varItem
        End Select
    Loop
    Set JsonParseArray = colResult

End Function


Private Function JsonParseString(strJson As String, lngPos As Long) As String

    Dim strOut As String
    Dim lngQuote As Long
    Dim lngSlash As Long

    lngPos = lngPos + 1
    Do
        lngQuote = InStr(lngPos, strJson, """", vbBinaryCompare)
        If lngQuote = 0 Then Err.Raise vbObjectError + 513, , "Unterminated string"
        lngSlash = InStr(lngPos, strJson, "\", vbBinaryCompare)
        If lngSlash = 0 Or lngSlash > lngQuote Then
            strOut = strOut & Mid$(strJson, lngPos, lngQuote - lngPos)
            lngPos = lngQuote + 1
            Exit Do
        End If
        strOut = strOut & Mid$(strJson, lngPos, lngSlash - lngPos)
        Select Case Mid$(strJson, lngSlash + 1, 1)
            Case "b": strOut = strOut & Chr$(8)
            Case "f": strOut = strOut & Chr$(12)
            Case "n": strOut = strOut & vbLf
            Case "r": strOut = strOut & vbCr
            Case "t": strOut = strOut & vbTab
            Case "u"
                strOut = strOut & ChrW$(CLng("&H" & Mid$(strJson, lngSlash + 2, 4)))
                lngSlash = lngSlash + 4
            Case Else
                strOut = strOut & Mid$(strJson, lngSlash + 1, 1)
        End Select
        lngPos = lngSlash + 2
    Loop
    JsonParseString = strOut

End Function


Private Function JsonParseNumber(strJson As String, lngPos As Long) As Variant

    Dim lngStart As Long
    Dim strNumber As String

    lngStart = lngPos
    Do While lngPos <= Len(strJson)
        If InStr(1, "+-0123456789.eE", Mid$(strJson, lngPos, 1), vbBinaryCompare) = 0 Then Exit Do
        lngPos = lngPos + 1
    Loop
    strNumber = Mid$(strJson, lngStart, lngPos - lngStart)
    If Len(strNumber) = 0 Then Err.Raise vbObjectError + 513, , "Invalid JSON value at position " & lngStart

    If Len(strNumber) >= 16 Then
        ' Keep very long numbers as text (same as the add-in)
        JsonParseNumber = strNumber
    ElseIf InStr(1, strNumber, ".") > 0 Or InStr(1, strNumber, "e", vbTextCompare) > 0 Then
        JsonParseNumber = Val(strNumber)
    ElseIf Abs(Val(strNumber)) <= 2147483647# Then
        JsonParseNumber = CLng(Val(strNumber))
    Else
        JsonParseNumber = Val(strNumber)
    End If

End Function

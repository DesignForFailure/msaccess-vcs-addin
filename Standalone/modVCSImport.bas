Attribute VB_Name = "modVCSImport"
'---------------------------------------------------------------------------------------
' Module    : modVCSImport
' Purpose   : Standalone importer for source files created by the MSAccess VCS add-in
'           : (export format 5.0.0) or by the companion module modVCSExport. The
'           : add-in does NOT need to be installed.
'           :
'           : Usage (Immediate window):
'           :   BuildFromSource "C:\Work\MyApp.accdb.src"
'           :       Creates a brand new database from the source files (in a separate
'           :       Access window), like the add-in's "Build From Source". The database
'           :       is created next to the source folder, using the original file name.
'           :       An existing file with that name is renamed to *_VCSBackup first.
'           :   BuildFromSource "C:\Work\MyApp.accdb.src", "C:\Work\MyApp_v2.accdb"
'           :       Same, with an explicit file name.
'           :   MergeFromSource "C:\Work\MyApp.accdb.src"
'           :       Imports all source files into the CURRENT database, replacing
'           :       forms, reports, queries, macros and modules. Existing local tables
'           :       (and their data) are kept unless ReplaceExistingTables:=True.
'           :   ImportObject "C:\Work\MyApp.accdb.src\forms\frmMain.form"
'           :       Imports one object into the current database (any of its files can
'           :       be given: .form/.cls/.json, .sql, .bas, tbldefs\*.xml, ...).
'           :
'           : Notes:
'           :  - The modules modVCSExport and modVCSImport are never replaced while they
'           :    are running in the current database. Run MergeFromSource/ImportObject
'           :    from the Immediate window (not from code that the import may replace).
'           :  - Conditional formatting that the add-in decoded into JSON
'           :    ("ConditionalFormatting" in a form/report .json) cannot be rebuilt by
'           :    this module. Inline ConditionalFormat blocks are fine.
'           :  - Queries are created through DAO from the .sql file (SQL view). The
'           :    Design View layout is recreated by Access the next time you save the
'           :    query in Design View.
'           :  - Navigation pane groups, command bars and VBE UserForms are skipped.
'           :
'           : Uses late binding only. No additional VBA references are required.
'---------------------------------------------------------------------------------------
Option Compare Binary
Option Explicit

' Modules belonging to this tool (never replaced in the running database)
Private Const SKIP_MODULES As String = "|modVCSExport|modVCSImport|"

' Access constants that may not exist in older versions
Private Const AC_TABLE_DATA_MACRO As Long = 12

' DAO constants (late bound)
Private Const DB_BOOLEAN As Long = 1
Private Const DB_BYTE As Long = 2
Private Const DB_INTEGER As Long = 3
Private Const DB_LONG As Long = 4
Private Const DB_CURRENCY As Long = 5
Private Const DB_SINGLE As Long = 6
Private Const DB_DOUBLE As Long = 7
Private Const DB_DATE As Long = 8
Private Const DB_TEXT As Long = 10
Private Const DB_MEMO As Long = 12
Private Const DB_GUID As Long = 15
Private Const DB_ATTACHMENT As Long = 101
Private Const DB_OPEN_TABLE As Long = 1
Private Const DB_OPEN_DYNASET As Long = 2
Private Const DB_FAIL_ON_ERROR As Long = 128
Private Const DB_QSQL_PASS_THROUGH As Long = 112
Private Const DB_QSPT_BULK As Long = 144
Private Const DB_SAFE_LINK_ATTRIBUTES As Long = &H10000 Or &H20000 Or &H1 Or &H80000002

' VBIDE constants (late bound)
Private Const VBEXT_CT_CLASS_MODULE As Long = 2

' ADODB constants (late bound)
Private Const AD_TYPE_BINARY As Long = 1
Private Const AD_TYPE_TEXT As Long = 2
Private Const AD_SAVE_CREATE_OVERWRITE As Long = 2
Private Const AD_READ_ALL As Long = -1

' 1x1 transparent PNG (used to create the MSysResources table when it is missing)
Private Const PNG_1X1 As String = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="

#If VBA7 Then
    Private Declare PtrSafe Function GetACP Lib "kernel32" () As Long
#Else
    Private Declare Function GetACP Lib "kernel32" () As Long
#End If

' Module state for the current import operation
Private m_FSO As Object             ' Scripting.FileSystemObject
Private m_App As Object             ' Target Access.Application (this one, or a new instance)
Private m_Dbs As Object             ' DAO.Database of the target
Private m_VBProject As Object       ' VBIDE.VBProject of the target
Private m_Source As String          ' Source folder, with trailing backslash
Private m_InPlace As Boolean        ' Importing into the database that runs this code
Private m_FullBuild As Boolean      ' Building a new, empty database
Private m_ReplaceTables As Boolean  ' Replace existing local tables (merge)
Private m_SingleObject As Boolean   ' ImportObject (asks before replacing data)
Private m_Ucs2 As Variant           ' Cached: LoadFromText expects UTF-16 files
Private m_AnsiCharset As String
Private m_Env As Object             ' Values from the .env file
Private m_NewTables As Object       ' Local tables created during this operation
Private m_Forms As Collection       ' Forms imported during this operation
Private m_Deferred As Object        ' Startup properties applied at the end of a build
Private m_DeferredMacros As Collection
Private m_PendingRefs As Collection ' References added at the end of an in-place merge
Private m_PendingColumns As Collection  ' Query column settings (applied after all queries exist)
Private m_StagedRelations As Collection ' Relationships removed while replacing tables
Private m_Log As Collection
Private m_Errors As Long
Private m_Warnings As Long


'---------------------------------------------------------------------------------------
' Procedure : BuildFromSource
' Purpose   : Build a new database from source files in a separate Access instance.
'           : SourceFolder - The export folder. Prompts for a folder if omitted (or
'           :                uses <current database>.src if it exists).
'           : DbPath       - Full path of the database to create. Defaults to the
'           :                original file name, in the parent of the source folder.
'           : LeaveOpen    - Leave the new database open when finished.
'           : Returns the path of the new database, or an empty string on failure.
'---------------------------------------------------------------------------------------
'
Public Function BuildFromSource(Optional ByVal SourceFolder As String, Optional ByVal DbPath As String, _
    Optional ByVal LeaveOpen As Boolean = True) As String

    Dim sngStart As Single
    Dim strBackup As String
    Dim lngFormat As Long
    Dim blnSuccess As Boolean

    sngStart = Timer
    If Not BeginImport(SourceFolder) Then Exit Function

    ' Determine the database to create
    If Len(DbPath) = 0 Then DbPath = GetDefaultDbPath
    If Len(DbPath) = 0 Then
        MsgBox "Unable to determine the database file name (vbe-project.json not found)." & vbCrLf & _
            "Please supply the DbPath argument.", vbExclamation, "Build From Source"
        Exit Function
    End If
    DbPath = m_FSO.GetAbsolutePathName(DbPath)
    If StrComp(DbPath, CurrentProject.FullName, vbTextCompare) = 0 Then
        DbPath = m_FSO.GetParentFolderName(DbPath) & "\" & m_FSO.GetBaseName(DbPath) & " (Built)." & _
            m_FSO.GetExtensionName(DbPath)
        LogWarning "The source database is the database running this code, so it cannot be replaced. " & _
            "Building " & DbPath & " instead."
    End If

    LogLine "Beginning build from source"
    LogLine "Source folder: " & m_Source
    LogLine "Database: " & DbPath
    LogLine CStr(Now)

    ' Back up any existing file
    If m_FSO.FileExists(DbPath) Then
        strBackup = GetBackupFileName(DbPath)
        On Error Resume Next
        m_FSO.MoveFile DbPath, strBackup
        If Err.Number <> 0 Then
            MsgBox "Unable to rename the existing database (is it open?)" & vbCrLf & DbPath & vbCrLf & vbCrLf & _
                Err.Description, vbExclamation, "Build From Source"
            Exit Function
        End If
        On Error GoTo 0
        LogLine "Saved existing database as " & m_FSO.GetFileName(strBackup)
    End If

    ' Create the new database in a separate instance of Access
    lngFormat = GetFileFormat(DbPath)
    On Error GoTo ErrCreate
    Set m_App = CreateObject("Access.Application")
    m_App.Visible = True
    m_App.NewCurrentDatabase DbPath, lngFormat
    LogLine "Created new database (format " & lngFormat & ")"

    m_InPlace = False
    m_FullBuild = True
    m_ReplaceTables = True
    If Not ConnectTarget Then GoTo CleanUp

    On Error GoTo ErrBuild
    RemoveNonBuiltInReferences
    ImportAll
    blnSuccess = True

CleanUp:
    FinishImport "Build", sngStart, False
    On Error Resume Next
    If Not m_App Is Nothing Then
        If LeaveOpen And blnSuccess Then
            m_App.Visible = True
            m_App.UserControl = True
        Else
            m_App.CloseCurrentDatabase
            m_App.Quit
        End If
    End If
    Set m_VBProject = Nothing
    Set m_Dbs = Nothing
    Set m_App = Nothing
    On Error GoTo 0

    If blnSuccess Then BuildFromSource = DbPath
    MsgBox IIf(blnSuccess, "Build complete.", "Build failed.") & vbCrLf & vbCrLf & DbPath & vbCrLf & vbCrLf & _
        m_Errors & " error(s), " & m_Warnings & " warning(s)." & vbCrLf & vbCrLf & _
        "Open the database and run Debug > Compile in the VBA editor to check the code." & _
        IIf(m_Errors + m_Warnings > 0, vbCrLf & vbCrLf & "See the log in " & m_Source & "logs\", ""), _
        IIf(m_Errors > 0 Or Not blnSuccess, vbExclamation, vbInformation), "Build From Source"
    Exit Function

ErrCreate:
    LogError "Unable to create the database " & DbPath
    Resume CleanUp

ErrBuild:
    LogError "The build stopped because of an unexpected error"
    Resume CleanUp

End Function


'---------------------------------------------------------------------------------------
' Procedure : MergeFromSource
' Purpose   : Import all source files into the current database, replacing existing
'           : objects. Local tables that already exist are left alone (with their
'           : data) unless ReplaceExistingTables is True. Table data is only loaded
'           : into tables created by this operation.
'---------------------------------------------------------------------------------------
'
Public Function MergeFromSource(Optional ByVal SourceFolder As String, _
    Optional ByVal ReplaceExistingTables As Boolean = False) As Boolean

    Dim sngStart As Single

    sngStart = Timer
    If Not BeginImport(SourceFolder) Then Exit Function

    LogLine "Beginning merge from source"
    LogLine "Source folder: " & m_Source
    LogLine "Database: " & CurrentProject.FullName
    LogLine CStr(Now)

    Set m_App = Application
    m_InPlace = True
    m_FullBuild = False
    m_ReplaceTables = ReplaceExistingTables
    On Error GoTo ErrHandler
    If ConnectTarget Then
        CloseAllObjects
        ImportAll
        MergeFromSource = (m_Errors = 0)
    End If

Finish:
    On Error GoTo 0
    FinishImport "Merge", sngStart, True
    AddPendingReferences
    Exit Function

ErrHandler:
    LogError "The merge stopped because of an unexpected error"
    Resume Finish

End Function


'---------------------------------------------------------------------------------------
' Procedure : ImportObject
' Purpose   : Import a single source file (one object) into the current database.
'---------------------------------------------------------------------------------------
'
Public Function ImportObject(ByVal FilePath As String) As Boolean

    Dim sngStart As Single
    Dim strRoot As String

    sngStart = Timer
    Set m_FSO = CreateObject("Scripting.FileSystemObject")
    FilePath = m_FSO.GetAbsolutePathName(FilePath)
    If Not m_FSO.FileExists(FilePath) Then
        MsgBox "File not found:" & vbCrLf & FilePath, vbExclamation, "Import Object"
        Exit Function
    End If
    strRoot = FindSourceRoot(FilePath)
    If Len(strRoot) = 0 Then
        MsgBox "Unable to find the source folder for:" & vbCrLf & FilePath, vbExclamation, "Import Object"
        Exit Function
    End If
    If Not BeginImport(strRoot) Then Exit Function

    LogLine "Importing " & Mid$(FilePath, Len(m_Source) + 1) & " into " & CurrentProject.Name

    Set m_App = Application
    m_InPlace = True
    m_FullBuild = False
    m_ReplaceTables = True
    m_SingleObject = True
    On Error GoTo ErrHandler
    If ConnectTarget Then
        ImportSingleFile FilePath
        InitializeForms
        ImportObject = (m_Errors = 0)
    End If

Finish:
    On Error GoTo 0
    FinishImport "Import", sngStart, (m_Errors > 0)
    AddPendingReferences
    Exit Function

ErrHandler:
    LogError "The import stopped because of an unexpected error"
    Resume Finish

End Function


'---------------------------------------------------------------------------------------
' Procedure : ImportAll
' Purpose   : Import every component type, in the same order as the add-in.
'---------------------------------------------------------------------------------------
'
Private Sub ImportAll()

    Dim blnThemes As Boolean

    ImportProject
    ImportVbeProject
    ImportVbeReferences
    ImportProjectProperties
    ImportSavedSpecs
    ' Modules replaced in the running database are imported last (see below)
    If Not m_InPlace Then ImportAllModules
    ImportSharedImages
    blnThemes = ImportThemes
    ImportDbProperties
    ImportImexSpecs
    ImportAllTableDefs
    ImportAllQueries
    ImportAllObjects "forms", acForm
    ImportAllObjects "macros", acMacro
    ImportAllObjects "reports", acReport
    ImportAllTableData
    ImportAllTableDataMacros
    ImportAllRelations
    RestoreStagedRelations
    ImportDocuments
    ImportLegacyHiddenAttributes
    SkipUnsupported

    ' Reopen a new database so imported themes are applied to the forms. (Startup
    ' properties and the AutoExec macro are applied afterwards, so nothing runs.)
    If m_FullBuild And blnThemes Then ReopenDatabase
    InitializeForms
    If m_FullBuild Then ApplyDeferredItems

    ' Replacing modules of the VBA project that is running this code is done last,
    ' in case it resets the running project.
    If m_InPlace Then ImportAllModules

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportSingleFile
' Purpose   : Route a source file to the correct importer.
'---------------------------------------------------------------------------------------
'
Private Sub ImportSingleFile(ByVal strFile As String)

    Dim strRelative As String
    Dim strCategory As String
    Dim strExt As String
    Dim strMain As String
    Dim dItems As Object

    strRelative = Mid$(strFile, Len(m_Source) + 1)
    strExt = LCase$(m_FSO.GetExtensionName(strFile))
    If InStr(1, strRelative, "\") > 0 Then strCategory = LCase$(Left$(strRelative, InStr(1, strRelative, "\") - 1))

    Select Case strCategory
        Case "forms", "reports"
            strMain = SwapExtension(strFile, IIf(strCategory = "forms", "form", "report"))
            If Not m_FSO.FileExists(strMain) Then strMain = SwapExtension(strFile, "bas")
            If m_FSO.FileExists(strMain) Then
                ImportFormOrReport IIf(strCategory = "forms", acForm, acReport), strMain
            Else
                LogError "Object definition file not found for " & strRelative
            End If
        Case "macros"
            strMain = SwapExtension(strFile, "macro")
            If Not m_FSO.FileExists(strMain) Then strMain = SwapExtension(strFile, "bas")
            If m_FSO.FileExists(strMain) Then
                ImportMacro strMain
            Else
                LogError "Macro definition file not found for " & strRelative
            End If
        Case "modules"
            If strExt = "json" Then
                strMain = SwapExtension(strFile, "bas")
                If Not m_FSO.FileExists(strMain) Then strMain = SwapExtension(strFile, "cls")
            Else
                strMain = strFile
            End If
            If m_FSO.FileExists(strMain) Then
                ImportModule strMain
            Else
                LogError "Module file (.bas/.cls) not found for " & strRelative
            End If
        Case "queries"
            strMain = strFile
            If strExt = "json" Then strMain = SwapExtension(strFile, "sql")
            If m_FSO.FileExists(strMain) Then
                ReportQueryResult GetObjectNameFromFileName(strMain), ImportQuery(strMain)
                ApplyPendingColumns
            Else
                LogError "Query file (.sql) not found for " & strRelative
            End If
        Case "tbldefs"
            If strExt = "json" Then
                Set dItems = ReadItems(strFile)
                If dItems.Exists("Connect") Then
                    ImportLinkedTable strFile
                ElseIf m_FSO.FileExists(SwapExtension(strFile, "xml")) Then
                    ImportLocalTable SwapExtension(strFile, "xml"), True
                Else
                    ImportObjectMetadata dItems, "Tables", GetObjectNameFromFileName(strFile), acTable
                End If
            Else
                ImportLocalTable SwapExtension(strFile, "xml"), True
            End If
        Case "tables"
            ImportTableData strFile, True
        Case "relations"
            ImportRelation strFile
        Case "tdmacros"
            ImportTableDataMacro strFile
        Case "themes"
            If ImportThemes(strFile) Then LogLine "Close and reopen the database to apply the theme."
        Case "images"
            If strExt = "json" Then
                ImportSharedImage strFile
            Else
                LogWarning "Select the .json file of the shared image to import it."
            End If
        Case "imexspecs"
            ImportImexSpec strFile
        Case "savedspecs"
            ImportSavedSpec strFile
        Case Else
            Select Case LCase$(m_FSO.GetFileName(strFile))
                Case "dbs-properties.json": ImportDbProperties
                Case "proj-properties.json": ImportProjectProperties
                Case "vbe-references.json": ImportVbeReferences
                Case "vbe-project.json": ImportVbeProject
                Case "project.json": ImportProject
                Case "documents.json": ImportDocuments
                Case "hidden-attributes.json": ImportLegacyHiddenAttributes
                Case Else
                    LogWarning "Unsupported source file: " & strRelative
            End Select
    End Select

End Sub


'=======================================================================================
' Setup and teardown
'=======================================================================================

'---------------------------------------------------------------------------------------
' Procedure : BeginImport
' Purpose   : Initialize module state and resolve the source folder.
'---------------------------------------------------------------------------------------
'
Private Function BeginImport(ByVal strFolder As String) As Boolean

    Set m_FSO = CreateObject("Scripting.FileSystemObject")
    Set m_Log = New Collection
    Set m_NewTables = NewDict
    Set m_Forms = New Collection
    Set m_Deferred = NewDict
    Set m_DeferredMacros = New Collection
    Set m_PendingRefs = New Collection
    Set m_PendingColumns = New Collection
    Set m_StagedRelations = New Collection
    m_SingleObject = False
    Set m_Env = Nothing
    Set m_App = Nothing
    Set m_Dbs = Nothing
    Set m_VBProject = Nothing
    m_Errors = 0
    m_Warnings = 0
    m_Ucs2 = Empty
    m_AnsiCharset = GetAnsiCharset

    ' Resolve the source folder
    If Len(strFolder) = 0 Then
        If m_FSO.FileExists(CurrentProject.FullName & ".src\vcs-options.json") Then
            strFolder = CurrentProject.FullName & ".src"
        Else
            strFolder = PickFolder
        End If
    End If
    If Len(strFolder) = 0 Then Exit Function
    strFolder = m_FSO.GetAbsolutePathName(strFolder)
    If Right$(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
    If Not m_FSO.FolderExists(strFolder) Then
        MsgBox "Source folder not found:" & vbCrLf & strFolder, vbExclamation, "Import Source"
        Exit Function
    End If
    If Not m_FSO.FileExists(strFolder & "vcs-options.json") Then
        If Not (m_FSO.FolderExists(strFolder & "modules") Or m_FSO.FolderExists(strFolder & "forms") _
            Or m_FSO.FolderExists(strFolder & "queries") Or m_FSO.FolderExists(strFolder & "tbldefs")) Then
            MsgBox "This folder does not look like an exported source folder (vcs-options.json not found):" & _
                vbCrLf & strFolder, vbExclamation, "Import Source"
            Exit Function
        End If
    End If
    m_Source = strFolder
    BeginImport = True

End Function


Private Function ConnectTarget() As Boolean

    On Error GoTo ErrHandler
    Set m_Dbs = m_App.CurrentDb
    Set m_VBProject = FindVBProject
    m_Ucs2 = Empty
    If m_VBProject Is Nothing Then LogWarning "Unable to access the VBA project. Modules will not be imported."
    ConnectTarget = True
    Exit Function

ErrHandler:
    LogError "Unable to open the target database"

End Function


Private Function FindVBProject() As Object

    Dim prj As Object
    Dim strFile As String
    Dim strTarget As String

    On Error Resume Next
    strTarget = m_App.CurrentProject.FullName
    For Each prj In m_App.VBE.VBProjects
        strFile = vbNullString
        strFile = prj.FileName
        If StrComp(strFile, strTarget, vbTextCompare) = 0 Then
            Set FindVBProject = prj
            Exit For
        End If
    Next prj
    If FindVBProject Is Nothing And Not m_InPlace Then Set FindVBProject = m_App.VBE.ActiveVBProject
    If FindVBProject Is Nothing And m_InPlace Then Set FindVBProject = Application.VBE.ActiveVBProject
    Err.Clear

End Function


Private Sub FinishImport(ByVal strOperation As String, ByVal sngStart As Single, ByVal blnShowMessage As Boolean)

    Dim strSummary As String
    Dim strText As String
    Dim strLogFile As String
    Dim varLine As Variant

    strSummary = "Done. (" & Format$(Timer - sngStart, "0.0") & " seconds) " & _
        m_Errors & " error(s), " & m_Warnings & " warning(s)."
    LogLine strSummary
    SysCmd acSysCmdClearStatus

    On Error Resume Next
    For Each varLine In m_Log
        strText = strText & varLine & vbCrLf
    Next varLine
    strLogFile = m_Source & "logs\" & strOperation & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".log"
    WriteTextFile strLogFile, strText
    Err.Clear
    On Error GoTo 0

    If blnShowMessage Then
        MsgBox strOperation & " complete." & vbCrLf & vbCrLf & strSummary & _
            IIf(m_Errors + m_Warnings > 0, vbCrLf & vbCrLf & "See the log file for details:" & vbCrLf & strLogFile, ""), _
            IIf(m_Errors > 0, vbExclamation, vbInformation), strOperation & " Source"
    End If

End Sub


Private Function PickFolder() As String
    On Error Resume Next
    With Application.FileDialog(4)      ' msoFileDialogFolderPicker
        .Title = "Select the exported source folder (contains vcs-options.json)"
        .AllowMultiSelect = False
        If .Show Then PickFolder = .SelectedItems(1)
    End With
    Err.Clear
End Function


Private Function FindSourceRoot(ByVal strFile As String) As String

    Dim strFolder As String
    Dim lngLevel As Long

    strFolder = m_FSO.GetParentFolderName(strFile)
    For lngLevel = 1 To 8
        If Len(strFolder) = 0 Then Exit For
        If m_FSO.FileExists(strFolder & "\vcs-options.json") Then
            FindSourceRoot = strFolder
            Exit Function
        End If
        strFolder = m_FSO.GetParentFolderName(strFolder)
    Next lngLevel

    ' No options file. Use the parent of the component folder.
    strFolder = m_FSO.GetParentFolderName(strFile)
    Do While Len(strFolder) > 0
        Select Case LCase$(m_FSO.GetFileName(strFolder))
            Case "forms", "reports", "macros", "modules", "queries", "tbldefs", "tables", "relations", _
                 "tdmacros", "themes", "images", "imexspecs", "savedspecs"
                FindSourceRoot = m_FSO.GetParentFolderName(strFolder)
                Exit Function
        End Select
        strFolder = m_FSO.GetParentFolderName(strFolder)
    Loop

End Function


Private Function GetDefaultDbPath() As String

    Dim dItems As Object
    Dim strFile As String

    Set dItems = ReadItems(m_Source & "vbe-project.json")
    If dItems.Exists("FileName") Then strFile = CStr(dItems("FileName"))
    If Left$(strFile, 4) = "rel:" Then strFile = Mid$(strFile, 5)
    If Len(strFile) = 0 Then Exit Function
    GetDefaultDbPath = m_FSO.GetParentFolderName(Left$(m_Source, Len(m_Source) - 1)) & "\" & strFile

End Function


Private Function GetBackupFileName(ByVal strPath As String) As String

    Dim lngCnt As Long
    Dim strTest As String
    Dim strBase As String

    strBase = m_FSO.GetParentFolderName(strPath) & "\" & m_FSO.GetBaseName(strPath) & "_VCSBackup"
    For lngCnt = 0 To 500
        strTest = strBase & IIf(lngCnt = 0, vbNullString, CStr(lngCnt)) & "." & m_FSO.GetExtensionName(strPath)
        If Not m_FSO.FileExists(strTest) Then
            GetBackupFileName = strTest
            Exit Function
        End If
    Next lngCnt

End Function


Private Function GetFileFormat(ByVal strPath As String) As Long

    Dim dItems As Object
    Dim lngFormat As Long

    Set dItems = ReadItems(m_Source & "project.json")
    If dItems.Exists("FileFormat") Then lngFormat = CLng(dItems("FileFormat"))

    ' Make sure the format matches the file extension
    If LCase$(m_FSO.GetExtensionName(strPath)) = "mdb" Then
        If lngFormat <> 9 And lngFormat <> 10 Then lngFormat = 10
    Else
        If lngFormat < 12 Then lngFormat = 12
    End If
    GetFileFormat = lngFormat

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReopenDatabase
' Purpose   : Close and reopen the new database so that imported themes are loaded.
'---------------------------------------------------------------------------------------
'
Private Sub ReopenDatabase()

    Dim strPath As String

    On Error GoTo ErrHandler
    LogLine "Reopening database to load themes..."
    strPath = m_App.CurrentProject.FullName
    Set m_VBProject = Nothing
    Set m_Dbs = Nothing
    m_App.CloseCurrentDatabase
    m_App.OpenCurrentDatabase strPath
    ConnectTarget
    Exit Sub

ErrHandler:
    LogError "Error reopening the database"

End Sub


Private Sub CloseAllObjects()

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
    Err.Clear

End Sub


Private Sub CloseObjectIfOpen(ByVal intType As Long, ByVal strName As String)
    On Error Resume Next
    If m_App.SysCmd(acSysCmdGetObjectState, intType, strName) <> 0 Then
        m_App.DoCmd.Close intType, strName, acSaveNo
    End If
    Err.Clear
End Sub


Private Sub SkipUnsupported()
    Dim varItem As Variant
    For Each varItem In Array("menus", "vbeforms")
        If m_FSO.FolderExists(m_Source & varItem) Then
            LogWarning "The " & varItem & " folder is not supported by this importer and was skipped."
        End If
    Next varItem
    If m_FSO.FileExists(m_Source & "nav-pane-groups.json") Then
        LogLine "Note: navigation pane groups (nav-pane-groups.json) are not imported."
    End If
End Sub


'=======================================================================================
' Project, VBA project and references
'=======================================================================================

Private Sub ImportProject()

    Dim dItems As Object

    On Error GoTo ErrHandler
    Set dItems = ReadItems(m_Source & "project.json")
    If dItems.Exists("RemovePersonalInformation") Then
        m_App.CurrentProject.RemovePersonalInformation = CBool(dItems("RemovePersonalInformation"))
    End If
    Exit Sub

ErrHandler:
    LogError "Error importing project.json"

End Sub


Private Sub ImportVbeProject()

    Dim dItems As Object

    On Error GoTo ErrHandler
    Set dItems = ReadItems(m_Source & "vbe-project.json")
    If dItems.Count = 0 Then Exit Sub

    If Not m_VBProject Is Nothing Then
        ' Do not rename the project that is running this code
        If dItems.Exists("Name") And Not m_InPlace Then
            If Len(dItems("Name")) > 0 And m_VBProject.Name <> dItems("Name") Then m_VBProject.Name = dItems("Name")
        End If
        If dItems.Exists("Description") Then m_VBProject.Description = Nz(dItems("Description"), vbNullString)
        On Error Resume Next
        If dItems.Exists("HelpFile") Then
            If Len(Nz(dItems("HelpFile"), vbNullString)) > 0 Then m_VBProject.HelpFile = dItems("HelpFile")
        End If
        If dItems.Exists("HelpContextId") Then m_VBProject.HelpContextID = CLng(dItems("HelpContextId"))
        Err.Clear
        On Error GoTo ErrHandler
    End If
    If dItems.Exists("ConditionalCompilationArguments") Then
        m_App.SetOption "Conditional Compilation Arguments", Nz(dItems("ConditionalCompilationArguments"), vbNullString)
    End If
    Exit Sub

ErrHandler:
    LogError "Error importing vbe-project.json"

End Sub


Private Sub RemoveNonBuiltInReferences()

    Dim colRefs As Collection
    Dim ref As Object

    On Error Resume Next
    Set colRefs = New Collection
    For Each ref In m_App.References
        If Not ref.BuiltIn Then colRefs.Add ref
    Next ref
    For Each ref In colRefs
        m_App.References.Remove ref
    Next ref
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportVbeReferences
' Purpose   : Add VBA references in the saved priority order. When importing into the
'           : running database, missing references are added at the very end (adding a
'           : reference can reset the state of running code).
'---------------------------------------------------------------------------------------
'
Private Sub ImportVbeReferences()

    Dim dItems As Object
    Dim varKey As Variant

    On Error GoTo ErrHandler
    Set dItems = ReadItems(m_Source & "vbe-references.json")
    For Each varKey In dItems.Keys
        If Not ReferenceExists(CStr(varKey)) Then
            If m_InPlace Then
                LogLine "  Reference to add: " & varKey
                m_PendingRefs.Add Array(CStr(varKey), dItems(varKey))
            Else
                AddReference CStr(varKey), dItems(varKey)
            End If
        End If
    Next varKey
    Exit Sub

ErrHandler:
    LogError "Error importing vbe-references.json"

End Sub


Private Function ReferenceExists(ByVal strName As String) As Boolean
    Dim ref As Object
    Dim strRef As String
    On Error Resume Next
    For Each ref In m_App.References
        strRef = vbNullString
        strRef = ref.Name
        If StrComp(strRef, strName, vbTextCompare) = 0 Then
            ReferenceExists = True
            Exit For
        End If
    Next ref
    Err.Clear
End Function


Private Function AddReference(ByVal strName As String, ByVal dRef As Object) As Boolean

    Dim varVersion As Variant
    Dim strPath As String

    On Error Resume Next
    If dRef.Exists("GUID") Then
        varVersion = Split(Nz(dRef("Version"), "0.0") & ".0", ".")
        m_App.References.AddFromGuid CStr(dRef("GUID")), CLng(Val(varVersion(0))), CLng(Val(varVersion(1)))
        If Err.Number <> 0 Then
            ' Try the latest installed version
            Err.Clear
            m_App.References.AddFromGuid CStr(dRef("GUID")), 0, 0
        End If
    ElseIf dRef.Exists("FullPath") Then
        strPath = ExpandRelativePath(CStr(dRef("FullPath")))
        If m_FSO.FileExists(strPath) Then
            m_App.References.AddFromFile strPath
        Else
            Err.Raise 53, , "File not found: " & strPath
        End If
    End If

    If Err.Number <> 0 Then
        LogWarning "Unable to add reference " & strName & " (" & Err.Description & ")"
        Err.Clear
    Else
        AddReference = True
    End If

End Function


' Add references queued during an in-place import. This runs last because adding a
' reference may reset the running VBA project.
Private Sub AddPendingReferences()

    Dim varItem As Variant
    Dim colPending As Collection

    If m_PendingRefs Is Nothing Then Exit Sub
    If m_PendingRefs.Count = 0 Then Exit Sub
    Set colPending = m_PendingRefs
    Set m_PendingRefs = New Collection
    For Each varItem In colPending
        If AddReference(CStr(varItem(0)), varItem(1)) Then Debug.Print "  Added reference " & varItem(0)
    Next varItem

End Sub


Private Sub ImportProjectProperties()

    Dim dItems As Object
    Dim varKey As Variant
    Dim varValue As Variant

    On Error GoTo ErrHandler
    Set dItems = ReadItems(m_Source & "proj-properties.json")
    For Each varKey In dItems.Keys
        Select Case varKey
            Case "Name", "Connection"
                ' Skip these
            Case Else
                varValue = dItems(varKey)
                If VarType(varValue) = vbString Then
                    If Left$(varValue, 4) = "rel:" Then varValue = ExpandRelativePath(CStr(varValue))
                End If
                On Error Resume Next
                m_App.CurrentProject.Properties(CStr(varKey)).Value = varValue
                If Err.Number <> 0 Then
                    Err.Clear
                    m_App.CurrentProject.Properties.Add CStr(varKey), varValue
                End If
                If Err.Number <> 0 Then LogWarning "Unable to set project property " & varKey & " (" & Err.Description & ")"
                Err.Clear
                On Error GoTo ErrHandler
        End Select
    Next varKey
    Exit Sub

ErrHandler:
    LogError "Error importing proj-properties.json"

End Sub


'=======================================================================================
' Database properties and documents
'=======================================================================================

'---------------------------------------------------------------------------------------
' Procedure : ImportDbProperties
' Purpose   : Restore DAO database properties. Engine-managed properties are skipped.
'           : During a build, startup properties are applied at the end so that no
'           : startup form or ribbon loads while the build is in progress.
'---------------------------------------------------------------------------------------
'
Private Sub ImportDbProperties()

    Dim dItems As Object
    Dim dProp As Object
    Dim varKey As Variant
    Dim varValue As Variant
    Dim lngType As Long

    On Error GoTo ErrHandler
    Set dItems = ReadItems(m_Source & "dbs-properties.json")
    If dItems.Count = 0 Then Exit Sub
    LogLine "Importing database properties..."

    For Each varKey In dItems.Keys
        Select Case varKey
            Case "Connection", "Name", "Version", "CollatingOrder", "Build", "AccessVersion", "ProjVer", _
                 "Updatable", "Transactions", "RecordsAffected", "Connect", "DesignMasterID", "ReplicaID", _
                 "HasOfflineLists", "Last VCS Export", "Last VCS Version"
                ' Managed by the database engine or by Access
            Case Else
                Set dProp = GetDict(dItems, CStr(varKey))
                If Not dProp Is Nothing Then
                    lngType = CLng(Val(GetText(dProp, "Type", CStr(DB_TEXT))))
                    varValue = ConvertJsonValue(GetValue(dProp, "Value"), lngType)
                    If m_FullBuild And IsStartupProperty(CStr(varKey)) Then
                        m_Deferred(CStr(varKey)) = Array(lngType, varValue)
                    Else
                        SetDaoProperty m_Dbs, CStr(varKey), lngType, varValue, "database property"
                    End If
                End If
        End Select
    Next varKey
    Exit Sub

ErrHandler:
    LogError "Error importing dbs-properties.json"

End Sub


Private Function IsStartupProperty(ByVal strName As String) As Boolean
    Select Case strName
        Case "StartUpForm", "CustomRibbonID", "StartUpMenuBar", "StartUpShortcutMenuBar"
            IsStartupProperty = True
    End Select
End Function


' Convert a parsed JSON value to the value expected for a DAO property.
Private Function ConvertJsonValue(ByVal varValue As Variant, ByVal lngType As Long) As Variant

    Dim bteData() As Byte
    Dim lngCnt As Long

    If IsObject(varValue) Then
        ' Byte array saved as a list of numbers
        If varValue.Count > 0 Then
            ReDim bteData(0 To varValue.Count - 1)
            For lngCnt = 1 To varValue.Count
                bteData(lngCnt - 1) = CByte(varValue(lngCnt))
            Next lngCnt
            ConvertJsonValue = bteData
        End If
    ElseIf VarType(varValue) = vbString Then
        If Left$(varValue, 4) = "rel:" Then
            ConvertJsonValue = ExpandRelativePath(CStr(varValue))
        ElseIf lngType = DB_DATE And Right$(varValue, 1) = "Z" And Not IsDate(varValue) Then
            ConvertJsonValue = FromIsoUtc(CStr(varValue))
        Else
            ConvertJsonValue = varValue
        End If
    Else
        ConvertJsonValue = varValue
    End If

End Function


Private Sub ImportDocuments()

    Dim dItems As Object
    Dim dDocs As Object
    Dim dProps As Object
    Dim dProp As Object
    Dim doc As Object
    Dim varDoc As Variant
    Dim varProp As Variant
    Dim lngType As Long

    On Error GoTo ErrHandler
    Set dItems = ReadItems(m_Source & "documents.json")
    Set dDocs = GetDict(dItems, "Databases")
    If dDocs Is Nothing Then Exit Sub

    m_Dbs.Containers("Databases").Documents.Refresh
    For Each varDoc In dDocs.Keys
        Set doc = Nothing
        On Error Resume Next
        Set doc = m_Dbs.Containers("Databases").Documents(CStr(varDoc))
        Err.Clear
        On Error GoTo ErrHandler
        If doc Is Nothing Then
            LogWarning "Database document " & varDoc & " not found."
        Else
            Set dProps = GetDict(dDocs, CStr(varDoc))
            For Each varProp In dProps.Keys
                Set dProp = GetDict(dProps, CStr(varProp))
                If dProp Is Nothing Then
                    SetDaoProperty doc, CStr(varProp), DB_TEXT, dProps(varProp), "document property"
                Else
                    lngType = CLng(Val(GetText(dProp, "Type", CStr(DB_TEXT))))
                    SetDaoProperty doc, CStr(varProp), lngType, ConvertJsonValue(GetValue(dProp, "Value"), lngType), "document property"
                End If
            Next varProp
        End If
    Next varDoc
    Exit Sub

ErrHandler:
    LogError "Error importing documents.json"

End Sub


' Hidden flags from older exports (hidden-attributes.json)
Private Sub ImportLegacyHiddenAttributes()

    Dim dItems As Object
    Dim varKey As Variant
    Dim varName As Variant
    Dim intType As Long

    On Error Resume Next
    Set dItems = ReadItems(m_Source & "hidden-attributes.json")
    For Each varKey In dItems.Keys
        Select Case varKey
            Case "Tables": intType = acTable
            Case "Queries": intType = acQuery
            Case "Forms": intType = acForm
            Case "Reports": intType = acReport
            Case "Scripts": intType = acMacro
            Case "Modules": intType = acModule
            Case Else: intType = -1
        End Select
        If intType <> -1 And IsObject(dItems(varKey)) Then
            For Each varName In dItems(varKey)
                m_App.SetHiddenAttribute intType, CStr(varName), True
                Err.Clear
            Next varName
        End If
    Next varKey
    Err.Clear

End Sub


Private Sub ApplyDeferredItems()

    Dim varKey As Variant
    Dim varItem As Variant

    For Each varItem In m_DeferredMacros
        ImportMacro CStr(varItem)
    Next varItem
    For Each varKey In m_Deferred.Keys
        varItem = m_Deferred(varKey)
        SetDaoProperty m_Dbs, CStr(varKey), CLng(varItem(0)), varItem(1), "database property"
    Next varKey

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportObjectMetadata
' Purpose   : Apply the "Properties" (Description) and "Hidden" keys of a companion
'           : json file to a database object.
'---------------------------------------------------------------------------------------
'
Private Sub ImportObjectMetadata(dItems As Object, ByVal strContainer As String, ByVal strName As String, ByVal intType As Long)

    Dim dProps As Object
    Dim dProp As Object
    Dim doc As Object
    Dim varKey As Variant
    Dim lngType As Long

    If dItems Is Nothing Then Exit Sub
    On Error GoTo ErrHandler

    Set dProps = GetDict(dItems, "Properties")
    If Not dProps Is Nothing Then
        m_Dbs.Containers(strContainer).Documents.Refresh
        Set doc = m_Dbs.Containers(strContainer).Documents(strName)
        For Each varKey In dProps.Keys
            Set dProp = GetDict(dProps, CStr(varKey))
            If dProp Is Nothing Then
                SetDaoProperty doc, CStr(varKey), DB_TEXT, dProps(varKey), strName & " property"
            Else
                lngType = CLng(Val(GetText(dProp, "Type", CStr(DB_TEXT))))
                SetDaoProperty doc, CStr(varKey), lngType, GetValue(dProp, "Value"), strName & " property"
            End If
        Next varKey
    End If

    If dItems.Exists("Hidden") Then
        If dItems("Hidden") = True Then m_App.SetHiddenAttribute intType, strName, True
    End If
    Exit Sub

ErrHandler:
    LogWarning "Unable to set description/hidden attribute of " & strName & " (" & Err.Description & ")"
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetDaoProperty
' Purpose   : Set a DAO property, creating it (or recreating it with the correct type)
'           : when needed. Returns True on success.
'---------------------------------------------------------------------------------------
'
Private Function SetDaoProperty(objParent As Object, ByVal strName As String, ByVal lngType As Long, _
    ByVal varValue As Variant, ByVal strLabel As String) As Boolean

    Dim prp As Object
    Dim blnExists As Boolean

    If IsNull(varValue) Or IsEmpty(varValue) Then Exit Function

    On Error Resume Next
    Set prp = objParent.Properties(strName)
    blnExists = (Err.Number = 0)
    Err.Clear

    If blnExists And lngType <> 0 Then
        If prp.Type <> lngType Then
            ' Recreate user-defined properties that have a different type
            objParent.Properties.Delete strName
            If Err.Number = 0 Then blnExists = False
            Err.Clear
        End If
    End If

    If blnExists Then
        If IsArray(varValue) Then
            prp.Value = varValue
        ElseIf CStr(Nz(prp.Value, vbNullString)) <> CStr(varValue) Then
            prp.Value = varValue
        End If
    Else
        If lngType = DB_TEXT And VarType(varValue) = vbString Then
            If Len(varValue) = 0 Then Exit Function
        End If
        If lngType = 0 Then lngType = DB_TEXT
        objParent.Properties.Append objParent.CreateProperty(strName, lngType, varValue)
    End If

    If Err.Number <> 0 Then
        LogWarning "Unable to set " & strLabel & " " & strName & " (" & Err.Description & ")"
        Err.Clear
    Else
        SetDaoProperty = True
    End If

End Function


'=======================================================================================
' VBA modules
'=======================================================================================

Private Sub ImportAllModules()

    Dim dFiles As Object
    Dim varKey As Variant

    Set dFiles = GetSourceFiles(m_Source & "modules", "bas|cls", True)
    If dFiles.Count = 0 Then Exit Sub
    LogLine "Importing modules..."
    For Each varKey In dFiles.Keys
        ImportModule CStr(dFiles(varKey))
    Next varKey

End Sub


Private Function IsSkippedModule(ByVal strName As String) As Boolean
    IsSkippedModule = (InStr(1, SKIP_MODULES, "|" & strName & "|", vbTextCompare) > 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ImportModule
' Purpose   : Import a .bas or .cls file through the VBE (which keeps the attributes,
'           : such as VB_PredeclaredId). The module header is rebuilt so that the module
'           : name always matches the file name.
'---------------------------------------------------------------------------------------
'
Private Sub ImportModule(ByVal strFile As String)

    Dim strName As String
    Dim strCode As String
    Dim strTemp As String
    Dim varLines As Variant
    Dim lngLine As Long
    Dim lngFirst As Long
    Dim blnClass As Boolean
    Dim blnClassAttributes As Boolean
    Dim blnCreatable As Boolean
    Dim blnExposed As Boolean
    Dim strHeader As String
    Dim strBody() As String
    Dim blnRemove() As Boolean
    Dim lngCnt As Long
    Dim cmp As Object

    strName = GetObjectNameFromFileName(strFile)
    If m_InPlace And IsSkippedModule(strName) Then
        LogLine "  Skipped " & strName & " (part of this tool)"
        Exit Sub
    End If
    If m_VBProject Is Nothing Then
        LogError "Cannot import module " & strName & " (VBA project not available)"
        Exit Sub
    End If
    Status "Importing module " & strName

    On Error GoTo ErrHandler
    strCode = ReadTextFile(strFile)
    varLines = Split(strCode, vbCrLf)
    ReDim blnRemove(0 To UBound(varLines) + 1)
    blnClass = (LCase$(m_FSO.GetExtensionName(strFile)) = "cls")

    ' Skip any existing header block and VB_Name attribute
    lngFirst = 0
    If UBound(varLines) >= 0 Then
        If Trim$(varLines(0)) = "VERSION 1.0 CLASS" Then
            blnClass = True
            For lngLine = 1 To UBound(varLines)
                If Trim$(varLines(lngLine)) = "END" Then
                    lngFirst = lngLine + 1
                    Exit For
                End If
            Next lngLine
        End If
    End If
    For lngLine = lngFirst To UBound(varLines)
        If lngLine > lngFirst + 12 Then Exit For
        If Left$(varLines(lngLine), 20) = "Attribute VB_Name = " Then
            blnRemove(lngLine) = True           ' Replaced by the rebuilt header
        ElseIf Left$(varLines(lngLine), 31) = "Attribute VB_GlobalNameSpace = " Then
            blnClass = True
            blnClassAttributes = True
        ElseIf varLines(lngLine) = "Attribute VB_Creatable = True" Then
            blnCreatable = True
        ElseIf varLines(lngLine) = "Attribute VB_Exposed = True" Then
            blnExposed = True
        End If
    Next lngLine

    ' Rebuild the header
    If blnClass Then
        strHeader = "VERSION 1.0 CLASS" & vbCrLf & "BEGIN" & vbCrLf & "  MultiUse = -1  'True" & vbCrLf & "END" & vbCrLf
    End If
    strHeader = strHeader & "Attribute VB_Name = """ & strName & """" & vbCrLf
    If blnClass And Not blnClassAttributes Then
        strHeader = strHeader & "Attribute VB_GlobalNameSpace = False" & vbCrLf & "Attribute VB_Creatable = False" & vbCrLf & _
            "Attribute VB_PredeclaredId = False" & vbCrLf & "Attribute VB_Exposed = False" & vbCrLf
    End If
    If lngFirst <= UBound(varLines) Then
        ReDim strBody(0 To UBound(varLines) - lngFirst)
        lngCnt = 0
        For lngLine = lngFirst To UBound(varLines)
            If Not blnRemove(lngLine) Then
                strBody(lngCnt) = varLines(lngLine)
                lngCnt = lngCnt + 1
            End If
        Next lngLine
        If lngCnt > 0 Then
            ReDim Preserve strBody(0 To lngCnt - 1)
            strCode = strHeader & Join(strBody, vbCrLf)
        Else
            strCode = strHeader
        End If
    Else
        strCode = strHeader
    End If

    ' The VBE imports modules in the system (ANSI) code page.
    strTemp = GetTempFile(IIf(blnClass, ".cls", ".bas"))
    WriteTextFile strTemp, strCode, m_AnsiCharset
    If StrComp(ReadTextFile(strTemp, m_AnsiCharset), NormalizeText(strCode), vbBinaryCompare) <> 0 Then
        LogWarning strName & " contains characters that are not supported by the VBA editor in the " & _
            m_AnsiCharset & " code page. They were replaced."
    End If

    RemoveModule strName
    Set cmp = m_VBProject.VBComponents.Import(strTemp)
    If cmp.Name <> strName Then cmp.Name = strName
    DeleteFile strTemp
    SaveModule strName

    ' Public creatable classes need the Instancing property set explicitly
    If blnClass And blnCreatable And blnExposed And cmp.Type = VBEXT_CT_CLASS_MODULE Then
        cmp.Properties("Instancing").Value = 5
        SaveModule strName
    End If

    ImportObjectMetadata ReadItems(SwapExtension(strFile, "json")), "Modules", strName, acModule
    Exit Sub

ErrHandler:
    LogError "Error importing module " & strName
    DeleteFile strTemp

End Sub


Private Sub RemoveModule(ByVal strName As String)

    Dim cmp As Object

    On Error Resume Next
    Set cmp = m_VBProject.VBComponents(strName)
    If Err.Number <> 0 Or cmp Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    m_App.DoCmd.DeleteObject acModule, strName
    If Err.Number <> 0 Then
        Err.Clear
        m_VBProject.VBComponents.Remove cmp
    End If
    Err.Clear

End Sub


Private Sub SaveModule(ByVal strName As String)
    On Error Resume Next
    Set m_App.VBE.ActiveVBProject = m_VBProject
    Err.Clear
    m_App.DoCmd.Save acModule, strName
    If Err.Number <> 0 Then LogWarning "Unable to save module " & strName & " (" & Err.Description & ")"
    Err.Clear
End Sub


'=======================================================================================
' Forms, reports and macros
'=======================================================================================

Private Sub ImportAllObjects(ByVal strFolder As String, ByVal intType As Long)

    Dim dFiles As Object
    Dim varKey As Variant

    Select Case intType
        Case acForm: Set dFiles = GetSourceFiles(m_Source & strFolder, "form|bas", True)
        Case acReport: Set dFiles = GetSourceFiles(m_Source & strFolder, "report|bas", True)
        Case acMacro: Set dFiles = GetSourceFiles(m_Source & strFolder, "macro|bas", False)
    End Select
    If dFiles.Count = 0 Then Exit Sub
    LogLine "Importing " & strFolder & "..."

    For Each varKey In dFiles.Keys
        If intType = acMacro Then
            If m_FullBuild And StrComp(varKey, "AutoExec", vbTextCompare) = 0 Then
                ' Import last, so it does not run while the build is in progress
                m_DeferredMacros.Add CStr(dFiles(varKey))
            Else
                ImportMacro CStr(dFiles(varKey))
            End If
        Else
            ImportFormOrReport intType, CStr(dFiles(varKey))
        End If
    Next varKey

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportFormOrReport
' Purpose   : Recombine .form/.report + .cls, load the object, then apply the print
'           : settings and metadata from the companion .json file.
'---------------------------------------------------------------------------------------
'
Private Sub ImportFormOrReport(ByVal intType As Long, ByVal strFile As String)

    Dim strName As String
    Dim strText As String
    Dim strClsFile As String
    Dim dItems As Object

    strName = GetObjectNameFromFileName(strFile)
    Status "Importing " & IIf(intType = acForm, "form ", "report ") & strName

    On Error GoTo ErrHandler
    strText = ReadTextFile(strFile)
    strClsFile = SwapExtension(strFile, "cls")
    If m_FSO.FileExists(strClsFile) Then
        strText = MergeVBA(strText, ReadTextFile(strClsFile))
    ElseIf InStr(1, strText, vbCrLf & "CodeBehindForm" & vbCrLf & "' See """, vbTextCompare) > 0 Then
        ' The code-behind file was removed, so the object no longer has a module.
        strText = MergeVBA(strText, vbNullString)
    End If

    Set dItems = ReadItems(SwapExtension(strFile, "json"))
    If dItems.Exists("ConditionalFormatting") Then
        LogWarning "Conditional formatting for " & strName & " is stored in JSON (decoded by the add-in) " & _
            "and cannot be rebuilt by this module. Re-create it in Access."
    End If

    CloseObjectIfOpen intType, strName
    If intType = acReport And m_InPlace Then DeleteObjectIfExists acReport, strName
    LoadFromTextEx intType, strName, strText

    If dItems.Exists("Printer") Or dItems.Exists("Margins") Or dItems.Exists("Device") Then
        ApplyPrintSettings intType, strName, dItems
    End If
    ImportObjectMetadata dItems, IIf(intType = acForm, "Forms", "Reports"), strName, intType
    If intType = acForm Then m_Forms.Add strName
    Exit Sub

ErrHandler:
    LogError "Error importing " & IIf(intType = acForm, "form ", "report ") & strName

End Sub


Private Sub ImportMacro(ByVal strFile As String)

    Dim strName As String

    strName = GetObjectNameFromFileName(strFile)
    Status "Importing macro " & strName
    On Error GoTo ErrHandler
    If m_InPlace Then DeleteObjectIfExists acMacro, strName
    LoadFromTextEx acMacro, strName, ReadTextFile(strFile)
    ImportObjectMetadata ReadItems(SwapExtension(strFile, "json")), "Scripts", strName, acMacro
    Exit Sub

ErrHandler:
    LogError "Error importing macro " & strName

End Sub


Private Sub DeleteObjectIfExists(ByVal intType As Long, ByVal strName As String)
    On Error Resume Next
    m_App.DoCmd.DeleteObject intType, strName
    Err.Clear
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MergeVBA
' Purpose   : Put the code-behind (.cls) back after the CodeBehindForm line.
'---------------------------------------------------------------------------------------
'
Private Function MergeVBA(ByVal strForm As String, ByVal strVBA As String) As String

    Const MARKER As String = "CodeBehindForm"

    Dim lngPos As Long
    Dim blnHasCode As Boolean

    strVBA = StripClassHeader(strVBA)
    blnHasCode = (Len(Trim$(Replace(strVBA, vbCrLf, vbNullString))) > 0)
    If Right$(strForm, 2) <> vbCrLf Then strForm = strForm & vbCrLf

    lngPos = InStr(1, vbCrLf & strForm, vbCrLf & MARKER & vbCrLf, vbTextCompare)
    If lngPos > 0 Then strForm = Left$(strForm, lngPos - 1)

    If blnHasCode Then
        If Right$(strVBA, 2) <> vbCrLf Then strVBA = strVBA & vbCrLf
        MergeVBA = strForm & MARKER & vbCrLf & strVBA
    Else
        MergeVBA = strForm
    End If

End Function


' Code-behind uses the SaveAsText layout, which starts with the class attributes and
' has no VERSION block or VB_Name. Normalize code that was written as a class module.
Private Function StripClassHeader(ByVal strVBA As String) As String

    Dim varLines As Variant
    Dim lngLine As Long
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim lngCnt As Long
    Dim strLines() As String
    Dim strOut As String
    Dim blnAttributes As Boolean

    strVBA = NormalizeText(strVBA)
    varLines = Split(strVBA, vbCrLf)
    If UBound(varLines) < 0 Then Exit Function

    If Trim$(varLines(0)) = "VERSION 1.0 CLASS" Then
        For lngLine = 1 To UBound(varLines)
            If Trim$(varLines(lngLine)) = "END" Then
                lngFirst = lngLine + 1
                Exit For
            End If
        Next lngLine
    End If

    lngLast = UBound(varLines)
    If varLines(lngLast) = vbNullString Then lngLast = lngLast - 1      ' (Text ends with a line break)
    ReDim strLines(0 To UBound(varLines) - lngFirst + 1)
    For lngLine = lngFirst To lngLast
        If Left$(varLines(lngLine), 20) <> "Attribute VB_Name = " Then
            If Left$(varLines(lngLine), 31) = "Attribute VB_GlobalNameSpace = " Then blnAttributes = True
            strLines(lngCnt) = varLines(lngLine)
            lngCnt = lngCnt + 1
        End If
    Next lngLine
    If lngCnt > 0 Then
        ReDim Preserve strLines(0 To lngCnt - 1)
        strOut = Join(strLines, vbCrLf) & vbCrLf
    End If

    If Not blnAttributes And Len(Trim$(Replace(strOut, vbCrLf, vbNullString))) > 0 Then
        strOut = "Attribute VB_GlobalNameSpace = False" & vbCrLf & "Attribute VB_Creatable = True" & vbCrLf & _
            "Attribute VB_PredeclaredId = True" & vbCrLf & "Attribute VB_Exposed = False" & vbCrLf & strOut
    End If
    StripClassHeader = strOut

End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadFromTextEx
' Purpose   : Write the text in the encoding expected by LoadFromText and load it.
'---------------------------------------------------------------------------------------
'
Private Sub LoadFromTextEx(ByVal intType As Long, ByVal strName As String, ByVal strText As String)

    Dim strTemp As String
    Dim lngErr As Long
    Dim strErr As String

    strTemp = GetTempFile(".txt")
    If RequiresUcs2 Then
        WriteTextFile strTemp, strText, "unicode"
    Else
        WriteTextFile strTemp, strText, "utf-8"
    End If

    On Error Resume Next
    m_App.LoadFromText intType, strName, strTemp
    lngErr = Err.Number
    strErr = Err.Description
    Err.Clear
    On Error GoTo 0

    DeleteFile strTemp
    If lngErr <> 0 Then Err.Raise lngErr, , strErr

End Sub


' True if LoadFromText expects UTF-16 files (.accdb). Detected by exporting a
' temporary query, like the add-in does.
Private Function RequiresUcs2() As Boolean

    Const PROBE_NAME As String = "zzVCS_Ucs2_Probe"

    Dim strTemp As String
    Dim bteData() As Byte

    If IsEmpty(m_Ucs2) Then
        On Error Resume Next
        m_Ucs2 = (m_App.CurrentProject.FileFormat >= 12)
        m_Dbs.QueryDefs.Delete PROBE_NAME
        Err.Clear
        m_Dbs.CreateQueryDef PROBE_NAME, "SELECT 1 AS Probe;"
        If Err.Number = 0 Then
            strTemp = GetTempFile(".txt")
            m_App.SaveAsText acQuery, PROBE_NAME, strTemp
            If Err.Number = 0 Then
                bteData = ReadBinaryFile(strTemp, 2)
                If UBound(bteData) >= 1 Then m_Ucs2 = (bteData(0) = &HFF And bteData(1) = &HFE)
            End If
            DeleteFile strTemp
            Err.Clear
            m_Dbs.QueryDefs.Delete PROBE_NAME
        End If
        Err.Clear
        On Error GoTo 0
    End If
    RequiresUcs2 = m_Ucs2

End Function


'---------------------------------------------------------------------------------------
' Procedure : ApplyPrintSettings
' Purpose   : Restore printer, paper and margin settings through the Printer object of
'           : the form or report (opened hidden in design view).
'---------------------------------------------------------------------------------------
'
Private Sub ApplyPrintSettings(ByVal intType As Long, ByVal strName As String, dItems As Object)

    Dim obj As Object
    Dim prt As Object
    Dim dSettings As Object
    Dim varKey As Variant
    Dim lngTwips As Long
    Dim strDevice As String

    On Error GoTo ErrHandler

    If intType = acReport Then
        m_App.DoCmd.OpenReport strName, acViewDesign, , , acHidden
        Set obj = m_App.Reports(strName)
    Else
        m_App.DoCmd.OpenForm strName, acDesign, , , , acHidden
        Set obj = m_App.Forms(strName)
    End If

    ' Specific printer
    Set dSettings = GetDict(dItems, "Device")
    If Not dSettings Is Nothing Then
        strDevice = GetText(dSettings, "DeviceName")
        Set prt = FindPrinter(strDevice)
        If prt Is Nothing Then
            LogWarning "Printer """ & strDevice & """ was not found for " & strName & ". Using the default printer."
        Else
            Set obj.Printer = prt
            obj.UseDefaultPrinter = False
        End If
    End If

    Set prt = obj.Printer

    Set dSettings = GetDict(dItems, "Printer")
    If Not dSettings Is Nothing Then
        For Each varKey In dSettings.Keys
            Select Case varKey
                Case "Orientation": prt.Orientation = EnumValue("Orientation", dSettings(varKey))
                Case "PaperSize": prt.PaperSize = EnumValue("PaperSize", dSettings(varKey))
                Case "Copies": prt.Copies = CLng(dSettings(varKey))
                Case "PrintQuality": prt.PrintQuality = EnumValue("PrintQuality", dSettings(varKey))
                Case "Color": prt.ColorMode = EnumValue("Color", dSettings(varKey))
                Case "Duplex": prt.Duplex = EnumValue("Duplex", dSettings(varKey))
                Case "DefaultSource": prt.PaperBin = EnumValue("DefaultSource", dSettings(varKey))
            End Select
        Next varKey
    End If

    Set dSettings = GetDict(dItems, "Margins")
    If Not dSettings Is Nothing Then
        For Each varKey In dSettings.Keys
            Select Case varKey
                Case "LeftMargin": prt.LeftMargin = InchesToTwips(dSettings(varKey))
                Case "TopMargin": prt.TopMargin = InchesToTwips(dSettings(varKey))
                Case "RightMargin": prt.RightMargin = InchesToTwips(dSettings(varKey))
                Case "BotMargin": prt.BottomMargin = InchesToTwips(dSettings(varKey))
                Case "DataOnly": prt.DataOnly = CBool(dSettings(varKey))
                Case "Columns": prt.ItemsAcross = CLng(dSettings(varKey))
                Case "ColumnSpacing": prt.ColumnSpacing = InchesToTwips(dSettings(varKey))
                Case "RowSpacing": prt.RowSpacing = InchesToTwips(dSettings(varKey))
                Case "ItemLayout": prt.ItemLayout = EnumValue("ItemLayout", dSettings(varKey))
                Case "DefaultSize": prt.DefaultSize = CBool(dSettings(varKey))
                Case "Width"
                    lngTwips = InchesToTwips(dSettings(varKey))
                    If prt.ItemSizeWidth <> lngTwips Then
                        If prt.DefaultSize Then prt.DefaultSize = False
                        prt.ItemSizeWidth = lngTwips
                    End If
                Case "Height"
                    lngTwips = InchesToTwips(dSettings(varKey))
                    If prt.ItemSizeHeight <> lngTwips Then
                        If prt.DefaultSize Then prt.DefaultSize = False
                        prt.ItemSizeHeight = lngTwips
                    End If
            End Select
        Next varKey
    End If

    ' Printer changes alone do not mark the design as changed.
    obj.Caption = obj.Caption
    m_App.DoCmd.Close intType, strName, acSaveYes
    Exit Sub

ErrHandler:
    LogWarning "Unable to apply print settings to " & strName & " (" & Err.Description & ")"
    Err.Clear
    On Error GoTo -1
    On Error Resume Next
    m_App.DoCmd.Close intType, strName, acSaveYes
    Err.Clear

End Sub


Private Function FindPrinter(ByVal strDevice As String) As Object
    Dim prt As Object
    On Error Resume Next
    For Each prt In m_App.Printers
        If StrComp(prt.DeviceName, strDevice, vbTextCompare) = 0 Then
            Set FindPrinter = prt
            Exit For
        End If
    Next prt
    Err.Clear
End Function


Private Function InchesToTwips(ByVal varInches As Variant) As Long
    InchesToTwips = CLng(Round(CDbl(varInches) * 1440, 0))
End Function


'---------------------------------------------------------------------------------------
' Procedure : EnumValue
' Purpose   : Convert the names used in the print settings json back to values.
'---------------------------------------------------------------------------------------
'
Private Function EnumValue(ByVal strEnum As String, ByVal varName As Variant) As Long

    If IsNumeric(varName) Then
        EnumValue = CLng(varName)
        Exit Function
    End If

    Select Case strEnum
        Case "Orientation"
            Select Case LCase$(CStr(varName))
                Case "portrait": EnumValue = acPRORPortrait
                Case "landscape": EnumValue = acPRORLandscape
            End Select
        Case "ItemLayout"
            Select Case LCase$(CStr(varName))
                Case "horizontal columns": EnumValue = acPRHorizontalColumnLayout
                Case "vertical columns": EnumValue = acPRVerticalColumnLayout
            End Select
        Case "Color"
            Select Case LCase$(CStr(varName))
                Case "color": EnumValue = acPRCMColor
                Case "monochrome": EnumValue = acPRCMMonochrome
            End Select
        Case "Duplex"
            Select Case LCase$(CStr(varName))
                Case "horizontal": EnumValue = acPRDPHorizontal
                Case "simplex": EnumValue = acPRDPSimplex
                Case "vertical": EnumValue = acPRDPVertical
            End Select
        Case "PrintQuality"
            Select Case LCase$(CStr(varName))
                Case "draft": EnumValue = acPRPQDraft
                Case "high": EnumValue = acPRPQHigh
                Case "low": EnumValue = acPRPQLow
                Case "medium": EnumValue = acPRPQMedium
            End Select
        Case "DefaultSource"
            Select Case LCase$(CStr(varName))
                Case "auto": EnumValue = acPRBNAuto
                Case "cassette": EnumValue = acPRBNCassette
                Case "envelope": EnumValue = acPRBNEnvelope
                Case "envelope manual": EnumValue = acPRBNEnvManual
                Case "form source": EnumValue = acPRBNFormSource
                Case "large capacity": EnumValue = acPRBNLargeCapacity
                Case "large format": EnumValue = acPRBNLargeFmt
                Case "lower": EnumValue = acPRBNLower
                Case "manual": EnumValue = acPRBNManual
                Case "middle": EnumValue = acPRBNMiddle
                Case "small format": EnumValue = acPRBNSmallFmt
                Case "tractor": EnumValue = acPRBNTractor
                Case "upper": EnumValue = acPRBNUpper
            End Select
        Case "PaperSize"
            Select Case LCase$(CStr(varName))
                Case "10x14": EnumValue = acPRPS10x14
                Case "11x17": EnumValue = acPRPS11x17
                Case "a3": EnumValue = acPRPSA3
                Case "a4": EnumValue = acPRPSA4
                Case "a4 small": EnumValue = acPRPSA4Small
                Case "a5": EnumValue = acPRPSA5
                Case "b4": EnumValue = acPRPSB4
                Case "b5": EnumValue = acPRPSB5
                Case "c size sheet": EnumValue = acPRPSCSheet
                Case "d size sheet": EnumValue = acPRPSDSheet
                Case "envelope #10": EnumValue = acPRPSEnv10
                Case "envelope #11": EnumValue = acPRPSEnv11
                Case "envelope #12": EnumValue = acPRPSEnv12
                Case "envelope #14": EnumValue = acPRPSEnv14
                Case "envelope #9": EnumValue = acPRPSEnv9
                Case "envelope b4": EnumValue = acPRPSEnvB4
                Case "envelope b5": EnumValue = acPRPSEnvB5
                Case "envelope b6": EnumValue = acPRPSEnvB6
                Case "envelope c3": EnumValue = acPRPSEnvC3
                Case "envelope c4": EnumValue = acPRPSEnvC4
                Case "envelope c5": EnumValue = acPRPSEnvC5
                Case "envelope c6": EnumValue = acPRPSEnvC6
                Case "envelope c65": EnumValue = acPRPSEnvC65
                Case "envelope dl": EnumValue = acPRPSEnvDL
                Case "italian envelope": EnumValue = acPRPSEnvItaly
                Case "monarch envelope": EnumValue = acPRPSEnvMonarch
                Case "envelope": EnumValue = acPRPSEnvPersonal
                Case "e size sheet": EnumValue = acPRPSESheet
                Case "executive": EnumValue = acPRPSExecutive
                Case "german legal fanfold": EnumValue = acPRPSFanfoldLglGerman
                Case "german standard fanfold": EnumValue = acPRPSFanfoldStdGerman
                Case "u.s. standard fanfold": EnumValue = acPRPSFanfoldUS
                Case "folio": EnumValue = acPRPSFolio
                Case "ledger": EnumValue = acPRPSLedger
                Case "legal": EnumValue = acPRPSLegal
                Case "letter": EnumValue = acPRPSLetter
                Case "letter small": EnumValue = acPRPSLetterSmall
                Case "note": EnumValue = acPRPSNote
                Case "quarto": EnumValue = acPRPSQuarto
                Case "statement": EnumValue = acPRPSStatement
                Case "tabloid": EnumValue = acPRPSTabloid
                Case "user-defined": EnumValue = acPRPSUser
            End Select
    End Select

    If EnumValue = 0 Then Err.Raise vbObjectError + 514, , "Unknown " & strEnum & " value: " & varName

End Function


'---------------------------------------------------------------------------------------
' Procedure : InitializeForms
' Purpose   : Open and save each imported form in design view so theme colors are
'           : rendered (same as the add-in does after a build).
'---------------------------------------------------------------------------------------
'
Private Sub InitializeForms()

    Dim varName As Variant
    Dim frm As Object

    If m_Forms Is Nothing Then Exit Sub
    If m_Forms.Count = 0 Then Exit Sub
    LogLine "Initializing forms..."
    For Each varName In m_Forms
        On Error Resume Next
        m_App.DoCmd.OpenForm CStr(varName), acDesign, , , , acHidden
        If Err.Number = 0 Then
            Set frm = m_App.Forms(CStr(varName))
            frm.Tag = frm.Tag
            Set frm = Nothing
            m_App.DoCmd.Close acForm, CStr(varName), acSaveYes
        End If
        If Err.Number <> 0 Then LogWarning "Unable to initialize form " & varName & " (" & Err.Description & ")"
        Err.Clear
        On Error GoTo 0
    Next varName

End Sub


'=======================================================================================
' Queries
'=======================================================================================

'---------------------------------------------------------------------------------------
' Procedure : ImportAllQueries
' Purpose   : Import all queries. Queries that fail are retried after the others, in
'           : case they depend on queries that did not exist yet.
'---------------------------------------------------------------------------------------
'
Private Sub ImportAllQueries()

    Dim dPending As Object
    Dim dFailed As Object
    Dim varKey As Variant
    Dim strError As String
    Dim lngPass As Long

    Set dPending = GetSourceFiles(m_Source & "queries", "sql|qdef|bas", False)
    If dPending.Count = 0 Then Exit Sub
    LogLine "Importing queries..."

    For lngPass = 1 To 4
        Set dFailed = NewDict
        For Each varKey In dPending.Keys
            strError = ImportQuery(CStr(dPending(varKey)))
            If Len(strError) > 0 Then dFailed.Add varKey, Array(dPending(varKey), strError)
        Next varKey
        If dFailed.Count = 0 Or dFailed.Count = dPending.Count Then Exit For
        Set dPending = NewDict
        For Each varKey In dFailed.Keys
            dPending.Add varKey, dFailed(varKey)(0)
        Next varKey
    Next lngPass

    For Each varKey In dFailed.Keys
        ReportQueryResult CStr(varKey), CStr(dFailed(varKey)(1))
    Next varKey
    m_Dbs.QueryDefs.Refresh
    ApplyPendingColumns

End Sub


Private Sub ApplyPendingColumns()

    Dim varItem As Variant
    Dim qdf As Object
    Dim dCols As Object

    For Each varItem In m_PendingColumns
        Set qdf = Nothing
        On Error Resume Next
        Set qdf = m_Dbs.QueryDefs(CStr(varItem(0)))
        Err.Clear
        On Error GoTo 0
        If Not qdf Is Nothing Then
            Set dCols = varItem(1)
            ApplyQueryColumns qdf, dCols, CStr(varItem(0))
        End If
    Next varItem
    Set m_PendingColumns = New Collection

End Sub


Private Sub ReportQueryResult(ByVal strName As String, ByVal strError As String)
    If Len(strError) > 0 Then
        m_Errors = m_Errors + 1
        LogLine "  ERROR: Error importing query " & strName & " (" & strError & ")"
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportQuery
' Purpose   : Create a query from its .sql and .json files (or a legacy .qdef/.bas
'           : file). Returns an error description, or an empty string on success.
'---------------------------------------------------------------------------------------
'
Private Function ImportQuery(ByVal strFile As String) As String

    Dim strName As String
    Dim strSql As String
    Dim strConnect As String
    Dim dItems As Object
    Dim dProps As Object
    Dim dCols As Object
    Dim qdf As Object
    Dim varKey As Variant
    Dim lngType As Long
    Dim blnPassThrough As Boolean
    Dim strError As String

    strName = GetObjectNameFromFileName(strFile)
    Status "Importing query " & strName
    On Error GoTo ErrHandler

    CloseObjectIfOpen acQuery, strName
    Set dItems = ReadItems(SwapExtension(strFile, "json"))

    Select Case LCase$(m_FSO.GetExtensionName(strFile))
        Case "qdef", "bas"
            ' Legacy format: SaveAsText output
            DeleteQuery strName
            LoadFromTextEx acQuery, strName, ReadTextFile(strFile)

        Case Else
            strSql = ReadTextFile(strFile)
            If Len(Trim$(Replace(strSql, vbCrLf, vbNullString))) = 0 Then Err.Raise vbObjectError + 515, , "The .sql file is empty"

            lngType = -1
            If dItems.Exists("QueryType") Then lngType = CLng(dItems("QueryType"))
            If dItems.Exists("Connect") Then strConnect = ResolveConnect(Nz(dItems("Connect"), vbNullString))
            blnPassThrough = (lngType = DB_QSQL_PASS_THROUGH Or lngType = DB_QSPT_BULK Or Len(strConnect) > 0)
            Set dProps = GetDict(dItems, "QueryProperties")

            DeleteQuery strName
            If blnPassThrough Then
                ' The connection must be set before the SQL, so Access does not parse it.
                Set qdf = m_Dbs.CreateQueryDef(strName)
                If Len(strConnect) = 0 Then strConnect = "ODBC;"
                qdf.Connect = strConnect
                qdf.SQL = strSql
                qdf.ReturnsRecords = CBool(GetPropertyValue(dProps, "ReturnsRecords", (lngType <> DB_QSPT_BULK)))
                If Not IsEmpty(GetPropertyValue(dProps, "ODBCTimeout", Empty)) Then
                    qdf.ODBCTimeout = CLng(GetPropertyValue(dProps, "ODBCTimeout", 60))
                End If
            Else
                Set qdf = m_Dbs.CreateQueryDef(strName, strSql)
            End If

            ' Other query properties
            If Not dProps Is Nothing Then
                For Each varKey In dProps.Keys
                    Select Case varKey
                        Case "ReturnsRecords", "Description"
                            ' Handled elsewhere
                        Case "ODBCTimeout"
                            If Not blnPassThrough Then SetTypedProperty qdf, CStr(varKey), dProps(varKey), "query property"
                        Case Else
                            SetTypedProperty qdf, CStr(varKey), dProps(varKey), "query property"
                    End Select
                Next varKey
            End If

            ' Column properties (width, caption, format...) are applied once all
            ' queries exist, since the fields of a query may depend on other queries.
            Set dCols = GetDict(dItems, "Columns")
            If Not dCols Is Nothing Then m_PendingColumns.Add Array(strName, dCols)
    End Select

    ImportObjectMetadata dItems, "Tables", strName, acQuery
    Exit Function

ErrHandler:
    strError = "Error " & Err.Number & ": " & Err.Description
    Err.Clear
    On Error GoTo -1
    On Error Resume Next
    ' Remove a partially created pass-through query
    If blnPassThrough Then m_Dbs.QueryDefs.Delete strName
    Err.Clear
    ImportQuery = strError

End Function


Private Sub DeleteQuery(ByVal strName As String)
    On Error Resume Next
    m_Dbs.QueryDefs.Delete strName
    Err.Clear
End Sub


' Value of a {"Type":..,"Value":..} entry in a properties dictionary
Private Function GetPropertyValue(dProps As Object, ByVal strName As String, ByVal varDefault As Variant) As Variant
    Dim dProp As Object
    GetPropertyValue = varDefault
    If dProps Is Nothing Then Exit Function
    If Not dProps.Exists(strName) Then Exit Function
    Set dProp = GetDict(dProps, strName)
    If dProp Is Nothing Then
        GetPropertyValue = dProps(strName)
    ElseIf dProp.Exists("Value") Then
        GetPropertyValue = dProp("Value")
    End If
End Function


' Set a property from a {"Type":..,"Value":..} entry (or a plain value)
Private Sub SetTypedProperty(objParent As Object, ByVal strName As String, ByVal varEntry As Variant, ByVal strLabel As String)
    If IsObject(varEntry) Then
        SetDaoProperty objParent, strName, DaoTypeFromName(GetValue(varEntry, "Type"), strName), GetValue(varEntry, "Value"), strLabel
    Else
        SetDaoProperty objParent, strName, DaoTypeFromName(Empty, strName), varEntry, strLabel
    End If
End Sub


Private Sub ApplyQueryColumns(qdf As Object, dCols As Object, ByVal strQuery As String)

    Dim varCol As Variant
    Dim varProp As Variant
    Dim fld As Object
    Dim dCol As Object

    On Error GoTo ErrHandler
    For Each varCol In dCols.Keys
        Set fld = FindQueryField(qdf, CStr(varCol))
        Set dCol = GetDict(dCols, CStr(varCol))
        If fld Is Nothing Or dCol Is Nothing Then
            LogLine "  Note: column " & varCol & " not found in query " & strQuery & " (column settings skipped)"
        Else
            For Each varProp In dCol.Keys
                If varProp <> "AggregateType" Then SetTypedProperty fld, CStr(varProp), dCol(varProp), "column property"
            Next varProp
        End If
    Next varCol
    Exit Sub

ErrHandler:
    LogWarning "Unable to apply column settings to query " & strQuery & " (" & Err.Description & ")"
    Err.Clear

End Sub


Private Function FindQueryField(qdf As Object, ByVal strName As String) As Object
    On Error Resume Next
    Set FindQueryField = qdf.Fields(strName)
    If FindQueryField Is Nothing And InStr(1, strName, ".") > 0 Then
        Err.Clear
        Set FindQueryField = qdf.Fields(Mid$(strName, InStrRev(strName, ".") + 1))
    End If
    Err.Clear
End Function


' DAO type from an integer, a "dbXxx" name, or a known property name.
Private Function DaoTypeFromName(ByVal varType As Variant, ByVal strProperty As String) As Long

    If IsNumeric(varType) And Not IsEmpty(varType) Then
        DaoTypeFromName = CLng(varType)
        Exit Function
    End If

    Select Case LCase$(Nz(varType, vbNullString))
        Case "dbboolean": DaoTypeFromName = DB_BOOLEAN
        Case "dbbyte": DaoTypeFromName = DB_BYTE
        Case "dbinteger": DaoTypeFromName = DB_INTEGER
        Case "dblong": DaoTypeFromName = DB_LONG
        Case "dbcurrency": DaoTypeFromName = DB_CURRENCY
        Case "dbsingle": DaoTypeFromName = DB_SINGLE
        Case "dbdouble": DaoTypeFromName = DB_DOUBLE
        Case "dbdate": DaoTypeFromName = DB_DATE
        Case "dbtext": DaoTypeFromName = DB_TEXT
        Case "dbmemo": DaoTypeFromName = DB_MEMO
        Case "dbguid": DaoTypeFromName = DB_GUID
        Case Else
            ' Known column properties
            Select Case strProperty
                Case "ColumnHidden": DaoTypeFromName = DB_BOOLEAN
                Case "ColumnWidth", "ColumnOrder", "DisplayControl", "ShowDatePicker": DaoTypeFromName = DB_INTEGER
                Case "TextAlign", "DecimalPlaces", "ResultType", "IMEMode", "IMESentenceMode": DaoTypeFromName = DB_BYTE
                Case "CurrencyLCID", "AggregateType": DaoTypeFromName = DB_LONG
                Case Else: DaoTypeFromName = DB_TEXT
            End Select
    End Select

End Function


'=======================================================================================
' Tables
'=======================================================================================

Private Function TableExists(ByVal strName As String) As Boolean
    Dim strTest As String
    On Error Resume Next
    strTest = m_Dbs.TableDefs(strName).Name
    TableExists = (Err.Number = 0)
    Err.Clear
End Function


Private Function IsLinkedTable(ByVal strName As String) As Boolean
    On Error Resume Next
    IsLinkedTable = (Len(m_Dbs.TableDefs(strName).Connect) > 0)
    Err.Clear
End Function


'---------------------------------------------------------------------------------------
' Procedure : ImportAllTableDefs
' Purpose   : Local tables (tbldefs\*.xml) first, then linked tables (*.json with a
'           : Connect value). The .sql files are informational and not imported.
'---------------------------------------------------------------------------------------
'
Private Sub ImportAllTableDefs()

    Dim dLocal As Object
    Dim dJson As Object
    Dim varKey As Variant
    Dim dItems As Object

    Set dLocal = GetSourceFiles(m_Source & "tbldefs", "xml", False)
    Set dJson = GetSourceFiles(m_Source & "tbldefs", "json", False)
    If dLocal.Count + dJson.Count = 0 Then Exit Sub
    LogLine "Importing tables..."

    For Each varKey In dLocal.Keys
        ImportLocalTable CStr(dLocal(varKey)), False
    Next varKey

    For Each varKey In dJson.Keys
        Set dItems = ReadItems(CStr(dJson(varKey)))
        If dItems.Exists("Connect") Then ImportLinkedTable CStr(dJson(varKey))
    Next varKey
    m_Dbs.TableDefs.Refresh

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportLocalTable
' Purpose   : Create a local table from its XML schema. An existing table is replaced
'           : only when replacing tables is allowed (its data is lost unless the
'           : source contains data for the table).
'---------------------------------------------------------------------------------------
'
Private Sub ImportLocalTable(ByVal strFile As String, ByVal blnSingleObject As Boolean)

    Dim strName As String
    Dim strPath As String
    Dim strData As String
    Dim colRelations As Collection
    Dim lngRecords As Long

    strName = GetObjectNameFromFileName(strFile)
    Status "Importing table " & strName
    On Error GoTo ErrHandler

    If Not m_FSO.FileExists(strFile) Then
        LogError "Table definition not found: " & strFile
        Exit Sub
    End If

    If TableExists(strName) Then
        If Not m_ReplaceTables Then
            LogLine "  Kept existing table " & strName & " (use ReplaceExistingTables:=True to replace it)"
            ImportObjectMetadata ReadItems(SwapExtension(strFile, "json")), "Tables", strName, acTable
            Exit Sub
        End If
        If blnSingleObject And m_InPlace And Not IsLinkedTable(strName) Then
            lngRecords = DCountTable(strName)
            strData = GetTableDataFile(strName)
            If lngRecords > 0 Then
                If MsgBox("Table '" & strName & "' contains " & lngRecords & " record(s)." & vbCrLf & vbCrLf & _
                    IIf(Len(strData) > 0, "The table will be recreated and its data reloaded from:" & vbCrLf & strData, _
                    "The table will be recreated and ALL OF ITS DATA WILL BE LOST.") & vbCrLf & vbCrLf & "Continue?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Replace Table") <> vbYes Then
                    LogLine "  Skipped table " & strName
                    Exit Sub
                End If
            End If
        End If
        CloseObjectIfOpen acTable, strName
        Set colRelations = StageRelations(strName)
        m_Dbs.TableDefs.Delete strName
        m_Dbs.TableDefs.Refresh
    End If

    ' ImportXML does not handle encoded characters (%) in the path
    strPath = strFile
    If InStr(1, strPath, "%") > 0 Then
        strPath = GetTempFile(".xml")
        m_FSO.CopyFile strFile, strPath, True
    End If
    m_App.ImportXML strPath, acStructureOnly
    If strPath <> strFile Then DeleteFile strPath
    m_Dbs.TableDefs.Refresh

    If TableExists(strName) Then
        m_NewTables(strName) = True
        ImportObjectMetadata ReadItems(SwapExtension(strFile, "json")), "Tables", strName, acTable
    Else
        LogError "Table " & strName & " was not created from " & m_FSO.GetFileName(strFile)
    End If

    ' Reload data and data macros before restoring relationships, so related data
    ' is in place when referential integrity is checked again.
    If blnSingleObject Then
        strData = GetTableDataFile(strName)
        If Len(strData) > 0 Then ImportTableData strData, True
        strData = m_Source & "tdmacros\" & GetSafeFileName(strName) & ".xml"
        If m_FSO.FileExists(strData) Then ImportTableDataMacro strData
    End If
    KeepStagedRelations colRelations, blnSingleObject
    Exit Sub

ErrHandler:
    LogError "Error importing table " & strName
    KeepStagedRelations colRelations, blnSingleObject

End Sub


' Restore removed relationships now (single object), or after all data has been
' loaded (merge).
Private Sub KeepStagedRelations(colRelations As Collection, ByVal blnRestoreNow As Boolean)
    Dim varItem As Variant
    If colRelations Is Nothing Then Exit Sub
    If blnRestoreNow Then
        RestoreRelations colRelations
    Else
        For Each varItem In colRelations
            m_StagedRelations.Add varItem
        Next varItem
    End If
End Sub


' Recreate relationships removed while replacing tables, unless the source files
' already recreated them.
Private Sub RestoreStagedRelations()

    Dim varItem As Variant
    Dim dRel As Object
    Dim strTest As String

    For Each varItem In m_StagedRelations
        Set dRel = varItem
        strTest = vbNullString
        On Error Resume Next
        strTest = m_Dbs.Relations(GetText(dRel, "Name")).Name
        Err.Clear
        On Error GoTo 0
        If Len(strTest) = 0 Then CreateRelation dRel
    Next varItem
    Set m_StagedRelations = New Collection

End Sub


Private Function DCountTable(ByVal strName As String) As Long
    Dim rst As Object
    On Error Resume Next
    Set rst = m_Dbs.OpenRecordset("SELECT Count(*) AS Total FROM [" & strName & "]")
    DCountTable = rst.Fields(0).Value
    rst.Close
    Err.Clear
End Function


Private Function GetTableDataFile(ByVal strName As String) As String
    Dim strBase As String
    strBase = m_Source & "tables\" & GetSafeFileName(strName)
    If m_FSO.FileExists(strBase & ".txt") Then
        GetTableDataFile = strBase & ".txt"
    ElseIf m_FSO.FileExists(strBase & ".xml") Then
        GetTableDataFile = strBase & ".xml"
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ImportLinkedTable
' Purpose   : Recreate a linked table from its .json definition.
'---------------------------------------------------------------------------------------
'
Private Sub ImportLinkedTable(ByVal strFile As String)

    Dim dItems As Object
    Dim dProps As Object
    Dim dField As Object
    Dim tdf As Object
    Dim fld As Object
    Dim idx As Object
    Dim varKey As Variant
    Dim varProp As Variant
    Dim strName As String
    Dim strConnect As String
    Dim blnUnique As Boolean
    Dim colRelations As Collection

    Set dItems = ReadItems(strFile)
    strName = GetText(dItems, "Name", GetObjectNameFromFileName(strFile))
    Status "Linking table " & strName
    On Error GoTo ErrHandler

    strConnect = ResolveConnect(GetText(dItems, "Connect"))

    If TableExists(strName) Then
        If Not IsLinkedTable(strName) Then
            ' Replacing a local table with a link discards the local data.
            If m_SingleObject Then
                If MsgBox("'" & strName & "' is a local table with " & DCountTable(strName) & " record(s)." & vbCrLf & vbCrLf & _
                    "Replace it with a linked table? The local table and ALL OF ITS DATA will be deleted.", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Replace Table") <> vbYes Then
                    LogLine "  Kept local table " & strName
                    Exit Sub
                End If
            ElseIf Not m_ReplaceTables Then
                LogWarning "A local table named " & strName & " exists. The linked table was not created."
                Exit Sub
            End If
            Set colRelations = StageRelations(strName)
        End If
        CloseObjectIfOpen acTable, strName
        m_Dbs.TableDefs.Delete strName
    End If

    Set tdf = m_Dbs.CreateTableDef(strName)
    tdf.Connect = strConnect
    tdf.SourceTableName = GetText(dItems, "SourceTableName", strName)
    tdf.Attributes = CLng(Nz(GetValue(dItems, "Attributes"), 0)) And DB_SAFE_LINK_ATTRIBUTES
    On Error Resume Next
    m_Dbs.TableDefs.Append tdf
    If Err.Number <> 0 Then
        LogWarning "Unable to link table " & strName & " (" & Err.Description & "). Check that the source is available: " & strConnect
        Err.Clear
        On Error GoTo 0
        If Not colRelations Is Nothing Then KeepStagedRelations colRelations, m_SingleObject
        Exit Sub
    End If
    On Error GoTo ErrHandler
    m_Dbs.TableDefs.Refresh
    Set tdf = m_Dbs.TableDefs(strName)

    ' Non-Access sources (ODBC views, text files) may need a unique index to be updatable
    If dItems.Exists("PrimaryKey") And StrComp(Left$(strConnect, 10), ";DATABASE=", vbTextCompare) <> 0 Then
        On Error Resume Next
        For Each idx In tdf.Indexes
            If idx.Unique Then blnUnique = True
        Next idx
        Err.Clear
        If Not blnUnique Then
            m_Dbs.Execute "CREATE UNIQUE INDEX __uniqueindex ON [" & strName & "] (" & dItems("PrimaryKey") & ") WITH PRIMARY"
            If Err.Number <> 0 Then LogWarning "Unable to create a primary key on linked table " & strName & " (" & Err.Description & ")"
            Err.Clear
        End If
        On Error GoTo ErrHandler
    End If

    ' Table and field properties
    Set dProps = GetDict(dItems, "TableProperties")
    If Not dProps Is Nothing Then
        For Each varProp In dProps.Keys
            SetTypedProperty tdf, CStr(varProp), dProps(varProp), "table property"
        Next varProp
    End If
    Set dProps = GetDict(dItems, "FieldProperties")
    If Not dProps Is Nothing Then
        For Each varKey In dProps.Keys
            Set fld = Nothing
            On Error Resume Next
            Set fld = tdf.Fields(CStr(varKey))
            Err.Clear
            On Error GoTo ErrHandler
            Set dField = GetDict(dProps, CStr(varKey))
            If Not fld Is Nothing And Not dField Is Nothing Then
                For Each varProp In dField.Keys
                    SetTypedProperty fld, CStr(varProp), dField(varProp), "field property"
                Next varProp
            End If
        Next varKey
    End If

    ImportObjectMetadata dItems, "Tables", strName, acTable
    If Not colRelations Is Nothing Then KeepStagedRelations colRelations, m_SingleObject
    Exit Sub

ErrHandler:
    LogError "Error linking table " & strName
    If Not colRelations Is Nothing Then KeepStagedRelations colRelations, m_SingleObject

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResolveConnect
' Purpose   : Resolve env: references (.env file) and rel: paths in a connection string.
'---------------------------------------------------------------------------------------
'
Private Function ResolveConnect(ByVal strConnect As String) As String

    Dim varParts As Variant
    Dim lngPart As Long
    Dim strPath As String
    Dim strKey As String

    If Left$(strConnect, 4) = "env:" Then
        strKey = Mid$(strConnect, 5)
        LoadEnv
        If m_Env.Exists(strKey) Then
            strConnect = m_Env(strKey)
        Else
            LogWarning "Connection key not found in .env file: " & strKey
            strConnect = "ODBC;"
        End If
    End If

    varParts = Split(strConnect, ";")
    For lngPart = 0 To UBound(varParts)
        If StrComp(Left$(varParts(lngPart), 9), "DATABASE=", vbTextCompare) = 0 Then
            strPath = Mid$(varParts(lngPart), 10)
            If Left$(strPath, 4) = "rel:" Then
                strPath = ExpandRelativePath(strPath)
                If Right$(strPath, 1) = "\" Then strPath = Left$(strPath, Len(strPath) - 1)
                varParts(lngPart) = Left$(varParts(lngPart), 9) & strPath
            End If
        End If
    Next lngPart
    ResolveConnect = Join(varParts, ";")

End Function


Private Sub LoadEnv()

    Dim varLine As Variant
    Dim strLine As String
    Dim strValue As String
    Dim lngPos As Long

    If Not m_Env Is Nothing Then Exit Sub
    Set m_Env = CreateObject("Scripting.Dictionary")
    If Not m_FSO.FileExists(m_Source & ".env") Then Exit Sub

    For Each varLine In Split(ReadTextFile(m_Source & ".env"), vbCrLf)
        strLine = Trim$(varLine)
        If Left$(strLine, 7) = "export " Then strLine = Trim$(Mid$(strLine, 8))
        lngPos = InStr(1, strLine, "=")
        If lngPos > 1 And Left$(strLine, 1) <> "#" Then
            strValue = Trim$(Mid$(strLine, lngPos + 1))
            If Len(strValue) >= 2 Then
                If (Left$(strValue, 1) = """" And Right$(strValue, 1) = """") Or _
                    (Left$(strValue, 1) = "'" And Right$(strValue, 1) = "'") Then
                    strValue = Mid$(strValue, 2, Len(strValue) - 2)
                End If
            End If
            m_Env(Trim$(Left$(strLine, lngPos - 1))) = strValue
        End If
    Next varLine

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportAllTableData
' Purpose   : Load table data. During a merge, data is only loaded into tables that
'           : were created by the merge (existing data is never overwritten).
'---------------------------------------------------------------------------------------
'
Private Sub ImportAllTableData()

    Dim dFiles As Object
    Dim varKey As Variant

    Set dFiles = GetSourceFiles(m_Source & "tables", "txt|xml", False)
    If dFiles.Count = 0 Then Exit Sub
    LogLine "Importing table data..."
    For Each varKey In dFiles.Keys
        ImportTableData CStr(dFiles(varKey)), False
    Next varKey

End Sub


Private Sub ImportTableData(ByVal strFile As String, ByVal blnForce As Boolean)

    Dim strName As String
    Dim strPath As String
    Dim colRelations As Collection

    strName = GetObjectNameFromFileName(strFile)
    Status "Importing data for " & strName
    On Error GoTo ErrHandler

    If Not TableExists(strName) Then
        LogWarning "Table " & strName & " does not exist. Data was not imported."
        Exit Sub
    End If
    If IsLinkedTable(strName) Then
        LogLine "  Skipped data for linked table " & strName
        Exit Sub
    End If
    If Not (m_FullBuild Or blnForce Or m_NewTables.Exists(strName)) Then
        LogLine "  Kept existing data in table " & strName
        Exit Sub
    End If
    If blnForce And m_InPlace And Not m_NewTables.Exists(strName) Then
        If DCountTable(strName) > 0 Then
            If MsgBox("Replace the " & DCountTable(strName) & " existing record(s) in '" & strName & "' with the data from:" & _
                vbCrLf & strFile & " ?", vbExclamation + vbYesNo + vbDefaultButton2, "Import Table Data") <> vbYes Then
                LogLine "  Skipped data for table " & strName
                Exit Sub
            End If
        End If
        ' Remove relationships while the data is replaced, so deleting the existing
        ' rows cannot cascade to related tables.
        Set colRelations = StageRelations(strName)
    End If

    Select Case LCase$(m_FSO.GetExtensionName(strFile))
        Case "txt"
            ImportTableDataTxt strName, strFile
        Case "xml"
            If blnForce Then m_Dbs.Execute "DELETE FROM [" & strName & "]", DB_FAIL_ON_ERROR
            strPath = strFile
            If InStr(1, strPath, "%") > 0 Then
                strPath = GetTempFile(".xml")
                m_FSO.CopyFile strFile, strPath, True
            End If
            m_App.ImportXML strPath, acAppendData
            If strPath <> strFile Then DeleteFile strPath
    End Select
    If Not colRelations Is Nothing Then RestoreRelations colRelations
    Exit Sub

ErrHandler:
    LogError "Error importing data for table " & strName
    If Not colRelations Is Nothing Then RestoreRelations colRelations

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportTableDataTxt
' Purpose   : Load the add-in's "Tab Delimited" format (header row, tab separated,
'           : with \t \r \n \\ escapes). Existing rows are deleted first.
'---------------------------------------------------------------------------------------
'
Private Sub ImportTableDataTxt(ByVal strTable As String, ByVal strFile As String)

    Dim varLines As Variant
    Dim varHeader As Variant
    Dim varValues As Variant
    Dim rst As Object
    Dim fld As Object
    Dim colFields() As Object
    Dim blnUse() As Boolean
    Dim lngLine As Long
    Dim lngCol As Long
    Dim lngRows As Long
    Dim lngFailed As Long
    Dim strValue As String

    varLines = Split(ReadTextFile(strFile), vbCrLf)
    If UBound(varLines) < 0 Then Exit Sub
    varHeader = Split(varLines(0), vbTab)
    If UBound(varHeader) < 0 Then Exit Sub

    m_Dbs.Execute "DELETE FROM [" & strTable & "]", DB_FAIL_ON_ERROR
    Set rst = m_Dbs.OpenRecordset(strTable, DB_OPEN_TABLE)

    ' Map columns to fields (skip unknown, calculated and binary fields)
    ReDim colFields(0 To UBound(varHeader))
    ReDim blnUse(0 To UBound(varHeader))
    For lngCol = 0 To UBound(varHeader)
        Set fld = Nothing
        On Error Resume Next
        Set fld = rst.Fields(CStr(varHeader(lngCol)))
        Err.Clear
        On Error GoTo 0
        If Not fld Is Nothing Then
            If fld.Type < DB_ATTACHMENT And Not IsCalculatedField(fld) Then
                Set colFields(lngCol) = fld
                blnUse(lngCol) = True
            End If
        End If
    Next lngCol

    For lngLine = 1 To UBound(varLines)
        If Len(varLines(lngLine)) = 0 And lngLine = UBound(varLines) Then Exit For
        varValues = Split(varLines(lngLine), vbTab)
        On Error GoTo RowError
        rst.AddNew
        For lngCol = 0 To UBound(varHeader)
            If blnUse(lngCol) And lngCol <= UBound(varValues) Then
                strValue = varValues(lngCol)
                If Len(strValue) = 0 Then
                    ' Null and empty strings are both saved as empty values
                    If colFields(lngCol).Required And colFields(lngCol).AllowZeroLength Then
                        If IsNull(colFields(lngCol).Value) Then colFields(lngCol).Value = vbNullString
                    End If
                ElseIf strValue <> "UNSUPPORTED DATA TYPE" Then
                    colFields(lngCol).Value = UnescapeDataValue(strValue)
                End If
            End If
        Next lngCol
        rst.Update
        lngRows = lngRows + 1
        GoTo NextRow
RowError:
        lngFailed = lngFailed + 1
        If lngFailed <= 5 Then LogWarning "Row " & lngLine & " of " & strTable & " was not imported (" & Err.Description & ")"
        Resume RowCleanup
RowCleanup:
        ' (Error handler state is reset by Resume, so cleanup errors are ignored here)
        On Error Resume Next
        rst.CancelUpdate
        Err.Clear
NextRow:
        On Error GoTo 0
    Next lngLine
    rst.Close

    If lngFailed > 0 Then
        m_Errors = m_Errors + 1
        LogLine "  ERROR: " & lngFailed & " row(s) of " & strTable & " could not be imported."
    End If
    LogLine "  " & strTable & ": " & lngRows & " row(s)"

End Sub


Private Function IsCalculatedField(fld As Object) As Boolean
    Dim strExpression As String
    On Error Resume Next
    strExpression = fld.Properties("Expression").Value
    IsCalculatedField = (Len(strExpression) > 0)
    Err.Clear
End Function


Private Function UnescapeDataValue(ByVal strValue As String) As String
    strValue = Replace(strValue, "\\", Chr$(26))
    strValue = Replace(strValue, "\r\n", vbCrLf)
    strValue = Replace(strValue, "\r", vbCr)
    strValue = Replace(strValue, "\n", vbLf)
    strValue = Replace(strValue, "\t", vbTab)
    UnescapeDataValue = Replace(strValue, Chr$(26), "\")
End Function


Private Sub ImportAllTableDataMacros()
    Dim dFiles As Object
    Dim varKey As Variant
    Set dFiles = GetSourceFiles(m_Source & "tdmacros", "xml", False)
    If dFiles.Count = 0 Then Exit Sub
    LogLine "Importing table data macros..."
    For Each varKey In dFiles.Keys
        ImportTableDataMacro CStr(dFiles(varKey))
    Next varKey
End Sub


Private Sub ImportTableDataMacro(ByVal strFile As String)

    Dim strName As String

    strName = GetObjectNameFromFileName(strFile)
    On Error GoTo ErrHandler
    If Not TableExists(strName) Then
        LogWarning "Table " & strName & " does not exist. Data macros were not imported."
        Exit Sub
    End If
    LoadFromTextEx AC_TABLE_DATA_MACRO, strName, ReadTextFile(strFile)
    Exit Sub

ErrHandler:
    LogError "Error importing data macros for " & strName

End Sub


'=======================================================================================
' Relationships
'=======================================================================================

Private Sub ImportAllRelations()
    Dim dFiles As Object
    Dim varKey As Variant
    Set dFiles = GetSourceFiles(m_Source & "relations", "json", False)
    If dFiles.Count = 0 Then Exit Sub
    LogLine "Importing relationships..."
    For Each varKey In dFiles.Keys
        ImportRelation CStr(dFiles(varKey))
    Next varKey
End Sub


Private Sub ImportRelation(ByVal strFile As String)
    Dim dItems As Object
    Set dItems = ReadItems(strFile)
    If dItems.Count = 0 Then
        LogError "Unable to read " & strFile
    Else
        CreateRelation dItems
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CreateRelation
' Purpose   : Create (or replace) a relationship from its dictionary definition.
'---------------------------------------------------------------------------------------
'
Private Sub CreateRelation(ByVal dItems As Object)

    Dim rel As Object
    Dim fld As Object
    Dim varField As Variant
    Dim strName As String
    Dim strTable As String
    Dim strForeign As String

    strName = GetText(dItems, "Name")
    strTable = GetText(dItems, "Table")
    strForeign = GetText(dItems, "ForeignTable")
    Status "Creating relationship " & strName

    ' Remove the existing relationship, and the foreign key index that the XML schema
    ' created with the same name (it would collide with the new relationship).
    On Error Resume Next
    m_Dbs.Relations.Delete strName
    Err.Clear
    m_Dbs.TableDefs(strTable).Indexes.Delete strName
    Err.Clear
    m_Dbs.TableDefs(strForeign).Indexes.Delete strName
    Err.Clear

    On Error GoTo ErrHandler
    Set rel = m_Dbs.CreateRelation(strName, strTable, strForeign, CLng(Nz(GetValue(dItems, "Attributes"), 0)))
    If IsObject(GetValue(dItems, "Fields")) Then
        For Each varField In dItems("Fields")
            Set fld = rel.CreateField(CStr(GetValue(varField, "Name")))
            fld.ForeignName = CStr(GetValue(varField, "ForeignName"))
            rel.Fields.Append fld
        Next varField
    End If
    m_Dbs.Relations.Append rel
    Exit Sub

ErrHandler:
    LogError "Error creating relationship " & strName & " (" & strTable & " -> " & strForeign & ")"

End Sub


' Remove (and return) the relationships that involve a table, so it can be replaced.
Private Function StageRelations(ByVal strTable As String) As Collection

    Dim colRelations As Collection
    Dim rel As Object
    Dim fld As Object
    Dim dRel As Object
    Dim dField As Object
    Dim colFields As Collection
    Dim varItem As Variant

    Set colRelations = New Collection
    On Error Resume Next
    For Each rel In m_Dbs.Relations
        If StrComp(rel.Table, strTable, vbTextCompare) = 0 Or StrComp(rel.ForeignTable, strTable, vbTextCompare) = 0 Then
            Set dRel = NewDict
            dRel.Add "Name", rel.Name
            dRel.Add "Attributes", rel.Attributes
            dRel.Add "Table", rel.Table
            dRel.Add "ForeignTable", rel.ForeignTable
            Set colFields = New Collection
            For Each fld In rel.Fields
                Set dField = NewDict
                dField.Add "Name", fld.Name
                dField.Add "ForeignName", fld.ForeignName
                colFields.Add dField
            Next fld
            dRel.Add "Fields", colFields
            colRelations.Add dRel
        End If
    Next rel
    For Each varItem In colRelations
        m_Dbs.Relations.Delete CStr(varItem("Name"))
    Next varItem
    Err.Clear
    Set StageRelations = colRelations

End Function


Private Sub RestoreRelations(colRelations As Collection)
    Dim varItem As Variant
    For Each varItem In colRelations
        CreateRelation varItem
    Next varItem
End Sub


'=======================================================================================
' Shared images, themes and specifications
'=======================================================================================

Private Sub EnsureResourcesTable()

    Dim strTemp As String
    Dim bteImage() As Byte

    If TableExists("MSysResources") Then Exit Sub
    On Error GoTo ErrHandler
    strTemp = GetTempFile(".png")
    bteImage = Base64ToBytes(PNG_1X1)
    WriteBinaryFile strTemp, bteImage
    m_App.CurrentProject.AddSharedImage "zzVCS_Temp_Image", strTemp
    m_Dbs.TableDefs.Refresh
    m_Dbs.Execute "DELETE FROM MSysResources WHERE [Name]='zzVCS_Temp_Image'"
    DeleteFile strTemp
    Exit Sub

ErrHandler:
    LogError "Unable to create the MSysResources table"
    DeleteFile strTemp

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportThemes
' Purpose   : Import themes\*.thmx (or a single file). Returns True if any theme was
'           : imported.
'---------------------------------------------------------------------------------------
'
Private Function ImportThemes(Optional ByVal strSingleFile As String) As Boolean

    Dim dFiles As Object
    Dim varKey As Variant
    Dim objFolder As Object

    If Len(strSingleFile) > 0 Then
        Set dFiles = NewDict
        dFiles.Add GetObjectNameFromFileName(strSingleFile), strSingleFile
    Else
        Set dFiles = GetSourceFiles(m_Source & "themes", "thmx", False)
        If m_FSO.FolderExists(m_Source & "themes") Then
            For Each objFolder In m_FSO.GetFolder(m_Source & "themes").SubFolders
                LogWarning "Extracted theme folders are not supported (" & objFolder.Name & "). Export with ExtractThemeFiles = false."
            Next objFolder
        End If
    End If
    If dFiles.Count = 0 Then Exit Function

    LogLine "Importing themes..."
    EnsureResourcesTable
    For Each varKey In dFiles.Keys
        If ImportTheme(CStr(varKey), CStr(dFiles(varKey))) Then ImportThemes = True
    Next varKey

End Function


Private Function ImportTheme(ByVal strName As String, ByVal strFile As String) As Boolean

    Dim rst As Object
    Dim rstFiles As Object

    Status "Importing theme " & strName
    On Error GoTo ErrHandler
    Set rst = m_Dbs.OpenRecordset("SELECT * FROM MSysResources WHERE [Type]='thmx' AND [Name]=" & SqlText(strName), DB_OPEN_DYNASET)
    If rst.EOF Then
        rst.AddNew
        rst.Fields("Name").Value = strName
        rst.Fields("Extension").Value = "thmx"
        rst.Fields("Type").Value = "thmx"
    Else
        rst.Edit
    End If
    Set rstFiles = rst.Fields("Data").Value
    Do While Not rstFiles.EOF
        rstFiles.Delete
        rstFiles.MoveNext
    Loop
    rstFiles.AddNew
    rstFiles.Fields("FileData").LoadFromFile strFile
    rstFiles.Update
    rstFiles.Close
    rst.Update
    rst.Close
    ImportTheme = True
    Exit Function

ErrHandler:
    LogError "Error importing theme " & strName

End Function


'---------------------------------------------------------------------------------------
' Procedure : ImportSharedImages
' Purpose   : Add images\*.json + image files to the shared image gallery.
'---------------------------------------------------------------------------------------
'
Private Sub ImportSharedImages()
    Dim dFiles As Object
    Dim varKey As Variant
    Set dFiles = GetSourceFiles(m_Source & "images", "json", False)
    If dFiles.Count = 0 Then Exit Sub
    LogLine "Importing shared images..."
    For Each varKey In dFiles.Keys
        ImportSharedImage CStr(dFiles(varKey))
    Next varKey
End Sub


Private Sub ImportSharedImage(ByVal strJsonFile As String)

    Dim dItems As Object
    Dim strName As String
    Dim strFileName As String
    Dim strImage As String
    Dim strTempFolder As String
    Dim strTempFile As String
    Dim objFile As Object

    Set dItems = ReadItems(strJsonFile)
    strName = GetText(dItems, "Name", GetObjectNameFromFileName(strJsonFile))
    strFileName = GetText(dItems, "FileName")
    Status "Importing image " & strName
    On Error GoTo ErrHandler

    ' Find the image file (same base name as the json file)
    strImage = SwapExtension(strJsonFile, m_FSO.GetExtensionName(strFileName))
    If Not m_FSO.FileExists(strImage) Or Len(strFileName) = 0 Then
        strImage = vbNullString
        For Each objFile In m_FSO.GetFolder(m_FSO.GetParentFolderName(strJsonFile)).Files
            If StrComp(m_FSO.GetBaseName(objFile.Name), m_FSO.GetBaseName(strJsonFile), vbTextCompare) = 0 _
                And LCase$(m_FSO.GetExtensionName(objFile.Name)) <> "json" Then
                strImage = objFile.Path
                Exit For
            End If
        Next objFile
    End If
    If Len(strImage) = 0 Then
        LogError "Image file not found for shared image " & strName
        Exit Sub
    End If
    If Len(strFileName) = 0 Then strFileName = m_FSO.GetFileName(strImage)

    ' Copy with the original file name (stored with the image)
    strTempFolder = m_FSO.GetSpecialFolder(2).Path & "\VCS_" & m_FSO.GetBaseName(m_FSO.GetTempName)
    m_FSO.CreateFolder strTempFolder
    strTempFile = strTempFolder & "\" & strFileName
    m_FSO.CopyFile strImage, strTempFile, True

    If TableExists("MSysResources") Then
        m_Dbs.Execute "DELETE FROM MSysResources WHERE [Type]='img' AND [Name]=" & SqlText(strName)
    End If
    m_App.CurrentProject.AddSharedImage strName, strTempFile
    m_FSO.DeleteFolder strTempFolder, True
    Exit Sub

ErrHandler:
    LogError "Error importing shared image " & strName
    On Error GoTo -1
    On Error Resume Next
    If Len(strTempFolder) > 0 Then m_FSO.DeleteFolder strTempFolder, True
    Err.Clear

End Sub


Private Sub ImportSavedSpecs()
    Dim dFiles As Object
    Dim varKey As Variant
    Set dFiles = GetSourceFiles(m_Source & "savedspecs", "json", False)
    For Each varKey In dFiles.Keys
        ImportSavedSpec CStr(dFiles(varKey))
    Next varKey
End Sub


Private Sub ImportSavedSpec(ByVal strFile As String)

    Dim dItems As Object
    Dim colSpecs As Object
    Dim objSpec As Object
    Dim strName As String
    Dim lngCnt As Long

    Set dItems = ReadItems(strFile)
    strName = GetText(dItems, "Name", GetObjectNameFromFileName(strFile))
    On Error GoTo ErrHandler

    Set colSpecs = m_App.CurrentProject.ImportExportSpecifications
    For lngCnt = colSpecs.Count - 1 To 0 Step -1
        If StrComp(colSpecs(lngCnt).Name, strName, vbTextCompare) = 0 Then colSpecs(lngCnt).Delete
    Next lngCnt
    Set objSpec = colSpecs.Add(strName, GetText(dItems, "XML"))
    If Len(GetText(dItems, "Description")) > 0 Then objSpec.Description = GetText(dItems, "Description")
    Exit Sub

ErrHandler:
    LogError "Error importing saved specification " & strName

End Sub


Private Sub ImportImexSpecs()
    Dim dFiles As Object
    Dim varKey As Variant
    Set dFiles = GetSourceFiles(m_Source & "imexspecs", "json", False)
    For Each varKey In dFiles.Keys
        ImportImexSpec CStr(dFiles(varKey))
    Next varKey
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportImexSpec
' Purpose   : Import/export specifications stored in MSysIMEXSpecs/MSysIMEXColumns.
'---------------------------------------------------------------------------------------
'
Private Sub ImportImexSpec(ByVal strFile As String)

    Dim dItems As Object
    Dim dCols As Object
    Dim dCol As Object
    Dim rst As Object
    Dim fld As Object
    Dim varKey As Variant
    Dim strName As String
    Dim lngSpecID As Long

    Set dItems = ReadItems(strFile)
    strName = GetText(dItems, "SpecName", GetObjectNameFromFileName(strFile))
    On Error GoTo ErrHandler

    If Not TableExists("MSysIMEXSpecs") Then
        m_App.SysCmd 555    ' Creates the MSysIMEX tables
        m_Dbs.TableDefs.Refresh
    End If

    ' Replace any existing specification with this name
    Set rst = m_Dbs.OpenRecordset("SELECT SpecID FROM MSysIMEXSpecs WHERE SpecName=" & SqlText(strName), DB_OPEN_DYNASET)
    Do While Not rst.EOF
        m_Dbs.Execute "DELETE FROM MSysIMEXColumns WHERE SpecID=" & rst.Fields("SpecID").Value
        rst.Delete
        rst.MoveNext
    Loop
    rst.Close

    Set rst = m_Dbs.OpenRecordset("MSysIMEXSpecs", DB_OPEN_DYNASET)
    rst.AddNew
    For Each fld In rst.Fields
        If fld.Name <> "SpecID" And dItems.Exists(fld.Name) Then
            If Not IsObject(dItems(fld.Name)) Then fld.Value = dItems(fld.Name)
        End If
    Next fld
    rst.Update
    rst.Bookmark = rst.LastModified
    lngSpecID = rst.Fields("SpecID").Value
    rst.Close

    Set dCols = GetDict(dItems, "Columns")
    If Not dCols Is Nothing Then
        Set rst = m_Dbs.OpenRecordset("MSysIMEXColumns", DB_OPEN_DYNASET)
        For Each varKey In dCols.Keys
            Set dCol = GetDict(dCols, CStr(varKey))
            rst.AddNew
            rst.Fields("SpecID").Value = lngSpecID
            rst.Fields("FieldName").Value = varKey
            For Each fld In rst.Fields
                If fld.Name <> "SpecID" And fld.Name <> "FieldName" Then
                    If Not dCol Is Nothing Then
                        If dCol.Exists(fld.Name) Then fld.Value = dCol(fld.Name)
                    End If
                End If
            Next fld
            rst.Update
        Next varKey
        rst.Close
    End If
    Exit Sub

ErrHandler:
    LogError "Error importing import/export specification " & strName

End Sub


'=======================================================================================
' Source file helpers
'=======================================================================================

'---------------------------------------------------------------------------------------
' Procedure : GetSourceFiles
' Purpose   : Return a dictionary of object name -> file path for the given extensions
'           : (in order of preference, separated by "|"). When the same object exists
'           : in more than one subfolder, the most recently modified file is used.
'---------------------------------------------------------------------------------------
'
Private Function GetSourceFiles(ByVal strFolder As String, ByVal strExtensions As String, ByVal blnRecursive As Boolean) As Object
    Set GetSourceFiles = NewDict
    If m_FSO.FolderExists(strFolder) Then
        CollectFiles m_FSO.GetFolder(strFolder), "|" & LCase$(strExtensions) & "|", blnRecursive, GetSourceFiles
    End If
End Function


Private Sub CollectFiles(objFolder As Object, ByVal strExtensions As String, ByVal blnRecursive As Boolean, dFiles As Object)

    Dim objFile As Object
    Dim objSub As Object
    Dim strExt As String
    Dim strName As String
    Dim strExisting As String
    Dim lngRankNew As Long
    Dim lngRankOld As Long

    For Each objFile In objFolder.Files
        strExt = LCase$(m_FSO.GetExtensionName(objFile.Name))
        lngRankNew = InStr(1, strExtensions, "|" & strExt & "|")
        If lngRankNew > 0 Then
            strName = GetObjectNameFromFileName(objFile.Path)
            If dFiles.Exists(strName) Then
                strExisting = dFiles(strName)
                lngRankOld = InStr(1, strExtensions, "|" & LCase$(m_FSO.GetExtensionName(strExisting)) & "|")
                If lngRankNew < lngRankOld Then
                    dFiles(strName) = objFile.Path
                ElseIf lngRankNew = lngRankOld Then
                    LogWarning "Duplicate source files for " & strName & ": " & Mid$(strExisting, Len(m_Source) + 1) & _
                        " and " & Mid$(objFile.Path, Len(m_Source) + 1) & ". Using the most recently modified file."
                    If objFile.DateLastModified > m_FSO.GetFile(strExisting).DateLastModified Then dFiles(strName) = objFile.Path
                End If
            Else
                dFiles.Add strName, objFile.Path
            End If
        End If
    Next objFile

    If blnRecursive Then
        For Each objSub In objFolder.SubFolders
            CollectFiles objSub, strExtensions, True, dFiles
        Next objSub
    End If

End Sub


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


' Object name from a source file name (reverses GetSafeFileName)
Private Function GetObjectNameFromFileName(ByVal strFile As String) As String
    Dim strName As String
    strName = m_FSO.GetBaseName(strFile)
    strName = Replace(strName, "%3C", "<", , , vbTextCompare)
    strName = Replace(strName, "%3E", ">", , , vbTextCompare)
    strName = Replace(strName, "%3A", ":", , , vbTextCompare)
    strName = Replace(strName, "%22", """", , , vbTextCompare)
    strName = Replace(strName, "%2F", "/", , , vbTextCompare)
    strName = Replace(strName, "%5C", "\", , , vbTextCompare)
    strName = Replace(strName, "%7C", "|", , , vbTextCompare)
    strName = Replace(strName, "%3F", "?", , , vbTextCompare)
    strName = Replace(strName, "%2A", "*", , , vbTextCompare)
    strName = Replace(strName, "%25", "%", , , vbTextCompare)   ' This must be done last.
    GetObjectNameFromFileName = strName
End Function


Private Function SwapExtension(ByVal strFile As String, ByVal strExtension As String) As String
    SwapExtension = m_FSO.GetParentFolderName(strFile) & "\" & m_FSO.GetBaseName(strFile) & "." & strExtension
End Function


' Expand a "rel:" path relative to the folder of the target database
Private Function ExpandRelativePath(ByVal strPath As String) As String
    If Left$(strPath, 4) = "rel:" Then
        strPath = Mid$(strPath, 5)
        If Len(strPath) = 0 Then
            ExpandRelativePath = m_App.CurrentProject.Path
        Else
            ExpandRelativePath = m_FSO.BuildPath(m_App.CurrentProject.Path, strPath)
        End If
    Else
        ExpandRelativePath = strPath
    End If
End Function


Private Function SqlText(ByVal strValue As String) As String
    SqlText = "'" & Replace(strValue, "'", "''") & "'"
End Function


'=======================================================================================
' File and text helpers
'=======================================================================================

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
    If Len(strFile) > 0 Then
        If m_FSO.FileExists(strFile) Then m_FSO.DeleteFile strFile, True
    End If
    Err.Clear
End Sub


Private Function NormalizeText(ByVal strText As String) As String
    strText = Replace(strText, vbCrLf, vbLf)
    strText = Replace(strText, vbCr, vbLf)
    strText = Replace(strText, vbLf, vbCrLf)
    If Right$(strText, 2) <> vbCrLf Then strText = strText & vbCrLf
    NormalizeText = strText
End Function


' Read a text file (UTF-8 by default). Line endings are normalized to CRLF and the
' byte order mark is removed.
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
    If Len(strText) > 0 Then ReadTextFile = NormalizeText(strText)

End Function


Private Sub WriteTextFile(ByVal strFile As String, ByVal strText As String, Optional ByVal strCharset As String = "utf-8")
    EnsureFolder m_FSO.GetParentFolderName(strFile)
    With CreateObject("ADODB.Stream")
        .Type = AD_TYPE_TEXT
        .Charset = strCharset
        .Open
        .WriteText NormalizeText(strText)
        .SaveToFile strFile, AD_SAVE_CREATE_OVERWRITE
        .Close
    End With
End Sub


Private Function ReadBinaryFile(ByVal strFile As String, Optional ByVal lngMaxBytes As Long = -1) As Byte()

    Dim bteEmpty() As Byte

    ReDim bteEmpty(0 To 0)
    ReadBinaryFile = bteEmpty
    If Not m_FSO.FileExists(strFile) Then Exit Function
    With CreateObject("ADODB.Stream")
        .Type = AD_TYPE_BINARY
        .Open
        .LoadFromFile strFile
        If .Size > 0 Then ReadBinaryFile = .Read(lngMaxBytes)
        .Close
    End With

End Function


Private Sub WriteBinaryFile(ByVal strFile As String, bteData() As Byte)
    With CreateObject("ADODB.Stream")
        .Type = AD_TYPE_BINARY
        .Open
        .Write bteData
        .SaveToFile strFile, AD_SAVE_CREATE_OVERWRITE
        .Close
    End With
End Sub


Private Function Base64ToBytes(ByVal strBase64 As String) As Byte()
    With CreateObject("MSXML2.DOMDocument.6.0").createElement("b64")
        .DataType = "bin.base64"
        .Text = strBase64
        Base64ToBytes = .nodeTypedValue
    End With
End Function


' Convert an ISO 8601 UTC time ("2023-11-17T19:35:41.000Z") to a local date/time
Private Function FromIsoUtc(ByVal strIso As String) As Variant

    Dim dteUtc As Date

    On Error GoTo ErrHandler
    dteUtc = DateSerial(CInt(Mid$(strIso, 1, 4)), CInt(Mid$(strIso, 6, 2)), CInt(Mid$(strIso, 9, 2))) + _
        TimeSerial(CInt(Mid$(strIso, 12, 2)), CInt(Mid$(strIso, 15, 2)), CInt(Mid$(strIso, 18, 2)))
    FromIsoUtc = dteUtc
    On Error Resume Next
    With CreateObject("WbemScripting.SWbemDateTime")
        .SetVarDate dteUtc, False
        FromIsoUtc = .GetVarDate(True)
    End With
    Err.Clear
    Exit Function

ErrHandler:
    FromIsoUtc = strIso
    Err.Clear

End Function


'=======================================================================================
' Logging
'=======================================================================================

Private Sub LogLine(ByVal strText As String)
    Debug.Print strText
    If Not m_Log Is Nothing Then m_Log.Add strText
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


Private Sub Status(ByVal strText As String)
    On Error Resume Next
    SysCmd acSysCmdSetStatus, Left$(strText, 200)
    DoEvents
    Err.Clear
End Sub


'=======================================================================================
' Dictionaries and JSON
'=======================================================================================

Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
    NewDict.CompareMode = vbTextCompare
End Function


' Child dictionary, or Nothing
Private Function GetDict(ByVal varParent As Variant, ByVal strKey As String) As Object
    If Not IsObject(varParent) Then Exit Function
    If varParent Is Nothing Then Exit Function
    If TypeName(varParent) <> "Dictionary" Then Exit Function
    If Not varParent.Exists(strKey) Then Exit Function
    If IsObject(varParent(strKey)) Then
        If TypeName(varParent(strKey)) = "Dictionary" Then Set GetDict = varParent(strKey)
    End If
End Function


' Value of a dictionary key (objects are returned as objects), or Empty if missing
Private Function GetValue(ByVal varParent As Variant, ByVal strKey As String) As Variant
    If Not IsObject(varParent) Then Exit Function
    If varParent Is Nothing Then Exit Function
    If TypeName(varParent) <> "Dictionary" Then Exit Function
    If Not varParent.Exists(strKey) Then Exit Function
    If IsObject(varParent(strKey)) Then
        Set GetValue = varParent(strKey)
    Else
        GetValue = varParent(strKey)
    End If
End Function


' Text value of a dictionary key, or the default when it is missing, Null or empty
Private Function GetText(ByVal varParent As Variant, ByVal strKey As String, Optional ByVal strDefault As String) As String
    GetText = strDefault
    If Not IsObject(varParent) Then Exit Function
    If varParent Is Nothing Then Exit Function
    If TypeName(varParent) <> "Dictionary" Then Exit Function
    If Not varParent.Exists(strKey) Then Exit Function
    If IsObject(varParent(strKey)) Then Exit Function
    If IsNull(varParent(strKey)) Then Exit Function
    If Len(CStr(varParent(strKey))) > 0 Then GetText = CStr(varParent(strKey))
End Function


' The "Items" section of a json source file (an empty dictionary if not available)
Private Function ReadItems(ByVal strFile As String) As Object
    Dim dFile As Object
    Set dFile = ReadJsonFile(strFile)
    Set ReadItems = GetDict(dFile, "Items")
    If ReadItems Is Nothing Then Set ReadItems = NewDict
End Function


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

Attribute VB_Name = "modABConstants"
Option Explicit

'===========================
' AutoBericht Workbook Schema
'===========================

Public Const SHEET_META As String = "Meta"
Public Const SHEET_CHAPTERS As String = "Chapters"
Public Const SHEET_ROWS As String = "Rows"
Public Const SHEET_PHOTOS As String = "Photos"
Public Const SHEET_LISTS As String = "Lists"
Public Const SHEET_EXPORT_LOG As String = "ExportLog"
Public Const SHEET_OVERRIDES_HISTORY As String = "OverridesHistory"

Public Const ROW_HEADER_ROW As Long = 1

Public Function HeaderMeta() As Variant
    HeaderMeta = Array("key", "value")
End Function

Public Function HeaderChapters() As Variant
    HeaderChapters = Array( _
        "chapterId", "parentId", "orderIndex", _
        "defaultTitle_de", "defaultTitle_fr", "defaultTitle_it", "defaultTitle_en", _
        "pageSize", "isActive")
End Function

Public Function HeaderRows() As Variant
    HeaderRows = Array( _
        "rowId", "chapterId", "titleOverride", _
        "masterFinding", "masterLevel1", "masterLevel2", "masterLevel3", "masterLevel4", _
        "overrideFinding", "useOverrideFinding", _
        "overrideLevel1", "overrideLevel2", "overrideLevel3", "overrideLevel4", _
        "useOverrideLevel1", "useOverrideLevel2", "useOverrideLevel3", "useOverrideLevel4", _
        "customerAnswer", "customerRemark", "customerPriority", _
        "selectedLevel", "includeFinding", "includeRecommendation", "overwriteMode", _
        "done", "notes", "lastEditedBy", "lastEditedAt")
End Function

Public Function HeaderPhotos() As Variant
    HeaderPhotos = Array( _
        "fileName", "filePath", "displayName", "notes", _
        "tagBericht", "tagSeminar", "tagTopic", _
        "preferredLocale", "capturedAt")
End Function

Public Function HeaderLists() As Variant
    HeaderLists = Array( _
        "listName", "value", "label_de", "label_fr", "label_it", "label_en", _
        "group", "sortOrder", "chapterId")
End Function

Public Function HeaderExportLog() As Variant
    HeaderExportLog = Array( _
        "exportTimestamp", "rowId", "renumberedId", _
        "includeFinding", "includeRecommendation", "selectedLevel", "notes")
End Function

Public Function HeaderOverridesHistory() As Variant
    HeaderOverridesHistory = Array( _
        "timestamp", "rowId", "fieldName", "oldValue", "newValue", "user")
End Function

Public Function SheetHeaders(sheetName As String) As Variant
    Select Case sheetName
        Case SHEET_META: SheetHeaders = HeaderMeta()
        Case SHEET_CHAPTERS: SheetHeaders = HeaderChapters()
        Case SHEET_ROWS: SheetHeaders = HeaderRows()
        Case SHEET_PHOTOS: SheetHeaders = HeaderPhotos()
        Case SHEET_LISTS: SheetHeaders = HeaderLists()
        Case SHEET_EXPORT_LOG: SheetHeaders = HeaderExportLog()
        Case SHEET_OVERRIDES_HISTORY: SheetHeaders = HeaderOverridesHistory()
        Case Else: SheetHeaders = Empty
    End Select
End Function

Public Function SheetList() As Variant
    SheetList = Array( _
        SHEET_META, SHEET_CHAPTERS, SHEET_ROWS, SHEET_PHOTOS, _
        SHEET_LISTS, SHEET_EXPORT_LOG, SHEET_OVERRIDES_HISTORY)
End Function

Public Function DefaultLocale() As String
    DefaultLocale = "de-CH"
End Function

Public Function OverwriteModes() As Variant
    OverwriteModes = Array("append", "replace")
End Function

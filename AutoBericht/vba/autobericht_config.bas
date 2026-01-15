Attribute VB_Name = "modAutoBerichtConfig"
Option Explicit

' === Global config ===
Public Const AB_DEFAULT_CHAPTER_IDS As String = "0,1,2,3,4,4.8,5,6,7,8,9,10,11,12,13,14"

' === Debug flags ===
Public Const AB_DEBUG_WORD_IMPORT As Boolean = True
Public Const AB_DEBUG_MARKDOWN As Boolean = True
Public Const AB_DEBUG_PPT_EXPORT As Boolean = True

' === Word styles (import + markdown) ===
Public Const AB_STYLE_BODY As String = "Normal"
Public Const AB_STYLE_SECTION As String = "Heading 2"
Public Const AB_STYLE_FINDING As String = "Heading 3"
Public Const AB_STYLE_TABLE As String = "Grid Table Light"
Public Const AB_STYLE_BULLET As String = "List Paragraph"
Public Const AB_STYLE_SKIP_HEADING2 As String = "Heading 2"
Public Const AB_STYLE_SKIP_HEADING3 As String = "Heading 3"
Public Const AB_STYLE_BOLD As String = ""
Public Const AB_STYLE_ITALIC As String = ""
Public Const AB_STYLE_BOLDITALIC As String = ""

' === Word prompts ===
Public Const AB_PROMPT_CHOOSE_CHAPTER_TITLE As String = "Choose chapter"
Public Const AB_PROMPT_CHAPTER_DEFAULT As String = "1"
Public Const AB_PROMPT_IMPORT_CHAPTER As String = "Import chapter (0, 1-14, 4.8):"
Public Const AB_PROMPT_MARKDOWN_CHAPTER As String = "Markdown for chapter (0, 1-14, 4.8):"

' === Word logo markers ===
Public Const AB_LOGO_MARKER_BIG As String = "LOGO_BIG$$"
Public Const AB_LOGO_MARKER_SMALL As String = "LOGO_SMALL$$"
Public Const AB_LOGO_HEIGHT_MAIN_CM As Double = 2#
Public Const AB_LOGO_HEIGHT_HEADER_CM As Double = 0.8#

' === Word text markers ===
Public Const AB_TEXT_MARKER_NAME As String = "NAME$$"
Public Const AB_TEXT_MARKER_COMPANY As String = "COMPANY$$"
Public Const AB_TEXT_MARKER_COMPANY_ID As String = "COMPANY_ID$$"
Public Const AB_TEXT_MARKER_AUTHOR As String = "AUTHOR$$"
Public Const AB_TEXT_MARKER_DATE As String = "DATE$$"
Public Const AB_TEXT_MARKER_MODERATOR As String = "MOD$$"
Public Const AB_TEXT_MARKER_CO_MODERATOR As String = "CO$$"

' === Word spider chart config ===
Public Const AB_SPIDER_MARKER As String = "SPIDER$$"
Public Const AB_SPIDER_SERIES_COMPANY As String = "Selbstbeurteilung"
Public Const AB_SPIDER_SERIES_CONSULTANT As String = "Beurteilung durch Suva"
Public Const AB_SPIDER_CHART_TYPE As Long = -4151 ' xlRadarMarkers
Public Const AB_SPIDER_AXIS_MIN As Double = 0
Public Const AB_SPIDER_AXIS_MAX As Double = 100
Public Const AB_SPIDER_SHOW_LEGEND As Boolean = True
Public Const AB_SPIDER_LEGEND_POS As Long = -4107 ' xlLegendPositionBottom
Public Const AB_SPIDER_PROMPT_WHEN_BOTH As Boolean = True
Public Const AB_SPIDER_PREFER_14 As Boolean = True

' === Word table config ===
Public Const AB_TABLE_COL1_WIDTH_PCT As Long = 35
Public Const AB_TABLE_COL2_WIDTH_PCT As Long = 58
Public Const AB_TABLE_COL3_WIDTH_PCT As Long = 7
Public Const AB_TABLE_HEADER_CHECKMARK As String = "✓"
Public Const AB_TABLE_HEADER_TITLE As String = "Systempunkte mit Verbesserungspotenzial"
Public Const AB_TABLE_HEADER_COL1 As String = "Ist-Zustand"
Public Const AB_TABLE_HEADER_COL2 As String = "Lösungsansätze"
Public Const AB_TABLE_HEADER_COL3 As String = "Prio"

Public Const AB_WD_FORMAT_DOCX As Long = 12

' === Shared file names ===
Public Const AB_SIDECAR_FILENAME As String = "project_sidecar.json"
Public Const AB_SIDECAR_DIALOG_TITLE As String = "Select project_sidecar.json"

' === PPT export config ===
Public Const AB_PPT_TEMPLATE_FOLDER As String = ""
Public Const AB_PPT_TEMPLATE As String = "Vorlage AutoBericht.pptx"
Public Const AB_PPT_OUTPUT_TRAINING_BASE As String = "Seminar_Slides"
Public Const AB_PPT_INCLUDE_SEMINAR_SLIDE As Boolean = True
Public Const AB_PPT_LAYOUT_CHAPTER As String = "chapterorange"
Public Const AB_PPT_LAYOUT_PICTURE As String = "picture"
Public Const AB_PPT_DEFAULT_LANG_SUFFIX As String = "d"
Public Const AB_PPT_SEMINAR_TITLE As String = "Seminar"
Public Const AB_PPT_TAG_ORDER As String = "Unterlassen|Dulden|Handeln|Vorbild|Iceberg|Pyramide|STOP|SOS|Verhindern|Audit|Risikobeurteilung|AVIVA"

' Training layout names (DE/FR) to allow manual swapping in one place.
Public Const AB_PPT_LAYOUT_SEMINAR_D As String = "seminar_d"
Public Const AB_PPT_LAYOUT_SEMINAR_F As String = "seminar_f"
Public Const AB_PPT_LAYOUT_UNTERLASSEN_D As String = "unterlassen_d"
Public Const AB_PPT_LAYOUT_UNTERLASSEN_F As String = "unterlassen_f"
Public Const AB_PPT_LAYOUT_DULDEN_D As String = "dulden_d"
Public Const AB_PPT_LAYOUT_DULDEN_F As String = "dulden_f"
Public Const AB_PPT_LAYOUT_HANDELN_D As String = "handeln_d"
Public Const AB_PPT_LAYOUT_HANDELN_F As String = "handeln_f"
Public Const AB_PPT_LAYOUT_VORBILD_D As String = "vorbild_d"
Public Const AB_PPT_LAYOUT_VORBILD_F As String = "vorbild_f"
Public Const AB_PPT_LAYOUT_VERHINDERN_D As String = "verhindern_d"
Public Const AB_PPT_LAYOUT_VERHINDERN_F As String = "verhindern_f"
Public Const AB_PPT_LAYOUT_AUDIT_D As String = "audit_d"
Public Const AB_PPT_LAYOUT_AUDIT_F As String = "audit_f"
Public Const AB_PPT_LAYOUT_RISIKOBEURTEILUNG_D As String = "risikobeurteilung_d"
Public Const AB_PPT_LAYOUT_RISIKOBEURTEILUNG_F As String = "risikobeurteilung_f"
Public Const AB_PPT_LAYOUT_AVIVA_D As String = "aviva_d"
Public Const AB_PPT_LAYOUT_AVIVA_F As String = "aviva_f"

' === PPT report export config (Bericht Besprechung) ===
Public Const AB_PPT_REPORT_TEMPLATE As String = "Vorlage AutoBericht.pptx"
Public Const AB_PPT_REPORT_OUTPUT_BASE As String = "Bericht_Besprechung"
Public Const AB_PPT_REPORT_INCLUDE_TITLE As Boolean = True
Public Const AB_PPT_REPORT_CH0_TITLE As String = "Management Summary"
Public Const AB_PPT_REPORT_FINDINGS_PER_SLIDE As Long = 3

' Report layout names (single language; all text injected)
Public Const AB_PPT_REPORT_LAYOUT_TITLE As String = "report_title"
Public Const AB_PPT_REPORT_LAYOUT_CHAPTER_SEPARATOR As String = "report_chapter_separator"
Public Const AB_PPT_REPORT_LAYOUT_CHAPTER_SCREENSHOT As String = "report_chapter_screenshot"
Public Const AB_PPT_REPORT_LAYOUT_SECTION_TEXT As String = "report_section_text"
Public Const AB_PPT_REPORT_LAYOUT_SECTION_PHOTO_3 As String = "report_section_photo_3"
Public Const AB_PPT_REPORT_LAYOUT_SECTION_PHOTO_6 As String = "report_section_photo_6"
Public Const AB_PPT_REPORT_LAYOUT_48_SEPARATOR As String = "report_section_48_separator"
Public Const AB_PPT_REPORT_LAYOUT_48_TEXT_PHOTO_3 As String = "report_section_48_text_photo_3"
Public Const AB_PPT_REPORT_LAYOUT_48_TEXT_PHOTO_6 As String = "report_section_48_text_photo_6"

' PPT paste constants
Public Const AB_PP_PASTE_ENHANCED_METAFILE As Long = 2

' PPT shape constants
Public Const AB_PP_PLACEHOLDER_PICTURE As Long = 18 ' fallback magic number
Public Const AB_MSO_PLACEHOLDER As Long = 14
Public Const AB_MSO_SHAPE_PICTURE As Long = 13

Public Function AB_NormalizeLangSuffix(ByVal langSuffix As String) As String
    Dim suffix As String
    suffix = LCase$(langSuffix)
    If suffix <> "d" And suffix <> "f" Then suffix = AB_PPT_DEFAULT_LANG_SUFFIX
    AB_NormalizeLangSuffix = suffix
End Function

Public Function AB_SeminarLayoutName(ByVal langSuffix As String) As String
    Dim suffix As String
    suffix = AB_NormalizeLangSuffix(langSuffix)
    If suffix = "f" Then
        AB_SeminarLayoutName = AB_PPT_LAYOUT_SEMINAR_F
    Else
        AB_SeminarLayoutName = AB_PPT_LAYOUT_SEMINAR_D
    End If
End Function

Public Function AB_PptTrainingLayoutForTag(ByVal tag As String, ByVal langSuffix As String) As String
    Dim suffix As String
    suffix = AB_NormalizeLangSuffix(langSuffix)
    Select Case LCase$(tag)
        Case "unterlassen"
            AB_PptTrainingLayoutForTag = IIf(suffix = "f", AB_PPT_LAYOUT_UNTERLASSEN_F, AB_PPT_LAYOUT_UNTERLASSEN_D)
        Case "dulden"
            AB_PptTrainingLayoutForTag = IIf(suffix = "f", AB_PPT_LAYOUT_DULDEN_F, AB_PPT_LAYOUT_DULDEN_D)
        Case "handeln"
            AB_PptTrainingLayoutForTag = IIf(suffix = "f", AB_PPT_LAYOUT_HANDELN_F, AB_PPT_LAYOUT_HANDELN_D)
        Case "vorbild"
            AB_PptTrainingLayoutForTag = IIf(suffix = "f", AB_PPT_LAYOUT_VORBILD_F, AB_PPT_LAYOUT_VORBILD_D)
        Case "verhindern"
            AB_PptTrainingLayoutForTag = IIf(suffix = "f", AB_PPT_LAYOUT_VERHINDERN_F, AB_PPT_LAYOUT_VERHINDERN_D)
        Case "audit"
            AB_PptTrainingLayoutForTag = IIf(suffix = "f", AB_PPT_LAYOUT_AUDIT_F, AB_PPT_LAYOUT_AUDIT_D)
        Case "risikobeurteilung"
            AB_PptTrainingLayoutForTag = IIf(suffix = "f", AB_PPT_LAYOUT_RISIKOBEURTEILUNG_F, AB_PPT_LAYOUT_RISIKOBEURTEILUNG_D)
        Case "aviva"
            AB_PptTrainingLayoutForTag = IIf(suffix = "f", AB_PPT_LAYOUT_AVIVA_F, AB_PPT_LAYOUT_AVIVA_D)
        Case Else
            AB_PptTrainingLayoutForTag = AB_PPT_LAYOUT_PICTURE
    End Select
End Function

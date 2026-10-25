Attribute VB_Name = "generator"
Option Explicit

' Named constants: keep API codes, wire values and physical dimensions here.

' Windows API and background window
Private Const WIN_METRIC_VIRTUAL_LEFT As Long = 76
Private Const WIN_METRIC_VIRTUAL_TOP As Long = 77
Private Const WIN_METRIC_VIRTUAL_WIDTH As Long = 78
Private Const WIN_METRIC_VIRTUAL_HEIGHT As Long = 79
Private Const BACKGROUND_SCREEN_OFFSET_PX As Long = 1024
Private Const BACKGROUND_WINDOW_WIDTH_PX As Long = 1280
Private Const BACKGROUND_WINDOW_HEIGHT_PX As Long = 900
Private Const BACKGROUND_WINDOW_POSITION_FLAGS As Long = 1108
Private Const BACKGROUND_RESET_DISTANCE_PX As Long = 1000
Private Const BACKGROUND_DEFAULT_POSITION_PX As Long = 100
Private Const BACKGROUND_MIN_WINDOW_SIZE_PX As Long = 300

' SolidWorks API values (preserved from the installed 2025 API)
Private Const SW_PREF_PDF_IN_COLOR As Long = 323
Private Const SW_PREF_PDF_EMBED_FONTS As Long = 324
Private Const SW_PREF_PDF_HIGH_QUALITY As Long = 325
Private Const SW_PREF_DIMENSION_STANDARD As Long = 13
Private Const SW_PREF_UNIT_SYSTEM As Long = 263
Private Const SW_PREF_LINEAR_UNITS As Long = 47
Private Const SW_PREF_LINEAR_PRECISION As Long = 24
Private Const SW_PREF_ANGULAR_PRECISION As Long = 28
Private Const SW_HIDE_AXES As Long = 4
Private Const SW_HIDE_PLANES As Long = 5
Private Const SW_HIDE_ORIGINS As Long = 6
Private Const SW_HIDE_TEMPORARY_AXES As Long = 7
Private Const SW_HIDE_COORDINATE_SYSTEMS As Long = 13
Private Const SW_HIDE_REFERENCE_POINTS As Long = 19
Private Const SW_HIDE_CURVES As Long = 195
Private Const SW_HIDE_SKETCHES As Long = 196
Private Const SW_HIDE_REFERENCE_POINTS_2 As Long = 219
Private Const SW_DOC_PART As Long = 1
Private Const SW_DOC_ASSEMBLY As Long = 2
Private Const SW_DOC_DRAWING As Long = 3
Private Const SW_OPEN_SILENT As Long = 1
Private Const SW_OPEN_SILENT_READ_ONLY As Long = 3
Private Const SW_SAVE_SILENT As Long = 1
Private Const SW_SAVE_CURRENT_VERSION As Long = 0
Private Const SW_EXPORT_PDF As Long = 1
Private Const SW_EXPORT_ALL_SHEETS As Long = 1
Private Const SW_CUSTOM_PAPER As Long = 12
Private Const SW_CUSTOM_SHEET_TEMPLATE As Long = 12
Private Const SW_REBUILD_USER_DECISION As Long = 0
Private Const SW_DISPLAY_HIDDEN_LINES_REMOVED As Long = 2
Private Const SW_NO_VIEW_ALIGNMENT As Long = 0
Private Const SW_CUSTOM_PROPERTY_TEXT As Long = 30
Private Const SW_PROPERTY_REPLACE_VALUE As Long = 2
Private Const SW_ANNOTATION_NOTE As Long = 6
Private Const SW_DRAFTING_STANDARD_GOST As Long = 6
Private Const SW_UNIT_SYSTEM_LEGACY_DEFAULT As Long = 0
Private Const SW_LENGTH_MILLIMETRES As Long = 0
Private Const SW_PREF_INPUT_DIMENSION As Long = 10
Private Const SW_TEXT_FORMAT_NOTE As Long = 0
Private Const SW_TEXT_FORMAT_DIMENSION As Long = 1
Private Const SW_BOM_TOP_LEVEL_ONLY As Long = 2
Private Const SW_BOM_ANCHOR_TOP_LEFT As Long = 1
Private Const SW_ARC_CENTER As Long = 1
Private Const SW_ARC_MAXIMUM As Long = 3
Private Const SW_MODEL_ITEMS_ALL_DIMENSIONS As Long = 8
Private Const SW_MODEL_ITEMS_MARKED_DIMENSIONS As Long = 32768
Private Const SW_MODEL_ITEMS_HOLE_PROFILES As Long = 65536
Private Const SW_MODEL_ITEMS_HOLE_LOCATIONS As Long = 131072
Private Const SW_SOLID_BODY As Long = 0
Private Const SW_VISIBLE_ENTITY_EDGE As Long = 1
Private Const SW_VIEW_FRONT As Long = 1
Private Const SW_VIEW_RIGHT As Long = 4
Private Const SW_VIEW_TOP As Long = 5
Private Const SW_VIEW_ISOMETRIC As Long = 7

' File modes, filenames and limits
Private Const FILE_FOR_READING As Long = 1
Private Const FILE_FOR_APPENDING As Long = 8
Private Const FILE_UNICODE As Long = -1
Private Const ATOMIC_REPLACE_WRITE_THROUGH As Long = 9
Private Const MAX_GENERATED_STEM_LENGTH As Long = 100
Private Const MAX_WINDOWS_FILENAME_LENGTH As Long = 255
Private Const ASCII_CONTROL_LIMIT As Long = 32
Private Const WINDOWS_DEVICE_SUPERSCRIPT_ONE As Long = &HB9
Private Const WINDOWS_DEVICE_SUPERSCRIPT_TWO As Long = &HB2
Private Const WINDOWS_DEVICE_SUPERSCRIPT_THREE As Long = &HB3
Private Const SETTINGS_RELATIVE_PATH As String = "\generator.ini"
Private Const BACKGROUND_STATE_RELATIVE_PATH As String = "\background-state.ini"
Private Const CHECKPOINT_RELATIVE_PATH As String = "\generation-state.tsv"
Private Const SHEET_FORMAT_RELATIVE_PATH As String = "\drawing-format.slddrt"
Private Const COVERAGE_RELATIVE_PATH As String = "\dimension-coverage.tsv"
Private Const SOURCE_BACKUP_INDEX_RELATIVE_PATH As String = "\source-backup.tsv"
Private Const RUN_LOG_RELATIVE_PATH As String = "\run.log"
Private Const MANIFEST_RELATIVE_PATH As String = "\manifest.tsv"
Private Const MODELS_RELATIVE_DIRECTORY As String = "\Models\"
Private Const DRAWINGS_RELATIVE_DIRECTORY As String = "\Drawings\"
Private Const PDF_RELATIVE_DIRECTORY As String = "\PDF\"
Private Const MODELS_RELATIVE_PATH As String = "\Models"
Private Const DRAWINGS_RELATIVE_PATH As String = "\Drawings"
Private Const PDF_RELATIVE_PATH As String = "\PDF"
Private Const DRAWING_EXTENSION As String = ".slddrw"
Private Const PDF_EXTENSION As String = ".pdf"
Private Const DRAWING_TEMPLATE_EXTENSION As String = ".drwdot"
Private Const ASSEMBLY_EXTENSION As String = "sldasm"
Private Const DRAWING_TEMPLATE_FILE_TYPE As String = "drwdot"
Private Const ATOMIC_TEMP_SUFFIX As String = ".new"
Private Const SNAPSHOT_PREFIX As String = "AD_"

' Checkpoint and progress protocol
Private Const STATE_PENDING As String = "PENDING"
Private Const STATE_RUNNING As String = "RUNNING"
Private Const STATE_DONE As String = "DONE"
Private Const STATE_FAILED As String = "FAILED"
Private Const STATE_COMPLETE As String = "COMPLETE"
Private Const STATE_PAUSED As String = "PAUSED"
Private Const CHECKPOINT_HEADER_TAG As String = "STATE"
Private Const CHECKPOINT_JOB_TAG As String = "JOB"
Private Const CHECKPOINT_VERSION As String = "2"
Private Const CHECKPOINT_HEADER_LAST_FIELD As Long = 7
Private Const CHECKPOINT_JOB_LAST_FIELD As Long = 14
Private Const HEADER_SOURCE_PATH As Long = 2
Private Const HEADER_SOURCE_STAMP As Long = 3
Private Const HEADER_SNAPSHOT_PATH As Long = 4
Private Const HEADER_SNAPSHOT_CONFIG As Long = 5
Private Const HEADER_PREFIX As Long = 6
Private Const HEADER_RENDER_SIGNATURE As Long = 7

' Defaults and formatting
Private Const DEFAULT_FONT As String = "GOST Type BU"
Private Const PAPER_A3 As String = "A3"
Private Const PAPER_A4 As String = "A4"
Private Const ORIENTATION_LANDSCAPE As String = "landscape"
Private Const ORIENTATION_PORTRAIT As String = "portrait"
Private Const DEFAULT_UP_AXIS As String = "+Y"
Private Const TIMESTAMP_FORMAT As String = "yyyymmdd_hhnnss"
Private Const FILE_STAMP_FORMAT As String = "yyyy-mm-dd hh:nn:ss"
Private Const JOB_NUMBER_FORMAT As String = "000"
Private Const GEOMETRY_KEY_FORMAT As String = "0.000000"
Private Const DIMENSION_NUMBER_FORMAT As String = "0.00"
Private Const DEFAULT_BATCH_SIZE_TEXT As String = "8"
Private Const DESCRIPTION_WRAP_CHARACTERS As Long = 28
Private Const MATERIAL_WRAP_CHARACTERS As Long = 38
Private Const LINEAR_DECIMAL_PLACES As Long = 2
Private Const ANGULAR_DECIMAL_PLACES As Long = 1

' Math and geometric tolerances
Private Const PI As Double = 3.14159265358979
Private Const MILLIMETRES_PER_METRE As Double = 1000
Private Const SEARCH_INITIAL_LIMIT As Double = 1E+30
Private Const MIN_PROJECTED_EXTENT_M As Double = 0.000001
Private Const CIRCLE_CLOSURE_TOLERANCE_M As Double = 0.0000001
Private Const FACE_NORMAL_ALIGNMENT_COSINE As Double = 0.999
Private Const MIN_DIMENSION_SPAN_M As Double = 0.00001
Private Const DIMENSION_MATCH_TOLERANCE_M As Double = 0.00001
Private Const UP_VECTOR_TOLERANCE As Double = 0.0000001
Private Const MIN_CIRCLE_RADIUS_M As Double = 0.00005
Private Const SCALE_REDUCTION_THRESHOLD As Double = 0.999
Private Const SCALE_COMPARISON_TOLERANCE As Double = 0.0000001
Private Const VIEW_FIT_SAFETY_FACTOR As Double = 0.95

' Sheet dimensions and reference layout (metres)
Private Const A4_SHORT_SIDE_M As Double = 0.21
Private Const A4_LONG_SIDE_AND_A3_SHORT_SIDE_M As Double = 0.297
Private Const A3_LONG_SIDE_M As Double = 0.42
Private Const SHEET_SIZE_TOLERANCE_M As Double = 0.0001
Private Const LAYOUT_HORIZONTAL_RESERVED_M As Double = 0.045
Private Const REFERENCE_USABLE_WIDTH_M As Double = 0.375
Private Const LAYOUT_VERTICAL_RESERVED_M As Double = 0.105
Private Const REFERENCE_USABLE_HEIGHT_M As Double = 0.192
Private Const LAYOUT_ORIGIN_X_M As Double = 0.03
Private Const LAYOUT_ORIGIN_Y_M As Double = 0.075
Private Const SMALL_SHEET_WIDTH_CLEARANCE_M As Double = 0.016
Private Const SMALL_SHEET_HEIGHT_CLEARANCE_M As Double = 0.01
Private Const ORTHOGRAPHIC_COLUMN_X_M As Double = 0.145
Private Const AUXILIARY_COLUMN_X_M As Double = 0.335
Private Const ORTHOGRAPHIC_ROW_Y_M As Double = 0.21
Private Const AUXILIARY_ROW_Y_M As Double = 0.115
Private Const ORTHOGRAPHIC_BOX_HEIGHT_M As Double = 0.09
Private Const BOTTOM_BOX_HEIGHT_M As Double = 0.05
Private Const SIDE_BOX_WIDTH_M As Double = 0.11
Private Const ISOMETRIC_BOX_HEIGHT_M As Double = 0.06
Private Const INITIAL_FLAT_VIEW_SCALE As Double = 0.1
Private Const DRAFT_MARKER_TOP_MARGIN_M As Double = 0.012
Private Const DRAFT_MARKER_TEXT_HEIGHT_M As Double = 0.002
Private Const FLAT_NOTICE_Y_M As Double = 0.065
Private Const AUXILIARY_CAPTION_X_M As Double = 0.28
Private Const AUXILIARY_CAPTION_Y_M As Double = 0.077
Private Const MAIN_BOX_WIDTH_M As Double = 0.21
Private Const AUXILIARY_BOX_WIDTH_M As Double = 0.115
Private Const DRAFT_MARKER_X_M As Double = 0.1
Private Const DETAIL_INITIAL_X_M As Double = 0.225
Private Const DETAIL_INITIAL_Y_M As Double = 0.17
Private Const DETAIL_HORIZONTAL_RESERVED_M As Double = 0.105
Private Const DETAIL_VERTICAL_RESERVED_M As Double = 0.175
Private Const DETAIL_CENTER_X_BIAS_M As Double = 0.035
Private Const DETAIL_CENTER_Y_BASE_M As Double = 0.095
Private Const DETAIL_CENTER_HEIGHT_REFERENCE_M As Double = 0.21

' Text and title block dimensions (metres)
Private Const CONTENT_LEFT_M As Double = 0.03
Private Const SECTION_HEADER_TOP_MARGIN_M As Double = 0.025
Private Const TEXT_HEIGHT_STANDARD_M As Double = 0.0035
Private Const TEXT_HEIGHT_SMALL_M As Double = 0.0025
Private Const TITLE_DESIGNATION_WIDTH_M As Double = 0.116
Private Const TITLE_TEXT_CELL_HEIGHT_M As Double = 0.013
Private Const TITLE_MAIN_CELL_WIDTH_M As Double = 0.066
Private Const TITLE_AUX_CODE_HEIGHT_M As Double = 0.012
Private Const TITLE_DESCRIPTION_HEIGHT_M As Double = 0.023
Private Const TITLE_ORGANIZATION_WIDTH_M As Double = 0.046
Private Const VERTICAL_NOTE_SINE_THRESHOLD As Double = 0.7
Private Const TITLE_MEDIUM_FONT_THRESHOLD_M As Double = 0.0036
Private Const TITLE_SMALL_FONT_THRESHOLD_M As Double = 0.0026
Private Const TITLE_LAST_FIT_PASS As Long = 3
Private Const TEXT_ESTIMATED_GLYPH_WIDTH_RATIO As Double = 0.55

' Technical note layout (metres)
Private Const TECHNICAL_TEXT_HORIZONTAL_RESERVED_M As Double = 0.06
Private Const TECHNICAL_TEXT_GLYPH_WIDTH_M As Double = 0.00175
Private Const TECHNICAL_FIRST_SHEET_Y_M As Double = 0.078
Private Const TECHNICAL_TEXT_HEIGHT_M As Double = 0.003
Private Const TECHNICAL_PAGE_RESERVED_HEIGHT_M As Double = 0.12
Private Const TECHNICAL_LINE_HEIGHT_M As Double = 0.0045
Private Const TECHNICAL_TEXT_TOP_MARGIN_M As Double = 0.04
Private Const TECHNICAL_HEADER_HEIGHT_M As Double = 0.004
Private Const TECHNICAL_FIRST_SHEET_MAX_LINES As Long = 3

' Native BOM layout (metres)
Private Const BOM_HORIZONTAL_RESERVED_M As Double = 0.06
Private Const BOM_VERTICAL_RESERVED_M As Double = 0.11
Private Const BOM_SOURCE_VIEW_Y_M As Double = 0.075
Private Const BOM_HIDDEN_VIEW_SCALE As Double = 0.01
Private Const BOM_TOP_MARGIN_M As Double = 0.04
Private Const BOM_POSITION_COLUMN_WIDTH_M As Double = 0.012
Private Const BOM_QUANTITY_COLUMN_WIDTH_M As Double = 0.02
Private Const BOM_FIXED_COLUMNS_WIDTH_M As Double = 0.032
Private Const BOM_DATA_ROW_HEIGHT_M As Double = 0.007
Private Const BOM_HEADER_ROW_HEIGHT_M As Double = 0.01
Private Const BOM_COMBINED_COLUMN_COUNT As Long = 3
Private Const BOM_SPLIT_COLUMN_COUNT As Long = 4
Private Const BOM_FIXED_COLUMN_COUNT As Long = 2

' Dimension label layout (metres)
Private Const DIAMETER_LABEL_OFFSET_M As Double = 0.012
Private Const DIAMETER_LABEL_PITCH_M As Double = 0.014
Private Const OVERALL_LABEL_OFFSET_M As Double = 0.009
Private Const DETAIL_OVERALL_LABEL_OFFSET_M As Double = 0.012
Private Const RADIUS_LABEL_OFFSET_M As Double = 0.008
Private Const RADIUS_LABEL_PITCH_M As Double = 0.009

' Coordinate dimension layout
Private Const COORDINATE_LAST_LANE As Long = 50
Private Const COORDINATE_BASE_OFFSET_M As Double = 0.016
Private Const COORDINATE_TOP_MARGIN_M As Double = 0.04
Private Const COORDINATE_LEFT_LIMIT_M As Double = 0.032
Private Const COORDINATE_LANE_PITCH_M As Double = 0.007
Private Const EMPTY_LANE_LIMIT As Double = 1E+20
Private Const COORDINATE_GLYPH_WIDTH_M As Double = 0.0018
Private Const COORDINATE_LABEL_PADDING_M As Double = 0.004
Private Const COORDINATE_RIGHT_MARGIN_M As Double = 0.015
Private Const COORDINATE_LABEL_GAP_M As Double = 0.002
Private Const COORDINATE_LANE_PENALTY_M As Double = 0.0001
Private Const COORDINATE_MIN_KEY_COUNT As Long = 3
Private Const COORDINATE_KEY_DECIMALS As Long = 5

' Supported drawing scale factors
Private Const SCALE_10_TO_1 As Double = 10#
Private Const SCALE_5_TO_1 As Double = 5#
Private Const SCALE_4_TO_1 As Double = 4#
Private Const SCALE_2_5_TO_1 As Double = 2.5
Private Const SCALE_2_TO_1 As Double = 2#
Private Const SCALE_1_TO_1 As Double = 1#
Private Const SCALE_1_TO_2 As Double = 0.5
Private Const SCALE_1_TO_2_5 As Double = 0.4
Private Const SCALE_1_TO_4 As Double = 0.25
Private Const SCALE_1_TO_5 As Double = 0.2
Private Const SCALE_1_TO_10 As Double = 0.1
Private Const SCALE_1_TO_20 As Double = 0.05
Private Const SCALE_1_TO_25 As Double = 0.04
Private Const SCALE_1_TO_40 As Double = 0.025
Private Const SCALE_1_TO_50 As Double = 0.02
Private Const SCALE_1_TO_100 As Double = 0.01
Private Const SCALE_1_TO_200 As Double = 0.005
Private Const SCALE_1_TO_400 As Double = 0.0025
Private Const SCALE_1_TO_500 As Double = 0.002
Private Const SCALE_1_TO_1000 As Double = 0.001

' Stable generator error offsets
Private Const ERR_SOURCE_NOT_OPEN As Long = 1
Private Const ERR_SOURCE_TYPE As Long = 2
Private Const ERR_SOURCE_NOT_SAVED As Long = 3
Private Const ERR_UNSAVED_SOURCE As Long = 4
Private Const ERR_DRAWING_SETTINGS As Long = 5
Private Const ERR_OUTPUT_DIRECTORY As Long = 6
Private Const ERR_OPEN_INPUT_OR_SNAPSHOT As Long = 7
Private Const ERR_APPLICATION_STATE As Long = 8
Private Const ERR_EMPTY_PLAN As Long = 9
Private Const ERR_DEPENDENCY_MISSING As Long = 10
Private Const ERR_DEPENDENCY_COLLISION As Long = 11
Private Const ERR_PACK_TARGETS As Long = 12
Private Const ERR_PACK_RESULT As Long = 13
Private Const ERR_PACK_STATUS As Long = 14
Private Const ERR_SNAPSHOT_MISSING As Long = 15
Private Const ERR_LIBRARY_COMPONENT As Long = 16
Private Const ERR_CONFIGURATION As Long = 17
Private Const ERR_OPEN_JOB_MODEL As Long = 18
Private Const ERR_VIEW_ACTIVATION As Long = 19
Private Const ERR_BODY_MISSING As Long = 20
Private Const ERR_NEW_DRAWING As Long = 21
Private Const ERR_DRAWING_TEMPLATE As Long = 22
Private Const ERR_MAIN_VIEW As Long = 23
Private Const ERR_SAVE_MODEL As Long = 24
Private Const ERR_SAVE_DRAWING As Long = 25
Private Const ERR_SAVE_PDF As Long = 26
Private Const ERR_REOPEN_DRAWING As Long = 27
Private Const ERR_UNSUPPORTED_SCALE As Long = 30
Private Const ERR_INSERT_NOTE As Long = 40
Private Const ERR_NATIVE_BOM As Long = 41
Private Const ERR_REMOVE_DIMENSION As Long = 50
Private Const ERR_DIMENSION_PARAMETER As Long = 51
Private Const ERR_FLAT_CONFIGURATION As Long = 63
Private Const ERR_FLAT_FACE As Long = 64
Private Const ERR_FLAT_ORIENTATION As Long = 65
Private Const ERR_CREATE_FLAT_VIEW As Long = 66
Private Const ERR_FLAT_EDGE As Long = 67
Private Const ERR_FLAT_MODEL_ACTIVATION As Long = 68
Private Const ERR_WRONG_ACTIVE_MODEL As Long = 69
Private Const ERR_GENERATION_DATA As Long = 70
Private Const ERR_CHECKPOINT_INVALID As Long = 71
Private Const ERR_UP_AXIS As Long = 73
Private Const ERR_DETAIL_SHEET As Long = 74
Private Const ERR_DIMENSION_LAYOUT As Long = 75

' Additional API and wire-format indices
Private Const WIN_SHOW_HIDE As Long = 0
Private Const WIN_INSERT_AFTER_DEFAULT As Long = 0
Private Const SW_WINDOW_NORMAL As Long = 0
Private Const SW_DIMENSION_INDEX_FIRST As Long = 0
Private Const SW_ANNOTATION_TEXT_FORMAT_FIRST As Long = 0
Private Const SW_USE_SHEET_SCALE As Long = 1
Private Const SW_USE_INDEPENDENT_SCALE As Long = 0
Private Const HEADER_TAG_FIELD As Long = 0
Private Const HEADER_VERSION_FIELD As Long = 1
Private Const COVERAGE_HEADER_PREFIX_LENGTH As Long = 8

' Geometry array layouts
Private Const OUTLINE_LEFT_INDEX As Long = 0
Private Const OUTLINE_BOTTOM_INDEX As Long = 1
Private Const OUTLINE_RIGHT_INDEX As Long = 2
Private Const OUTLINE_TOP_INDEX As Long = 3
Private Const BBOX_X_MIN_INDEX As Long = 0
Private Const BBOX_Y_MIN_INDEX As Long = 1
Private Const BBOX_Z_MIN_INDEX As Long = 2
Private Const BBOX_X_MAX_INDEX As Long = 3
Private Const BBOX_Y_MAX_INDEX As Long = 4
Private Const BBOX_Z_MAX_INDEX As Long = 5
Private Const CIRCLE_CENTER_X_INDEX As Long = 0
Private Const CIRCLE_CENTER_Y_INDEX As Long = 1
Private Const CIRCLE_CENTER_Z_INDEX As Long = 2
Private Const CIRCLE_AXIS_X_INDEX As Long = 3
Private Const CIRCLE_AXIS_Y_INDEX As Long = 4
Private Const CIRCLE_AXIS_Z_INDEX As Long = 5
Private Const CIRCLE_RADIUS_INDEX As Long = 6
Private Const MATRIX_XX_INDEX As Long = 0
Private Const MATRIX_XY_INDEX As Long = 1
Private Const MATRIX_XZ_INDEX As Long = 2
Private Const MATRIX_YX_INDEX As Long = 3
Private Const MATRIX_YY_INDEX As Long = 4
Private Const MATRIX_YZ_INDEX As Long = 5
Private Const MATRIX_ZX_INDEX As Long = 6
Private Const MATRIX_ZY_INDEX As Long = 7
Private Const MATRIX_ZZ_INDEX As Long = 8
Private Const MATRIX_SCALE_INDEX As Long = 12
Private Const MATRIX_LAST_INDEX As Long = 15

' Math and geometric tolerances
Private Const MIN_PROJECTED_UP_VECTOR_LENGTH As Double = 0.00001

' File modes, filenames and limits
Private Const RESERVED_DEVICE_NAME_LENGTH As Long = 4
Private Const RESERVED_DEVICE_PREFIX_LENGTH As Long = 3

' Files and serialized protocol values
Private Const OUTPUT_FOLDER_RELATIVE_PREFIX As String = "\DrawingPack_"
Private Const SOURCE_BACKUP_RELATIVE_DIRECTORY As String = "\SourceBackup\"
Private Const NAMED_VIEW_PREFIX As String = "AD_VIEW_"
Private Const NAMED_FLAT_VIEW_PREFIX As String = "AD_FLAT_"
Private Const RENDER_SIGNATURE_PREFIX As String = "20261025-v3-templates|PaperFormat="

' API array field indices
Private Const SHEET_WIDTH_PROPERTY As Long = 5
Private Const SHEET_HEIGHT_PROPERTY As Long = 6

Private Type WindowBounds
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

#If VBA7 Then
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal windowHandle As LongPtr, ByVal command As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal windowHandle As LongPtr) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal metric As Long) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal windowHandle As LongPtr, ByVal insertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal windowHandle As LongPtr, ByRef bounds As WindowBounds) As Long
#Else
Private Declare Function ShowWindow Lib "user32" (ByVal windowHandle As Long, ByVal command As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal windowHandle As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal metric As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal windowHandle As Long, ByVal insertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal flags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal windowHandle As Long, ByRef bounds As WindowBounds) As Long
#End If

' SolidWorks 2025 VBA. All geometry remains native and associative.
' Entry point: main. Source documents are not passed to Save3.
' Pack and Go may reserialize an assembly; exact source backups are retained first.
' A Pack and Go snapshot is made first; derived flat configurations are saved there.
#If VBA7 Then
Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal existingName As LongPtr, ByVal newName As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Function MoveFileExW Lib "kernel32" (ByVal existingName As Long, ByVal newName As Long, ByVal flags As Long) As Long
#End If
Private checkpointPath As String, snapshotPath As String, sourcePath As String, sourceStamp As String, snapshotConfig As String
Private attempts As Long, resumeRun As Boolean, renderSignature As String
Private sw As SldWorks.SldWorks, fs As Object, logFile As Object, manifest As Object
Private jobs As Object, quantities As Object, settings As Object
Private runDir As String, sourceDir As String, prefix As String, templatePath As String
Private paperFormat As String, paperOrientation As String, runtimeFormatPath As String
Private sheetWidth As Double, sheetHeight As Double, layoutSx As Double, layoutSy As Double
Private original As SldWorks.ModelDoc2, rootModel As SldWorks.ModelDoc2, currentDrawing As SldWorks.ModelDoc2
Private failures As Long, completed As Long, warningCount As Long, jobNumber As Long
Private inputDimPreference As Boolean, preferenceChanged As Boolean
Private oldEmbedFonts As Boolean, oldPdfColor As Boolean, oldPdfQuality As Boolean
Private viewCounter As Long
Private flatBody As SldWorks.Body2
Private progressPath As String, plannedJobs As Long
Private backgroundOwned As Boolean
Private backgroundLeft As Long, backgroundTop As Long, backgroundWidth As Long, backgroundHeight As Long



Private Sub PrepareDocumentWindow(ByVal document As SldWorks.ModelDoc2)
    Dim frame As SldWorks.Frame
    ' Drawing export needs a document window and graphics context, even with a hidden frame.
    document.Visible = True
    If backgroundOwned Then
        sw.UserControlBackground = True
        Set frame = sw.Frame: frame.KeepInvisible = True
        HideBackgroundFrame
    End If
End Sub

Private Sub HideBackgroundFrame()
    Dim frame As SldWorks.Frame
    If backgroundOwned = False Then
        Exit Sub
    End If
    On Error Resume Next
    Set frame = sw.Frame
    ' This SolidWorks build needs a visible graphics window for PDF export.
    ' Place only our dedicated frame beyond the entire virtual desktop, without activation.
    If FrameIsOffscreen() And IsWindowVisible(frame.GetHWndx64) <> 0 Then
        Exit Sub
    End If
    sw.FrameState = SW_WINDOW_NORMAL
    ' NOZORDER | NOACTIVATE | SHOWWINDOW | NOSENDCHANGING. SolidWorks FrameLeft/Top clamp to the desktop.
    SetWindowPos frame.GetHWndx64, WIN_INSERT_AFTER_DEFAULT, GetSystemMetrics(WIN_METRIC_VIRTUAL_LEFT) + GetSystemMetrics(WIN_METRIC_VIRTUAL_WIDTH) + BACKGROUND_SCREEN_OFFSET_PX, GetSystemMetrics(WIN_METRIC_VIRTUAL_TOP) + GetSystemMetrics(WIN_METRIC_VIRTUAL_HEIGHT) + BACKGROUND_SCREEN_OFFSET_PX, BACKGROUND_WINDOW_WIDTH_PX, BACKGROUND_WINDOW_HEIGHT_PX, BACKGROUND_WINDOW_POSITION_FLAGS
End Sub

Private Function FrameIsOffscreen() As Boolean
    Dim bounds As WindowBounds, frame As SldWorks.Frame
    Set frame = sw.Frame
    If GetWindowRect(frame.GetHWndx64, bounds) = 0 Then
        Exit Function
    End If
    FrameIsOffscreen = bounds.Left >= GetSystemMetrics(WIN_METRIC_VIRTUAL_LEFT) + GetSystemMetrics(WIN_METRIC_VIRTUAL_WIDTH) And bounds.Top >= GetSystemMetrics(WIN_METRIC_VIRTUAL_TOP) + GetSystemMetrics(WIN_METRIC_VIRTUAL_HEIGHT)
End Function

Private Sub FinishBackground()
    Dim state As Object, frame As SldWorks.Frame, bounds As WindowBounds
    If backgroundOwned = False Then
        Exit Sub
    End If
    On Error Resume Next
    Set frame = sw.Frame
    HideBackgroundFrame
    Set state = fs.CreateTextFile(sw.GetCurrentMacroPathFolder & BACKGROUND_STATE_RELATIVE_PATH, True, True)
    state.WriteLine "PID=" & sw.GetProcessID
    state.WriteLine "Hidden=" & CStr(IsWindowVisible(frame.GetHWndx64) = 0)
    state.WriteLine "Offscreen=" & CStr(FrameIsOffscreen())
    GetWindowRect frame.GetHWndx64, bounds
    state.WriteLine "FrameLeft=" & bounds.Left
    state.WriteLine "FrameTop=" & bounds.Top
    state.WriteLine "ApplicationVisibleFlag=" & CStr(sw.Visible)
    state.Close
    sw.CloseAllDocuments True
    ShowWindow frame.GetHWndx64, WIN_SHOW_HIDE
    ' Avoid persisting an inaccessible position for the next interactive launch.
    If backgroundLeft < -BACKGROUND_RESET_DISTANCE_PX Or backgroundTop < -BACKGROUND_RESET_DISTANCE_PX Then
        backgroundLeft = BACKGROUND_DEFAULT_POSITION_PX
        backgroundTop = BACKGROUND_DEFAULT_POSITION_PX
    End If
    If backgroundWidth < BACKGROUND_MIN_WINDOW_SIZE_PX Then
        backgroundWidth = BACKGROUND_WINDOW_WIDTH_PX
    End If
    If backgroundHeight < BACKGROUND_MIN_WINDOW_SIZE_PX Then
        backgroundHeight = BACKGROUND_WINDOW_HEIGHT_PX
    End If
    sw.FrameLeft = backgroundLeft: sw.FrameTop = backgroundTop
    sw.FrameWidth = backgroundWidth: sw.FrameHeight = backgroundHeight
    Set currentDrawing = Nothing: Set rootModel = Nothing: Set original = Nothing
    Set jobs = Nothing: Set quantities = Nothing: Set flatBody = Nothing
    ' Document writes and window restoration are complete before this marker.
    Set state = fs.OpenTextFile(sw.GetCurrentMacroPathFolder & BACKGROUND_STATE_RELATIVE_PATH, FILE_FOR_APPENDING, False, FILE_UNICODE)
    state.WriteLine "ReadyToExit=1"
    state.Close
    sw.ExitApp
    Set sw = Nothing
End Sub

Private Function ShouldRunJob(ByVal job As Object) As Boolean
    ShouldRunJob = Len(Setting("JobFilter", "")) = 0 Or InStr(1, CStr(job("path")) & "|" & CStr(job("body")), Setting("JobFilter", ""), vbTextCompare) > 0
End Function

Private Sub EnsureFolder(ByVal path As String)
    Dim parent As String
    If fs.FolderExists(path) Then
        Exit Sub
    End If
    parent = fs.GetParentFolderName(path)
    If Len(parent) = 0 Or parent = path Then
        Err.Raise vbObjectError + ERR_OUTPUT_DIRECTORY, , "Cannot create output folder: " & path
    End If
    EnsureFolder parent
    fs.CreateFolder path
End Sub

Private Function MakeSnapshot(ByVal doc As SldWorks.ModelDoc2) As String
    Dim extension As SldWorks.ModelDocExtension, pg As SldWorks.PackAndGo, names As Variant, results As Variant, n As Variant, seen As Object
    Dim backupList As Object, backupComplete As Boolean, savedError As Long, savedText As String
    On Error GoTo SnapshotFailed
    Log "PACK get object"
    Set extension = doc.Extension
    Set pg = extension.GetPackAndGo()

    pg.IncludeDrawings = False
    pg.IncludeSimulationResults = False
    pg.IncludeToolboxComponents = True
    pg.IncludeSuppressed = False
    pg.FlattenToSingleFolder = True
    Log "PACK get names"
    pg.GetDocumentNames names
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = 1
    For Each n In names
        If Not fs.FileExists(CStr(n)) Then
            Err.Raise vbObjectError + ERR_DEPENDENCY_MISSING, , "Missing reference: " & CStr(n)
        End If
        If seen.Exists(fs.GetFileName(CStr(n))) Then
            Err.Raise vbObjectError + ERR_DEPENDENCY_COLLISION, , "Duplicate dependency filenames. Use unique filenames before Pack and Go: " & CStr(n)
        End If
        seen.Add fs.GetFileName(CStr(n)), True
    Next n
    prefix = SNAPSHOT_PREFIX & Format$(Now, TIMESTAMP_FORMAT) & "_"
    EnsureFolder runDir & SOURCE_BACKUP_RELATIVE_DIRECTORY & prefix
    Set backupList = fs.CreateTextFile(runDir & SOURCE_BACKUP_INDEX_RELATIVE_PATH, True, True)
    backupList.WriteLine "Original" & vbTab & "Backup"
    For Each n In names
        fs.CopyFile CStr(n), runDir & SOURCE_BACKUP_RELATIVE_DIRECTORY & prefix & "\" & fs.GetFileName(CStr(n)), False
        backupList.WriteLine CStr(n) & vbTab & "SourceBackup\" & prefix & "\" & fs.GetFileName(CStr(n))
    Next n
    backupList.Close
    backupComplete = True
    Log "SOURCE exact file backups created before Pack and Go"
    Dim destinationNames() As String, fileIndex As Long
    ReDim destinationNames(LBound(names) To UBound(names))
    For fileIndex = LBound(names) To UBound(names)
        destinationNames(fileIndex) = runDir & MODELS_RELATIVE_DIRECTORY & prefix & fs.GetFileName(CStr(names(fileIndex)))
    Next fileIndex
    If pg.SetDocumentSaveToNames(destinationNames) = False Then
        Err.Raise vbObjectError + ERR_PACK_TARGETS, , "Pack and Go names rejected."
    End If
    Log "PACK save"
    results = doc.Extension.SavePackAndGo(pg)
    If IsEmpty(results) Then
        Err.Raise vbObjectError + ERR_PACK_RESULT, , "Pack and Go returned no status."
    End If
    For Each n In results
        If CLng(n) <> 0 Then
            Err.Raise vbObjectError + ERR_PACK_STATUS, , "Pack and Go error " & CStr(n)
        End If
    Next n
    ' SavePackAndGo can serialize source documents even when opened read-only.
    ' Restore their exact bytes from the pre-operation backup before committing a plan.
    RestoreSnapshotSources names
    sourceStamp = FileStamp(sourcePath)
    MakeSnapshot = runDir & MODELS_RELATIVE_DIRECTORY & prefix & fs.GetFileName(doc.GetPathName)
    If Not fs.FileExists(MakeSnapshot) Then
        Err.Raise vbObjectError + ERR_SNAPSHOT_MISSING, , "Snapshot assembly is absent."
    End If
    Log "SNAPSHOT " & MakeSnapshot
    Exit Function
SnapshotFailed:
    savedError = Err.Number: savedText = Err.Description
    On Error Resume Next
    If Not backupList Is Nothing Then
        backupList.Close
    End If
    If backupComplete Then
        RestoreSnapshotSources names
    End If
    On Error GoTo 0
    Err.Raise savedError, "MakeSnapshot", savedText
End Function

Private Sub RestoreSnapshotSources(ByVal names As Variant)
    Dim n As Variant, exactBackup As String
    For Each n In names
        exactBackup = runDir & SOURCE_BACKUP_RELATIVE_DIRECTORY & prefix & "\" & fs.GetFileName(CStr(n))
        If FileStamp(CStr(n)) <> FileStamp(exactBackup) Then
            fs.CopyFile exactBackup, CStr(n), True
        End If
    Next n
End Sub

Private Sub Inventory(ByVal asm As SldWorks.ModelDoc2)
    Dim comps As Variant, i As Long, c As Object, doc As Object, j As Object, k As String
    comps = AsAssembly(asm).GetComponents(False)
    If IsEmpty(comps) Then
        Exit Sub
    End If
    For i = LBound(comps) To UBound(comps)
        Set c = comps(i)
        If c.IsSuppressed = False Then
            Set doc = c.GetModelDoc2
            If doc Is Nothing Then
                Warn "Unresolved component: " & c.Name2
            Else
                If InStr(1, doc.GetPathName, runDir & MODELS_RELATIVE_DIRECTORY, vbTextCompare) <> 1 Then
                    Dim localCopy As String, openError As Long, openWarning As Long
                    localCopy = runDir & MODELS_RELATIVE_DIRECTORY & prefix & fs.GetFileName(doc.GetPathName)
                    Log "LIBRARY using copied model: " & localCopy
                    If fs.FileExists(localCopy) = False Then
                        Err.Raise vbObjectError + ERR_LIBRARY_COMPONENT, , "Local copy absent: " & localCopy
                    End If
                    Set doc = sw.OpenDoc6(localCopy, doc.GetType, 1, c.ReferencedConfiguration, openError, openWarning)
                    If doc Is Nothing Then
                        Err.Raise vbObjectError + ERR_LIBRARY_COMPONENT, , "Cannot open copied library component: " & localCopy
                    End If
                End If
                k = LCase$(doc.GetPathName & "|" & c.ReferencedConfiguration)
                If quantities.Exists(k) Then
                    Set j = quantities(k): j("qty") = CLng(j("qty")) + 1
                Else
                    Set j = CreateObject("Scripting.Dictionary")
                    j.Add "doc", doc: j.Add "cfg", c.ReferencedConfiguration: j.Add "qty", 1
                    quantities.Add k, j
                End If
            End If
        End If
    Next i
End Sub

Private Sub AddModelJobs(ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal qty As Long)
    Dim bodies As Variant, i As Long, b As Object, k As String
    If doc.GetConfigurationByName(cfg) Is Nothing Then
        Err.Raise vbObjectError + ERR_CONFIGURATION, , "Configuration absent: " & cfg
    End If
    doc.ShowConfiguration2 cfg
    If doc.ConfigurationManager.ActiveConfiguration.Name <> cfg Then
        Err.Raise vbObjectError + ERR_CONFIGURATION, , "Configuration could not be activated: " & cfg
    End If
    If doc.GetType = SW_DOC_ASSEMBLY Then
        AddJob doc, cfg, "", "", qty, False
        Exit Sub
    End If
    bodies = AsPart(doc).GetBodies2(SW_SOLID_BODY, True)
    If IsEmpty(bodies) Then
        Warn "No visible solid bodies: " & doc.GetPathName
        Exit Sub
    End If
    If UBound(bodies) = LBound(bodies) Then
        Set b = bodies(LBound(bodies))
        AddJob doc, cfg, b.Name, "", qty, b.IsSheetMetal
    Else
        For i = LBound(bodies) To UBound(bodies)
            Set b = bodies(i)
            AddJob doc, cfg, b.Name, CutListName(doc, b.Name), qty, b.IsSheetMetal, True
        Next i
    End If
End Sub

Private Sub AddJob(ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal bodyName As String, ByVal cutName As String, ByVal qty As Long, ByVal sheetMetal As Boolean, Optional ByVal multiBody As Boolean = False)
    Dim j As Object, k As String
    k = LCase$(doc.GetPathName & "|" & cfg & "|" & bodyName)
    If jobs.Exists(k) Then
        Exit Sub
    End If
    Set j = CreateObject("Scripting.Dictionary")
    j.Add "doc", doc: j.Add "cfg", cfg: j.Add "body", bodyName
    j.Add "path", doc.GetPathName: j.Add "type", doc.GetType
    j.Add "cut", cutName: j.Add "qty", qty: j.Add "sheet", sheetMetal: j.Add "multi", multiBody
    jobs.Add k, j
End Sub

Private Function CutListName(ByVal doc As SldWorks.ModelDoc2, ByVal bodyName As String) As String
    Dim f As Object, bf As Object, bodies As Variant, b As Variant
    Set f = doc.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "CutListFolder" Then
            Set bf = f.GetSpecificFeature2
            bodies = bf.GetBodies
            If Not IsEmpty(bodies) Then
                For Each b In bodies
                    If b.Name = bodyName Then
                        CutListName = f.Name
                        Exit Function
                    End If
                Next b
            End If
        End If
        Set f = f.GetNextFeature
    Loop
End Function

Private Function FindBody(ByVal doc As SldWorks.ModelDoc2, ByVal bodyName As String) As Object
    Dim a As Variant, b As Variant
    a = AsPart(doc).GetBodies2(SW_SOLID_BODY, True)
    If IsEmpty(a) Then
        Exit Function
    End If
    For Each b In a
        If b.Name = bodyName Then
            Set FindBody = b
            Exit Function
        End If
    Next b
End Function

Private Sub ProcessJob(ByVal j As Object)
    Dim doc As SldWorks.ModelDoc2, d As SldWorks.ModelDoc2, mainView As Object, bottomView As Object, sideView As Object, isoView As Object, flatView As SldWorks.View
    Dim b As Object, cfg As String, oldCfg As String, name As String, code As String, material As String
    Dim outName As String, best As String, dims As Long, rc As Long, er As Long, wr As Long, success As Boolean
    Dim sheet As Object, viewScaleValue As Double, flatScale As Double, bodies As Variant, oneBody(0) As Object
    Dim mark As String, bodyName As String, mass As String, isAssembly As Boolean, multi As Boolean, active As Object
    Dim firstSheetName As String
    On Error GoTo Failed
    outName = CStr(j("out"))
    Log "JOB " & outName
    Dim previousDrawing As SldWorks.ModelDoc2
    Set previousDrawing = sw.GetOpenDocumentByName(runDir & DRAWINGS_RELATIVE_DIRECTORY & outName & DRAWING_EXTENSION)
    If Not previousDrawing Is Nothing Then
        sw.CloseDoc previousDrawing.GetTitle
    End If
    Set previousDrawing = Nothing
    cfg = CStr(j("cfg")): bodyName = CStr(j("body"))
    Set doc = sw.OpenDoc6(CStr(j("path")), CLng(j("type")), SW_OPEN_SILENT, cfg, er, wr)
    If doc Is Nothing Then
        Err.Raise vbObjectError + ERR_OPEN_JOB_MODEL, , "Cannot open copied model for drawing: " & er
    End If
    PrepareDocumentWindow doc
    oldCfg = doc.ConfigurationManager.ActiveConfiguration.Name
    doc.ShowConfiguration2 cfg
    isAssembly = (doc.GetType = SW_DOC_ASSEMBLY)
    If isAssembly Then
        AsAssembly(doc).ResolveAllLightWeightComponents False
    End If
    name = PropertyValue(doc, cfg, "Наименование|Найменування|Description|SW-Title")
    If name = "" Then
        name = OriginalBase(doc.GetPathName)
    End If
    code = PropertyValue(doc, cfg, "Обозначение|Позначення|Designation|PartNumber|Number")
    If code = "" Then
        code = OriginalBase(doc.GetPathName)
    End If
    multi = False
    If Not isAssembly Then
        bodies = AsPart(doc).GetBodies2(SW_SOLID_BODY, True)
        multi = UBound(bodies) > LBound(bodies)
        Set b = FindBody(doc, bodyName)
        If b Is Nothing Then
            Err.Raise vbObjectError + ERR_BODY_MISSING, , "Body not found: " & bodyName
        End If
        If multi Then
            name = name & " / " & IIf(CStr(j("cut")) = "", bodyName, CStr(j("cut")))
            code = code & "." & Format$(BodyOrdinal(doc, bodyName), JOB_NUMBER_FORMAT)
        End If
    Else
        code = code & " СБ"
    End If
    material = PropertyValue(doc, cfg, "Материал|Матеріал|Material")
    If material = "" And Not isAssembly Then
        Dim matDb As String
        material = AsPart(doc).GetMaterialPropertyName2(cfg, matDb)
    End If
    mass = MassText(doc, b)
    JobStage j, "Creating drawing"
    Set d = sw.NewDocument(templatePath, 0, 0, 0)
    If d Is Nothing Then
        Err.Raise vbObjectError + ERR_NEW_DRAWING, , "Cannot create drawing."
    End If
    HideBackgroundFrame
    Set currentDrawing = d
    Set sheet = AsDrawing(d).GetCurrentSheet
    firstSheetName = sheet.GetName
    InitializePaper d
    FillTemplateProperties d, code, name, material, mass
    ConfigureDrawing d
    best = BestOrientation(doc, b)
    Set mainView = MakeView(d, doc, cfg, b, best, LX(ORTHOGRAPHIC_COLUMN_X_M), LY(ORTHOGRAPHIC_ROW_Y_M))
    If mainView Is Nothing Then
        Err.Raise vbObjectError + ERR_MAIN_VIEW, , "Main view failed."
    End If
    mainView.UseSheetScale = SW_USE_INDEPENDENT_SCALE
    mainView.ScaleDecimal = 1
    d.ForceRebuild3 False
    Dim box As Variant
    box = mainView.GetOutline
    If Not isAssembly Then
        If CBool(j("sheet")) And box(OUTLINE_TOP_INDEX) - box(1) > box(2) - box(0) Then
            mainView.Angle = PI / 2
        End If
    End If
    If Not CBool(j("sheet")) Then
        UprightView mainView
    End If
    d.ForceRebuild3 False
    viewScaleValue = FitScale(mainView, FitWidth(MAIN_BOX_WIDTH_M), FitHeight(ORTHOGRAPHIC_BOX_HEIGHT_M))
    SetPaperScale sheet, viewScaleValue
    mainView.UseSheetScale = SW_USE_SHEET_SCALE
    CenterView mainView, LX(ORTHOGRAPHIC_COLUMN_X_M), LY(ORTHOGRAPHIC_ROW_Y_M)
    d.ClearSelection2 True
    AsDrawing(d).ActivateView ""
    d.Extension.SelectByID2 mainView.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
    Set bottomView = AsDrawing(d).CreateUnfoldedViewAt3(LX(ORTHOGRAPHIC_COLUMN_X_M), LY(AUXILIARY_ROW_Y_M), 0, False)
    If Not bottomView Is Nothing Then
        ApplyBody bottomView, b
        bottomView.AlignWithView SW_NO_VIEW_ALIGNMENT, Nothing
        bottomView.UseParentScale = True
    End If
    d.ClearSelection2 True
    d.Extension.SelectByID2 mainView.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
    Set sideView = AsDrawing(d).CreateUnfoldedViewAt3(LX(AUXILIARY_COLUMN_X_M), LY(ORTHOGRAPHIC_ROW_Y_M), 0, False)
    If Not sideView Is Nothing Then
        ApplyBody sideView, b
        sideView.AlignWithView SW_NO_VIEW_ALIGNMENT, Nothing
        sideView.UseParentScale = True
    End If
    If Not CBool(j("sheet")) Then
        If Not bottomView Is Nothing Then
            UprightView bottomView
        End If
        If Not sideView Is Nothing Then
            UprightView sideView
        End If
    End If
    d.ForceRebuild3 False
    ' Reduce all orthographic views together if a projected view is too large.
    Dim reduction As Double
    reduction = 1
    If Not bottomView Is Nothing Then
        reduction = MinD(reduction, FitFactor(bottomView, FitWidth(MAIN_BOX_WIDTH_M), FitHeight(BOTTOM_BOX_HEIGHT_M)))
    End If
    If Not sideView Is Nothing Then
        reduction = MinD(reduction, FitFactor(sideView, FitWidth(SIDE_BOX_WIDTH_M), FitHeight(ORTHOGRAPHIC_BOX_HEIGHT_M)))
    End If
    If reduction < SCALE_REDUCTION_THRESHOLD Then
        viewScaleValue = StandardScale(viewScaleValue * reduction)
        SetPaperScale sheet, viewScaleValue
        d.ForceRebuild3 False
    End If
    CenterView mainView, LX(ORTHOGRAPHIC_COLUMN_X_M), LY(ORTHOGRAPHIC_ROW_Y_M)
    If Not bottomView Is Nothing Then
        CenterView bottomView, LX(ORTHOGRAPHIC_COLUMN_X_M), LY(AUXILIARY_ROW_Y_M)
    End If
    If Not sideView Is Nothing Then
        CenterView sideView, LX(AUXILIARY_COLUMN_X_M), LY(ORTHOGRAPHIC_ROW_Y_M)
    End If
    If CBool(j("sheet")) Then
        Set flatView = MakeFlatView(d, doc, cfg, b)
        If Not flatView Is Nothing Then
            flatView.UseSheetScale = False
            flatView.ScaleDecimal = INITIAL_FLAT_VIEW_SCALE
        Else
            Warn "Flat pattern not created for " & bodyName
        End If
    Else
        Set isoView = MakeView(d, doc, cfg, b, "*Isometric", LX(AUXILIARY_COLUMN_X_M), LY(AUXILIARY_ROW_Y_M))
        If Not isoView Is Nothing Then
            isoView.UseSheetScale = SW_USE_INDEPENDENT_SCALE: isoView.ScaleDecimal = 1
            d.ForceRebuild3 False
            isoView.ScaleDecimal = FitScale(isoView, FitWidth(AUXILIARY_BOX_WIDTH_M), FitHeight(ISOMETRIC_BOX_HEIGHT_M))
            CenterView isoView, LX(AUXILIARY_COLUMN_X_M), LY(AUXILIARY_ROW_Y_M)
        End If
    End If
    d.ForceRebuild3 False
    ' Flat view insertion can change the sheet scale; restore the calculated layout.
    SetPaperScale sheet, viewScaleValue
    mainView.UseSheetScale = SW_USE_SHEET_SCALE
    If Not bottomView Is Nothing Then
        bottomView.UseParentScale = True
    End If
    If Not sideView Is Nothing Then
        sideView.UseParentScale = True
    End If
    d.ForceRebuild3 False
    CenterView mainView, LX(ORTHOGRAPHIC_COLUMN_X_M), LY(ORTHOGRAPHIC_ROW_Y_M)
    If Not bottomView Is Nothing Then
        CenterView bottomView, LX(ORTHOGRAPHIC_COLUMN_X_M), LY(AUXILIARY_ROW_Y_M)
    End If
    If Not sideView Is Nothing Then
        CenterView sideView, LX(AUXILIARY_COLUMN_X_M), LY(ORTHOGRAPHIC_ROW_Y_M)
    End If
    d.GraphicsRedraw2
    box = mainView.Position
    Log "VIEW before dimensions scale=" & mainView.ScaleDecimal & " position=" & box(0)
    JobStage j, "Dimensioning folded views"
    If Not isAssembly Then
        ImportModelItems d
    End If
    dims = OverallDimensions(d, mainView)
    If Not bottomView Is Nothing Then
        rc = OverallDimensions(d, bottomView)
    End If
    If Not sideView Is Nothing Then
        rc = OverallDimensions(d, sideView)
    End If
    If Not isAssembly Then
        ' Feature dimensions on every orthographic view, including sheet-metal parts.
        If Not CBool(j("sheet")) Then
            DimensionFeatures d, mainView, False
            If Not bottomView Is Nothing Then
                DimensionFeatures d, bottomView, False
            End If
            If Not sideView Is Nothing Then
                DimensionFeatures d, sideView, False
            End If
        Else
            rc = DiameterDimensions(d, mainView): rc = RadiusDimensions(d, mainView)
            If Not bottomView Is Nothing Then
                rc = DiameterDimensions(d, bottomView)
            End If
            If Not sideView Is Nothing Then
                rc = DiameterDimensions(d, sideView)
            End If
        End If
    End If
    box = mainView.Position
    Log "VIEW after dimensions scale=" & mainView.ScaleDecimal & " position=" & box(0)
    AsDrawing(d).ActivateView "": d.ClearSelection2 True
    RefreshTemplateFields d
    AddText d, "АВТОМАТИЧНИЙ ПРОЄКТ - ПЕРЕВІРИТИ ПЕРЕД ВИПУСКОМ", DRAFT_MARKER_X_M, sheetHeight - DRAFT_MARKER_TOP_MARGIN_M, DRAFT_MARKER_TEXT_HEIGHT_M, 0
    If Not flatView Is Nothing Then
        AddText d, "Розміри розгортки - на окремому аркуші.", CONTENT_LEFT_M, FLAT_NOTICE_Y_M, TEXT_HEIGHT_SMALL_M, 0
    End If
    If Not isoView Is Nothing Then
        AddText d, "Ізометрія (" & ScaleText(isoView.ScaleDecimal) & ")", LX(AUXILIARY_CAPTION_X_M), AUXILIARY_CAPTION_Y_M, TEXT_HEIGHT_SMALL_M, 0
    End If
    AddTechnicalNotes d, doc, cfg, b, CLng(j("qty")), isAssembly
    If isAssembly Then
        AddAssemblyList d, doc, cfg, code, name, mass
    End If
    NormalizeDimensions d
    AsDrawing(d).ActivateSheet sheet.GetName
    d.ActiveView.EnableGraphicsUpdate = True
    d.ForceRebuild3 False
    d.ViewZoomtofit2
    ' Persist derived flat configurations in the snapshot only.
    doc.ShowConfiguration2 oldCfg
    If doc.Save3(SW_SAVE_SILENT, er, wr) = False Then
        Err.Raise vbObjectError + ERR_SAVE_MODEL, , "Cannot save snapshot model: " & er
    End If
    success = d.Extension.SaveAs(runDir & DRAWINGS_RELATIVE_DIRECTORY & outName & DRAWING_EXTENSION, SW_SAVE_CURRENT_VERSION, SW_SAVE_SILENT, Nothing, er, wr)
    If success = False Then
        Err.Raise vbObjectError + ERR_SAVE_DRAWING, , "Drawing save failed: " & er
    End If
    ' Reopen to refresh cached drawing display data before PDF export.
    sw.CloseDoc d.GetTitle
    Set d = sw.OpenDoc6(runDir & DRAWINGS_RELATIVE_DIRECTORY & outName & DRAWING_EXTENSION, SW_DOC_DRAWING, SW_OPEN_SILENT, "", er, wr)
    If d Is Nothing Then
        Err.Raise vbObjectError + ERR_REOPEN_DRAWING, , "Cannot reopen saved drawing: " & er
    End If
    Set currentDrawing = d
    If backgroundOwned Then
        PrepareDocumentWindow d
    End If
    d.ForceRebuild3 False
    Set flatView = AsDrawing(d).GetFirstView
    Set flatView = flatView.GetNextView
    Do While Not flatView Is Nothing
        If flatView.IsFlatPatternView Then
            JobStage j, "Building and dimensioning flat pattern"
            flatScale = LayoutFlatView(d, doc, flatView)
            Set active = sw.ActivateDoc3(d.GetTitle, False, SW_REBUILD_USER_DECISION, er)
            AddText d, "Розгортка (" & ScaleText(flatScale) & ")", LX(AUXILIARY_CAPTION_X_M), AUXILIARY_CAPTION_Y_M, TEXT_HEIGHT_SMALL_M, 0
            AddFlatDimensionSheet d, doc, flatView, code, name, material, mass
            AsDrawing(d).ActivateSheet firstSheetName
            Exit Do
        End If
        Set flatView = flatView.GetNextView
    Loop
    d.ForceRebuild3 False
    d.GraphicsRedraw2
    RefreshTemplateFields d
    JobStage j, "Saving rebuilt drawing"
    NormalizeDimensions d
    WriteDimensionCoverage d, outName
    AsDrawing(d).ActivateSheet firstSheetName
    Log "SAVE rebuilt drawing " & outName
    If d.Save3(SW_SAVE_SILENT, er, wr) = False Then
        Err.Raise vbObjectError + ERR_SAVE_DRAWING, , "Rebuilt drawing save failed: " & er
    End If
    Dim pdf As Object
    Set pdf = sw.GetExportFileData(SW_EXPORT_PDF)
    pdf.ViewPdfAfterSaving = False
    pdf.SetSheets SW_EXPORT_ALL_SHEETS, Empty
    JobStage j, "Exporting PDF"
    Log "EXPORT PDF " & outName
    success = d.Extension.SaveAs(runDir & PDF_RELATIVE_DIRECTORY & outName & PDF_EXTENSION, SW_SAVE_CURRENT_VERSION, SW_SAVE_SILENT, pdf, er, wr)
    If success = False Then
        Err.Raise vbObjectError + ERR_SAVE_PDF, , "PDF save failed: " & er
    End If
    AppendManifest jobNumber & vbTab & OriginalBase(doc.GetPathName) & vbTab & cfg & vbTab & bodyName & vbTab & j("qty") & vbTab & outName & DRAWING_EXTENSION & vbTab & "OK; REVIEW"
    j("status") = STATE_DONE: j("stage") = "Saved"
    j("drawingBytes") = CStr(fs.GetFile(runDir & DRAWINGS_RELATIVE_DIRECTORY & outName & DRAWING_EXTENSION).Size)
    j("pdfBytes") = CStr(fs.GetFile(runDir & PDF_RELATIVE_DIRECTORY & outName & PDF_EXTENSION).Size)
    SaveCheckpoint
    completed = completed + 1
    Log "OK " & outName & " dims(main)=" & dims
    sw.CloseDoc d.GetTitle
    Set currentDrawing = Nothing
    Exit Sub
Failed:
    Dim msg As String
    msg = Err.Number & " " & Err.Description
    failures = failures + 1
    On Error Resume Next
    j("status") = STATE_FAILED: j("stage") = msg
    SaveCheckpoint
    Log "FAIL " & outName & " | " & msg
    AppendManifest jobNumber & vbTab & OriginalBase(CStr(j("path"))) & vbTab & cfg & vbTab & bodyName & vbTab & j("qty") & vbTab & outName & vbTab & "FAILED " & msg
    On Error Resume Next
    If Not doc Is Nothing Then
        doc.ShowConfiguration2 oldCfg
    End If
    If Not d Is Nothing Then
        sw.CloseDoc d.GetTitle
    End If
    Set currentDrawing = Nothing
End Sub

Private Sub ConfigureDrawing(ByVal d As SldWorks.ModelDoc2)
    Dim tf As Object, i As Long
    d.Extension.SetUserPreferenceInteger SW_PREF_DIMENSION_STANDARD, 0, SW_DRAFTING_STANDARD_GOST ' GOST
    d.Extension.SetUserPreferenceInteger SW_PREF_UNIT_SYSTEM, 0, SW_UNIT_SYSTEM_LEGACY_DEFAULT ' Keep the existing preset; linear units are set explicitly below.
    d.Extension.SetUserPreferenceInteger SW_PREF_LINEAR_UNITS, 0, SW_LENGTH_MILLIMETRES ' millimetres
    d.Extension.SetUserPreferenceInteger SW_PREF_LINEAR_PRECISION, 0, LINEAR_DECIMAL_PLACES
    d.Extension.SetUserPreferenceInteger SW_PREF_ANGULAR_PRECISION, 0, ANGULAR_DECIMAL_PLACES
    Dim toggleId As Variant
    For Each toggleId In Array(SW_HIDE_AXES, SW_HIDE_PLANES, SW_HIDE_ORIGINS, SW_HIDE_TEMPORARY_AXES, SW_HIDE_COORDINATE_SYSTEMS, SW_HIDE_REFERENCE_POINTS, SW_HIDE_CURVES, SW_HIDE_SKETCHES, SW_HIDE_REFERENCE_POINTS_2)
        d.SetUserPreferenceToggle CLng(toggleId), False
    Next toggleId
    For i = SW_TEXT_FORMAT_NOTE To SW_TEXT_FORMAT_DIMENSION
        Set tf = d.GetUserPreferenceTextFormat(i)
        tf.TypeFaceName = Setting("Font", DEFAULT_FONT)
        tf.CharHeight = TEXT_HEIGHT_STANDARD_M: tf.Italic = False
        d.SetUserPreferenceTextFormat i, tf
    Next i
End Sub

Private Sub ResolveDrawingTemplate()
    Dim base As String, folder As String
    base = Setting("LauncherSettingsDirectory", sw.GetCurrentMacroPathFolder)
    folder = Setting("TemplatesDirectory", Setting("LauncherScriptDirectory", sw.GetCurrentMacroPathFolder) & "\templates")
    If InStr(folder, ":") = 0 And Left$(folder, 2) <> "\\" Then
        folder = fs.BuildPath(base, folder)
    End If
    paperFormat = UCase$(Setting("PaperFormat", PAPER_A3))
    paperOrientation = LCase$(Setting("PaperOrientation", ORIENTATION_LANDSCAPE))
    If paperFormat <> PAPER_A4 And paperFormat <> PAPER_A3 Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "PaperFormat must be A4 or A3."
    End If
    If paperOrientation <> ORIENTATION_LANDSCAPE And paperOrientation <> ORIENTATION_PORTRAIT Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "PaperOrientation must be landscape or portrait."
    End If
    templatePath = Setting("DrawingTemplate", "")
    If templatePath = "" Then
        templatePath = fs.BuildPath(folder, paperFormat & "." & paperOrientation & DRAWING_TEMPLATE_EXTENSION)
    End If
    If InStr(templatePath, ":") = 0 And Left$(templatePath, 2) <> "\\" Then
        templatePath = fs.BuildPath(base, templatePath)
    End If
    templatePath = fs.GetAbsolutePathName(templatePath)
    If LCase$(fs.GetExtensionName(templatePath)) <> DRAWING_TEMPLATE_FILE_TYPE Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "DrawingTemplate must be a .drwdot file."
    End If
    If Not fs.FileExists(templatePath) Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "Drawing template not found: " & templatePath
    End If
End Sub

Private Sub ValidateDrawingTemplate()
    Dim candidate As SldWorks.ModelDoc2, errorNumber As Long, errorText As String
    On Error GoTo Failed
    Set candidate = sw.NewDocument(templatePath, 0, 0, 0)
    If candidate Is Nothing Then
        Err.Raise vbObjectError + ERR_DRAWING_TEMPLATE, , "Cannot open selected drawing template."
    End If
    InitializePaper candidate, False
    sw.CloseDoc candidate.GetTitle
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If Not candidate Is Nothing Then
        sw.CloseDoc candidate.GetTitle
    End If
    On Error GoTo 0
    Err.Raise errorNumber, , errorText
End Sub

Private Sub InitializePaper(ByVal d As SldWorks.ModelDoc2, Optional ByVal saveFormat As Boolean = True)
    Dim sheet As SldWorks.Sheet, properties As Variant, shortSide As Double, longSide As Double
    Dim active As SldWorks.ModelDoc2, error As Long, annotations As Variant, noteCount As Long, formatResult As Long
    PrepareDocumentWindow d
    Set active = sw.ActivateDoc3(d.GetTitle, False, SW_REBUILD_USER_DECISION, error)
    If AsDrawing(d).GetSheetCount <> 1 Then
        Err.Raise vbObjectError + ERR_DRAWING_TEMPLATE, , "Drawing template must contain one sheet."
    End If
    Set sheet = AsDrawing(d).GetCurrentSheet
    properties = sheet.GetProperties
    sheetWidth = CDbl(properties(SHEET_WIDTH_PROPERTY)): sheetHeight = CDbl(properties(SHEET_HEIGHT_PROPERTY))
    shortSide = MinD(sheetWidth, sheetHeight): longSide = -MinD(-sheetWidth, -sheetHeight)
    If paperFormat = PAPER_A4 Then
        If Abs(shortSide - A4_SHORT_SIDE_M) > SHEET_SIZE_TOLERANCE_M Or Abs(longSide - A4_LONG_SIDE_AND_A3_SHORT_SIDE_M) > SHEET_SIZE_TOLERANCE_M Then
            Err.Raise vbObjectError + ERR_DRAWING_TEMPLATE, , "Selected template size does not match A4 (210 x 297 mm)."
        End If
    Else
        If Abs(shortSide - A4_LONG_SIDE_AND_A3_SHORT_SIDE_M) > SHEET_SIZE_TOLERANCE_M Or Abs(longSide - A3_LONG_SIDE_M) > SHEET_SIZE_TOLERANCE_M Then
            Err.Raise vbObjectError + ERR_DRAWING_TEMPLATE, , "Selected template size does not match A3 (297 x 420 mm)."
        End If
    End If
    layoutSx = (sheetWidth - LAYOUT_HORIZONTAL_RESERVED_M) / REFERENCE_USABLE_WIDTH_M
    layoutSy = (sheetHeight - LAYOUT_VERTICAL_RESERVED_M) / REFERENCE_USABLE_HEIGHT_M
    ' Export the EMBEDDED format. Renamed/copied .drwdot files need no old external .slddrt.
    If saveFormat Then
        AsDrawing(d).EditTemplate
        AsDrawing(d).EditSheet
        d.ForceRebuild3 False
        annotations = AsDrawing(d).GetFirstView.GetAnnotations
        If Not IsEmpty(annotations) Then
            noteCount = UBound(annotations) + 1
        End If
        Log "TEMPLATE format=" & sheet.GetTemplateName & "; notes=" & noteCount
        formatResult = sheet.SaveFormat(runtimeFormatPath)
        Log "TEMPLATE SaveFormat result=" & formatResult
        If formatResult = 0 Then
            Err.Raise vbObjectError + ERR_DRAWING_TEMPLATE, , "Cannot save embedded drawing sheet format."
        End If
        sheet.SetTemplateName runtimeFormatPath
    End If
    Log "PAPER " & paperFormat & " " & Format$(sheetWidth * MILLIMETRES_PER_METRE, "0") & "x" & Format$(sheetHeight * MILLIMETRES_PER_METRE, "0") & " mm | " & templatePath
End Sub

Private Function LX(ByVal reference As Double) As Double
    LX = LAYOUT_ORIGIN_X_M + (reference - LAYOUT_ORIGIN_X_M) * layoutSx
End Function

Private Function LY(ByVal reference As Double) As Double
    LY = LAYOUT_ORIGIN_Y_M + (reference - LAYOUT_ORIGIN_Y_M) * layoutSy
End Function

Private Function FitWidth(ByVal reference As Double) As Double
    FitWidth = reference * layoutSx - SMALL_SHEET_WIDTH_CLEARANCE_M * (1 - MinD(1, layoutSx))
End Function

Private Function FitHeight(ByVal reference As Double) As Double
    FitHeight = reference * layoutSy - SMALL_SHEET_HEIGHT_CLEARANCE_M * (1 - MinD(1, layoutSy))
End Function

Private Sub NewFormattedSheet(ByVal d As SldWorks.ModelDoc2, ByVal name As String)
    Dim sheet As SldWorks.Sheet
    If AsDrawing(d).NewSheet4(name, SW_CUSTOM_PAPER, SW_CUSTOM_SHEET_TEMPLATE, 1, 1, True, runtimeFormatPath, sheetWidth, sheetHeight, "", 0, 0, 0, 0, 0, 0) = False Then
        Err.Raise vbObjectError + ERR_DETAIL_SHEET, , "Cannot create formatted sheet: " & name
    End If
    AsDrawing(d).ActivateSheet name
    If AsDrawing(d).SetupSheet5(name, SW_CUSTOM_PAPER, SW_CUSTOM_SHEET_TEMPLATE, 1, 1, True, runtimeFormatPath, sheetWidth, sheetHeight, "", True) = False Then
        Err.Raise vbObjectError + ERR_DETAIL_SHEET, , "Cannot apply sheet format: " & name
    End If
    Set sheet = AsDrawing(d).GetCurrentSheet
    sheet.SetTemplateName runtimeFormatPath
    If sheet.ReloadTemplate(False) <> 0 Then
        Err.Raise vbObjectError + ERR_DETAIL_SHEET, , "Cannot reload sheet format: " & name
    End If
    d.ForceRebuild3 False
End Sub

Private Sub SetPaperScale(ByVal sheet As Object, ByVal scaleValue As Double)
    If scaleValue >= 1 Then
        sheet.SetScale scaleValue, 1, False, False
    Else
        sheet.SetScale 1, 1 / scaleValue, False, False
    End If
End Sub

Private Sub SetDrawingProperty(ByVal d As SldWorks.ModelDoc2, ByVal name As String, ByVal value As String)
    d.Extension.CustomPropertyManager("").Add3 name, SW_CUSTOM_PROPERTY_TEXT, value, SW_PROPERTY_REPLACE_VALUE
End Sub

Private Sub FillTemplateProperties(ByVal d As SldWorks.ModelDoc2, ByVal code As String, ByVal name As String, ByVal material As String, ByVal mass As String)
    Dim key As Variant
    SetDrawingProperty d, "Designation", code
    SetDrawingProperty d, "Description", WrapText(name, DESCRIPTION_WRAP_CHARACTERS)
    SetDrawingProperty d, "Material", WrapText(material, MATERIAL_WRAP_CHARACTERS)
    SetDrawingProperty d, "Mass", mass
    If Setting("Organization", "") <> "" Then
        SetDrawingProperty d, "Organization", Setting("Organization", "")
    End If
    For Each key In Array("Liter1", "Liter2", "Liter3", "DrawnBy", "DrawnDate", "CheckedBy", "CheckedDate", "TechnicalControlBy", "TechnicalControlDate", "NormControlBy", "NormControlDate", "ApprovedBy", "ApprovedDate")
        If settings.Exists(CStr(key)) Then
            SetDrawingProperty d, CStr(key), Setting(CStr(key), "")
        End If
    Next key
    d.ForceRebuild3 False
End Sub

Private Sub RefreshTemplateFields(ByVal d As SldWorks.ModelDoc2)
    Dim sheets As Variant, name As Variant, view As SldWorks.View, annotations As Variant, item As Variant
    Dim ann As SldWorks.Annotation, note As SldWorks.Note, tf As SldWorks.TextFormat, box As Variant
    Dim w As Double, h As Double, actualW As Double, actualH As Double, attempt As Long
    sheets = AsDrawing(d).GetSheetNames
    For Each name In sheets
        AsDrawing(d).ActivateSheet CStr(name)
        Set view = AsDrawing(d).GetFirstView
        annotations = view.GetAnnotations
        If Not IsEmpty(annotations) Then
            For Each item In annotations
                Set ann = item
                If ann.GetType = SW_ANNOTATION_NOTE Then
                    Set note = ann.GetSpecificAnnotation
                    If note.GetName = "DSTU_SheetNumber" Then
                        If UBound(sheets) = 0 Then
                            note.SetText "Аркуш "
                        Else
                            note.SetText "Аркуш $PRP:""SW-Current Sheet"""
                        End If
                    End If
                    w = 0: h = 0
                    Select Case note.GetName
                        Case "DSTU_Designation": w = TITLE_DESIGNATION_WIDTH_M: h = TITLE_TEXT_CELL_HEIGHT_M
                        Case "DSTU_Designation26": w = TITLE_MAIN_CELL_WIDTH_M: h = TITLE_AUX_CODE_HEIGHT_M
                        Case "DSTU_Description": w = TITLE_MAIN_CELL_WIDTH_M: h = TITLE_DESCRIPTION_HEIGHT_M
                        Case "DSTU_Material": w = TITLE_MAIN_CELL_WIDTH_M: h = TITLE_TEXT_CELL_HEIGHT_M
                        Case "DSTU_Organization": w = TITLE_ORGANIZATION_WIDTH_M: h = TITLE_TEXT_CELL_HEIGHT_M
                    End Select
                    If w > 0 And Len(note.GetText) > 0 Then
                        For attempt = 0 To TITLE_LAST_FIT_PASS
                            box = note.GetExtent
                            actualW = box(BBOX_X_MAX_INDEX) - box(BBOX_X_MIN_INDEX): actualH = box(BBOX_Y_MAX_INDEX) - box(BBOX_Y_MIN_INDEX)
                            If Abs(Sin(note.Angle)) > VERTICAL_NOTE_SINE_THRESHOLD Then
                                actualW = box(BBOX_Y_MAX_INDEX) - box(BBOX_Y_MIN_INDEX)
                                actualH = box(BBOX_X_MAX_INDEX) - box(BBOX_X_MIN_INDEX)
                            End If
                            If actualW <= w And actualH <= h Then
                                Exit For
                            End If
                            Set tf = ann.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
                            If tf.CharHeight > TITLE_MEDIUM_FONT_THRESHOLD_M Then
                                tf.CharHeight = TEXT_HEIGHT_STANDARD_M
                            ElseIf tf.CharHeight > TITLE_SMALL_FONT_THRESHOLD_M Then
                                tf.CharHeight = TEXT_HEIGHT_SMALL_M
                            Else
                                tf.WidthFactor = tf.WidthFactor * MinD(1, w / actualW)
                            End If
                            ann.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, tf
                        Next attempt
                    End If
                End If
            Next item
        End If
    Next name
End Sub

Private Function BodyOrdinal(ByVal doc As SldWorks.ModelDoc2, ByVal bodyName As String) As Long
    Dim bodies As Variant, b As Variant
    BodyOrdinal = 1
    bodies = AsPart(doc).GetBodies2(SW_SOLID_BODY, True)
    For Each b In bodies
        If StrComp(b.Name, bodyName, vbTextCompare) < 0 Then
            BodyOrdinal = BodyOrdinal + 1
        End If
    Next b
End Function

Private Function MakeView(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal body As Object, ByVal orientation As String, ByVal x As Double, ByVal y As Double) As Object
    Dim v As Object, active As SldWorks.ModelDoc2, nativeName As String, standardId As Long, er As Long
    Select Case orientation
        Case "*Front": standardId = SW_VIEW_FRONT
        Case "*Top": standardId = SW_VIEW_TOP
        Case "*Right": standardId = SW_VIEW_RIGHT
        Case "*Isometric": standardId = SW_VIEW_ISOMETRIC
    End Select
    nativeName = orientation
    If standardId > 0 Then
        Set active = sw.ActivateDoc3(fs.GetFileName(doc.GetPathName), False, SW_REBUILD_USER_DECISION, er)
        If active Is Nothing Then
            Err.Raise vbObjectError + ERR_VIEW_ACTIVATION, , "Cannot activate model for named view: " & er
        End If
        doc.ShowConfiguration2 cfg
        doc.ShowNamedView2 "", standardId
        viewCounter = viewCounter + 1
        nativeName = NAMED_VIEW_PREFIX & CStr(viewCounter)
        doc.NameView nativeName
        Set active = sw.ActivateDoc3(d.GetTitle, False, SW_REBUILD_USER_DECISION, er)
        If active Is Nothing Then
            Err.Raise vbObjectError + ERR_VIEW_ACTIVATION, , "Cannot reactivate drawing: " & er & " " & d.GetTitle
        End If
    End If
    AsDrawing(d).ActivateView "": d.ClearSelection2 True
    Set v = AsDrawing(d).CreateDrawViewFromModelView3(doc.GetPathName, nativeName, x, y, 0)
    If v Is Nothing Then
        Warn "View API returned no view: " & doc.GetPathName & " | " & nativeName
        Exit Function
    End If
    v.ReferencedConfiguration = cfg
    ApplyBody v, body
    v.SetDisplayMode3 False, SW_DISPLAY_HIDDEN_LINES_REMOVED, False, False ' hidden lines removed
    If orientation = "*Isometric" Then
        UprightView v
    End If
    Set MakeView = v
End Function

Private Sub ApplyBody(ByVal v As Object, ByVal b As Object)
    Dim a(0) As Object
    If b Is Nothing Then
        Exit Sub
    End If
    Set a(0) = b
    v.Bodies = a
End Sub

Private Function MakeFlatView(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal b As Object) As Object
    Dim er As Long, active As Object, v As Object
    On Error GoTo Failed
    Set active = sw.ActivateDoc3(fs.GetFileName(doc.GetPathName), False, SW_REBUILD_USER_DECISION, er)
    doc.ShowConfiguration2 cfg
    doc.ClearSelection2 True
    If b.Select2(False, Nothing) = False Then
        GoTo Failed
    End If
    Set active = sw.ActivateDoc3(d.GetTitle, False, SW_REBUILD_USER_DECISION, er)
    Set v = AsDrawing(d).CreateFlatPatternViewFromModelView3(doc.GetPathName, cfg, LX(AUXILIARY_COLUMN_X_M), LY(AUXILIARY_ROW_Y_M), 0, False, False)
    If Not v Is Nothing Then
        If v.IsFlatPatternView = False Then
            GoTo Failed
        End If
        v.SetDisplayMode3 False, SW_DISPLAY_HIDDEN_LINES_REMOVED, False, False
        v.ShowSheetMetalBendNotes = (Setting("ShowBendNotes", "0") = "1")
        d.ForceRebuild3 False
        Set MakeFlatView = v
    End If
    Exit Function
Failed:
    Warn "Flat API: " & Err.Description
End Function

Private Function LayoutFlatView(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByRef v As SldWorks.View) As Double
    Dim bounds As Variant, position As Variant
    Set v = CreatePlanarFlatView(d, doc, v)
    v.UseSheetScale = False: v.ScaleDecimal = 1: v.Angle = 0
    d.ForceRebuild3 False
    bounds = ProjectedFlatBounds(v)
    Log "FLAT bounds=" & bounds(OUTLINE_LEFT_INDEX) & "," & bounds(OUTLINE_BOTTOM_INDEX) & "," & bounds(OUTLINE_RIGHT_INDEX) & "," & bounds(OUTLINE_TOP_INDEX)
    If bounds(OUTLINE_RIGHT_INDEX) - bounds(OUTLINE_LEFT_INDEX) < MIN_PROJECTED_EXTENT_M Or bounds(OUTLINE_TOP_INDEX) - bounds(OUTLINE_BOTTOM_INDEX) < MIN_PROJECTED_EXTENT_M Then
        Err.Raise vbObjectError + ERR_GENERATION_DATA, , "Empty projected flat geometry."
    End If
    If bounds(OUTLINE_TOP_INDEX) - bounds(OUTLINE_BOTTOM_INDEX) > bounds(OUTLINE_RIGHT_INDEX) - bounds(OUTLINE_LEFT_INDEX) Then
        v.Angle = PI / 2
        d.ForceRebuild3 False
        bounds = ProjectedFlatBounds(v)
    End If
    LayoutFlatView = StandardScale(VIEW_FIT_SAFETY_FACTOR * MinD(FitWidth(AUXILIARY_BOX_WIDTH_M) / (bounds(OUTLINE_RIGHT_INDEX) - bounds(OUTLINE_LEFT_INDEX)), FitHeight(BOTTOM_BOX_HEIGHT_M) / (bounds(OUTLINE_TOP_INDEX) - bounds(OUTLINE_BOTTOM_INDEX))))
    v.ScaleDecimal = LayoutFlatView
    d.ForceRebuild3 False
    bounds = ProjectedFlatBounds(v): position = v.Position
    position(0) = position(0) + LX(AUXILIARY_COLUMN_X_M) - (bounds(OUTLINE_LEFT_INDEX) + bounds(OUTLINE_RIGHT_INDEX)) / 2
    position(1) = position(1) + LY(AUXILIARY_ROW_Y_M) - (bounds(OUTLINE_BOTTOM_INDEX) + bounds(OUTLINE_TOP_INDEX)) / 2
    v.Position = position
    Log "FLAT layout scale=" & LayoutFlatView
End Function

Private Function ProjectedFlatBounds(ByVal v As SldWorks.View) As Variant
    Dim util As SldWorks.MathUtility, inverse As SldWorks.MathTransform, vector As SldWorks.MathVector
    Dim point As SldWorks.MathPoint, directions As Variant, direction(2) As Double, coords(2) As Double, axis As Variant, p As Variant
    Dim x As Double, y As Double, z As Double, bounds(OUTLINE_TOP_INDEX) As Double, i As Long
    Set util = sw.GetMathUtility: Set inverse = v.ModelToViewTransform.Inverse
    directions = Array(Array(-1#, 0#, 0#), Array(0#, -1#, 0#), Array(1#, 0#, 0#), Array(0#, 1#, 0#))
    For i = OUTLINE_LEFT_INDEX To OUTLINE_TOP_INDEX
        direction(0) = directions(i)(0): direction(1) = directions(i)(1): direction(2) = directions(i)(2)
        Set vector = util.CreateVector(direction): Set vector = vector.MultiplyTransform(inverse)
        axis = vector.ArrayData
        If flatBody.GetExtremePoint(axis(0), axis(1), axis(2), x, y, z) = False Then
            Err.Raise vbObjectError + ERR_GENERATION_DATA, , "Cannot calculate exact flat bounds."
        End If
        coords(0) = x: coords(1) = y: coords(2) = z
        Set point = util.CreatePoint(coords): Set point = point.MultiplyTransform(v.ModelToViewTransform)
        p = point.ArrayData: bounds(i) = p(i Mod 2)
    Next i
    ProjectedFlatBounds = bounds
End Function

Private Function CreatePlanarFlatView(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal v As SldWorks.View) As SldWorks.View
    Dim cfg As String, oldCfg As String, er As Long, active As Object, bodies As Variant, body As SldWorks.Body2
    Dim faces As Variant, item As Variant, face As SldWorks.Face2, bestFace As SldWorks.Face2, area As Double
    Dim viewName As String, oldName As String, wr As Long
    Dim edges As Variant, edgeItem As Variant, edge As SldWorks.Edge, first As SldWorks.Vertex, last As SldWorks.Vertex
    Dim n As Variant, p As Variant, q As Variant, rotation(MATRIX_LAST_INDEX) As Double, length As Double, longest As Double
    Dim ax As Double, ay As Double, az As Double, dx As Double, dy As Double, dz As Double
    Dim util As SldWorks.MathUtility, camera As SldWorks.ModelView, transform As SldWorks.MathTransform, normal As SldWorks.MathVector
    Dim curve As SldWorks.Curve, parameters As SldWorks.CurveParamData, sample As Long, xyz As Variant
    Set doc = v.ReferencedDocument
    PrepareDocumentWindow doc
    oldCfg = doc.ConfigurationManager.ActiveConfiguration.Name: cfg = v.ReferencedConfiguration
    Set active = sw.ActivateDoc3(fs.GetFileName(doc.GetPathName), False, SW_REBUILD_USER_DECISION, er)
    If active Is Nothing Then
        Err.Raise vbObjectError + ERR_FLAT_MODEL_ACTIVATION, , "Cannot activate flat source: " & er
    End If
    If LCase$(active.GetPathName) <> LCase$(doc.GetPathName) Then
        Err.Raise vbObjectError + ERR_WRONG_ACTIVE_MODEL, , "Wrong active model."
    End If
    doc.ShowConfiguration2 cfg: doc.ForceRebuild3 False
    If doc.ConfigurationManager.ActiveConfiguration.Name <> cfg Then
        Err.Raise vbObjectError + ERR_FLAT_CONFIGURATION, , "Missing flat configuration."
    End If
    bodies = AsPart(doc).GetBodies2(SW_SOLID_BODY, True)
    Log "FLAT source config=" & cfg & " bodies=" & (UBound(bodies) + 1)
    For Each item In bodies
        Set body = item
        If InStr(1, body.Name, "Flat-Pattern", vbTextCompare) > 0 Or UBound(bodies) = 0 Then
            Exit For
        End If
    Next item
    Set flatBody = body.Copy
    xyz = body.GetBodyBox
    Log "FLAT exact body=" & body.Name & " box=" & xyz(BBOX_X_MIN_INDEX) & "," & xyz(BBOX_Y_MIN_INDEX) & "," & xyz(BBOX_Z_MIN_INDEX) & "," & xyz(BBOX_X_MAX_INDEX) & "," & xyz(BBOX_Y_MAX_INDEX) & "," & xyz(BBOX_Z_MAX_INDEX)
    faces = body.GetFaces
    For Each item In faces
        Set face = item
        If face.GetSurface.IsPlane Then
            If face.GetArea > area Then
                area = face.GetArea
                Set bestFace = face
            End If
        End If
    Next item
    If bestFace Is Nothing Then
        Err.Raise vbObjectError + ERR_FLAT_FACE, , "No flat plane."
    End If
    n = bestFace.Normal: edges = bestFace.GetEdges
    For Each edgeItem In edges
        Set edge = edgeItem
        If edge.GetCurve.IsLine Then
            Set first = edge.GetStartVertex: Set last = edge.GetEndVertex
            If Not first Is Nothing And Not last Is Nothing Then
                p = first.GetPoint: q = last.GetPoint
                dx = q(0) - p(0): dy = q(1) - p(1): dz = q(2) - p(2)
                length = Sqr(dx * dx + dy * dy + dz * dz)
                If length > longest Then
                    longest = length
                    ax = dx / length
                    ay = dy / length
                    az = dz / length
                End If
            End If
        End If
    Next edgeItem
    If longest = 0 Then
        Err.Raise vbObjectError + ERR_FLAT_EDGE, , "No straight flat edge."
    End If
    rotation(MATRIX_XX_INDEX) = ax: rotation(MATRIX_YX_INDEX) = ay: rotation(MATRIX_ZX_INDEX) = az
    rotation(MATRIX_XY_INDEX) = n(1) * az - n(2) * ay
    rotation(MATRIX_YY_INDEX) = n(2) * ax - n(0) * az
    rotation(MATRIX_ZY_INDEX) = n(0) * ay - n(1) * ax
    rotation(MATRIX_XZ_INDEX) = n(0): rotation(MATRIX_YZ_INDEX) = n(1): rotation(MATRIX_ZZ_INDEX) = n(2): rotation(MATRIX_SCALE_INDEX) = 1
    Set util = sw.GetMathUtility: Set camera = doc.ActiveView
    Set transform = util.CreateTransform(rotation)
    camera.Orientation3 = transform
    doc.ViewZoomtofit2
    doc.ClearSelection2 True
    viewCounter = viewCounter + 1
    viewName = NAMED_FLAT_VIEW_PREFIX & CStr(viewCounter) & "_" & cfg
    doc.NameView viewName
    doc.Save3 SW_SAVE_SILENT, er, wr
    oldName = v.GetName2
    Set active = sw.ActivateDoc3(d.GetTitle, False, SW_REBUILD_USER_DECISION, er)
    d.ClearSelection2 True
    If d.Extension.SelectByID2(oldName, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0) = False Then
        Err.Raise vbObjectError + ERR_FLAT_ORIENTATION, , "Cannot replace flat orientation."
    End If
    d.EditDelete
    Set v = AsDrawing(d).CreateDrawViewFromModelView3(doc.GetPathName, viewName, LX(AUXILIARY_COLUMN_X_M), LY(AUXILIARY_ROW_Y_M), 0)
    If v Is Nothing Then
        Err.Raise vbObjectError + ERR_CREATE_FLAT_VIEW, , "Cannot create planar flat view."
    End If
    v.ReferencedConfiguration = cfg
    v.SetDisplayMode3 False, SW_DISPLAY_HIDDEN_LINES_REMOVED, False, False
    d.ForceRebuild3 False
    Set normal = util.CreateVector(n): Set normal = normal.MultiplyTransform(v.ModelToViewTransform)
    p = normal.ArrayData
    If Abs(p(2)) / Sqr(p(0) * p(0) + p(1) * p(1) + p(2) * p(2)) < FACE_NORMAL_ALIGNMENT_COSINE Then
        Err.Raise vbObjectError + ERR_GENERATION_DATA, , "Flat view is not normal to developed face."
    End If
    Set CreatePlanarFlatView = v
End Function

Private Function BestOrientation(ByVal doc As SldWorks.ModelDoc2, ByVal b As Object) As String
    Dim box As Variant, x As Double, y As Double, z As Double
    If doc.GetType = SW_DOC_ASSEMBLY Then
        box = AsAssembly(doc).GetBox(0)
    Else
        box = b.GetBodyBox
    End If
    x = box(BBOX_X_MAX_INDEX) - box(BBOX_X_MIN_INDEX): y = box(BBOX_Y_MAX_INDEX) - box(BBOX_Y_MIN_INDEX): z = box(BBOX_Z_MAX_INDEX) - box(BBOX_Z_MIN_INDEX)
    If doc.GetType = SW_DOC_ASSEMBLY Then
        Select Case Right$(UCase$(Setting("UpAxis", DEFAULT_UP_AXIS)), 1)
            Case "Y": BestOrientation = IIf(x >= z, "*Front", "*Right")
            Case "Z": BestOrientation = IIf(x >= y, "*Top", "*Right")
            Case "X": BestOrientation = IIf(z >= y, "*Top", "*Front")
            Case Else: Err.Raise vbObjectError + ERR_UP_AXIS, , "UpAxis must be +X, -X, +Y, -Y, +Z or -Z."
        End Select
        Exit Function
    End If
    BestOrientation = "*Front"
    If x * z > x * y And x * z >= y * z Then
        BestOrientation = "*Top"
    End If
    If y * z > x * y And y * z > x * z Then
        BestOrientation = "*Right"
    End If
End Function

Private Function FitFactor(ByVal v As Object, ByVal w As Double, ByVal h As Double) As Double
    Dim a As Variant
    a = v.GetOutline
    If a(OUTLINE_RIGHT_INDEX) - a(OUTLINE_LEFT_INDEX) < MIN_PROJECTED_EXTENT_M Or a(OUTLINE_TOP_INDEX) - a(OUTLINE_BOTTOM_INDEX) < MIN_PROJECTED_EXTENT_M Then
        FitFactor = 1
        Exit Function
    End If
    FitFactor = MinD(w / (a(OUTLINE_RIGHT_INDEX) - a(OUTLINE_LEFT_INDEX)), h / (a(OUTLINE_TOP_INDEX) - a(OUTLINE_BOTTOM_INDEX)))
End Function

Private Function FitScale(ByVal v As Object, ByVal w As Double, ByVal h As Double) As Double
    FitScale = StandardScale(v.ScaleDecimal * FitFactor(v, w, h))
End Function

Private Function StandardScale(ByVal limit As Double) As Double
    Dim a As Variant, i As Long
    a = Array(SCALE_10_TO_1, SCALE_5_TO_1, SCALE_4_TO_1, SCALE_2_5_TO_1, SCALE_2_TO_1, SCALE_1_TO_1, SCALE_1_TO_2, SCALE_1_TO_2_5, SCALE_1_TO_4, SCALE_1_TO_5, SCALE_1_TO_10, SCALE_1_TO_20, SCALE_1_TO_25, SCALE_1_TO_40, SCALE_1_TO_50, SCALE_1_TO_100, SCALE_1_TO_200, SCALE_1_TO_400, SCALE_1_TO_500, SCALE_1_TO_1000)
    For i = 0 To UBound(a)
        If CDbl(a(i)) <= limit + SCALE_COMPARISON_TOLERANCE Then
            StandardScale = a(i)
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + ERR_UNSUPPORTED_SCALE, , "Model does not fit the supported drawing scale range."
End Function

Private Sub CenterView(ByVal v As Object, ByVal x As Double, ByVal y As Double)
    Dim a As Variant, p As Variant
    a = v.GetOutline: p = v.Position
    p(0) = p(0) + x - (a(OUTLINE_LEFT_INDEX) + a(OUTLINE_RIGHT_INDEX)) / 2
    p(1) = p(1) + y - (a(OUTLINE_BOTTOM_INDEX) + a(OUTLINE_TOP_INDEX)) / 2
    v.Position = p
End Sub

Private Function VisibleEdgeCount(ByVal v As Object) As Long
    Dim c As Variant, entities As Variant, i As Long
    c = v.GetVisibleComponents
    If IsEmpty(c) Then
        Exit Function
    End If
    For i = LBound(c) To UBound(c)
        entities = v.GetVisibleEntities2(c(i), 1)
        If Not IsEmpty(entities) Then
            VisibleEdgeCount = VisibleEdgeCount + UBound(entities) - LBound(entities) + 1
        End If
    Next i
End Function

Private Function DiameterDimensions(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View) As Long
    Dim comps As Variant, edges As Variant, ci As Variant, ei As Variant, edge As SldWorks.Edge
    Dim curve As SldWorks.Curve, cp As Variant, start As SldWorks.Vertex, finish As SldWorks.Vertex
    Dim p As Variant, q As Variant, axis(2) As Double, vector As SldWorks.MathVector, util As SldWorks.MathUtility
    Dim seen As Object, key As String, radius As Double, outline As Variant, dd As SldWorks.DisplayDimension
    On Error GoTo Failed
    Set seen = CreateObject("Scripting.Dictionary")
    Set util = sw.GetMathUtility
    comps = v.GetVisibleComponents: outline = v.GetOutline
    If IsEmpty(comps) Then
        Exit Function
    End If
    For Each ci In comps
        edges = v.GetVisibleEntities2(ci, SW_VISIBLE_ENTITY_EDGE)
        If Not IsEmpty(edges) Then
            For Each ei In edges
                Set edge = ei: Set curve = edge.GetCurve
                If curve.IsCircle Then
                    cp = curve.CircleParams: radius = cp(CIRCLE_RADIUS_INDEX)
                    If radius < MIN_CIRCLE_RADIUS_M Then
                        GoTo NextCircle
                    End If
                    Set start = edge.GetStartVertex: Set finish = edge.GetEndVertex
                    If Not start Is Nothing And Not finish Is Nothing Then
                        p = start.GetPoint: q = finish.GetPoint
                        If Abs(p(0) - q(0)) + Abs(p(1) - q(1)) + Abs(p(2) - q(2)) > CIRCLE_CLOSURE_TOLERANCE_M Then
                            GoTo NextCircle
                        End If
                    End If
                    axis(0) = cp(CIRCLE_AXIS_X_INDEX): axis(1) = cp(CIRCLE_AXIS_Y_INDEX): axis(2) = cp(CIRCLE_AXIS_Z_INDEX)
                    Set vector = util.CreateVector(axis)
                    Set vector = vector.MultiplyTransform(v.ModelToViewTransform)
                    p = vector.ArrayData
                    If Abs(p(2)) < FACE_NORMAL_ALIGNMENT_COSINE * Sqr(p(0) ^ 2 + p(1) ^ 2 + p(2) ^ 2) Then
                        GoTo NextCircle
                    End If
                    key = Format$(radius, GEOMETRY_KEY_FORMAT)
                    If seen.Exists(key) Then
                        GoTo NextCircle
                    End If
                    seen.Add key, True
                    d.ClearSelection2 True: AsDrawing(d).ActivateView ""
                    If v.SelectEntity(edge, False) Then
                        Set dd = d.AddDiameterDimension2(outline(OUTLINE_RIGHT_INDEX) + DIAMETER_LABEL_OFFSET_M, outline(OUTLINE_TOP_INDEX) - DIAMETER_LABEL_PITCH_M * DiameterDimensions, 0)
                        If Not dd Is Nothing Then
                            If ValidateDimension(d, v, dd, 2 * radius * v.ScaleDecimal) Then
                                DiameterDimensions = DiameterDimensions + 1
                            End If
                        End If
                    End If
                End If
NextCircle:
            Next ei
        End If
    Next ci
    d.ClearSelection2 True
    Exit Function
Failed:
    Warn "Diameter dimensions " & v.Name & ": " & Err.Description
End Function

Private Function OverallDimensions(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View, Optional ByVal lowerRight As Boolean = False) As Long
    Dim comps As Variant, verts As Variant, i As Long, k As Long, p As Variant, mp As SldWorks.MathPoint
    Dim edges As Variant, edgeItem As Variant, edge As SldWorks.Edge, vertexList As Collection, endpoint As SldWorks.Vertex
    Dim leftV As Object, rightV As Object, topV As Object, bottomV As Object, vertex As SldWorks.Vertex
    Dim comp As SldWorks.Component2, math As SldWorks.MathUtility, stage As String
    Dim refDoc As SldWorks.ModelDoc2, realComp As SldWorks.Component2, modelComps As Object
    Dim ca As Variant, ci As Variant, compName As String, partDoc As SldWorks.ModelDoc2, sourceEntity As Object
    Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, outline As Variant, dimObj As Object
    Dim horizontalY As Double, verticalX As Double
    On Error GoTo Failed
    stage = "components"
    Set math = sw.GetMathUtility
    Set refDoc = v.ReferencedDocument
    Set modelComps = CreateObject("Scripting.Dictionary"): modelComps.CompareMode = 1
    If refDoc.GetType = SW_DOC_ASSEMBLY Then
        ca = AsAssembly(refDoc).GetComponents(False)
        For Each ci In ca
            Set realComp = ci
            modelComps.Add realComp.Name2, realComp
        Next ci
    End If
    xmin = SEARCH_INITIAL_LIMIT: ymin = SEARCH_INITIAL_LIMIT: xmax = -SEARCH_INITIAL_LIMIT: ymax = -SEARCH_INITIAL_LIMIT
    comps = v.GetVisibleComponents
    If IsEmpty(comps) Then
        Exit Function
    End If
    For i = LBound(comps) To UBound(comps)
        Set comp = comps(i)
        Set realComp = Nothing
        If refDoc.GetType = SW_DOC_ASSEMBLY Then
            compName = Mid$(comp.Name2, InStr(comp.Name2, "/") + 1)
            If modelComps.Exists(compName) Then
                Set realComp = modelComps(compName)
            End If
        End If
        Set vertexList = New Collection
            edges = v.GetVisibleEntities2(comp, SW_VISIBLE_ENTITY_EDGE)
            If Not IsEmpty(edges) Then
                For Each edgeItem In edges
                    Set edge = edgeItem
                    Set endpoint = edge.GetStartVertex
                    If Not endpoint Is Nothing Then
                        vertexList.Add endpoint
                    End If
                    Set endpoint = edge.GetEndVertex
                    If Not endpoint Is Nothing Then
                        vertexList.Add endpoint
                    End If
                Next edgeItem
            End If
        If vertexList.Count > 0 Then
            For k = 1 To vertexList.Count
                Set vertex = vertexList(k)
                stage = "point"
                Set mp = math.CreatePoint(vertex.GetPoint)
                Dim corresponding As Object
                stage = "transform"
                Set corresponding = vertex
                If refDoc.GetType = SW_DOC_ASSEMBLY Then
                    If realComp Is Nothing Then
                        GoTo NextVertex
                    End If
                    Set mp = mp.MultiplyTransform(realComp.Transform2)
                End If
                Set mp = mp.MultiplyTransform(v.ModelToViewTransform)
                p = mp.ArrayData
                stage = "corresponding"
                If refDoc.GetType = SW_DOC_ASSEMBLY Then
                    If p(0) >= xmin And p(0) <= xmax And p(1) >= ymin And p(1) <= ymax Then
                        GoTo NextVertex
                    End If
                    Set partDoc = realComp.GetModelDoc2
                    Set sourceEntity = partDoc.Extension.GetCorresponding2(vertex)
                    If sourceEntity Is Nothing Then
                        Set sourceEntity = vertex
                    End If
                    Set corresponding = realComp.GetCorrespondingEntity(sourceEntity)
                    If corresponding Is Nothing Then
                        GoTo NextVertex
                    End If
                End If
                If p(0) < xmin Then
                    xmin = p(0)
                    Set leftV = corresponding
                End If
                If p(0) > xmax Then
                    xmax = p(0)
                    Set rightV = corresponding
                End If
                If p(1) < ymin Then
                    ymin = p(1)
                    Set bottomV = corresponding
                End If
                If p(1) > ymax Then
                    ymax = p(1)
                    Set topV = corresponding
                End If
NextVertex:
            Next k
        End If
        IncludeCircleExtremes v, edges, realComp, refDoc, xmin, xmax, ymin, ymax, leftV, rightV, bottomV, topV
    Next i
    outline = v.GetOutline
    horizontalY = outline(OUTLINE_TOP_INDEX) + OVERALL_LABEL_OFFSET_M: verticalX = outline(OUTLINE_LEFT_INDEX) - OVERALL_LABEL_OFFSET_M
    If lowerRight Then
        outline = ProjectedFlatBounds(v)
        horizontalY = outline(OUTLINE_BOTTOM_INDEX) - DETAIL_OVERALL_LABEL_OFFSET_M: verticalX = outline(OUTLINE_RIGHT_INDEX) + DETAIL_OVERALL_LABEL_OFFSET_M
    End If
    stage = "select horizontal"
    AsDrawing(d).ActivateView "": d.ClearSelection2 True
    If Not leftV Is Nothing And xmax - xmin > MIN_DIMENSION_SPAN_M Then
        If SelectInView(d, v, leftV, False) And SelectInView(d, v, rightV, True) Then
            stage = "add horizontal"
            Set dimObj = d.AddHorizontalDimension2((outline(OUTLINE_LEFT_INDEX) + outline(OUTLINE_RIGHT_INDEX)) / 2, horizontalY, 0)
            If Not dimObj Is Nothing Then
                SetMaximumArcEnds dimObj
                stage = "precision horizontal"
                dimObj.SetPrecision3 LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES
                stage = "validate horizontal"
                If ValidateDimension(d, v, dimObj, xmax - xmin) Then
                    OverallDimensions = OverallDimensions + 1
                End If
            End If
        End If
    End If
    d.ClearSelection2 True
    stage = "select vertical"
    If Not bottomV Is Nothing And ymax - ymin > MIN_DIMENSION_SPAN_M Then
        If SelectInView(d, v, bottomV, False) And SelectInView(d, v, topV, True) Then
            stage = "add vertical"
            Set dimObj = d.AddVerticalDimension2(verticalX, (outline(OUTLINE_BOTTOM_INDEX) + outline(OUTLINE_TOP_INDEX)) / 2, 0)
            If Not dimObj Is Nothing Then
                SetMaximumArcEnds dimObj
                stage = "precision vertical"
                dimObj.SetPrecision3 LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES
                stage = "validate vertical"
                If ValidateDimension(d, v, dimObj, ymax - ymin) Then
                    OverallDimensions = OverallDimensions + 1
                End If
            End If
        End If
    End If
    d.ClearSelection2 True
    Exit Function
Failed:
    Warn "Overall dimensions " & v.Name & " [" & stage & "]: " & Err.Description
End Function

Private Sub SetMaximumArcEnds(ByVal display As SldWorks.DisplayDimension)
    Dim dimension As SldWorks.Dimension
    Set dimension = display.GetDimension2(SW_DIMENSION_INDEX_FIRST)
    If dimension Is Nothing Then
        Exit Sub
    End If
    dimension.SetArcEndCondition 1, SW_ARC_MAXIMUM
    dimension.SetArcEndCondition 2, SW_ARC_MAXIMUM
End Sub

Private Sub IncludeCircleExtremes(ByVal v As SldWorks.View, ByVal edges As Variant, ByVal comp As SldWorks.Component2, ByVal doc As SldWorks.ModelDoc2, ByRef xmin As Double, ByRef xmax As Double, ByRef ymin As Double, ByRef ymax As Double, ByRef leftV As Object, ByRef rightV As Object, ByRef bottomV As Object, ByRef topV As Object)
    Dim item As Variant, edge As SldWorks.Edge, curve As SldWorks.Curve, cp As Variant, p As Variant, q As Variant
    Dim first As SldWorks.Vertex, last As SldWorks.Vertex, util As SldWorks.MathUtility, point As SldWorks.MathPoint
    Dim axis As SldWorks.MathVector, entity As Object, sourceEntity As Object, part As SldWorks.ModelDoc2, radius As Double
    Dim center(2) As Double, direction(2) As Double
    If IsEmpty(edges) Then
        Exit Sub
    End If
    Set util = sw.GetMathUtility
    For Each item In edges
        Set edge = item: Set curve = edge.GetCurve
        If curve.IsCircle Then
            Set first = edge.GetStartVertex: Set last = edge.GetEndVertex
            If Not first Is Nothing And Not last Is Nothing Then
                p = first.GetPoint: q = last.GetPoint
                If Abs(p(0) - q(0)) + Abs(p(1) - q(1)) + Abs(p(2) - q(2)) > CIRCLE_CLOSURE_TOLERANCE_M Then
                    GoTo NextCircleBound
                End If
            End If
            cp = curve.CircleParams
            center(0) = cp(CIRCLE_CENTER_X_INDEX): center(1) = cp(CIRCLE_CENTER_Y_INDEX): center(2) = cp(CIRCLE_CENTER_Z_INDEX)
            direction(0) = cp(CIRCLE_AXIS_X_INDEX): direction(1) = cp(CIRCLE_AXIS_Y_INDEX): direction(2) = cp(CIRCLE_AXIS_Z_INDEX)
            Set point = util.CreatePoint(center)
            Set axis = util.CreateVector(direction)
            Set entity = edge
            If doc.GetType = SW_DOC_ASSEMBLY Then
                If comp Is Nothing Then
                    GoTo NextCircleBound
                End If
                Set point = point.MultiplyTransform(comp.Transform2)
                Set axis = axis.MultiplyTransform(comp.Transform2)
                Set part = comp.GetModelDoc2
                Set sourceEntity = part.Extension.GetCorresponding2(edge)
                If sourceEntity Is Nothing Then
                    Set sourceEntity = edge
                End If
                Set entity = comp.GetCorrespondingEntity(sourceEntity)
                If entity Is Nothing Then
                    GoTo NextCircleBound
                End If
            End If
            Set point = point.MultiplyTransform(v.ModelToViewTransform): p = point.ArrayData
            Set axis = axis.MultiplyTransform(v.ModelToViewTransform): q = axis.ArrayData
            If Abs(q(2)) < FACE_NORMAL_ALIGNMENT_COSINE * Sqr(q(0) ^ 2 + q(1) ^ 2 + q(2) ^ 2) Then
                GoTo NextCircleBound
            End If
            radius = cp(CIRCLE_RADIUS_INDEX) * v.ScaleDecimal
            If p(0) - radius < xmin Then
                xmin = p(0) - radius
                Set leftV = entity
            End If
            If p(0) + radius > xmax Then
                xmax = p(0) + radius
                Set rightV = entity
            End If
            If p(1) - radius < ymin Then
                ymin = p(1) - radius
                Set bottomV = entity
            End If
            If p(1) + radius > ymax Then
                ymax = p(1) + radius
                Set topV = entity
            End If
        End If
NextCircleBound:
    Next item
End Sub

Private Function SelectInView(ByVal d As SldWorks.ModelDoc2, ByVal v As Object, ByVal entity As Object, ByVal append As Boolean) As Boolean
    Dim e As SldWorks.Entity, typedView As SldWorks.View
    Set typedView = v
    SelectInView = typedView.SelectEntity(entity, append)
End Function

Private Function ValidateDimension(ByVal d As SldWorks.ModelDoc2, ByVal v As Object, ByVal display As Object, ByVal paperSpan As Double) As Boolean
    Dim dimension As SldWorks.Dimension, actual As Double, expected As Double
    Dim typedDisplay As SldWorks.DisplayDimension
    Set typedDisplay = display
    Set dimension = typedDisplay.GetDimension2(SW_DIMENSION_INDEX_FIRST)
    If dimension Is Nothing Then
        d.ForceRebuild3 False
        Set dimension = typedDisplay.GetDimension2(SW_DIMENSION_INDEX_FIRST)
    End If
    If dimension Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_PARAMETER, , "New dimension has no parameter."
    End If
    actual = dimension.SystemValue
    expected = paperSpan / v.ScaleDecimal
    If Abs(actual - expected) <= DIMENSION_MATCH_TOLERANCE_M Then
        ValidateDimension = True
    Else
        Warn "Dimension rejected: " & v.Name & " measured=" & actual & " expected=" & expected
        d.ClearSelection2 True
        Dim ann As SldWorks.Annotation
        Set ann = typedDisplay.GetAnnotation
        If ann.Select3(False, Nothing) = False Then
            Err.Raise vbObjectError + ERR_REMOVE_DIMENSION, , "Cannot select an invalid dimension for removal."
        End If
        d.EditDelete
    End If
End Function

Private Sub NormalizeDimensions(ByVal d As SldWorks.ModelDoc2)
    Dim sheets As Variant, sheetName As Variant, v As SldWorks.View
    Dim display As SldWorks.DisplayDimension, ann As SldWorks.Annotation, tf As SldWorks.TextFormat
    sheets = AsDrawing(d).GetSheetNames
    Set tf = d.GetUserPreferenceTextFormat(SW_TEXT_FORMAT_DIMENSION)
    tf.TypeFaceName = Setting("Font", DEFAULT_FONT)
    tf.CharHeight = TEXT_HEIGHT_STANDARD_M
    For Each sheetName In sheets
        AsDrawing(d).ActivateSheet CStr(sheetName)
        Set v = AsDrawing(d).GetFirstView
        Set v = v.GetNextView
        Do While Not v Is Nothing
            Set display = v.GetFirstDisplayDimension5
            Do While Not display Is Nothing
                display.SetPrecision3 LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES
                Set ann = display.GetAnnotation
                ann.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, tf
                Set display = display.GetNext5
            Loop
            Set v = v.GetNextView
        Loop
    Next sheetName
End Sub





Private Sub AddText(ByVal d As SldWorks.ModelDoc2, ByVal text As String, ByVal x As Double, ByVal y As Double, ByVal height As Double, ByVal angle As Double)
    Dim note As Object, ann As Object, tf As Object
    If Len(text) = 0 Then
        Exit Sub
    End If
    d.ClearSelection2 True
    Set note = d.InsertNote(text)
    If note Is Nothing Then
        Err.Raise vbObjectError + ERR_INSERT_NOTE, , "Cannot insert note."
    End If
    Set ann = note.GetAnnotation
    Set tf = ann.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
    tf.TypeFaceName = Setting("Font", DEFAULT_FONT): tf.CharHeight = height
    tf.Italic = False: tf.Escapement = angle
    ann.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, tf
    ann.SetPosition2 x, y, 0
End Sub

Private Sub AddTechnicalNotes(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal b As Object, ByVal qty As Long, ByVal assembly As Boolean)
    Dim text As String, notes As String, lines As Variant, firstSheet As String, chunk As String
    Dim i As Long, page As Long, perPage As Long, last As Long, row As Long
    text = "Кількість у виробі: " & qty & "."
    If Not b Is Nothing Then
        text = text & vbCrLf & "Тіло: " & b.Name
    End If
    notes = PropertyValue(doc, cfg, "Технические требования|Технічні вимоги|TechnicalRequirements")
    If notes <> "" Then
        text = text & vbCrLf & notes
    End If
    text = WrapText(text, CLng((sheetWidth - TECHNICAL_TEXT_HORIZONTAL_RESERVED_M) / TECHNICAL_TEXT_GLYPH_WIDTH_M))
    lines = Split(text, vbCrLf)
    If UBound(lines) < TECHNICAL_FIRST_SHEET_MAX_LINES Then
        AddText d, text, CONTENT_LEFT_M, TECHNICAL_FIRST_SHEET_Y_M, TECHNICAL_TEXT_HEIGHT_M, 0
        Exit Sub
    End If
    firstSheet = AsDrawing(d).GetCurrentSheet.GetName
    perPage = Int((sheetHeight - TECHNICAL_PAGE_RESERVED_HEIGHT_M) / TECHNICAL_LINE_HEIGHT_M)
    For i = 0 To UBound(lines) Step perPage
        page = page + 1: chunk = ""
        NewFormattedSheet d, "Технічні вимоги " & page
        last = CLng(MinD(UBound(lines), i + perPage - 1))
        For row = i To last
            If chunk <> "" Then
                chunk = chunk & vbCrLf
            End If
            chunk = chunk & CStr(lines(row))
        Next row
        AddText d, "ТЕХНІЧНІ ВИМОГИ", CONTENT_LEFT_M, sheetHeight - SECTION_HEADER_TOP_MARGIN_M, TECHNICAL_HEADER_HEIGHT_M, 0
        AddText d, chunk, CONTENT_LEFT_M, sheetHeight - TECHNICAL_TEXT_TOP_MARGIN_M, TECHNICAL_TEXT_HEIGHT_M, 0
    Next i
    AsDrawing(d).ActivateSheet firstSheet
End Sub

Private Sub AddAssemblyList(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal code As String, ByVal name As String, ByVal mass As String)
    Dim v As Object, bom As Object, tf As Object, table As SldWorks.TableAnnotation
    Dim page As Long, firstRow As Long, lastRow As Long, row As Long, col As Long
    Dim height As Double, rowHeight As Double, width As Double, available As Double
    firstRow = 1: width = sheetWidth - BOM_HORIZONTAL_RESERVED_M: available = sheetHeight - BOM_VERTICAL_RESERVED_M
    Do
        page = page + 1
        NewFormattedSheet d, "Склад " & page
        Set v = MakeView(d, doc, cfg, Nothing, "*Isometric", CONTENT_LEFT_M, BOM_SOURCE_VIEW_Y_M)
        If v Is Nothing Then
            Err.Raise vbObjectError + ERR_NATIVE_BOM, , "Cannot create native BOM source view."
        End If
        v.UseSheetScale = False: v.ScaleDecimal = BOM_HIDDEN_VIEW_SCALE
        Set bom = v.InsertBomTable4(False, CONTENT_LEFT_M, sheetHeight - BOM_TOP_MARGIN_M, SW_BOM_ANCHOR_TOP_LEFT, SW_BOM_TOP_LEVEL_ONLY, cfg, "", False, 0, False)
        If bom Is Nothing Then
            Err.Raise vbObjectError + ERR_NATIVE_BOM, , "Native BOM insertion failed."
        End If
        Set table = bom
        Set tf = table.GetTextFormat
        tf.CharHeight = TEXT_HEIGHT_SMALL_M: tf.TypeFaceName = Setting("Font", DEFAULT_FONT)
        table.SetTextFormat False, tf
        For col = 0 To table.ColumnCount - 1
            If col = 0 Then
                table.SetColumnWidth col, BOM_POSITION_COLUMN_WIDTH_M, 0
            ElseIf col = table.ColumnCount - 1 Then
                table.SetColumnWidth col, BOM_QUANTITY_COLUMN_WIDTH_M, 0
            Else
                table.SetColumnWidth col, (width - BOM_FIXED_COLUMNS_WIDTH_M) / (table.ColumnCount - BOM_FIXED_COLUMN_COUNT), 0
            End If
        Next col
        table.Text2(0, 0, False) = "Поз."
        If table.ColumnCount = BOM_COMBINED_COLUMN_COUNT Then
            table.Text2(0, 1, False) = "Позначення / найменування"
        End If
        If table.ColumnCount = BOM_SPLIT_COLUMN_COUNT Then
            table.Text2(0, 1, False) = "Позначення"
            table.Text2(0, 2, False) = "Найменування"
        End If
        table.Text2(0, table.ColumnCount - 1, False) = "Кільк."
        For row = 0 To table.RowCount - 1
            table.SetRowHeight row, IIf(row = 0, BOM_HEADER_ROW_HEIGHT_M, BOM_DATA_ROW_HEIGHT_M), 0
        Next row
        d.ForceRebuild3 False
        height = table.GetRowHeight(0): lastRow = firstRow - 1
        For row = firstRow To table.RowCount - 1
            rowHeight = table.GetRowHeight(row)
            If height + rowHeight > available And lastRow >= firstRow Then
                Exit For
            End If
            If height + rowHeight > available Then
                Err.Raise vbObjectError + ERR_NATIVE_BOM, , "A BOM row is taller than the selected sheet."
            End If
            height = height + rowHeight: lastRow = row
        Next row
        For row = 1 To table.RowCount - 1
            table.RowHidden(row) = (row < firstRow Or row > lastRow)
        Next row
        v.SetVisible False, False
        AsDrawing(d).HiddenViewsVisible = False
        AddText d, "Склад виробу (таблиця компонентів)", CONTENT_LEFT_M, sheetHeight - SECTION_HEADER_TOP_MARGIN_M, TEXT_HEIGHT_STANDARD_M, 0
        Log "BOM page=" & page & " rows=" & firstRow & ".." & lastRow
        firstRow = lastRow + 1
    Loop While firstRow < table.RowCount
End Sub

Private Function PropertyValue(ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal aliases As String) As String
    Dim a As Variant, n As Variant, pm As Object, raw As String, value As String
    Dim resolved As Boolean, linked As Boolean, which As Long
    a = Split(aliases, "|")
    For which = 0 To 1
        If which = 0 Then
            Set pm = doc.Extension.CustomPropertyManager(cfg)
        Else
            Set pm = doc.Extension.CustomPropertyManager("")
        End If
        For Each n In a
            raw = "": value = ""
            pm.Get6 CStr(n), False, raw, value, resolved, linked
            If Len(Trim$(value)) > 0 Then
                PropertyValue = value
                Exit Function
            End If
        Next n
    Next which
End Function

Private Function MassText(ByVal doc As SldWorks.ModelDoc2, ByVal body As Object) As String
    Dim mp As SldWorks.MassProperty2, a(0) As Object
    On Error GoTo Unknown
    Set mp = doc.Extension.CreateMassProperty2
    mp.UseSystemUnits = True
    mp.IncludeHiddenBodiesOrComponents = False
    If Not body Is Nothing Then
        Set a(0) = body
        mp.SelectedItems = a
    End If
    mp.Recalculate
    MassText = Format$(mp.Mass, "0.000")
    Exit Function
Unknown:
    MassText = ""
    Warn "Mass unavailable for " & doc.GetTitle
End Function

Private Function OriginalBase(ByVal p As String) As String
    OriginalBase = fs.GetBaseName(p)
    If Left$(OriginalBase, Len(prefix)) = prefix Then
        OriginalBase = Mid$(OriginalBase, Len(prefix) + 1)
    End If
End Function

Private Function SafeFile(ByVal text As String) As String
    Dim bad As Variant, ch As Variant
    bad = Array("<", ">", ":", """", "/", "\", "|", "?", "*", vbTab, vbCr, vbLf)
    For Each ch In bad: text = Replace(text, CStr(ch), "_"): Next ch
    text = Trim$(text)
    Do While Len(text) > 0
        If Right$(text, 1) <> "." And Right$(text, 1) <> " " Then
            Exit Do
        End If
        text = Left$(text, Len(text) - 1)
    Loop
    If Len(text) = 0 Then
        text = "unnamed"
    End If
    SafeFile = Left$(text, MAX_GENERATED_STEM_LENGTH)
End Function

Private Function ValidArtifactStem(ByVal text As String) As Boolean
    ' Validate the persisted basename without sanitizing, truncating or renaming it.
    ' SafeFile limits individual name segments; their combined basename can be longer.
    Dim i As Long, ch As String, code As Long, deviceName As String
    If Len(text) = 0 Or Len(text) + Len(DRAWING_EXTENSION) > MAX_WINDOWS_FILENAME_LENGTH Then
        Exit Function
    End If
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1): code = AscW(ch)
        If code >= 0 And code < ASCII_CONTROL_LIMIT Then
            Exit Function
        End If
        If InStr(1, "<>:""/\|?*", ch, vbBinaryCompare) > 0 Then
            Exit Function
        End If
    Next i
    ' The actual filenames include an extension, so a stem may end in a dot/space.
    deviceName = UCase$(Trim$(Split(text, ".")(0)))
    Select Case deviceName
        Case "", "CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$": Exit Function
    End Select
    If Len(deviceName) = RESERVED_DEVICE_NAME_LENGTH Then
        If Left$(deviceName, RESERVED_DEVICE_PREFIX_LENGTH) = "COM" Or Left$(deviceName, RESERVED_DEVICE_PREFIX_LENGTH) = "LPT" Then
            If InStr(1, "123456789" & ChrW$(WINDOWS_DEVICE_SUPERSCRIPT_ONE) & ChrW$(WINDOWS_DEVICE_SUPERSCRIPT_TWO) & ChrW$(WINDOWS_DEVICE_SUPERSCRIPT_THREE), Right$(deviceName, 1), vbBinaryCompare) > 0 Then
                Exit Function
            End If
        End If
    End If
    ValidArtifactStem = True
End Function

Private Function ScaleText(ByVal viewScaleValue As Double) As String
    If viewScaleValue >= 1 Then
        ScaleText = CStr(viewScaleValue) & ":1"
    Else
        ScaleText = "1:" & CStr(1 / viewScaleValue)
    End If
End Function

Private Function TextFit(ByVal text As String, ByVal width As Double, ByVal height As Double) As Double
    If Len(text) = 0 Then
        TextFit = height
        Exit Function
    End If
    TextFit = MinD(height, width / (Len(text) * TEXT_ESTIMATED_GLYPH_WIDTH_RATIO))
    If TextFit < TEXT_HEIGHT_SMALL_M Then
        TextFit = TEXT_HEIGHT_SMALL_M
    End If
End Function

Private Function WrapText(ByVal text As String, ByVal maxChars As Long) As String
    Dim lines As Variant, line As Variant, s As String, cut As Long, result As String
    lines = Split(Replace(text, vbCr, ""), vbLf)
    For Each line In lines
        s = CStr(line)
        Do While Len(s) > maxChars
            cut = InStrRev(Left$(s, maxChars), " ")
            If cut < maxChars \ 2 Then
                cut = maxChars
            End If
            result = result & Left$(s, cut) & vbCrLf
            s = LTrim$(Mid$(s, cut + 1))
        Loop
        result = result & s & vbCrLf
    Next line
    If Len(result) > 1 Then
        result = Left$(result, Len(result) - 2)
    End If
    WrapText = result
End Function

Private Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then
        MinD = a
    Else
        MinD = b
    End If
End Function

Private Sub ReadSettings(ByVal p As String)
    Dim f As Object, line As String, i As Long
    If Not fs.FileExists(p) Then
        Exit Sub
    End If
    Set f = fs.OpenTextFile(p, FILE_FOR_READING, False, FILE_UNICODE)
    Do Until f.AtEndOfStream
        line = Trim$(f.ReadLine)
        If Len(line) > 0 And Left$(line, 1) <> ";" And Left$(line, 1) <> "[" Then
            i = InStr(line, "=")
            If i > 1 Then
                settings(Trim$(Left$(line, i - 1))) = Trim$(Mid$(line, i + 1))
            End If
        End If
    Loop
    f.Close
End Sub

Private Function Setting(ByVal key As String, ByVal fallback As String) As String
    Setting = fallback
    If settings.Exists(key) Then
        Setting = CStr(settings(key))
    End If
End Function



Private Sub RestorePreferences()
    On Error Resume Next
    If preferenceChanged Then
        sw.SetUserPreferenceToggle SW_PREF_INPUT_DIMENSION, inputDimPreference
        sw.SetUserPreferenceToggle SW_PREF_PDF_EMBED_FONTS, oldEmbedFonts
        sw.SetUserPreferenceToggle SW_PREF_PDF_IN_COLOR, oldPdfColor
        sw.SetUserPreferenceToggle SW_PREF_PDF_HIGH_QUALITY, oldPdfQuality
    End If
    preferenceChanged = False
End Sub

Private Sub CloseSnapshots()
    Dim docs As Variant, item As Variant, doc As SldWorks.ModelDoc2, pass As Long
    On Error Resume Next
    For pass = SW_DOC_ASSEMBLY To SW_DOC_PART Step -1
        docs = sw.GetDocuments
        If IsEmpty(docs) Then
            Exit Sub
        End If
        For Each item In docs
            Set doc = item
            If InStr(1, doc.GetPathName, runDir & MODELS_RELATIVE_DIRECTORY, vbTextCompare) = 1 Then
                If doc.GetType = pass Then
                    sw.CloseDoc doc.GetTitle
                End If
            End If
        Next item
    Next pass
End Sub

Private Sub Log(ByVal text As String)
    HideBackgroundFrame
    If Len(runDir) > 0 Then
        Set logFile = fs.OpenTextFile(runDir & RUN_LOG_RELATIVE_PATH, FILE_FOR_APPENDING, True, FILE_UNICODE)
        logFile.WriteLine Format$(Now, "hh:nn:ss") & " " & text
        logFile.Close: Set logFile = Nothing
    End If
    ReportProgress STATE_RUNNING, text
    Debug.Print text
End Sub

Private Sub ReportProgress(ByVal state As String, ByVal message As String)
    Dim stream As Object
    If Len(progressPath) = 0 Then
        Exit Sub
    End If
    ' Close after every record: the VBS monitor reads while RunMacro2 is blocked.
    On Error Resume Next
    message = Replace(Replace(Replace(message, vbTab, " "), vbCr, " "), vbLf, " ")
    Set stream = fs.OpenTextFile(progressPath, FILE_FOR_APPENDING, True, FILE_UNICODE)
    stream.WriteLine Format$(Now, "hh:nn:ss") & vbTab & state & vbTab & (completed + failures) & vbTab & plannedJobs & vbTab & completed & vbTab & failures & vbTab & warningCount & vbTab & runDir & vbTab & message
    stream.Close
End Sub

Private Sub Warn(ByVal text As String)
    warningCount = warningCount + 1
    Log "WARN " & text
End Sub

Private Function AsDrawing(ByVal doc As SldWorks.ModelDoc2) As SldWorks.DrawingDoc
    Set AsDrawing = doc
End Function
Private Function AsPart(ByVal doc As SldWorks.ModelDoc2) As SldWorks.PartDoc
    Set AsPart = doc
End Function
Private Function AsAssembly(ByVal doc As SldWorks.ModelDoc2) As SldWorks.AssemblyDoc
    Set AsAssembly = doc
End Function


Private Function FileStamp(ByVal path As String) As String
    Dim f As Object
    Set f = fs.GetFile(path)
    FileStamp = CStr(f.Size) & "|" & Format$(f.DateLastModified, FILE_STAMP_FORMAT)
End Function

Private Function RenderSettingsSignature() As String
    Dim key As Variant
    RenderSettingsSignature = RENDER_SIGNATURE_PREFIX & paperFormat & "|Template=" & templatePath & "|TemplateStamp=" & FileStamp(templatePath)
    For Each key In Array("Font", "Organization", "AutoDimensions", "MaxEdgesForAutoDimensions", "ImportModelDimensions", "UpAxis", "DimensionSpacing", "Liter1", "Liter2", "Liter3", "DrawnBy", "DrawnDate", "CheckedBy", "CheckedDate", "TechnicalControlBy", "TechnicalControlDate", "NormControlBy", "NormControlDate", "ApprovedBy", "ApprovedDate")
        RenderSettingsSignature = RenderSettingsSignature & "|" & CStr(key) & "=" & Setting(CStr(key), "")
    Next key
End Function

Private Sub SelectPlan()
    Dim selected As Object, key As Variant, j As Object, n As Long
    Set selected = CreateObject("Scripting.Dictionary"): selected.CompareMode = 1
    For Each key In jobs.Keys
        Set j = jobs(key)
        If ShouldRunJob(j) Then
            n = n + 1
            j.Add "number", n: j.Add "status", STATE_PENDING: j.Add "stage", "Queued"
            j.Add "drawingBytes", "0": j.Add "pdfBytes", "0"
            j.Add "out", Format$(n, JOB_NUMBER_FORMAT) & "_" & SafeFile(OriginalBase(CStr(j("path"))) & "_" & CStr(j("cfg")))
            If CBool(j("multi")) Then
                j("out") = j("out") & "_" & SafeFile(CStr(j("cut")) & "_" & CStr(j("body")))
            End If
            selected.Add CStr(n), j
            If CLng(Setting("MaxJobs", "0")) > 0 Then
                If n >= CLng(Setting("MaxJobs", "0")) Then
                    Exit For
                End If
            End If
        End If
    Next key
    Set jobs = selected
End Sub

Private Function Cell(ByVal value As Variant) As String
    Cell = Replace(Replace(Replace(CStr(value), vbTab, " "), vbCr, " "), vbLf, " ")
End Function

Private Sub SaveCheckpoint()
    Dim stream As Object, tempPath As String, key As Variant, j As Object, line As String, field As Variant
    tempPath = checkpointPath & ATOMIC_TEMP_SUFFIX
    Set stream = fs.CreateTextFile(tempPath, True, True)
    stream.WriteLine CHECKPOINT_HEADER_TAG & vbTab & CHECKPOINT_VERSION & vbTab & sourcePath & vbTab & sourceStamp & vbTab & snapshotPath & vbTab & snapshotConfig & vbTab & prefix & vbTab & renderSignature
    For Each key In jobs.Keys
        Set j = jobs(key): line = CHECKPOINT_JOB_TAG
        For Each field In Array("number", "path", "cfg", "body", "cut", "qty", "sheet", "multi", "type", "status", "out", "stage", "drawingBytes", "pdfBytes")
            line = line & vbTab & Cell(j(CStr(field)))
        Next field
        stream.WriteLine line
    Next key
    stream.Close
    ' Atomic replacement on the same volume. Old checkpoint survives an interrupted write.
    If MoveFileExW(StrPtr(tempPath), StrPtr(checkpointPath), ATOMIC_REPLACE_WRITE_THROUGH) = 0 Then
        Err.Raise vbObjectError + ERR_GENERATION_DATA, , "Cannot commit generation-state.tsv"
    End If
End Sub

Private Sub LoadCheckpoint()
    Dim stream As Object, fields As Variant, names As Variant, j As Object, i As Long
    Set stream = fs.OpenTextFile(checkpointPath, FILE_FOR_READING, False, FILE_UNICODE)
    If stream.AtEndOfStream Then
        Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Empty checkpoint. Use /reset to generate again."
    End If
    fields = Split(stream.ReadLine, vbTab)
    If UBound(fields) <> CHECKPOINT_HEADER_LAST_FIELD Then
        Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Invalid checkpoint header."
    End If
    If fields(HEADER_TAG_FIELD) <> CHECKPOINT_HEADER_TAG Or fields(HEADER_VERSION_FIELD) <> CHECKPOINT_VERSION Then
        Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Unsupported checkpoint version."
    End If
    If StrComp(fields(HEADER_SOURCE_PATH), sourcePath, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Checkpoint belongs to another input. Choose another output folder."
    End If
    If fields(HEADER_SOURCE_STAMP) <> sourceStamp Then
        Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Input file changed. Use /reset for a new snapshot."
    End If
    If fields(HEADER_RENDER_SIGNATURE) <> renderSignature Then
        Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Drawing settings or generator version changed. Use /reset."
    End If
    snapshotPath = fields(HEADER_SNAPSHOT_PATH): snapshotConfig = fields(HEADER_SNAPSHOT_CONFIG): prefix = fields(HEADER_PREFIX)
    names = Array("number", "path", "cfg", "body", "cut", "qty", "sheet", "multi", "type", "status", "out", "stage", "drawingBytes", "pdfBytes")
    Do While Not stream.AtEndOfStream
        fields = Split(stream.ReadLine, vbTab)
        If UBound(fields) <> CHECKPOINT_JOB_LAST_FIELD Then
            Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Invalid checkpoint job."
        End If
        Set j = CreateObject("Scripting.Dictionary")
        For i = 0 To UBound(names): j.Add CStr(names(i)), fields(i + 1): Next i
        j("sheet") = CBool(j("sheet")): j("multi") = CBool(j("multi"))
        j("path") = fs.GetAbsolutePathName(CStr(j("path")))
        If Not ValidArtifactStem(CStr(j("out"))) Then
            Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Invalid checkpoint artifact name for job " & CStr(j("number")) & ": " & CStr(j("out"))
        End If
        If InStr(1, CStr(j("path")), runDir & MODELS_RELATIVE_DIRECTORY, vbTextCompare) <> 1 Then
            Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Snapshot path is outside output/Models."
        End If
        If Not fs.FileExists(CStr(j("path"))) Then
            Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Snapshot model missing. Use /reset."
        End If
        If j("status") = STATE_DONE Then
            If ArtifactsValid(j) Then
                completed = completed + 1
            Else
                j("status") = STATE_PENDING: j("stage") = "Artifact missing or changed; regenerate"
            End If
        Else
            j("status") = STATE_PENDING
        End If
        jobs.Add CStr(j("number")), j
    Loop
    stream.Close
    If jobs.Count = 0 Then
        Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Checkpoint has no jobs."
    End If
End Sub

Private Function ArtifactsValid(ByVal j As Object) As Boolean
    Dim native As String, pdf As String
    native = runDir & DRAWINGS_RELATIVE_DIRECTORY & CStr(j("out")) & DRAWING_EXTENSION
    pdf = runDir & PDF_RELATIVE_DIRECTORY & CStr(j("out")) & PDF_EXTENSION
    If Not fs.FileExists(native) Or Not fs.FileExists(pdf) Then
        Exit Function
    End If
    ArtifactsValid = (CStr(fs.GetFile(native).Size) = CStr(j("drawingBytes")) And CStr(fs.GetFile(pdf).Size) = CStr(j("pdfBytes")) And CDbl(j("pdfBytes")) > 0)
End Function

Private Sub JobStage(ByVal j As Object, ByVal stage As String)
    j("stage") = stage
    SaveCheckpoint
    Log "STAGE " & CStr(j("out")) & " | " & stage
End Sub

Private Sub AppendManifest(ByVal line As String)
    Set manifest = fs.OpenTextFile(runDir & MANIFEST_RELATIVE_PATH, FILE_FOR_APPENDING, True, FILE_UNICODE)
    manifest.WriteLine line
    manifest.Close: Set manifest = Nothing
End Sub

Private Sub UprightView(ByVal v As SldWorks.View)
    Dim axis(2) As Double, vector As SldWorks.MathVector, util As SldWorks.MathUtility
    Dim p As Variant, up As String, angle As Double, direction As Double
    up = UCase$(Setting("UpAxis", DEFAULT_UP_AXIS)): direction = 1
    If Left$(up, 1) = "-" Then
        direction = -1
    End If
    Select Case Right$(up, 1)
        Case "X": axis(0) = direction
        Case "Y": axis(1) = direction
        Case "Z": axis(2) = direction
        Case Else: Err.Raise vbObjectError + ERR_UP_AXIS, , "Invalid UpAxis."
    End Select
    Set util = sw.GetMathUtility: Set vector = util.CreateVector(axis)
    Set vector = vector.MultiplyTransform(v.ModelToViewTransform): p = vector.ArrayData
    If Sqr(p(0) ^ 2 + p(1) ^ 2) < MIN_PROJECTED_UP_VECTOR_LENGTH Then ' looking along up axis
        Exit Sub
    End If
    If Abs(p(1)) < UP_VECTOR_TOLERANCE Then
        angle = Sgn(p(0)) * PI / 2
    Else
        angle = Atn(p(0) / p(1))
        If p(1) < 0 Then
            angle = angle + PI
        End If
    End If
    v.Angle = v.Angle + angle
End Sub

Private Sub ImportModelItems(ByVal d As SldWorks.ModelDoc2)
    Dim annotations As Variant, types As Long
    On Error GoTo Failed
    If Setting("ImportModelDimensions", "0") = "0" Then
        Exit Sub
    End If
    types = SW_MODEL_ITEMS_MARKED_DIMENSIONS + SW_MODEL_ITEMS_HOLE_PROFILES + SW_MODEL_ITEMS_HOLE_LOCATIONS ' marked dimensions and Hole Wizard dimensions
    If LCase$(Setting("ImportModelDimensions", "0")) = "all" Then
        types = SW_MODEL_ITEMS_ALL_DIMENSIONS
    End If
    AsDrawing(d).ActivateView "": d.ClearSelection2 True
    annotations = AsDrawing(d).InsertModelAnnotations3(0, types, True, True, False, False)
    Exit Sub
Failed:
    Warn "Import model dimensions: " & Err.Description
End Sub

Private Sub DimensionFeatures(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View, ByVal detail As Boolean)
    Dim count As Long
    If Setting("AutoDimensions", "1") <> "1" Then
        Exit Sub
    End If
    count = CoordinateDimensions(d, v, detail)
    count = DiameterDimensions(d, v)
    count = RadiusDimensions(d, v)
End Sub

Private Function RadiusDimensions(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View) As Long
    Dim comps As Variant, edges As Variant, ci As Variant, ei As Variant, edge As SldWorks.Edge
    Dim curve As SldWorks.Curve, cp As Variant, first As SldWorks.Vertex, last As SldWorks.Vertex
    Dim p As Variant, q As Variant, seen As Object, key As String, outline As Variant
    Dim util As SldWorks.MathUtility, vector As SldWorks.MathVector, axis(2) As Double, dd As SldWorks.DisplayDimension
    On Error GoTo Failed
    Set seen = CreateObject("Scripting.Dictionary"): Set util = sw.GetMathUtility
    comps = v.GetVisibleComponents: outline = v.GetOutline
    If IsEmpty(comps) Then
        Exit Function
    End If
    For Each ci In comps
        edges = v.GetVisibleEntities2(ci, SW_VISIBLE_ENTITY_EDGE)
        If Not IsEmpty(edges) Then
            For Each ei In edges
                Set edge = ei: Set curve = edge.GetCurve
                If curve.IsCircle Then
                    Set first = edge.GetStartVertex: Set last = edge.GetEndVertex
                    If first Is Nothing Or last Is Nothing Then
                        GoTo NextArc
                    End If
                    p = first.GetPoint: q = last.GetPoint
                    If Abs(p(0) - q(0)) + Abs(p(1) - q(1)) + Abs(p(2) - q(2)) < CIRCLE_CLOSURE_TOLERANCE_M Then
                        GoTo NextArc
                    End If
                    cp = curve.CircleParams
                    axis(0) = cp(CIRCLE_AXIS_X_INDEX): axis(1) = cp(CIRCLE_AXIS_Y_INDEX): axis(2) = cp(CIRCLE_AXIS_Z_INDEX)
                    Set vector = util.CreateVector(axis): Set vector = vector.MultiplyTransform(v.ModelToViewTransform)
                    p = vector.ArrayData
                    If Abs(p(2)) < FACE_NORMAL_ALIGNMENT_COSINE * Sqr(p(0) ^ 2 + p(1) ^ 2 + p(2) ^ 2) Then
                        GoTo NextArc
                    End If
                    key = Format$(cp(CIRCLE_RADIUS_INDEX), GEOMETRY_KEY_FORMAT)
                    If seen.Exists(key) Then
                        GoTo NextArc
                    End If
                    seen.Add key, True
                    AsDrawing(d).ActivateView "": d.ClearSelection2 True
                    If v.SelectEntity(edge, False) Then
                        Set dd = d.AddRadialDimension2(outline(OUTLINE_RIGHT_INDEX) + RADIUS_LABEL_OFFSET_M, outline(OUTLINE_BOTTOM_INDEX) + RADIUS_LABEL_PITCH_M * RadiusDimensions, 0)
                        If Not dd Is Nothing Then
                            If ValidateDimension(d, v, dd, cp(CIRCLE_RADIUS_INDEX) * v.ScaleDecimal) Then
                                RadiusDimensions = RadiusDimensions + 1
                            End If
                        End If
                    End If
                End If
NextArc:
            Next ei
        End If
    Next ci
    Exit Function
Failed:
    Warn "Radius dimensions " & v.Name & ": " & Err.Description
End Function

Private Sub AddFlatDimensionSheet(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal sourceView As SldWorks.View, ByVal code As String, ByVal name As String, ByVal material As String, ByVal mass As String)
    Dim v As SldWorks.View, detailScale As Double, sheet As SldWorks.Sheet, rc As Long, bounds As Variant, position As Variant
    If Setting("AutoDimensions", "1") <> "1" Then
        Exit Sub
    End If
    NewFormattedSheet d, "Розміри розгортки"
    Set v = AsDrawing(d).CreateDrawViewFromModelView3(doc.GetPathName, sourceView.GetOrientationName, LX(DETAIL_INITIAL_X_M), LY(DETAIL_INITIAL_Y_M), 0)
    If v Is Nothing Then
        Err.Raise vbObjectError + ERR_DETAIL_SHEET, , "Cannot create flat dimension view."
    End If
    v.ReferencedConfiguration = sourceView.ReferencedConfiguration
    v.UseSheetScale = False: v.ScaleDecimal = 1: v.Angle = sourceView.Angle
    d.ForceRebuild3 False
    bounds = ProjectedFlatBounds(v)
    v.ScaleDecimal = StandardScale(VIEW_FIT_SAFETY_FACTOR * MinD((sheetWidth - DETAIL_HORIZONTAL_RESERVED_M) / (bounds(OUTLINE_RIGHT_INDEX) - bounds(OUTLINE_LEFT_INDEX)), (sheetHeight - DETAIL_VERTICAL_RESERVED_M) / (bounds(OUTLINE_TOP_INDEX) - bounds(OUTLINE_BOTTOM_INDEX))))
    bounds = ProjectedFlatBounds(v): position = v.Position
    position(0) = position(0) + (sheetWidth + DETAIL_CENTER_X_BIAS_M) / 2 - (bounds(OUTLINE_LEFT_INDEX) + bounds(OUTLINE_RIGHT_INDEX)) / 2
    position(1) = position(1) + DETAIL_CENTER_Y_BASE_M + (sheetHeight - DETAIL_CENTER_HEIGHT_REFERENCE_M) / 2 - (bounds(OUTLINE_BOTTOM_INDEX) + bounds(OUTLINE_TOP_INDEX)) / 2
    v.Position = position
    Set sheet = AsDrawing(d).GetCurrentSheet
    detailScale = v.ScaleDecimal: SetPaperScale sheet, detailScale
    v.SetDisplayMode3 False, SW_DISPLAY_HIDDEN_LINES_REMOVED, False, False
    d.ForceRebuild3 False
    rc = OverallDimensions(d, v, True)
    DimensionFeatures d, v, True
    AsDrawing(d).ActivateView "": d.ClearSelection2 True
    RefreshTemplateFields d
    AddText d, "РОЗГОРТКА - КООРДИНАТНІ РОЗМІРИ ТА ОТВОРИ", CONTENT_LEFT_M, sheetHeight - SECTION_HEADER_TOP_MARGIN_M, TEXT_HEIGHT_SMALL_M, 0
End Sub

Private Sub WriteDimensionCoverage(ByVal d As SldWorks.ModelDoc2, ByVal outName As String)
    Dim names As Variant, name As Variant, v As SldWorks.View, stream As Object, path As String, temp As String
    Dim previous As String, line As Variant
    path = runDir & COVERAGE_RELATIVE_PATH
    If fs.FileExists(path) Then
        Set stream = fs.OpenTextFile(path, FILE_FOR_READING, False, FILE_UNICODE)
        previous = stream.ReadAll
        stream.Close
    End If
    temp = path & ATOMIC_TEMP_SUFFIX
    Set stream = fs.CreateTextFile(temp, True, True)
    stream.WriteLine "Drawing" & vbTab & "Sheet" & vbTab & "View" & vbTab & "Edges" & vbTab & "Dimensions" & vbTab & "Configuration"
    For Each line In Split(previous, vbCrLf)
        If Len(CStr(line)) > 0 And Left$(CStr(line), COVERAGE_HEADER_PREFIX_LENGTH) <> "Drawing" & vbTab Then
            If Left$(CStr(line), Len(outName) + 1) <> outName & vbTab Then
                stream.WriteLine line
            End If
        End If
    Next line
    names = AsDrawing(d).GetSheetNames
    For Each name In names
        AsDrawing(d).ActivateSheet CStr(name)
        Set v = AsDrawing(d).GetFirstView: Set v = v.GetNextView
        Do While Not v Is Nothing
            stream.WriteLine outName & vbTab & name & vbTab & v.Name & vbTab & VisibleEdgeCount(v) & vbTab & v.GetDisplayDimensionCount & vbTab & v.ReferencedConfiguration
            Set v = v.GetNextView
        Loop
    Next name
    stream.Close
    If MoveFileExW(StrPtr(temp), StrPtr(path), ATOMIC_REPLACE_WRITE_THROUGH) = 0 Then
        Err.Raise vbObjectError + ERR_GENERATION_DATA, , "Cannot commit dimension-coverage.tsv"
    End If
End Sub
' Geometry-driven native dimensions. Every unique visible contour coordinate is included.
' Circle centers participate as real circular edges, never as loose sketch points.
Private Function CoordinateDimensions(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View, ByVal detail As Boolean) As Long
    Dim xs As Object, ys As Object, comps As Variant, ci As Variant, edges As Variant, ei As Variant
    Dim edge As SldWorks.Edge, curve As SldWorks.Curve, vertex As SldWorks.Vertex, util As SldWorks.MathUtility
    Dim point As SldWorks.MathPoint, cp As Variant, p As Variant, q As Variant, a(2) As Double, endpoint As Long
    Dim normal As SldWorks.MathVector, radius As Double, fullCircle As Boolean
    On Error GoTo Failed
    Set xs = CreateObject("Scripting.Dictionary"): Set ys = CreateObject("Scripting.Dictionary")
    Set util = sw.GetMathUtility
    comps = v.GetVisibleComponents
    If IsEmpty(comps) Then
        Exit Function
    End If
    For Each ci In comps
        edges = v.GetVisibleEntities2(ci, SW_VISIBLE_ENTITY_EDGE)
        If Not IsEmpty(edges) Then
            For Each ei In edges
                Set edge = ei: Set curve = edge.GetCurve
                fullCircle = False
                If curve.IsCircle Then
                    Dim first As SldWorks.Vertex, last As SldWorks.Vertex
                    Set first = edge.GetStartVertex: Set last = edge.GetEndVertex
                    fullCircle = (first Is Nothing Or last Is Nothing)
                    If Not first Is Nothing And Not last Is Nothing Then
                        p = first.GetPoint: q = last.GetPoint
                        fullCircle = Abs(p(0) - q(0)) + Abs(p(1) - q(1)) + Abs(p(2) - q(2)) < CIRCLE_CLOSURE_TOLERANCE_M
                    End If
                    If fullCircle Then
                        cp = curve.CircleParams
                        a(0) = cp(CIRCLE_AXIS_X_INDEX): a(1) = cp(CIRCLE_AXIS_Y_INDEX): a(2) = cp(CIRCLE_AXIS_Z_INDEX)
                        Set normal = util.CreateVector(a): Set normal = normal.MultiplyTransform(v.ModelToViewTransform)
                        p = normal.ArrayData
                        If Abs(p(2)) > FACE_NORMAL_ALIGNMENT_COSINE * Sqr(p(0) ^ 2 + p(1) ^ 2 + p(2) ^ 2) Then
                            a(0) = cp(CIRCLE_CENTER_X_INDEX): a(1) = cp(CIRCLE_CENTER_Y_INDEX): a(2) = cp(CIRCLE_CENTER_Z_INDEX)
                            Set point = util.CreatePoint(a): Set point = point.MultiplyTransform(v.ModelToViewTransform)
                            p = point.ArrayData
                            AddCoordinate xs, edge, p(0), v.ScaleDecimal
                            AddCoordinate ys, edge, p(1), v.ScaleDecimal
                        End If
                    End If
                End If
                If Not fullCircle Then
                    For endpoint = 0 To 1
                        Set vertex = Nothing
                        If endpoint = 0 Then
                            Set vertex = edge.GetStartVertex
                        Else
                            Set vertex = edge.GetEndVertex
                        End If
                        If Not vertex Is Nothing Then
                            Set point = util.CreatePoint(vertex.GetPoint): Set point = point.MultiplyTransform(v.ModelToViewTransform)
                            p = point.ArrayData
                            AddCoordinate xs, vertex, p(0), v.ScaleDecimal
                            AddCoordinate ys, vertex, p(1), v.ScaleDecimal
                        End If
                    Next endpoint
                End If
            Next ei
        End If
    Next ci
    Dim geometryBounds As Variant, key As Variant, record As Object, xmin As Double, xmax As Double, ymin As Double, ymax As Double
    xmin = SEARCH_INITIAL_LIMIT: xmax = -SEARCH_INITIAL_LIMIT: ymin = SEARCH_INITIAL_LIMIT: ymax = -SEARCH_INITIAL_LIMIT
    For Each key In xs.Keys
        Set record = xs(key)
        xmin = MinD(xmin, CDbl(record("coordinate")))
        xmax = -MinD(-xmax, -CDbl(record("coordinate")))
    Next key
    For Each key In ys.Keys
        Set record = ys(key)
        ymin = MinD(ymin, CDbl(record("coordinate")))
        ymax = -MinD(-ymax, -CDbl(record("coordinate")))
    Next key
    geometryBounds = Array(xmin, ymin, xmax, ymax)
    If detail Then
        geometryBounds = ProjectedFlatBounds(v)
    End If
    CoordinateDimensions = CoordinateAxis(d, v, xs, 0, geometryBounds) + CoordinateAxis(d, v, ys, 1, geometryBounds)
    Log "COORDINATES " & v.Name & " x=" & xs.Count & " y=" & ys.Count & " dimensions=" & CoordinateDimensions
    Exit Function
Failed:
    Warn "Coordinate dimensions " & v.Name & ": " & Err.Description
End Function

Private Sub AddCoordinate(ByVal coordinates As Object, ByVal entity As Object, ByVal value As Double, ByVal viewScale As Double)
    Dim key As String, record As Object
    key = CStr(Round(value / viewScale, COORDINATE_KEY_DECIMALS)) ' 0.01 mm, matching displayed precision
    If coordinates.Exists(key) Then
        Exit Sub
    End If
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "entity", entity: record.Add "coordinate", value
    coordinates.Add key, record
End Sub

Private Function CoordinateAxis(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View, ByVal coordinates As Object, ByVal axis As Long, ByVal geometryBounds As Variant) As Long
    Dim keys As Variant, i As Long, j As Long, temp As Variant, first As Object, last As Object, entity As Object
    Dim outline As Variant, lane As Long, occupied(COORDINATE_LAST_LANE) As Double, textWidth As Double, center As Double
    Dim base As Double, location As Double, span As Double, maxLane As Long, lanes As Long, rc As Long
    Dim candidate As Double, bestCenter As Double, bestCost As Double, cost As Double, bestLane As Long, limit As Double
    Dim dd As SldWorks.DisplayDimension, dimension As SldWorks.Dimension, text As String
    If coordinates.Count < COORDINATE_MIN_KEY_COUNT Then ' two coordinates are already dimensioned by OverallDimensions
        Exit Function
    End If
    keys = coordinates.Keys
    For i = LBound(keys) + 1 To UBound(keys)
        temp = keys(i): j = i - 1
        Do While j >= LBound(keys)
            If CDbl(keys(j)) <= CDbl(temp) Then
                Exit Do
            End If
            keys(j + 1) = keys(j): j = j - 1
        Loop
        keys(j + 1) = temp
    Next i
    outline = geometryBounds
    For i = 0 To COORDINATE_LAST_LANE: occupied(i) = -EMPTY_LANE_LIMIT: Next i
    If axis = 0 Then
        base = outline(OUTLINE_TOP_INDEX) + COORDINATE_BASE_OFFSET_M
        maxLane = Int((sheetHeight - COORDINATE_TOP_MARGIN_M - base) / COORDINATE_LANE_PITCH_M)
    Else
        base = outline(OUTLINE_LEFT_INDEX) - COORDINATE_BASE_OFFSET_M
        maxLane = Int((base - COORDINATE_LEFT_LIMIT_M) / COORDINATE_LANE_PITCH_M)
    End If
    If maxLane > COORDINATE_LAST_LANE Then
        maxLane = COORDINATE_LAST_LANE
    End If
    If maxLane < 0 Then
        maxLane = 0
    End If
    For i = LBound(keys) To UBound(keys) - 1
        Set first = coordinates(keys(i)): Set last = coordinates(keys(i + 1))
        span = CDbl(last("coordinate")) - CDbl(first("coordinate"))
        If span < v.ScaleDecimal * MIN_DIMENSION_SPAN_M Then
            GoTo NextCoordinate
        End If
        center = (CDbl(first("coordinate")) + CDbl(last("coordinate"))) / 2
        text = Format$(span / v.ScaleDecimal * MILLIMETRES_PER_METRE, DIMENSION_NUMBER_FORMAT)
        textWidth = Len(text) * COORDINATE_GLYPH_WIDTH_M + COORDINATE_LABEL_PADDING_M
        If axis = 0 Then
            limit = sheetWidth - COORDINATE_RIGHT_MARGIN_M
        Else
            limit = sheetHeight - COORDINATE_TOP_MARGIN_M
        End If
        bestLane = -1: bestCost = SEARCH_INITIAL_LIMIT
        For lane = 0 To maxLane
            candidate = -MinD(-center, -(occupied(lane) + textWidth / 2 + COORDINATE_LABEL_GAP_M))
            If candidate + textWidth / 2 <= limit Then
                cost = Abs(candidate - center) + lane * COORDINATE_LANE_PENALTY_M
                If cost < bestCost Then
                    bestLane = lane
                    bestCenter = candidate
                    bestCost = cost
                End If
            End If
        Next lane
        If bestLane < 0 Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Dimension labels do not fit selected sheet: " & v.Name
        End If
        lane = bestLane: center = bestCenter
        occupied(lane) = center + textWidth / 2
        AsDrawing(d).ActivateView "": d.ClearSelection2 True
        If v.SelectEntity(first("entity"), False) And v.SelectEntity(last("entity"), True) Then
            If axis = 0 Then
                location = base + lane * COORDINATE_LANE_PITCH_M
                Set dd = d.AddHorizontalDimension2(center, location, 0)
            Else
                location = base - lane * COORDINATE_LANE_PITCH_M
                Set dd = d.AddVerticalDimension2(location, center, 0)
            End If
            If Not dd Is Nothing Then
                Set dimension = dd.GetDimension2(SW_DIMENSION_INDEX_FIRST)
                If Not dimension Is Nothing Then
                    dimension.SetArcEndCondition 1, SW_ARC_CENTER: dimension.SetArcEndCondition 2, SW_ARC_CENTER
                End If
                If ValidateDimension(d, v, dd, span) Then
                    CoordinateAxis = CoordinateAxis + 1
                End If
            Else
                Warn "Coordinate could not be dimensioned: " & v.Name & " span=" & text
            End If
        Else
            Warn "Coordinate selection failed: " & v.Name & " span=" & text
        End If
NextCoordinate:
    Next i
    d.ClearSelection2 True
End Function


Public Sub main()
    Dim source As SldWorks.ModelDoc2, p As String, cfg As String, key As Variant, job As Object
    Dim errNumber As Long, errText As String
    On Error GoTo Fatal
    Set sw = Application.SldWorks
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set settings = CreateObject("Scripting.Dictionary")
    settings.CompareMode = 1
    progressPath = "": runDir = "": plannedJobs = 0
    backgroundOwned = False: attempts = 0: resumeRun = False
    completed = 0: failures = 0: warningCount = 0: jobNumber = 0: viewCounter = 0
    ReadSettings sw.GetCurrentMacroPathFolder & SETTINGS_RELATIVE_PATH
    progressPath = Setting("LauncherProgress", "")
    ReportProgress STATE_RUNNING, "VBA started"
    If Setting("LauncherBackground", "0") = "1" Then
        If CStr(sw.GetProcessID) = Setting("LauncherForbiddenPID", "0") Then
            Err.Raise vbObjectError + ERR_APPLICATION_STATE, , "Background startup reused the user's SolidWorks process."
        End If
        backgroundOwned = True
        ' Retain the background application when the last visible drawing closes.
        sw.UserControlBackground = True
        Dim hiddenFrame As SldWorks.Frame, inputType As Long, inputError As Long, inputWarning As Long
        Set hiddenFrame = sw.Frame: hiddenFrame.KeepInvisible = True
        backgroundLeft = sw.FrameLeft: backgroundTop = sw.FrameTop
        backgroundWidth = sw.FrameWidth: backgroundHeight = sw.FrameHeight
        HideBackgroundFrame
        inputType = SW_DOC_PART
        If LCase$(fs.GetExtensionName(Setting("LauncherInput", ""))) = ASSEMBLY_EXTENSION Then
            inputType = SW_DOC_ASSEMBLY
        End If
        sw.SetCurrentWorkingDirectory fs.GetParentFolderName(Setting("LauncherInput", ""))
        Set original = sw.OpenDoc6(Setting("LauncherInput", ""), inputType, SW_OPEN_SILENT_READ_ONLY, "", inputError, inputWarning)
        If original Is Nothing Then
            Err.Raise vbObjectError + ERR_OPEN_INPUT_OR_SNAPSHOT, , "Cannot open background input: " & inputError
        End If
        PrepareDocumentWindow original
        Set source = sw.ActivateDoc3(fs.GetFileName(original.GetPathName), False, SW_REBUILD_USER_DECISION, inputError)
        ReportProgress STATE_RUNNING, "Hidden SolidWorks PID=" & sw.GetProcessID
    Else
        Set original = sw.ActiveDoc
    End If
    If original Is Nothing Then
        Err.Raise vbObjectError + ERR_SOURCE_NOT_OPEN, , "Open a saved part or assembly first."
    End If
    If original.GetType <> SW_DOC_PART And original.GetType <> SW_DOC_ASSEMBLY Then
        Err.Raise vbObjectError + ERR_SOURCE_TYPE, , "Activate a part or assembly, not a drawing."
    End If
    p = original.GetPathName
    If Len(p) = 0 Then
        Err.Raise vbObjectError + ERR_SOURCE_NOT_SAVED, , "Save the source model before running."
    End If
    If Len(Setting("LauncherInput", "")) > 0 Then
        If StrComp(p, Setting("LauncherInput", ""), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + ERR_APPLICATION_STATE, , "The requested input model is not active."
        End If
    End If
    sourceDir = fs.GetParentFolderName(p)
    ResolveDrawingTemplate
    ValidateDrawingTemplate
    runDir = Setting("LauncherOutput", "")
    If Len(runDir) = 0 Then
        runDir = sourceDir & OUTPUT_FOLDER_RELATIVE_PREFIX & Format$(Now, TIMESTAMP_FORMAT)
        If fs.FolderExists(runDir) Then
            Err.Raise vbObjectError + ERR_OUTPUT_DIRECTORY, , "Output folder already exists. Try again in a second."
        End If
    Else
        runDir = fs.GetAbsolutePathName(runDir)
    End If
    EnsureFolder runDir
    EnsureFolder runDir & MODELS_RELATIVE_PATH
    EnsureFolder runDir & DRAWINGS_RELATIVE_PATH
    EnsureFolder runDir & PDF_RELATIVE_PATH
    runtimeFormatPath = runDir & SHEET_FORMAT_RELATIVE_PATH
    sourcePath = original.GetPathName
    sourceStamp = FileStamp(sourcePath)
    checkpointPath = runDir & CHECKPOINT_RELATIVE_PATH
    renderSignature = RenderSettingsSignature()
    Set jobs = CreateObject("Scripting.Dictionary"): jobs.CompareMode = 1
    resumeRun = fs.FileExists(checkpointPath)
    Dim er As Long, wr As Long
    If resumeRun Then
        LoadCheckpoint
        Log "RESUME existing plan; completed=" & completed & " total=" & jobs.Count
    Else
        ' A new snapshot must reflect a saved source. Resume uses only saved Models/.
        If original.GetSaveFlag Then
            Err.Raise vbObjectError + ERR_UNSAVED_SOURCE, , "Source has unsaved changes. Save or discard them yourself before creating a new snapshot."
        End If
        Set logFile = fs.CreateTextFile(runDir & RUN_LOG_RELATIVE_PATH, True, True)
        logFile.Close: Set logFile = Nothing
        Set manifest = fs.CreateTextFile(runDir & MANIFEST_RELATIVE_PATH, True, True)
        manifest.WriteLine "Number" & vbTab & "Model" & vbTab & "Configuration" & vbTab & "Body" & vbTab & "Quantity" & vbTab & "Drawing" & vbTab & "Status"
        manifest.Close: Set manifest = Nothing
        Log "START SolidWorks " & sw.RevisionNumber & " | " & sourcePath
        cfg = original.ConfigurationManager.ActiveConfiguration.Name
        snapshotConfig = cfg
        snapshotPath = MakeSnapshot(original)
        Set rootModel = sw.OpenDoc6(snapshotPath, original.GetType, SW_OPEN_SILENT, cfg, er, wr)
        If rootModel Is Nothing Then
            Err.Raise vbObjectError + ERR_OPEN_INPUT_OR_SNAPSHOT, , "Cannot open snapshot: " & er
        End If
        Set quantities = CreateObject("Scripting.Dictionary"): quantities.CompareMode = 1
        If rootModel.GetType = SW_DOC_ASSEMBLY Then
            AsAssembly(rootModel).ResolveAllLightWeightComponents False
            Inventory rootModel
        End If
        AddModelJobs rootModel, cfg, 1
        For Each key In quantities.Keys
            Set job = quantities(key): Set source = job("doc")
            AddModelJobs source, CStr(job("cfg")), CLng(job("qty"))
        Next key
        SelectPlan
        SaveCheckpoint
    End If
    plannedJobs = jobs.Count
    Log "JOBS selected=" & plannedJobs
    If plannedJobs = 0 Then
        Err.Raise vbObjectError + ERR_EMPTY_PLAN, , "No drawing jobs match the settings."
    End If
    ' Inventory is now plain data. Release all model COM references before drawing.
    Set source = Nothing: Set rootModel = Nothing: Set quantities = Nothing
    For Each key In jobs.Keys
        Set job = jobs(key)
        If job.Exists("doc") Then
            job.Remove "doc"
        End If
    Next key
    CloseSnapshots
    If backgroundOwned Then
        sw.CloseDoc original.GetTitle
        Set original = Nothing
    End If
    inputDimPreference = sw.GetUserPreferenceToggle(SW_PREF_INPUT_DIMENSION)
    oldEmbedFonts = sw.GetUserPreferenceToggle(SW_PREF_PDF_EMBED_FONTS)
    oldPdfColor = sw.GetUserPreferenceToggle(SW_PREF_PDF_IN_COLOR)
    oldPdfQuality = sw.GetUserPreferenceToggle(SW_PREF_PDF_HIGH_QUALITY)
    sw.SetUserPreferenceToggle SW_PREF_INPUT_DIMENSION, False
    sw.SetUserPreferenceToggle SW_PREF_PDF_EMBED_FONTS, True
    sw.SetUserPreferenceToggle SW_PREF_PDF_IN_COLOR, False
    sw.SetUserPreferenceToggle SW_PREF_PDF_HIGH_QUALITY, True
    preferenceChanged = True
    For Each key In jobs.Keys
        Set job = jobs(key)
        If CStr(job("status")) <> STATE_DONE Then
            jobNumber = CLng(job("number"))
            job("status") = STATE_RUNNING: job("stage") = "Opening model"
            SaveCheckpoint
            attempts = attempts + 1
            ProcessJob job
            Set flatBody = Nothing
            ' Close copied assemblies before copied parts, after ProcessJob releases locals.
            CloseSnapshots
            If CLng(Setting("JobsPerBatch", DEFAULT_BATCH_SIZE_TEXT)) > 0 Then
                If attempts >= CLng(Setting("JobsPerBatch", DEFAULT_BATCH_SIZE_TEXT)) Then
                    Exit For
                End If
            End If
            If CLng(Setting("MaxFailures", "0")) > 0 Then
                If failures >= CLng(Setting("MaxFailures", "0")) Then
                    Exit For
                End If
            End If
        End If
        DoEvents
    Next key
    RestorePreferences
    CloseSnapshots
    Log "COMPLETE drawings=" & completed & " failed=" & failures & " warnings=" & warningCount
    If Not logFile Is Nothing Then
        logFile.Close
    End If
    If Not manifest Is Nothing Then
        manifest.Close
    End If
    Set logFile = Nothing: Set manifest = Nothing
    Set currentDrawing = Nothing
    If Not original Is Nothing Then
        Set source = sw.ActivateDoc3(fs.GetFileName(original.GetPathName), False, SW_REBUILD_USER_DECISION, er)
    End If
    If backgroundOwned Then
        HideBackgroundFrame
        If FrameIsOffscreen() = False Then
            Err.Raise vbObjectError + ERR_APPLICATION_STATE, , "Background frame is on screen."
        End If
    End If
    If failures = 0 And completed = plannedJobs Then
        ReportProgress STATE_COMPLETE, "Drawings=" & completed & "; warnings=" & warningCount
    ElseIf failures = 0 Then
        ReportProgress STATE_PAUSED, "Batch saved; drawings=" & completed & "/" & plannedJobs
    Else
        ReportProgress STATE_FAILED, "Drawings=" & completed & "; failed=" & failures & "; selected=" & plannedJobs
    End If
    Debug.Print "Drawing pack: " & completed & " drawings; " & failures & " failures. " & runDir
    If Setting("ShowSummary", "0") = "1" Then
        MsgBox "Drawings: " & completed & vbCrLf & "Failures: " & failures & vbCrLf & "Review warnings: " & warningCount & vbCrLf & runDir, vbInformation, "Drawing pack"
    End If
    FinishBackground
    Exit Sub
Fatal:
    errNumber = Err.Number: errText = Err.Description
    On Error Resume Next
    RestorePreferences
    If Not logFile Is Nothing Then
        logFile.WriteLine "FATAL " & errNumber & " " & errText
        logFile.Close
    End If
    If Not manifest Is Nothing Then
        manifest.Close
    End If
    ReportProgress STATE_FAILED, "VBA error " & errNumber & ": " & errText
    If Setting("ShowSummary", "0") = "1" Then
        MsgBox "Macro stopped: " & errText & vbCrLf & runDir, vbExclamation, "Drawing pack"
    End If
    FinishBackground
End Sub

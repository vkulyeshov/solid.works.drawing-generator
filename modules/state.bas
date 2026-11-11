Attribute VB_Name = "state"
Option Explicit
Private planarLayoutCompatibility As Object
Private planarLayoutMaterial As Object
Private planarLayoutVisibleMaterial As Object
Private planarLayoutOutlines As Object
Private planarDetailLabels As Object
Private planarDetailLabelBoxes As Object
Private planarSplitOverview As Boolean, planarSplitPending As Boolean, planarSplitGroups As Object, planarSplitRequested As Object
Private planarOverviewName As String, jobAttempt As Long
Private planarModelCandidates As Boolean, planarModelDerived As Long
Private planarLayoutRejections As Object
Private planarTextMetricCache As Object
Private checkpointPath As String, snapshotPath As String, sourcePath As String, sourceStamp As String, snapshotConfig As String
Private attempts As Long, resumeRun As Boolean, renderSignature As String
Private excludeSuppressed As Boolean
Private maxJobAttempts As Long, batchFailures As Long, interruptedJobNumber As Long
Private sw As SldWorks.SldWorks, fs As Object, logFile As Object, manifest As Object
Private dimensionKeys As Object
Private featureDimensionOwners As Object
Private featureDimensionRepeatedKeys As Object
Private featureDimensionViewPriorities As Object
Private featureDimensionOccurrenceCounts As Object
Private featureCircleClassifications As Object
Private featureDimensionPlacement As Object
Private featureDimensionLanes As Object
Private profileRoleDrawing As Boolean
Private profilePagedMode As Boolean
Private profilePageRegions As Object, profilePageCrops As Object
Private profilePageViews As Collection, profilePageSheets As Collection
Private profilePagePlan As Collection, profilePageGroups As Collection
Private profilePageFirstSheet As String, profilePageModel As String, profilePageConfiguration As String
Private profilePageBody As String, profilePageOrientation As String
Private profilePageAngle As Double, profilePageMatrix As Variant, profilePageLimit As Long
Private profilePageInitialView As SldWorks.View
Private profilePageFailureView As String
Private profilePageThickness As Double
Private profilePageSplitSerial As Long
Private profilePanelFrame As Collection, profilePanelFeatures As Collection
Private profileCompactMode As Boolean
Private profileCompactEdgeCount As Long
Private profileFamilyMode As Boolean
Private profilePanelOpenings As Collection
Private semanticLayoutObstacles As Collection
Private semanticLayoutGeometry As Object
Private planarCandidateBudget As Long
Private planarCandidateResume As Object
Private planarStrictCrossings As Boolean, planarSearchLimit As Long
Private planarSampleCount As Long, planarSampleSkips As Long
Private planarDimensionRequests As Object
Private planarContourPoints As Collection
Private profileRawLoops As Collection
Private profileGeometryViews As Object
Private profileDimensionRequests As Object
Private profileSheetContours As Object
Private autoExpandedViews As Boolean
Private featurePatternGraphics As Collection
Private routingXMinimum As Double, routingXMaximum As Double
Private coordinateSelection As Object
Private coordinateRegion As Variant
Private coordinateSourceView As SldWorks.View
Private coordinateSourceX As Object, coordinateSourceY As Object
Private coordinateSourceOrigin As Variant, coordinateSourceScale As Double
Private coordinateAssignedRequests As Collection
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




Attribute VB_Name = "state"
Option Explicit
Private checkpointPath As String, snapshotPath As String, sourcePath As String, sourceStamp As String, snapshotConfig As String
Private attempts As Long, resumeRun As Boolean, renderSignature As String
Private maxJobAttempts As Long, batchFailures As Long, interruptedJobNumber As Long
Private sw As SldWorks.SldWorks, fs As Object, logFile As Object, manifest As Object
Private dimensionKeys As Object
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




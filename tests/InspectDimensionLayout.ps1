param(
    [Parameter(Mandatory=$true)][string]$DrawingPath,
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [switch]$IncludeGeometry
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestPaths.ps1')
$DrawingPath = Get-ExternalTestPath $DrawingPath
$OutputPath = Get-ExternalTestPath $OutputPath
$api = 'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist\SolidWorks.Interop.sldworks.dll'
Add-Type -Path $api
$source = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using SolidWorks.Interop.sldworks;
public static class DimensionLayoutInspection {
    const int DrawingDocument = 3;
    const int SilentReadOnly = 3;
    const int VirtualLeft = 76, VirtualTop = 77, VirtualWidth = 78, VirtualHeight = 79;
    const int OffscreenOffset = 1024;
    [DllImport("oleaut32.dll", PreserveSig=false)]
    static extern void GetActiveObject(ref Guid id, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object app);
    [DllImport("user32.dll")] static extern int GetSystemMetrics(int index);
    public static void Run(string path, string output, bool includeGeometry) {
        Type type = Type.GetTypeFromProgID("SldWorks.Application");
        Guid id = type.GUID;
        int existing = 0;
        try {
            object prior;
            GetActiveObject(ref id, IntPtr.Zero, out prior);
            existing = ((SldWorks)prior).GetProcessID();
        } catch (COMException) { }
        SldWorks sw = (SldWorks)Activator.CreateInstance(type);
        int pid = sw.GetProcessID();
        if (pid == existing) throw new Exception("Inspection requires its own SolidWorks process.");
        ModelDoc2 doc = null;
        try {
            Console.WriteLine("Inspector owns SolidWorks PID=" + pid);
            sw.UserControlBackground = true;
            ((Frame)sw.Frame()).KeepInvisible = true;
            sw.FrameLeft = GetSystemMetrics(VirtualLeft) + GetSystemMetrics(VirtualWidth) + OffscreenOffset;
            sw.FrameTop = GetSystemMetrics(VirtualTop) + GetSystemMetrics(VirtualHeight) + OffscreenOffset;
            int error = 0, warning = 0;
            doc = sw.OpenDoc6(path, DrawingDocument, SilentReadOnly, "", ref error, ref warning);
            Console.WriteLine("Opened drawing");
            if (doc == null) throw new Exception("Cannot inspect drawing: " + error);
            doc.Visible = true;
            DrawingDoc drawing = (DrawingDoc)doc;
            var records = new List<object>();
            var geometry = new List<object>();
            foreach (string sheet in (string[])drawing.GetSheetNames()) {
                drawing.ActivateSheet(sheet);
                doc.ForceRebuild3(false);
                doc.GraphicsRedraw2();
                View view = (View)drawing.GetFirstView();
                view = (View)view.GetNextView();
                while (view != null) {
                    Console.WriteLine("Inspecting " + view.Name);
                    var vertices = new List<object>();
                    object[] components = includeGeometry ? view.GetVisibleComponents() as object[] : null;
                    if (components != null) foreach (Component2 component in components) {
                        object[] edges = view.GetVisibleEntities2(component, 1) as object[];
                        if (edges == null) continue;
                        foreach (Edge edge in edges) {
                            foreach (Vertex vertex in new Vertex[] { (Vertex)edge.GetStartVertex(), (Vertex)edge.GetEndVertex() }) {
                                if (vertex == null) continue;
                                MathPoint point = (MathPoint)((MathUtility)sw.GetMathUtility()).CreatePoint(vertex.GetPoint());
                                point = (MathPoint)point.MultiplyTransform(view.ModelToViewTransform);
                                vertices.Add(point.ArrayData);
                            }
                        }
                    }
                    geometry.Add(new { sheet, view=view.Name, scale=view.ScaleDecimal, vertices });
                    DisplayDimension dimension = (DisplayDimension)view.GetFirstDisplayDimension5();
                    while (dimension != null) {
                        Annotation annotation = (Annotation)dimension.GetAnnotation();
                        DisplayData data = (DisplayData)dimension.GetDisplayData();
                        var lines = new List<object>();
                        var arrows = new List<object>();
                        var text = new List<object>();
                        for (int i = 0; i < data.GetLineCount(); i++) lines.Add(data.GetLineAtIndex2(i));
                        for (int i = 0; i < data.GetArrowHeadCount(); i++) arrows.Add(data.GetArrowHeadAtIndex2(i));
                        for (int i = 0; i < data.GetTextCount(); i++) text.Add(new {
                            value=data.GetTextAtIndex(i), position=data.GetTextPositionAtIndex(i),
                            angle=data.GetTextAngleAtIndex(i), height=data.GetTextHeightAtIndex(i)
                        });
                        records.Add(new { sheet, view=view.Name, scale=view.ScaleDecimal,
                            outline=view.GetOutline(), name=dimension.GetNameForSelection(),
                            type=dimension.Type2, smartWitness=dimension.SmartWitness,
                            arrowSide=dimension.ArrowSide, position=annotation.GetPosition(), lines, arrows, text });
                        dimension = (DisplayDimension)dimension.GetNext5();
                    }
                    view = (View)view.GetNextView();
                }
            }
            var serializer = new JavaScriptSerializer { MaxJsonLength=int.MaxValue };
            File.WriteAllText(output, serializer.Serialize(new { drawing=path, pid, dimensions=records, geometry }),
                              new System.Text.UTF8Encoding(false));
        } finally {
            if (doc != null) sw.CloseDoc(doc.GetTitle());
            sw.ExitApp();
        }
    }
}
'@
Add-Type -TypeDefinition $source -ReferencedAssemblies $api, 'System.Web.Extensions'
[DimensionLayoutInspection]::Run($DrawingPath, $OutputPath, [bool]$IncludeGeometry)
Write-Output "Native dimension inspection saved: $OutputPath"

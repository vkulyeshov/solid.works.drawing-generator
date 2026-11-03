param([Parameter(Mandatory=$true)][string]$DrawingPath,
      [Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestPaths.ps1')
$work = Get-ExternalTestPath $OutputDirectory
New-Item -ItemType Directory -Force -Path $work | Out-Null
$api = 'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist\SolidWorks.Interop.sldworks.dll'
Add-Type -Path $api
$source = @'
using System;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Globalization;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
public static class DrawingGeometryInspection {
    const int DrawingDocument=3, SilentReadOnly=3;
    static string Values(object value) {
        var numbers=value as double[];
        if(numbers==null) return "";
        return string.Join("\t",Array.ConvertAll(numbers,x=>x.ToString("R",CultureInfo.InvariantCulture)));
    }
    public static void Run(string input,string output) {
        var prior=new HashSet<int>();
        foreach(var process in Process.GetProcessesByName("SLDWORKS")) prior.Add(process.Id);
        var sw=(SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
        if(prior.Contains(sw.GetProcessID())) throw new Exception("Inspection must own its process");
        try {
            sw.UserControlBackground=true;
            sw.Visible=false;
            int error=0,warning=0;
            var doc=(ModelDoc2)sw.OpenDoc6(input,DrawingDocument,SilentReadOnly,"",ref error,ref warning);
            if(doc==null) throw new Exception("Cannot open drawing: "+error);
            var report=new StringBuilder();
            var drawing=(DrawingDoc)doc;
            var utility=(MathUtility)sw.GetMathUtility();
            foreach(string sheet in (string[])drawing.GetSheetNames()) {
                drawing.ActivateSheet(sheet);
                var view=(View)((View)drawing.GetFirstView()).GetNextView();
                while(view!=null) {
                    report.AppendLine("VIEW\t"+sheet+"\t"+view.Name+"\t"+view.ReferencedConfiguration+"\t"+view.GetOrientationName()+"\t"+Values(view.GetOutline()));
                    report.AppendLine("MATRIX\t"+view.Name+"\t"+Values(view.ModelToViewTransform.ArrayData));
                    var sketch=view.GetSketch() as Sketch;
                    if(sketch!=null) {
                        var segments=sketch.GetSketchSegments() as object[];
                        if(segments!=null) foreach(SketchSegment segment in segments) {
                            if(segment.GetType()!=0) continue;
                            var sketchLine=(SketchLine)segment;
                            var firstPoint=(SketchPoint)sketchLine.GetStartPoint2();
                            var lastPoint=(SketchPoint)sketchLine.GetEndPoint2();
                            var inverse=(MathTransform)sketch.ModelToSketchTransform.Inverse();
                            var a=(MathPoint)((MathPoint)utility.CreatePoint(new double[]{firstPoint.X,firstPoint.Y,firstPoint.Z})).MultiplyTransform(inverse);
                            var b=(MathPoint)((MathPoint)utility.CreatePoint(new double[]{lastPoint.X,lastPoint.Y,lastPoint.Z})).MultiplyTransform(inverse);
                            report.AppendLine("SKETCHLINE\t"+view.Name+"\t"+segment.ConstructionGeometry+"\t"+sketchLine.Infinite+"\t"+segment.GetLength()+"\t"+Values(a.ArrayData)+"\t"+Values(b.ArrayData));
                        }
                    }
                    var components=view.GetVisibleComponents() as object[];
                    if(components!=null) foreach(Component2 component in components) {
                        var edges=view.GetVisibleEntities2(component,1) as object[];
                        if(edges==null) continue;
                        foreach(Edge edge in edges) {
                            var curve=(Curve)edge.GetCurve();
                            var first=edge.GetStartVertex() as Vertex;
                            var last=edge.GetEndVertex() as Vertex;
                            if(curve.IsLine() && first!=null && last!=null) {
                                var pa=(MathPoint)((MathPoint)utility.CreatePoint(first.GetPoint())).MultiplyTransform(view.ModelToViewTransform);
                                var pb=(MathPoint)((MathPoint)utility.CreatePoint(last.GetPoint())).MultiplyTransform(view.ModelToViewTransform);
                                double[] a=(double[])pa.ArrayData,b=(double[])pb.ArrayData;
                                report.AppendLine("EDGE\t"+sheet+"\t"+view.Name+"\t"+Values(a)+"\t"+Values(b));
                            }
                        }
                    }
                    view=(View)view.GetNextView();
                }
            }
            File.WriteAllText(Path.Combine(output,"drawing-geometry.tsv"),report.ToString(),Encoding.UTF8);
            Console.WriteLine(output);
        } finally { sw.CloseAllDocuments(true);sw.ExitApp(); }
    }
}
'@
Add-Type -TypeDefinition $source -ReferencedAssemblies $api
[DrawingGeometryInspection]::Run((Resolve-Path -LiteralPath $DrawingPath).Path,$work)

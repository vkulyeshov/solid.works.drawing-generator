param([Parameter(Mandatory=$true)][string]$ModelPath,
      [Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestPaths.ps1')
$work = Get-ExternalTestPath $OutputDirectory
New-Item -ItemType Directory -Force -Path $work | Out-Null
$inputCopy = Join-Path $work 'inspection.SLDPRT'
Copy-Item -LiteralPath $ModelPath -Destination $inputCopy
$api = 'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist\SolidWorks.Interop.sldworks.dll'
Add-Type -Path $api
$source = @'
using System;
using System.IO;
using System.Text;
using System.Globalization;
using System.Diagnostics;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
public static class FeatureInspection {
    const int PartDocument=1, SilentReadOnly=3;
    static StringBuilder report=new StringBuilder();
    static string Values(object value) {
        var values=value as double[];
        if(values==null) return "";
        var output=new List<string>();
        foreach(double item in values) output.Add(item.ToString("R",CultureInfo.InvariantCulture));
        return string.Join("\t",output);
    }
    static void Feature(Feature feature, string parent) {
        report.AppendLine("FEATURE\t"+parent+"\t"+feature.Name+"\t"+feature.GetTypeName2());
        report.AppendLine("FEATUREKIND\t"+feature.Name+"\t"+feature.GetTypeName());
        if(feature.GetTypeName2()=="ProfileFeature") {
            var sketch=feature.GetSpecificFeature2() as Sketch;
            report.AppendLine("SKETCHTRANSFORM\t"+feature.Name+"\t"+Values(sketch.ModelToSketchTransform.ArrayData));
            var segments=sketch.GetSketchSegments() as object[];
            if(segments!=null) foreach(SketchSegment segment in segments) {
                int type=((ISketchSegment)segment).GetType();
                if(type==0) {
                    var line=(SketchLine)segment;
                    var a=(SketchPoint)line.GetStartPoint2(); var b=(SketchPoint)line.GetEndPoint2();
                    report.AppendLine("SKETCHLINE\t"+feature.Name+"\t"+segment.ConstructionGeometry+"\t"+Values(new double[] {a.X,a.Y,a.Z,b.X,b.Y,b.Z}));
                } else if(type==1) {
                    var arc=(SketchArc)segment;
                    var a=(SketchPoint)arc.GetStartPoint2(); var b=(SketchPoint)arc.GetEndPoint2();var c=(SketchPoint)arc.GetCenterPoint2();
                    report.AppendLine("SKETCHARC\t"+feature.Name+"\t"+segment.ConstructionGeometry+"\t"+Values(new double[] {a.X,a.Y,a.Z,b.X,b.Y,b.Z,c.X,c.Y,c.Z,arc.GetRadius(),segment.GetLength()}));
                }
            }
        }
        var display=feature.GetFirstDisplayDimension() as DisplayDimension;
        while(display!=null) {
            var dimension=(Dimension)display.GetDimension2(0);
            report.AppendLine("DIM\t"+feature.Name+"\t"+dimension.FullName+"\t"+display.Type2+"\t"+dimension.SystemValue.ToString("R",CultureInfo.InvariantCulture));
            var references=dimension.ReferencePoints as object[];
            if(references!=null) foreach(MathPoint point in references) report.AppendLine("DIMREF\t"+feature.Name+"\t"+dimension.FullName+"\t"+Values(point.ArrayData));
            var attached=((Annotation)display.GetAnnotation()).GetAttachedEntities3() as object[];
            if(attached!=null) foreach(object entity in attached) {
                var segment=entity as SketchSegment;
                if(segment!=null) report.AppendLine("DIMATT\t"+dimension.FullName+"\tsketch\t"+segment.ConstructionGeometry+"\t"+((ISketchSegment)segment).GetType()+"\t"+segment.GetLength().ToString("R",CultureInfo.InvariantCulture));
                else report.AppendLine("DIMATT\t"+dimension.FullName+"\t"+(entity is SketchPoint ? "point" : entity is Edge ? "edge" : entity is Vertex ? "vertex" : "other"));
            }
            display=feature.GetNextDisplayDimension(display) as DisplayDimension;
        }
        var sub=feature.GetFirstSubFeature() as Feature;
        while(sub!=null) { Feature(sub,feature.Name); sub=sub.GetNextSubFeature() as Feature; }
    }
    public static void Run(string input,string output) {
        var prior=new HashSet<int>();
        foreach(var p in Process.GetProcessesByName("SLDWORKS")) prior.Add(p.Id);
        var sw=(SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
        if(prior.Contains(sw.GetProcessID())) throw new Exception("Inspection must own its process");
        try {
            sw.UserControlBackground=true;
            int error=0,warning=0;
            var doc=(ModelDoc2)sw.OpenDoc6(input,PartDocument,SilentReadOnly,"",ref error,ref warning);
            if(doc==null) throw new Exception("Cannot open copied model: "+error);
            var feature=doc.FirstFeature() as Feature;
            while(feature!=null) { Feature(feature,""); feature=feature.GetNextFeature() as Feature; }
            foreach(Body2 body in (object[])((PartDoc)doc).GetBodies2(0,true)) {
                foreach(Face2 face in (object[])body.GetFaces()) {
                    var surface=(Surface)face.GetSurface();
                    report.AppendLine("SURFACE\t"+surface.Identity()+"\t"+surface.IsPlane()+"\t"+surface.IsCylinder()+"\t"+Values(face.GetBox())+"\t"+Values(face.Normal));
                    var owner=(Feature)face.GetFeature();
                    if(surface.IsCylinder()) report.AppendLine("CYLINDER\t"+owner.Name+"\t"+Values(surface.CylinderParams));
                    if(surface.IsCylinder() || surface.IsCone() || surface.IsPlane())
                        report.AppendLine("FACE\t"+owner.Name+"\t"+(surface.IsCylinder() ? "cylinder" : surface.IsCone() ? "cone" : "plane")+"\t"+face.FaceInSurfaceSense()+"\t"+Values(face.GetBox())+"\t"+Values(face.Normal));
                }
                foreach(Edge edge in (object[])body.GetEdges()) {
                    var curve=(Curve)edge.GetCurve();
                    var a=edge.GetStartVertex() as Vertex;
                    var b=edge.GetEndVertex() as Vertex;
                    if(curve.IsCircle()) report.AppendLine("CIRCLE\t"+Values(curve.CircleParams)+"\t"+(a==null ? "" : Values(a.GetPoint()))+"\t"+(b==null ? "" : Values(b.GetPoint())));
                    else if(curve.IsLine()) report.AppendLine("LINE\t"+(a==null ? "" : Values(a.GetPoint()))+"\t"+(b==null ? "" : Values(b.GetPoint())));
                }
            }
            File.WriteAllText(output,report.ToString(),Encoding.UTF8);
            Console.WriteLine(output);
        } finally { sw.CloseAllDocuments(true); sw.ExitApp(); }
    }
}
'@
Add-Type -TypeDefinition $source -ReferencedAssemblies $api
[FeatureInspection]::Run($inputCopy,(Join-Path $work 'features.tsv'))

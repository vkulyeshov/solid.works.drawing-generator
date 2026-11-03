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
public static class PitchCircleProbe {
    static string Values(object data) {
        var result=new List<string>();
        foreach(double value in (double[])data) result.Add(value.ToString("R",CultureInfo.InvariantCulture));
        return string.Join("\t",result);
    }
    public static void Run(string source,string work) {
        var prior=new HashSet<int>();
        foreach(var process in Process.GetProcessesByName("SLDWORKS")) prior.Add(process.Id);
        var sw=(SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
        int pid=sw.GetProcessID();
        if(prior.Contains(pid)) throw new Exception("The probe must own its process");
        Console.WriteLine("Probe owns PID="+pid);
        try {
            sw.UserControlBackground=true;sw.Visible=true;
            string copy=Path.Combine(work,"probe.slddrw");File.Copy(source,copy,true);
            var dependencies=(string[])sw.GetDocumentDependencies2(copy,true,true,false);
            for(int i=1;i<dependencies.Length;i+=2) {
                if(!Path.GetExtension(dependencies[i]).Equals(".sldprt",StringComparison.OrdinalIgnoreCase)) continue;
                string model=Path.Combine(work,Path.GetFileName(dependencies[i]));File.Copy(dependencies[i],model,true);
                if(!sw.ReplaceReferencedDocument(copy,dependencies[i],model)) throw new Exception("Cannot isolate the model reference");
            }
            int error=0,warning=0;
            var doc=(ModelDoc2)sw.OpenDoc6(copy,3,1,"",ref error,ref warning);
            if(doc==null) throw new Exception("Cannot open the isolated drawing");
            sw.ActivateDoc3(doc.GetTitle(),false,0,ref error);
            var drawing=(DrawingDoc)doc;drawing.ActivateSheet(((string[])drawing.GetSheetNames())[1]);drawing.EditSheet();
            var view=(View)((View)drawing.GetFirstView()).GetNextView();
            var sketch=(Sketch)view.GetSketch();
            var report=new StringBuilder();
            report.AppendLine("VIEW\t"+Values(view.ModelToViewTransform.ArrayData));
            report.AppendLine("SKETCH\t"+Values(sketch.ModelToSketchTransform.ArrayData));
            foreach(bool local in new bool[] {true,false}) {
                drawing.ActivateView(view.Name);doc.ClearSelection2(true);
                double[] position=(double[])view.Position;
                double x=local ? 0 : position[0]/view.ScaleDecimal;
                double y=local ? 0 : position[1]/view.ScaleDecimal;
                var circle=(SketchSegment)doc.SketchManager.CreateCircleByRadius(x,y,0,0.05);
                circle.ConstructionGeometry=true;
                circle.Select4(false,null);
                var display=(DisplayDimension)doc.AddDiameterDimension2(position[0]+0.035,position[1]+0.025,0);
                if(display==null) throw new Exception("No diameter dimension created");
                report.AppendLine("CASE\t"+local+"\t"+display.GetNameForSelection()+"\t"+((Dimension)display.GetDimension2(0)).SystemValue.ToString("R",CultureInfo.InvariantCulture));
                var owner=(View)drawing.GetFirstView();
                while(owner!=null) {
                    var item=(DisplayDimension)owner.GetFirstDisplayDimension5();
                    while(item!=null) {
                        if(item.GetNameForSelection()==display.GetNameForSelection()) report.AppendLine("OWNER\t"+owner.Name);
                        item=(DisplayDimension)item.GetNext5();
                    }
                    owner=(View)owner.GetNextView();
                }
                var data=(DisplayData)display.GetDisplayData();
                for(int i=0;i<data.GetArrowHeadCount();i++) report.AppendLine("ARROW\t"+Values(data.GetArrowHeadAtIndex2(i)));
                ((Annotation)display.GetAnnotation()).Select3(false,null);doc.EditDelete();
                circle.Select4(false,null);doc.EditDelete();
            }
            File.WriteAllText(Path.Combine(work,"probe.tsv"),report.ToString(),Encoding.UTF8);
            Console.WriteLine(report.ToString());
        } finally { sw.CloseAllDocuments(true);sw.ExitApp(); }
    }
}
'@
Add-Type -TypeDefinition $source -ReferencedAssemblies $api
[PitchCircleProbe]::Run((Resolve-Path -LiteralPath $DrawingPath).Path,$work)

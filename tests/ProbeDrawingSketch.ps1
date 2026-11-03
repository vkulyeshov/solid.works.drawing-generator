param([Parameter(Mandatory=$true)][string]$DrawingPath,
      [Parameter(Mandatory=$true)][string]$OutputDirectory,
      [switch]$CreateDetail)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestPaths.ps1')
$work = Get-ExternalTestPath $OutputDirectory
New-Item -ItemType Directory -Force -Path $work | Out-Null
$api = 'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist\SolidWorks.Interop.sldworks.dll'
Add-Type -Path $api
$source = @'
using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using SolidWorks.Interop.sldworks;
public static class DrawingSketchProbe {
    const int DrawingDocument=3, SilentReadOnly=3, PdfExport=1, AllSheets=1, SaveSilent=1, ColorPdf=323, Red=255;
    const double RadiusM=0.005;
    public static void Run(string path,string output,bool createDetail) {
        var prior=new HashSet<int>();
        foreach(var process in Process.GetProcessesByName("SLDWORKS")) prior.Add(process.Id);
        var sw=(SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
        if(prior.Contains(sw.GetProcessID())) throw new Exception("Probe must own its process");
        bool oldColor=sw.GetUserPreferenceToggle(ColorPdf);
        try {
            sw.UserControlBackground=true;
            sw.Visible=false;
            sw.SetUserPreferenceToggle(ColorPdf,true);
            var utility=(MathUtility)sw.GetMathUtility();
            for(int variant=createDetail ? 3 : 0;variant<4;variant++) {
                int error=0,warning=0;
                var doc=(ModelDoc2)sw.OpenDoc6(path,DrawingDocument,SilentReadOnly,"",ref error,ref warning);
                var drawing=(DrawingDoc)doc;
                var view=(View)((View)drawing.GetFirstView()).GetNextView();
                double[] box=(double[])view.GetOutline();
                double x=(box[0]+box[2])/2,y=(box[1]+box[3])/2;
                if(createDetail) { x=box[0]+0.008; y=box[3]-0.012; }
                var sketch=(Sketch)view.GetSketch();
                var point=(MathPoint)utility.CreatePoint(new double[]{x,y,0});
                if(variant==1) {
                    var origin=(MathPoint)((MathPoint)utility.CreatePoint(new double[]{0,0,0})).MultiplyTransform(view.ModelToViewTransform);
                    double[] p=(double[])origin.ArrayData;
                    point=(MathPoint)utility.CreatePoint(new double[]{(x-p[0])/view.ScaleDecimal,(y-p[1])/view.ScaleDecimal,0});
                } else if(variant==2) {
                    point=(MathPoint)point.MultiplyTransform(view.ModelToViewTransform.Inverse());
                    point=(MathPoint)point.MultiplyTransform(sketch.ModelToSketchTransform);
                } else if(variant==3) point=(MathPoint)point.MultiplyTransform(sketch.ModelToSketchTransform);
                double[] coordinate=(double[])point.ArrayData;
                Console.WriteLine("Input coordinates="+coordinate[0].ToString("R",CultureInfo.InvariantCulture)+","+coordinate[1].ToString("R",CultureInfo.InvariantCulture));
                drawing.ActivateView(view.Name);
                doc.ClearSelection2(true);
                var manager=(SketchManager)doc.SketchManager;
                manager.AddToDB=true;
                var circle=(SketchSegment)manager.CreateCircleByRadius(coordinate[0],coordinate[1],0,variant==0 ? RadiusM : RadiusM/view.ScaleDecimal);
                manager.AddToDB=false;
                circle.Color=Red;
                if(createDetail) {
                    var selection=(SelectData)((SelectionMgr)doc.SelectionManager).CreateSelectData();
                    selection.View=view;
                    Console.WriteLine("Circle selection="+circle.Select4(false,selection));
                    var detail=(View)drawing.CreateDetailViewAt4(0.12,0.15,0,0,1,1,"Z",1,false,false,false,5);
                    if(detail==null) throw new Exception("Empty detail result");
                    var parent=detail.GetBaseView() as View;
                    Console.WriteLine("Detail="+detail.Name+"; parent="+(parent==null ? "NULL" : parent.Name)+"; visible components="+detail.GetVisibleComponentCount());
                }
                doc.GraphicsRedraw2();
                var pdf=(ExportPdfData)sw.GetExportFileData(PdfExport);
                pdf.SetSheets(AllSheets,null);
                pdf.ViewPdfAfterSaving=false;
                if(!doc.Extension.SaveAs(Path.Combine(output,"variant-"+variant+".pdf"),0,SaveSilent,pdf,ref error,ref warning)) throw new Exception("Probe PDF failed: "+error);
                Console.WriteLine(variant+" expected "+x.ToString("R",CultureInfo.InvariantCulture)+","+y.ToString("R",CultureInfo.InvariantCulture));
                sw.CloseAllDocuments(true);
            }
        } finally {
            sw.SetUserPreferenceToggle(ColorPdf,oldColor);
            sw.CloseAllDocuments(true);
            sw.ExitApp();
        }
    }
}
'@
Add-Type -TypeDefinition $source -ReferencedAssemblies $api
[DrawingSketchProbe]::Run((Resolve-Path -LiteralPath $DrawingPath).Path,$work,[bool]$CreateDetail)

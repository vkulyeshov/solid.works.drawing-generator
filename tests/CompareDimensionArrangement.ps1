param(
    [Parameter(Mandatory=$true)][string]$DrawingPath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [ValidateSet('placement','creation','postprocess','inspect','all')][string]$Cases = 'postprocess'
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestPaths.ps1')
$DrawingPath = Get-ExternalTestPath $DrawingPath
$OutputDirectory = Get-ExternalTestPath $OutputDirectory
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$api = 'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist\SolidWorks.Interop.sldworks.dll'
Add-Type -Path $api
$source = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using SolidWorks.Interop.sldworks;
public static class DimensionArrangementComparison {
    const int DrawingDocument=3, SilentReadOnly=3, SilentSave=1;
    const int AutoArrange=0, RadialDimension=5, DiameterDimension=6, HorizontalBentText=2;
    const int AllEntities=1, BaselineScheme=1, ChainScheme=3, Above=1, Left=-1;
    const int VirtualLeft=76, VirtualTop=77, VirtualWidth=78, VirtualHeight=79, OffscreenOffset=1024;
    const int MacroTimeoutMilliseconds=120000;
    const double ShoulderLengthM=0.004, ValueToleranceM=0.000000001;
    [DllImport("oleaut32.dll", PreserveSig=false)]
    static extern void GetActiveObject(ref Guid id, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object app);
    [DllImport("user32.dll")] static extern int GetSystemMetrics(int index);
    static object ReferenceData(DisplayDimension display) {
        var points=new List<object>();
        Dimension dimension=(Dimension)display.GetDimension2(0);
        try {
            Array references=dimension.ReferencePoints as Array;
            if(references!=null) foreach(object reference in references) {
                MathPoint point=reference as MathPoint;
                points.Add(point==null ? reference : point.ArrayData);
            }
        } catch(COMException) { }
        return points;
    }
    static object Geometry(DrawingDoc drawing) {
        var result=new List<object>();
        View view=(View)((View)drawing.GetFirstView()).GetNextView();
        while(view!=null) {
            var points=new List<object>(); var vertices=new List<object>(); var bodyInfo=new List<object>();
            ModelDoc2 model=view.ReferencedDocument;
            if(model!=null && model.GetType()==1) {
                object[] bodies=view.Bodies as object[];
                if(bodies==null || bodies.Length==0) bodies=((PartDoc)model).GetBodies2(0,true) as object[];
                if(bodies!=null) foreach(Body2 body in bodies) {
                    bodyInfo.Add(new {name=body.Name,box=body.GetBodyBox()});
                    object[] edges=body.GetEdges() as object[];
                    if(edges==null) continue;
                    foreach(Edge edge in edges) {
                        foreach(Vertex vertex in new Vertex[] {(Vertex)edge.GetStartVertex(),(Vertex)edge.GetEndVertex()})
                            if(vertex!=null) {points.Add(vertex.GetPoint());vertices.Add(vertex.GetPoint());}
                        Curve curve=(Curve)edge.GetCurve();
                        CurveParamData parameters=(CurveParamData)edge.GetCurveParams3();
                        for(int step=1;step<=3;step++) {
                            double parameter=parameters.UMinValue+(parameters.UMaxValue-parameters.UMinValue)*step/4.0;
                            double[] point=(double[])edge.Evaluate2(parameter,0);
                            points.Add(new double[] {point[0],point[1],point[2]});
                        }
                    }
                }
            }
            result.Add(new {view=view.Name,outline=view.GetOutline(),transform=view.ModelToViewTransform.ArrayData,bodyInfo,vertices,points});
            view=(View)view.GetNextView();
        }
        return result;
    }
    static Dictionary<string,double> Values(DrawingDoc drawing) {
        var result=new Dictionary<string,double>();
        foreach(string sheet in (string[])drawing.GetSheetNames()) {
            drawing.ActivateSheet(sheet);
            View view=(View)((View)drawing.GetFirstView()).GetNextView();
            while(view!=null) {
                DisplayDimension display=(DisplayDimension)view.GetFirstDisplayDimension5();
                while(display!=null) {
                    Dimension dimension=(Dimension)display.GetDimension2(0);
                    result.Add(sheet+"|"+view.Name+"|"+display.GetNameForSelection(),dimension.SystemValue);
                    display=(DisplayDimension)display.GetNext5();
                }
                view=(View)view.GetNextView();
            }
        }
        return result;
    }
    static List<object> DisplayRecords(DrawingDoc drawing) {
        var result=new List<object>();
        foreach(string sheet in (string[])drawing.GetSheetNames()) {
            drawing.ActivateSheet(sheet);
            View view=(View)((View)drawing.GetFirstView()).GetNextView();
            while(view!=null) {
                DisplayDimension display=(DisplayDimension)view.GetFirstDisplayDimension5();
                while(display!=null) {
                    DisplayData data=(DisplayData)display.GetDisplayData();
                    var lines=new List<object>(); var text=new List<object>(); var arrows=new List<object>();
                    for(int i=0;i<data.GetLineCount();i++) lines.Add(data.GetLineAtIndex2(i));
                    for(int i=0;i<data.GetArrowHeadCount();i++) arrows.Add(data.GetArrowHeadAtIndex2(i));
                    for(int i=0;i<data.GetTextCount();i++) text.Add(new {
                        value=data.GetTextAtIndex(i),position=data.GetTextPositionAtIndex(i),
                        angle=data.GetTextAngleAtIndex(i),height=data.GetTextHeightAtIndex(i)
                    });
                    result.Add(new {sheet,view=view.Name,name=display.GetNameForSelection(),
                        type=display.Type2,leader=display.GetBrokenLeader2(),offsetText=display.OffsetText,references=ReferenceData(display),
                        value=((Dimension)display.GetDimension2(0)).SystemValue,arrows,lines,text});
                    display=(DisplayDimension)display.GetNext5();
                }
                view=(View)view.GetNextView();
            }
        }
        return result;
    }
    public static void Run(string input,string output,string cases) {
        var priorProcesses=new HashSet<int>();
        foreach(Process process in Process.GetProcessesByName("SLDWORKS")) priorProcesses.Add(process.Id);
        Type appType=Type.GetTypeFromProgID("SldWorks.Application");
        Console.WriteLine("Starting an isolated SolidWorks process");
        SldWorks sw=(SldWorks)Activator.CreateInstance(appType);
        int pid=sw.GetProcessID();
        if(priorProcesses.Contains(pid)) throw new Exception("The comparison requires its own SolidWorks process.");
        ModelDoc2 doc=null;
        var report=new List<object>();
        string layoutFailure=null;
        try {
            Console.WriteLine("Comparison owns SolidWorks PID="+pid);
            sw.UserControlBackground=true;
            ((Frame)sw.Frame()).KeepInvisible=true;
            sw.FrameLeft=GetSystemMetrics(VirtualLeft)+GetSystemMetrics(VirtualWidth)+OffscreenOffset;
            sw.FrameTop=GetSystemMetrics(VirtualTop)+GetSystemMetrics(VirtualHeight)+OffscreenOffset;
            if(cases=="postprocess" || cases=="inspect") input=Path.Combine(output,"autodimension_chain.slddrw");
            // The comparison does not acquire the generation lock: it only edits isolated test copies.
            string isolated=Path.Combine(output,"input_"+Guid.NewGuid().ToString("N")); Directory.CreateDirectory(isolated);
            string working=Path.Combine(isolated,Path.GetFileName(input));
            File.Copy(input,working,true);
            string[] dependencies=(string[])sw.GetDocumentDependencies2(working,true,false,false);
            var models=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if(dependencies!=null) for(int i=1;i<dependencies.Length;i+=2) {
                if(Path.GetExtension(dependencies[i]).Equals(".sldprt",StringComparison.OrdinalIgnoreCase)) models.Add(dependencies[i]);
                if(Path.GetExtension(dependencies[i]).Equals(".sldasm",StringComparison.OrdinalIgnoreCase))
                    throw new Exception("This fixture must reference one part, not an assembly.");
            }
            if(models.Count!=1) throw new Exception("The comparison expects exactly one source part.");
            foreach(string model in models) {
                string copy=Path.Combine(isolated,Path.GetFileName(model));
                File.Copy(model,copy,true);
                if(!sw.ReplaceReferencedDocument(working,model,copy)) throw new Exception("Cannot isolate the model reference.");
            }
            input=working;
            foreach(string mode in new string[] {"native","native_bent","native_by_view",
                "autodimension_chain","autodimension_baseline","postprocessed_chain"}) {
                bool regenerate=mode.StartsWith("autodimension_");
                if((cases=="postprocess" || cases=="inspect") && mode!="postprocessed_chain") continue;
                if(cases!="postprocess" && cases!="inspect" && mode=="postprocessed_chain") continue;
                if(cases=="placement" && regenerate) continue;
                if(cases=="creation" && !regenerate) continue;
                int error=0,warning=0;
                doc=sw.OpenDoc6(input,DrawingDocument,SilentReadOnly,"",ref error,ref warning);
                if(doc==null) throw new Exception("Cannot open copied drawing: "+error);
                doc.Visible=true;
                DrawingDoc drawing=(DrawingDoc)doc;
                var before=Values(drawing);
                var calls=new List<object>();
                if(cases=="inspect") {
                    File.WriteAllText(Path.Combine(output,"input-geometry.json"),
                        new JavaScriptSerializer {MaxJsonLength=int.MaxValue}.Serialize(new {dimensions=DisplayRecords(drawing),geometry=Geometry(drawing)}));
                    sw.CloseDoc(doc.GetTitle()); doc=null;
                    Console.WriteLine("Native dimension references exported.");
                    continue;
                }
                if(mode=="postprocessed_chain") {
                    if(!File.Exists(Path.Combine(output,"input-geometry.json"))) {
                        File.WriteAllText(Path.Combine(output,"input-geometry.json"),
                            new JavaScriptSerializer {MaxJsonLength=int.MaxValue}.Serialize(new {dimensions=DisplayRecords(drawing),geometry=Geometry(drawing)}));
                    }
                    string macro=Path.Combine(Path.GetDirectoryName(output),"postprocess.swp");
                    string result=Path.Combine(Path.GetDirectoryName(output),"postprocess-result.txt");
                    string control=Path.Combine(Path.GetDirectoryName(output),"postprocess-control.txt");
                    File.WriteAllText(control,pid+"\r\n"+doc.GetPathName()+"\r\n",System.Text.Encoding.Unicode);
                    if(File.Exists(result)) File.Delete(result);
                    foreach(string artifact in new string[] {"native-dimension-audit.tsv","native-view-regions.tsv","native-symmetry-lines.tsv"}) {
                        string auditPath=Path.Combine(Path.GetDirectoryName(output),artifact);
                        if(File.Exists(auditPath)) File.Delete(auditPath);
                    }
                    using(var watchdog=new System.Threading.Timer(delegate(object unused) {
                        try { Process.GetProcessById(pid).Kill(); } catch(ArgumentException) { }
                    },null,MacroTimeoutMilliseconds,System.Threading.Timeout.Infinite)) {
                        if(!sw.RunMacro(macro,"Macro11","NativeLayoutTest")) throw new Exception("Postprocessing macro failed to run.");
                    }
                    if(!File.Exists(result)) throw new Exception("Postprocessing macro produced no result.");
                    string message=File.ReadAllText(result).Trim();
                    Console.WriteLine(message);
                    if(message!="OK") layoutFailure=message;
                }
                if(mode!="postprocessed_chain") foreach(string sheet in (string[])drawing.GetSheetNames()) {
                    drawing.ActivateSheet(sheet);
                    drawing.ActivateView(""); doc.ClearSelection2(true);
                    if(regenerate) {
                        var targetViews=new List<string>();
                        View target=(View)((View)drawing.GetFirstView()).GetNextView();
                        while(target!=null) {
                            DisplayDimension existingDimension=(DisplayDimension)target.GetFirstDisplayDimension5();
                            if(existingDimension!=null) targetViews.Add(target.Name);
                            while(existingDimension!=null) {
                                ((Annotation)existingDimension.GetAnnotation()).Select3(true,null);
                                existingDimension=(DisplayDimension)existingDimension.GetNext5();
                            }
                            target=(View)target.GetNextView();
                        }
                        if(!doc.Extension.DeleteSelection2(0)) throw new Exception("Cannot remove copied drawing dimensions.");
                        foreach(string targetName in targetViews) {
                            doc.ClearSelection2(true);
                            drawing.ActivateView(targetName);
                            bool picked=doc.Extension.SelectByID2(targetName,"DRAWINGVIEW",0,0,0,false,0,null,0);
                            if(!picked) throw new Exception("Cannot select drawing view "+targetName);
                            int scheme=mode=="autodimension_chain" ? ChainScheme : BaselineScheme;
                            int status=drawing.AutoDimension(AllEntities,scheme,Above,scheme,Left);
                            calls.Add(new {sheet,view=targetName,operation="AutoDimension",status});
                            Console.WriteLine(mode+": "+targetName+"; AutoDimension status="+status);
                        }
                        drawing.ActivateView(""); doc.ClearSelection2(true);
                    }
                    int selected=0;
                    View view=(View)((View)drawing.GetFirstView()).GetNextView();
                    while(view!=null) {
                        int viewSelected=0;
                        DisplayDimension display=(DisplayDimension)view.GetFirstDisplayDimension5();
                        while(display!=null) {
                            if(mode=="native_bent" && (display.Type2==RadialDimension || display.Type2==DiameterDimension)) {
                                display.SetBrokenLeader2(false,HorizontalBentText);
                                display.SolidLeader=false;
                                display.SetBentLeaderLength(false,ShoulderLengthM);
                            }
                            if(((Annotation)display.GetAnnotation()).Select3(true,null)) {selected++;viewSelected++;}
                            display=(DisplayDimension)display.GetNext5();
                        }
                        if(mode=="native_by_view" && viewSelected>1) {
                            bool ok=doc.Extension.AlignDimensions(AutoArrange,0);
                            calls.Add(new {sheet,view=view.Name,selected=viewSelected,ok});
                            doc.ClearSelection2(true);
                        }
                        view=(View)view.GetNextView();
                    }
                    if(mode!="native_by_view" && selected>1) {
                        bool ok=doc.Extension.AlignDimensions(AutoArrange,0);
                        calls.Add(new {sheet,view="all",selected,ok});
                    }
                    doc.ClearSelection2(true);
                    doc.ForceRebuild3(false); doc.GraphicsRedraw2();
                }
                var after=Values(drawing);
                if(!regenerate) {
                    if(mode!="postprocessed_chain" && before.Count!=after.Count) throw new Exception(mode+": the dimension count changed.");
                    foreach(var item in after) {
                        if(!before.ContainsKey(item.Key) || Math.Abs(before[item.Key]-item.Value)>ValueToleranceM)
                            throw new Exception(mode+": a dimension changed: "+item.Key);
                    }
                    if(mode=="postprocessed_chain") {
                        string auditPath=Path.Combine(Path.GetDirectoryName(output),"native-dimension-audit.tsv");
                        var removed=new Dictionary<string,string>();
                        if(File.Exists(auditPath)) foreach(string row in File.ReadAllLines(auditPath,System.Text.Encoding.Unicode)) {
                            string[] fields=row.Split('\t');
                            if(fields.Length>=5) removed.Add(fields[2],fields[3]);
                        }
                        int missing=0;
                        foreach(var item in before) if(!after.ContainsKey(item.Key)) {
                            missing++;
                            string name=item.Key.Substring(item.Key.LastIndexOf('|')+1);
                            if(!removed.ContainsKey(name)) throw new Exception("An unaudited dimension was removed: "+name);
                            bool matched=false;
                            foreach(var kept in after) if(kept.Key.EndsWith("|"+removed[name]) && Math.Abs(kept.Value-item.Value)<=ValueToleranceM) matched=true;
                            if(!matched) throw new Exception("A removed dimension has no unchanged representative: "+name);
                        }
                        if(missing!=removed.Count) throw new Exception("Dimension deduplication audit count mismatch.");
                    }
                }
                var data=DisplayRecords(drawing);
                string suffix=layoutFailure==null ? "" : "_failed";
                string native=Path.Combine(output,mode+suffix+".slddrw");
                string pdf=Path.Combine(output,mode+suffix+".pdf");
                if(!doc.Extension.SaveAs(native,0,SilentSave,null,ref error,ref warning)) throw new Exception("Cannot save drawing: "+error);
                if(!doc.Extension.SaveAs(pdf,0,SilentSave,null,ref error,ref warning)) throw new Exception("Cannot export PDF: "+error);
                report.Add(new {mode,originalDimensions=before.Count,dimensions=after.Count,
                    valuesUnchanged=!regenerate,calls,dimensionsData=data,pdf,native});
                Console.WriteLine(mode+": dimensions="+after.Count+"; "+pdf);
                sw.CloseDoc(doc.GetTitle()); doc=null;
            }
            File.WriteAllText(Path.Combine(output,cases+"-comparison.json"),
                new JavaScriptSerializer {MaxJsonLength=int.MaxValue}.Serialize(report),new System.Text.UTF8Encoding(false));
            if(layoutFailure!=null) throw new Exception(layoutFailure);
        } finally {
            try { if(doc!=null) sw.CloseDoc(doc.GetTitle()); } catch(COMException) { }
            try { sw.ExitApp(); } catch(COMException) { }
        }
    }
}
'@
Add-Type -TypeDefinition $source -ReferencedAssemblies $api, 'System.Web.Extensions'
[DimensionArrangementComparison]::Run($DrawingPath, $OutputDirectory, $Cases)

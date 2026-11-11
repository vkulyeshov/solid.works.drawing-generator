param([Parameter(Mandatory=$true)][string]$WorkDirectory,
      [string]$ModelFileName = 'Matochina4x100.SLDPRT',
      [switch]$InventoryOnly,
      [switch]$RemapInventoryReferences,
      [switch]$ProfileGeometryOnly,
      [switch]$InspectOnly,
      [ValidateRange(60,3600)][int]$MacroTimeoutSeconds = 480)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestPaths.ps1')
$work = Get-ExternalTestPath $WorkDirectory
$api = 'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist\SolidWorks.Interop.sldworks.dll'
Add-Type -Path $api
$source = @'
using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Globalization;
using Microsoft.Win32;
using System.Security.Cryptography;
using SolidWorks.Interop.sldworks;
public static class SixViewGenerationTest {
    const int PartDocument=1, AssemblyDocument=2, DrawingDocument=3, SilentReadOnly=3, OffscreenOffset=1024;
    const int SolidBody=0;
    const int CalloutAbove=3, CalloutAboveDefinition=7, CalloutBelowDefinition=8;
    const int VirtualLeft=76, VirtualTop=77, VirtualWidth=78, VirtualHeight=79;
    const int InventoryTimeoutMilliseconds=900000;
    const int StartupTimeoutMilliseconds=120000, StartupPollMilliseconds=250;
    const int ShutdownGraceMilliseconds=5000;
    const int KeepMacroLoaded=0;
    static Process ownedProcess;
    [DllImport("user32.dll")] static extern int GetSystemMetrics(int index);
    [DllImport("ole32.dll")] static extern int CreateBindCtx(int reserved,out IBindCtx context);
    static SldWorks FindProcessApplication(int pid) {
        IBindCtx context=null;
        IRunningObjectTable table=null;
        IEnumMoniker iterator=null;
        try {
            Marshal.ThrowExceptionForHR(CreateBindCtx(0,out context));
            context.GetRunningObjectTable(out table);
            table.EnumRunning(out iterator);
            var entry=new IMoniker[1];
            while(iterator.Next(1,entry,IntPtr.Zero)==0) {
                try {
                    string name;
                    entry[0].GetDisplayName(context,null,out name);
                    if(name.EndsWith("SolidWorks_PID_"+pid,StringComparison.OrdinalIgnoreCase)) {
                        object application;
                        table.GetObject(entry[0],out application);
                        var sw=application as SldWorks;
                        if(sw!=null && sw.GetProcessID()==pid) return sw;
                    }
                } catch(COMException) { }
                finally { Marshal.ReleaseComObject(entry[0]); }
            }
            return null;
        } finally {
            if(iterator!=null) Marshal.ReleaseComObject(iterator);
            if(table!=null) Marshal.ReleaseComObject(table);
            if(context!=null) Marshal.ReleaseComObject(context);
        }
    }
    static SldWorks StartApplication() {
        var type=Type.GetTypeFromProgID("SldWorks.Application");
        string server=(string)Registry.GetValue("HKEY_CLASSES_ROOT\\CLSID\\{"+type.GUID+"}\\LocalServer32",null,null);
        if(string.IsNullOrWhiteSpace(server)) throw new Exception("SolidWorks server registration is missing.");
        server=server.Trim();
        string executable=server.StartsWith("\"") ? server.Substring(1,server.IndexOf('"',1)-1) : server.Substring(0,server.IndexOf(".exe",StringComparison.OrdinalIgnoreCase)+4);
        var start=new ProcessStartInfo(executable,"/b");
        start.UseShellExecute=false;
        start.CreateNoWindow=true;
        start.WindowStyle=ProcessWindowStyle.Hidden;
        var process=Process.Start(start);
        ownedProcess=process;
        Console.WriteLine("Started dedicated test process PID="+process.Id);
        var timer=Stopwatch.StartNew();
        while(timer.ElapsedMilliseconds<StartupTimeoutMilliseconds && !process.HasExited) {
            var application=FindProcessApplication(process.Id);
            if(application!=null) return application;
            System.Threading.Thread.Sleep(StartupPollMilliseconds);
        }
        throw new Exception("Dedicated SolidWorks did not register its PID-specific application object.");
    }
    static string Number(double value) { return value.ToString("R",CultureInfo.InvariantCulture); }
    static string Clean(string value) { return value.Replace("\t"," ").Replace("\r"," ").Replace("\n"," "); }
    static void RemapCopiedReferences(SldWorks sw,string work) {
        string models=Path.Combine(work,"Models");
        string prefix="";
        foreach(string line in File.ReadAllLines(Path.Combine(work,"generator.ini"),Encoding.Unicode))
            if(line.StartsWith("SnapshotPrefix=",StringComparison.OrdinalIgnoreCase)) prefix=line.Substring("SnapshotPrefix=".Length).Trim();
        int changed=0;
        foreach(string path in Directory.GetFiles(models)) {
            string extension=Path.GetExtension(path);
            if(!extension.Equals(".sldasm",StringComparison.OrdinalIgnoreCase) && !extension.Equals(".sldprt",StringComparison.OrdinalIgnoreCase)) continue;
            var dependencies=sw.GetDocumentDependencies2(path,false,false,false) as Array;
            if(dependencies==null) continue;
            for(int index=1;index<dependencies.Length;index+=2) {
                string previous=Convert.ToString(dependencies.GetValue(index));
                if(string.IsNullOrWhiteSpace(previous)) continue;
                string next=Path.Combine(models,Path.GetFileName(previous));
                if(!File.Exists(next)) next=Path.Combine(models,prefix+Path.GetFileName(previous));
                if(!File.Exists(next) || string.Equals(next,previous,StringComparison.OrdinalIgnoreCase)) continue;
                File.SetAttributes(path,File.GetAttributes(path)&~FileAttributes.ReadOnly);
                if(!sw.ReplaceReferencedDocument(path,previous,next)) {
                    string kind=Path.GetExtension(previous);
                    if(kind.Equals(".sldasm",StringComparison.OrdinalIgnoreCase) || kind.Equals(".sldprt",StringComparison.OrdinalIgnoreCase))
                        throw new Exception("Cannot remap copied reference: "+path+" -> "+previous);
                } else changed++;
            }
        }
        var report=new StringBuilder();
        using(var hash=SHA256.Create()) foreach(string path in Directory.GetFiles(models)) {
            using(var stream=File.OpenRead(path)) report.AppendLine(Path.GetFileName(path)+"\t"+BitConverter.ToString(hash.ComputeHash(stream)).Replace("-","").ToLowerInvariant());
        }
        File.WriteAllText(Path.Combine(work,"inventory-input-hashes.tsv"),report.ToString(),Encoding.Unicode);
        Console.WriteLine("Remapped copied references="+changed);
    }
    static void CloseTestDocuments(SldWorks sw,string work) {
        var documents=sw.GetDocuments() as object[];
        if(documents==null) return;
        string root=Path.GetFullPath(work).TrimEnd('\\')+"\\";
        foreach(ModelDoc2 document in documents) {
            string path=document.GetPathName();
            if(!string.IsNullOrEmpty(path) && Path.GetFullPath(path).StartsWith(root,StringComparison.OrdinalIgnoreCase))
                sw.CloseDoc(document.GetTitle());
        }
    }
    static void FinishTestProcess(SldWorks sw,string work) {
        // Only the process explicitly started above is eligible for forced
        // cleanup. SolidWorks can retain its background OLE server on ExitApp.
        var cleanup=new System.Threading.Thread(delegate() {
            try { CloseTestDocuments(sw,work);sw.ExitApp(); } catch(COMException) { }
        });
        cleanup.IsBackground=true;
        cleanup.Start();
        if(ownedProcess!=null && !ownedProcess.WaitForExit(ShutdownGraceMilliseconds)) {
            ownedProcess.Kill();
            ownedProcess.WaitForExit(ShutdownGraceMilliseconds);
        }
        cleanup.Join(ShutdownGraceMilliseconds);
    }
    static void InspectResult(SldWorks sw,string work) {
        string output=Path.Combine(work,"output");
        string[] files=Directory.GetFiles(Path.Combine(output,"Drawings"),"*.SLDDRW");
        if(files.Length!=1) throw new Exception("Expected one test drawing");
        // Keep the input document open. Closing every referenced document can
        // invalidate enumerated COM objects or release a background OLE server.
        var open=sw.GetOpenDocumentByName(files[0]) as ModelDoc2;
        if(open!=null) sw.CloseDoc(open.GetTitle());
        int error=0,warning=0;
        var doc=(ModelDoc2)sw.OpenDoc6(files[0],DrawingDocument,SilentReadOnly,"",ref error,ref warning);
        if(doc==null) throw new Exception("Cannot reopen generated drawing: "+error);
        var drawing=(DrawingDoc)doc;
        var report=new StringBuilder();
        foreach(string name in (string[])drawing.GetSheetNames()) {
            drawing.ActivateSheet(name);
            var sheet=(Sheet)drawing.GetCurrentSheet();
            var properties=(double[])sheet.GetProperties();
            report.AppendLine("SHEET\t"+name+"\t"+sheet.SheetFormatVisible+"\t"+Number(properties[2])+"\t"+Number(properties[3]));
            var view=(View)((View)drawing.GetFirstView()).GetNextView();
            while(view!=null) {
                double[] box=(double[])view.GetOutline();
                report.AppendLine("VIEW\t"+name+"\t"+view.Name+"\t"+Number(box[0])+"\t"+Number(box[1])+"\t"+Number(box[2])+"\t"+Number(box[3]));
                var transform=(double[])((MathTransform)view.ModelToViewTransform).ArrayData;
                report.AppendLine("TRANSFORM\t"+name+"\t"+view.Name+"\t"+string.Join("\t",Array.ConvertAll(transform,Number)));
                var notes=view.GetNotes() as object[];
                if(notes!=null) foreach(Note note in notes) {
                    var extent=note.GetExtent() as double[];
                    report.AppendLine("VIEWNOTE\t"+view.Name+"\t"+Clean(note.GetText())+"\t"+
                        (extent==null ? "" : string.Join("\t",Array.ConvertAll(extent,Number))));
                }
                var display=(DisplayDimension)view.GetFirstDisplayDimension5();
                while(display!=null) {
                    var dimension=(Dimension)display.GetDimension2(0);
                    var data=(DisplayData)display.GetDisplayData();
                    var text=new List<string>();
                    for(int i=0;i<data.GetTextCount();i++) text.Add(Clean(data.GetTextAtIndex(i)));
                    string owner=name+"\t"+view.Name+"\t"+display.GetNameForSelection();
                    var annotation=(Annotation)display.GetAnnotation();
                    report.AppendLine("APPEARANCE\t"+owner+"\t"+annotation.Color+"\t"+annotation.LayerOverride);
                    report.AppendLine("ARROWMODE\t"+owner+"\t"+display.ArrowSide);
                    report.AppendLine("DIM\t"+owner+"\t"+display.Type2+"\t"+Number(dimension.SystemValue)+"\t"+dimension.FullName+"\t"+display.IsHoleCallout()+"\t"+string.Join(" ",text)+"\t"+data.GetArrowHeadCount());
                    if(display.IsHoleCallout()) report.AppendLine("CALLOUT\t"+owner+"\t"+Clean(display.GetText(CalloutAboveDefinition))+"\t"+Clean(display.GetText(CalloutAbove))+"\t"+Clean(display.GetText(CalloutBelowDefinition)));
                    foreach(var dataKind in new[]{"display","annotation"}) {
                        var textData=dataKind=="display" ? data : (DisplayData)annotation.GetDisplayData();
                        if(textData==null) continue;
                        for(int i=0;i<textData.GetTextCount();i++) {
                            double[] textPosition=(double[])textData.GetTextPositionAtIndex(i);
                            report.AppendLine("TEXTGEOMETRY\t"+owner+"\t"+dataKind+"\t"+i+"\t"+textData.GetTextRefPositionAtIndex(i)+"\t"+
                                Number(textPosition[0])+"\t"+Number(textPosition[1])+"\t"+Number(textData.GetTextHeightAtIndex(i))+"\t"+
                                Number(textData.GetTextAngleAtIndex(i))+"\t"+textData.GetTextFontAtIndex(i)+"\t"+Clean(textData.GetTextAtIndex(i)));
                        }
                    }
                    if(data.GetTextCount()>0) {
                        double[] position=(double[])data.GetTextPositionAtIndex(0);
                        report.AppendLine("TEXT\t"+owner+"\t"+Number(position[0])+"\t"+Number(position[1]));
                    }
                    var references=dimension.ReferencePoints as object[];
                    if(references!=null) foreach(MathPoint reference in references) {
                        double[] point=(double[])reference.ArrayData;
                        report.AppendLine("REFERENCE\t"+owner+"\t"+Number(point[0])+"\t"+Number(point[1])+"\t"+Number(point[2]));
                    }
                    for(int i=0;i<data.GetArrowHeadCount();i++) {
                        double[] arrow=(double[])data.GetArrowHeadAtIndex2(i);
                        report.AppendLine("ARROW\t"+owner+"\t"+Number(arrow[0])+"\t"+Number(arrow[1]));
                        report.AppendLine("ARROWGEOMETRY\t"+owner+"\t"+string.Join("\t",Array.ConvertAll(arrow,Number)));
                    }
                    for(int i=0;i<data.GetLineCount();i++) {
                        double[] line=(double[])data.GetLineAtIndex2(i);
                        report.AppendLine("LINE\t"+owner+"\t"+Number(line[4])+"\t"+Number(line[5])+"\t"+Number(line[7])+"\t"+Number(line[8]));
                    }
                    for(int i=0;i<data.GetArcCount();i++) {
                        var arc=(double[])data.GetArcAtIndex2(i);
                        report.AppendLine("ARC\t"+owner+"\t"+string.Join("\t",Array.ConvertAll(arc,Number)));
                    }
                    display=(DisplayDimension)display.GetNext5();
                }
                view=(View)view.GetNextView();
            }
        }
        File.WriteAllText(Path.Combine(output,"saved-dimensions.tsv"),report.ToString(),Encoding.Unicode);
    }
    static void InspectProfileGeometry(ModelDoc2 model,string work) {
        var part=model as PartDoc;
        if(part==null) throw new Exception("Profile geometry inspection requires a copied part.");
        var bodies=part.GetBodies2(SolidBody,true) as object[];
        if(bodies==null || bodies.Length!=1) throw new Exception("Profile inspection requires one solid body.");
        Face2 cap=null;
        foreach(Face2 face in (object[])((Body2)bodies[0]).GetFaces())
            if(((Surface)face.GetSurface()).IsPlane() && (cap==null || face.GetArea()>cap.GetArea())) cap=face;
        if(cap==null) throw new Exception("The copied profile has no planar cap.");
        var result=new StringBuilder("{\"loops\":[");
        bool firstLoop=true;
        foreach(Loop2 loop in (object[])cap.GetLoops()) {
            if(!firstLoop) result.Append(',');firstLoop=false;
            result.Append("{\"outer\":").Append(loop.IsOuter() ? "true" : "false").Append(",\"segments\":[");
            bool firstEdge=true;
            foreach(Edge edge in (object[])loop.GetEdges()) {
                if(!((Curve)edge.GetCurve()).IsLine()) throw new Exception("The inspected profile contains a curved boundary.");
                if(!firstEdge) result.Append(',');firstEdge=false;
                var a=(double[])((Vertex)edge.GetStartVertex()).GetPoint();
                var b=(double[])((Vertex)edge.GetEndVertex()).GetPoint();
                result.Append("[[").Append(string.Join(",",Array.ConvertAll(a,Number))).Append("],[").Append(string.Join(",",Array.ConvertAll(b,Number))).Append("]]");
            }
            result.Append("]}");
        }
        result.Append("]}");
        File.WriteAllText(Path.Combine(work,"profile-geometry.json"),result.ToString(),Encoding.UTF8);
    }
    public static void Run(string work,string modelFileName,bool inventoryOnly,bool remapReferences,bool profileGeometryOnly,int macroTimeoutMilliseconds,bool inspectOnly) {
        if(inspectOnly && (inventoryOnly || remapReferences || profileGeometryOnly))
            throw new ArgumentException("InspectOnly requires a generated single-part fixture.");
        var prior=new HashSet<int>();
        foreach(Process process in Process.GetProcessesByName("SLDWORKS")) prior.Add(process.Id);
        Console.WriteLine("Starting an isolated SolidWorks process");
        SldWorks sw=StartApplication();
        int pid=sw.GetProcessID();
        if(prior.Contains(pid)) throw new Exception("The test must own its SolidWorks process.");
        Console.WriteLine("Six-view test owns SolidWorks PID="+pid);
        File.WriteAllText(Path.Combine(work,"owned-pid.txt"),pid.ToString());
        try {
            sw.UserControlBackground=true;
            sw.Visible=true;
            sw.FrameLeft=GetSystemMetrics(VirtualLeft)+GetSystemMetrics(VirtualWidth)+OffscreenOffset;
            sw.FrameTop=GetSystemMetrics(VirtualTop)+GetSystemMetrics(VirtualHeight)+OffscreenOffset;
            if(remapReferences) {
                if(!inventoryOnly) throw new Exception("Reference remapping is restricted to copied inventory fixtures.");
                RemapCopiedReferences(sw,work);
            }
            string input=Path.Combine(work,inventoryOnly ? "Models" : "input",modelFileName);
            sw.SetCurrentWorkingDirectory(Path.GetDirectoryName(input));
            int error=0,warning=0;
            int documentType=modelFileName.EndsWith(".sldasm",StringComparison.OrdinalIgnoreCase) ? AssemblyDocument : PartDocument;
            ModelDoc2 model=sw.OpenDoc6(input,documentType,SilentReadOnly,"",ref error,ref warning);
            if(model==null) throw new Exception("Cannot open the copied test model: "+error);
            var active=(ModelDoc2)sw.ActivateDoc3(Path.GetFileName(input),false,0,ref error);
            if(active==null || !string.Equals(active.GetPathName(),input,StringComparison.OrdinalIgnoreCase))
                throw new Exception("The copied test model did not become active.");
            if(inspectOnly) {
                InspectResult(sw,work);
                Console.WriteLine("Persisted drawing inspection completed: "+Path.Combine(work,"output"));
                return;
            }
            if(profileGeometryOnly) {
                InspectProfileGeometry(model,work);
                Console.WriteLine("Copied profile geometry inspection completed.");
                return;
            }
            // Each invocation must validate its own events, including fixture retries.
            File.WriteAllText(Path.Combine(work,"progress.tsv"),string.Empty,Encoding.Unicode);
            Console.WriteLine(inventoryOnly ? "Running copied saved-reference inventory" : "Running the canonical six-view generator");
            using(var watchdog=new System.Threading.Timer(delegate(object unused) {
                try { if(ownedProcess!=null && !ownedProcess.HasExited) ownedProcess.Kill(); } catch(ArgumentException) { }
            },null,inventoryOnly ? InventoryTimeoutMilliseconds : macroTimeoutMilliseconds,System.Threading.Timeout.Infinite)) {
                int macroError=0;
                if(!sw.RunMacro2(Path.Combine(work,"generator.swp"),"Macro11",inventoryOnly ? "InspectInventory" : "main",KeepMacroLoaded,out macroError))
                    throw new Exception("RunMacro2 returned false; native macro error="+macroError);
            }
            string events=File.ReadAllText(Path.Combine(work,"progress.tsv"),System.Text.Encoding.Unicode);
            if(!events.Contains("\tCOMPLETE\t")) {
                if(File.Exists(Path.Combine(work,"output","diagnostic-failed.pdf"))) InspectResult(sw,work);
                throw new Exception("Generation did not complete. Read progress.tsv and output/run.log.");
            }
            if(!inventoryOnly) InspectResult(sw,work);
            Console.WriteLine("Six-view generation completed: "+Path.Combine(work,"output"));
        } finally {
            FinishTestProcess(sw,work);
        }
    }
}
'@
Add-Type -TypeDefinition $source -ReferencedAssemblies $api
[SixViewGenerationTest]::Run($work,$ModelFileName,$InventoryOnly.IsPresent,$RemapInventoryReferences.IsPresent,$ProfileGeometryOnly.IsPresent,[int][TimeSpan]::FromSeconds($MacroTimeoutSeconds).TotalMilliseconds,$InspectOnly.IsPresent)

param([string]$AuditDirectory = '')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'TestPaths.ps1')
$api='C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist\SolidWorks.Interop.sldworks.dll'
Add-Type -Path $api
$code=@'
using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using SolidWorks.Interop.sldworks;
public static class PaperResultInspection {
 [DllImport("oleaut32.dll",PreserveSig=false)] static extern void GetActiveObject(ref Guid id,IntPtr reserved,[MarshalAs(UnmanagedType.IUnknown)]out object app);
 public static void Run(string root) {
  Type t=Type.GetTypeFromProgID("SldWorks.Application");Guid id=t.GUID;object app;bool owned=false;
  try{GetActiveObject(ref id,IntPtr.Zero,out app);}catch(COMException){app=Activator.CreateInstance(t);owned=true;}
  SldWorks sw=(SldWorks)app;sw.Visible=true;
  ModelDoc2 prior=(ModelDoc2)sw.ActiveDoc;string priorPath=prior==null?"":prior.GetPathName();
  var result=new List<object>();ModelDoc2 d=null;
  try {
   foreach(string name in new[]{"A4.landscape","A4.portrait","A3.landscape","A3.portrait","assembly"}) {
    string folder=Path.Combine(root,name,"out","Drawings");
    foreach(string path in Directory.GetFiles(folder,"*.slddrw")) {
     int error=0,warning=0;d=sw.OpenDoc6(path,3,3,"",ref error,ref warning);
     if(d==null)throw new Exception("Cannot inspect "+path+": "+error);
     d.Visible=true;DrawingDoc dr=(DrawingDoc)d;var pages=new List<object>();
     foreach(string sheetName in (string[])dr.GetSheetNames()) {
      dr.ActivateSheet(sheetName);d.ForceRebuild3(false);
      Sheet sheet=(Sheet)dr.GetCurrentSheet();double[] props=(double[])sheet.GetProperties();
      View view=(View)dr.GetFirstView();object[] annotations=(object[])view.GetAnnotations();
      var notes=new Dictionary<string,string>();
      if(annotations!=null)foreach(Annotation a in annotations) {
       Note n=a.GetSpecificAnnotation() as Note;
       if(n!=null&&n.GetName().StartsWith("DSTU_")) {
        if(notes.ContainsKey(n.GetName()))throw new Exception("Duplicate template note: "+n.GetName());
        notes.Add(n.GetName(),n.GetText());
       }
      }
      if(notes.Count!=51)throw new Exception("Template frame note count "+notes.Count+" on "+name+" / "+sheetName);
      if(!notes["DSTU_SheetNumber"].Contains((pages.Count+1).ToString()))throw new Exception("Wrong per-sheet number");
      if(notes["DSTU_Scale"].StartsWith("0"))throw new Exception("Unnormalized drawing scale");
      var views=new List<object>();view=(View)view.GetNextView();
      while(view!=null){views.Add(new {name=view.Name,visible=view.GetVisible(),bounds=view.GetOutline(),dimensions=view.GetDisplayDimensionCount()});view=(View)view.GetNextView();}
      pages.Add(new {name=sheetName,width=props[5],height=props[6],scaleNumerator=props[2],scaleDenominator=props[3],format=sheet.GetTemplateName(),notes=notes,views=views});
     }
     result.Add(new {test=name,path=path,pages=pages});
     sw.CloseDoc(d.GetTitle());d=null;
    }
   }
   var serializer=new JavaScriptSerializer();serializer.MaxJsonLength=4000000;
   File.WriteAllText(Path.Combine(root,"native-inspection.json"),serializer.Serialize(result),System.Text.Encoding.UTF8);
   Console.WriteLine("PASS: all generated native sheets preserve 51 template notes and correct sheet numbers/scales");
  }finally {
   if(d!=null)sw.CloseDoc(d.GetTitle());
   object[] docs=(object[])sw.GetDocuments();
   if(docs!=null)foreach(ModelDoc2 doc in docs)if(doc.GetPathName().StartsWith(root,StringComparison.OrdinalIgnoreCase))sw.CloseDoc(doc.GetTitle());
   if(priorPath!=""){int error=0;sw.ActivateDoc3(priorPath,false,0,ref error);}
   if(owned)sw.ExitApp();Marshal.ReleaseComObject(sw);
  }
 }
}
'@
Add-Type -TypeDefinition $code -ReferencedAssemblies $api,'System.Web.Extensions.dll'
if ($AuditDirectory -eq '') { $AuditDirectory = 'paper_formats' }
[PaperResultInspection]::Run((Get-ExternalTestPath $AuditDirectory))

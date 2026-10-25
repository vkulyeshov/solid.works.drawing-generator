using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using SolidWorks.Interop.sldworks;

// Form 1, DSTU GOST 2.104:2006, from the supplied rules.pdf, pp. 11-14, 18.
// Dimensions are millimetres. All entities are native sheet-format geometry/notes.
public static class DstuTemplateBuilder {
 [DllImport("oleaut32.dll", PreserveSig=false)]
 static extern void GetActiveObject(ref Guid id,IntPtr reserved,[MarshalAs(UnmanagedType.IUnknown)]out object app);
 const string Font="GOST Type BU";
 class Spec {
  public string Name,Format; public double W,H,CW,CH; public bool Rotate,PortraitA3; public int Paper;
  public Spec(string name,string fmt,double w,double h,bool rotate,bool portrait,int paper) {
   Name=name;Format=fmt;W=w;H=h;Rotate=rotate;PortraitA3=portrait;Paper=paper;
   CW=rotate?h:w;CH=rotate?w:h;
  }
 }
 class Cell {
  public string Name,Text; public double X,Y,W,H,Size,Angle;
  public Cell(string name,string text,double x,double y,double w,double h,double size,double angle) {
   Name=name;Text=text;X=x;Y=y;W=w;H=h;Size=size;Angle=angle;
  }
 }
 static readonly Spec[] Specs={
  new Spec("DSTU_A4_Portrait","A4",210,297,false,false,7),
  new Spec("DSTU_A4_Landscape","A4",297,210,false,false,6),
  new Spec("DSTU_A3_Portrait","A3",297,420,false,true,12),
  new Spec("DSTU_A3_Landscape","A3",420,297,false,false,8)
 };
 static readonly string[] Properties={"Designation","Description","Material","Mass","Liter1","Liter2","Liter3","Organization","SheetNumber","DrawnBy","DrawnDate","CheckedBy","CheckedDate","TechnicalControlBy","TechnicalControlDate","NormControlBy","NormControlDate","ApprovedBy","ApprovedDate","InventoryOriginal","OriginalAcceptance","InventoryReplacement","InventoryDuplicate","DuplicateAcceptance","ReferenceDesignation","FirstUse","CopiedBy"};
 static string Link(string name){return "$PRP:\""+name+"\"";}
 static double UpperAuxHeight(Spec s){return s.CH<270?50:120;}
 static double[] Point(Spec s,double x,double y){return s.Rotate?new[]{(s.CH-y)/1000,x/1000,0}:new[]{x/1000,y/1000,0};}
 static void Line(ModelDoc2 d,Spec s,string layer,double x1,double y1,double x2,double y2) {
  double[] a=Point(s,x1,y1),b=Point(s,x2,y2);
  SketchSegment seg=(SketchSegment)d.SketchManager.CreateLine(a[0],a[1],0,b[0],b[1],0);
  if(seg==null)throw new Exception("Cannot create format line");seg.Layer=layer;seg.LayerOverride=0;
 }
 static void Rect(ModelDoc2 d,Spec s,string layer,double x,double y,double w,double h){
  Line(d,s,layer,x,y,x+w,y);Line(d,s,layer,x+w,y,x+w,y+h);
  Line(d,s,layer,x+w,y+h,x,y+h);Line(d,s,layer,x,y+h,x,y);
 }
 static void DrawGeometry(ModelDoc2 d,Spec s){
  string main="DSTU_MAIN",thin="DSTU_THIN";double x=s.CW-190,y=5;
  SketchManager sm=d.SketchManager;sm.AddToDB=true;
  try{
   Rect(d,s,main,20,5,s.CW-25,s.CH-10);
   // Main inscription: 185 x 55. Its bottom/right edges coincide with the frame.
   Line(d,s,main,x,y,x,y+55);Line(d,s,main,x,y+55,x+185,y+55);
   Line(d,s,main,x+65,y,x+65,y+55);
   Line(d,s,main,x+65,y+40,x+185,y+40);
   Line(d,s,main,x+65,y+15,x+185,y+15);
   Line(d,s,main,x+135,y,x+135,y+40);
   foreach(double xx in new[]{17.0,40.0,55.0})Line(d,s,thin,x+xx,y,x+xx,y+55);
   Line(d,s,thin,x+7,y+30,x+7,y+55);
   for(int yy=5;yy<55;yy+=5)Line(d,s,thin,x,y+yy,x+65,y+yy);
   Line(d,s,thin,x+135,y+35,x+185,y+35);
   Line(d,s,thin,x+135,y+20,x+185,y+20);
   foreach(double xx in new[]{150.0,167.0})Line(d,s,thin,x+xx,y+20,x+xx,y+40);
   foreach(double xx in new[]{140.0,145.0})Line(d,s,thin,x+xx,y+20,x+xx,y+35);
   Line(d,s,thin,x+155,y+15,x+155,y+20);
   // Graph 26: 70 x 14, rotated per clause 6.1 and appendix A.
   if(s.PortraitA3){Line(d,s,main,s.CW-19,s.CH-75,s.CW-19,s.CH-5);Line(d,s,main,s.CW-19,s.CH-75,s.CW-5,s.CH-75);}
   else {Line(d,s,main,20,s.CH-19,90,s.CH-19);Line(d,s,main,90,s.CH-19,90,s.CH-5);}
   // Additional inventory/acceptance columns 19-25, matching form 1 and reference.
   Rect(d,s,main,8,5,12,145);Line(d,s,thin,13,5,13,150);
   foreach(double yy in new[]{30.0,65.0,90.0,115.0})Line(d,s,thin,8,yy,20,yy);
   double ah=UpperAuxHeight(s);
   Rect(d,s,main,8,s.CH-5-ah,12,ah);Line(d,s,thin,13,s.CH-5-ah,13,s.CH-5);
   Line(d,s,thin,8,s.CH-5-ah/2,20,s.CH-5-ah/2);
  }finally{sm.AddToDB=false;}
 }
 static List<Cell> Cells(Spec s){
  var c=new List<Cell>();double x=s.CW-190,y=5;
  c.Add(new Cell("Designation",Link("Designation"),x+65,y+40,120,15,5,0));
  c.Add(new Cell("Description",Link("Description"),x+65,y+15,70,25,5,0));
  c.Add(new Cell("Material",Link("Material"),x+65,y,70,15,3.5,0));
  c.Add(new Cell("Organization",Link("Organization"),x+135,y,50,15,3.5,0));
  string[] heads={"\u041b\u0456\u0442.","\u041c\u0430\u0441\u0430","\u041c\u0430\u0441\u0448\u0442\u0430\u0431"};double[] starts={135,150,167},widths={15,17,18};
  for(int i=0;i<3;i++)c.Add(new Cell("Header_"+i,heads[i],x+starts[i],y+35,widths[i],5,2.5,0));
  for(int i=0;i<3;i++)c.Add(new Cell("Liter"+(i+1),Link("Liter"+(i+1)),x+135+5*i,y+20,5,15,3.5,0));
  c.Add(new Cell("Mass",Link("Mass"),x+150,y+20,17,15,3.5,0));
  c.Add(new Cell("Scale",Link("SW-Sheet Scale"),x+167,y+20,18,15,3.5,0));
  c.Add(new Cell("SheetNumber","\u0410\u0440\u043a\u0443\u0448 "+Link("SheetNumber"),x+135,y+15,20,5,2.5,0));
  c.Add(new Cell("SheetCount","\u0410\u0440\u043a\u0443\u0448\u0456\u0432 "+Link("SW-Total Sheets"),x+155,y+15,30,5,2.5,0));
  string[] rh={"\u0417\u043c.","\u0410\u0440\u043a\u0443\u0448","\u2116 \u0434\u043e\u043a\u0443\u043c.","\u041f\u0456\u0434\u043f.","\u0414\u0430\u0442\u0430"};double[] rs={0,7,17,40,55},rw={7,10,23,15,10};
  for(int i=0;i<5;i++)c.Add(new Cell("RevisionHeader_"+i,rh[i],x+rs[i],y+30,rw[i],5,2.5,0));
  string[] roles={"\u0417\u0430\u0442\u0432.","\u041d. \u043a\u043e\u043d\u0442\u0440.","","\u0422. \u043a\u043e\u043d\u0442\u0440.","\u041f\u0435\u0440\u0435\u0432.","\u0420\u043e\u0437\u0440\u043e\u0431."};
  string[] names={"ApprovedBy","NormControlBy","","TechnicalControlBy","CheckedBy","DrawnBy"};
  string[] dates={"ApprovedDate","NormControlDate","","TechnicalControlDate","CheckedDate","DrawnDate"};
  for(int i=0;i<6;i++){
   if(roles[i]=="")continue;
   c.Add(new Cell("Role_"+i,roles[i],x,y+5*i,17,5,2.5,0));
   c.Add(new Cell(names[i],Link(names[i]),x+17,y+5*i,23,5,2.5,0));
   c.Add(new Cell(dates[i],Link(dates[i]),x+55,y+5*i,10,5,2.5,0));
  }
  if(s.PortraitA3)c.Add(new Cell("Designation26",Link("Designation"),s.CW-19,s.CH-75,14,70,5,90));
  else c.Add(new Cell("Designation26",Link("Designation"),20,s.CH-19,70,14,5,180));
  string[] inv={"\u0406\u043d\u0432. \u2116 \u043e\u0440\u0438\u0433.","\u041f\u0456\u0434\u043f. \u0442\u0430 \u0434\u0430\u0442\u0430","\u0417\u0430\u043c. \u0456\u043d\u0432. \u2116","\u0406\u043d\u0432. \u2116 \u0434\u0443\u0431\u043b.","\u041f\u0456\u0434\u043f. \u0442\u0430 \u0434\u0430\u0442\u0430"};
  string[] ip={"InventoryOriginal","OriginalAcceptance","InventoryReplacement","InventoryDuplicate","DuplicateAcceptance"};
  double[] iy={5,30,65,90,115},ih={25,35,25,25,35};
  for(int i=0;i<5;i++){
   c.Add(new Cell("InventoryLabel_"+i,inv[i],8,iy[i],5,ih[i],2.5,90));
   c.Add(new Cell(ip[i],Link(ip[i]),13,iy[i],7,ih[i],2.5,90));
  }
  double ah=UpperAuxHeight(s),half=ah/2;
  c.Add(new Cell("ReferenceLabel","\u0414\u043e\u0432\u0456\u0434. \u2116",8,s.CH-5-ah,5,half,2.5,90));
  c.Add(new Cell("ReferenceDesignation",Link("ReferenceDesignation"),13,s.CH-5-ah,7,half,3.5,90));
  c.Add(new Cell("FirstUseLabel","\u041f\u0435\u0440\u0432. \u0437\u0430\u0441\u0442\u043e\u0441.",8,s.CH-5-half,5,half,2.5,90));
  c.Add(new Cell("FirstUse",Link("FirstUse"),13,s.CH-5-half,7,half,3.5,90));
  c.Add(new Cell("CopiedBy","\u041a\u043e\u043f\u0456\u044e\u0432\u0430\u0432 "+Link("CopiedBy"),x+65,0,70,5,2.5,0));
  c.Add(new Cell("Format","\u0424\u043e\u0440\u043c\u0430\u0442 "+s.Format,x+135,0,50,5,2.5,0));
  return c;
 }
 static Note AddCell(ModelDoc2 d,Spec s,Cell cell){
  d.ClearSelection2(true);Note n=(Note)d.InsertNote(cell.Text);
  if(n==null)throw new Exception("Cannot insert "+cell.Name);
  n.SetName("DSTU_"+cell.Name);Annotation a=(Annotation)n.GetAnnotation();a.Layer="DSTU_TEXT";
  TextFormat tf=(TextFormat)a.GetTextFormat(0);tf.TypeFaceName=Font;tf.CharHeight=cell.Size/1000;tf.Italic=true;tf.Bold=false;tf.WidthFactor=1;
  // Eight-character DD.MM.YY dates must fit the 10 mm form-1 date column.
  if(cell.Name.EndsWith("Date"))tf.WidthFactor=.8;
  if(!a.SetTextFormat(0,false,tf))throw new Exception("Cannot format "+cell.Name);
  n.SetTextJustification(2);n.SetTextVerticalJustification(1);
  n.Angle=(cell.Angle+(s.Rotate?90:0))*Math.PI/180;
  double[] p=Point(s,cell.X+cell.W/2,cell.Y+cell.H/2);a.SetPosition2(p[0],p[1],0);
  return n;
 }
 static void Save(ModelDoc2 d,string path){int error=0,warning=0;if(!d.Extension.SaveAs(path,0,1,null,ref error,ref warning))throw new Exception("Save failed "+path+": "+error);}
 static void Configure(ModelDoc2 d){
  d.Extension.SetUserPreferenceInteger(13,0,6);d.Extension.SetUserPreferenceInteger(263,0,0);d.Extension.SetUserPreferenceInteger(47,0,0);
  d.Extension.SetUserPreferenceInteger(24,0,2);d.Extension.SetUserPreferenceInteger(28,0,1);
  foreach(int i in new[]{0,1}){TextFormat tf=(TextFormat)d.GetUserPreferenceTextFormat(i);tf.TypeFaceName=Font;tf.CharHeight=.0035;tf.Italic=true;d.SetUserPreferenceTextFormat(i,tf);}
  LayerMgr layers=(LayerMgr)d.GetLayerManager();layers.AddLayer("DSTU_MAIN","\u041e\u0441\u043d\u043e\u0432\u043d\u0456 \u043b\u0456\u043d\u0456\u0457 \u0444\u043e\u0440\u043c\u0438 1",0,0,3);layers.AddLayer("DSTU_THIN","\u0422\u043e\u043d\u043a\u0456 \u043b\u0456\u043d\u0456\u0457 \u0444\u043e\u0440\u043c\u0438 1",0,0,1);layers.AddLayer("DSTU_TEXT","\u041d\u0430\u043f\u0438\u0441\u0438 \u0442\u0430 \u043f\u043e\u0432'\u044f\u0437\u0430\u043d\u0456 \u0432\u043b\u0430\u0441\u0442\u0438\u0432\u043e\u0441\u0442\u0456",0,0,1);
  CustomPropertyManager cp=d.Extension.CustomPropertyManager[""];
  foreach(string p in Properties)cp.Add3(p,30,"",2);
  cp.Add3("TemplateStandard",30,"\u0414\u0421\u0422\u0423 \u0413\u041e\u0421\u0422 2.104:2006; \u0444\u043e\u0440\u043c\u0430 1",2);
 }
 static void Demo(ModelDoc2 d){
  CustomPropertyManager cp=d.Extension.CustomPropertyManager[""];
  string[] p={"Designation","Description","Material","Mass","Organization","DrawnBy","CheckedBy","NormControlBy","ApprovedBy","DrawnDate","InventoryOriginal","FirstUse"};
  string[] v={"\u041a\u0420\u041d 02.04.002","\u0411\u043e\u043a\u043e\u0432\u0430 \u0441\u0442\u0456\u043d\u043a\u0430","\u041b\u0438\u0441\u0442 2,0 \u0413\u041e\u0421\u0422 19903-2015\n\u0421\u04423\u0441\u043f \u0413\u041e\u0421\u0422 16523-97","1,42","\u041f\u0440\u0438\u043a\u043b\u0430\u0434 \u0437\u0430\u043f\u043e\u0432\u043d\u0435\u043d\u043d\u044f","\u041c\u0430\u0441\u043b\u044e\u043a","\u0411\u0443\u0440\u043d\u0430\u0442\u043d\u0438\u0439","\u0411\u0443\u0440\u043d\u0430\u0442\u043d\u0438\u0439","\u0411\u0443\u0440\u043d\u0430\u0442\u043d\u0438\u0439","25.10.26","12345","\u041a\u0420\u041d 02.04.000"};
  for(int i=0;i<p.Length;i++)cp.Add3(p[i],30,v[i],2);
 }
 public static void Run(string output,bool verifyOnly,string only){
  Directory.CreateDirectory(output);Directory.CreateDirectory(Path.Combine(output,"Formats"));Directory.CreateDirectory(Path.Combine(output,"Preview"));Directory.CreateDirectory(Path.Combine(output,"Verification"));
  Type type=Type.GetTypeFromProgID("SldWorks.Application");Guid id=type.GUID;object obj;bool owned=false;
  try{GetActiveObject(ref id,IntPtr.Zero,out obj);}catch(COMException){obj=Activator.CreateInstance(type);owned=true;}
  SldWorks sw=(SldWorks)obj;ModelDoc2 previous=(ModelDoc2)sw.ActiveDoc;string previousPath=previous==null?"":previous.GetPathName();
  sw.Visible=true;Console.WriteLine("SolidWorks "+sw.RevisionNumber()+" PID="+sw.GetProcessID()+" owned="+owned);
  bool oldEmbed=sw.GetUserPreferenceToggle(324),oldColor=sw.GetUserPreferenceToggle(323);sw.SetUserPreferenceToggle(324,true);sw.SetUserPreferenceToggle(323,false);
  var records=new List<object>();string seed="C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS 2025\\templates\\gost-part drw.drwdot";
  ModelDoc2 d=null;
  try{
   foreach(Spec s in Specs){
    if(only!="" && s.Name!=only)continue;
    string templateFile=s.Name=="DSTU_A4_Landscape"?"DSTU_A4_Landscape_Horizontal.drwdot":s.Name+".drwdot";
    string template=Path.Combine(output,templateFile),format=Path.Combine(output,"Formats",s.Name+".slddrt");
    string proofTemplate=template;
    Console.WriteLine("Creating/verifying "+s.Name);
    if(!verifyOnly){
     d=(ModelDoc2)sw.NewDocument(seed,s.Paper,s.W/1000,s.H/1000);if(d==null)throw new Exception("NewDocument failed");d.Visible=true;
     DrawingDoc dr=(DrawingDoc)d;Sheet sh=(Sheet)dr.GetCurrentSheet();
     if(!dr.SetupSheet5(sh.GetName(),s.Paper,13,1,1,true,"",s.W/1000,s.H/1000,"",true))throw new Exception("SetupSheet failed");
     Configure(d);dr.ActivateView("");d.ClearSelection2(true);dr.EditTemplate();DrawGeometry(d,s);
     foreach(Cell cell in Cells(s))AddCell(d,s,cell);
     ((LayerMgr)d.GetLayerManager()).SetCurrentLayer("-None-");dr.EditSheet();d.ForceRebuild3(false);d.ViewZoomtofit2();
     if(!sh.SaveFormat(format))throw new Exception("SaveFormat failed");
     // Associate the saved format through the custom-format API, not just its path.
     if(!dr.SetupSheet5(sh.GetName(),s.Paper,12,1,1,true,format,s.W/1000,s.H/1000,"",true))throw new Exception("Cannot apply saved format");
     d.ForceRebuild3(false);
     // Saving via a separate file also works when the user has the old template open.
     string staged=Path.Combine(output,"Verification","Stage_"+s.Name+".drwdot");
     Save(d,staged);sw.CloseDoc(d.GetTitle());d=null;
     File.Copy(staged,template,true);proofTemplate=staged;
    }
    // Proof is made from a new document instantiated from the SAVED template.
    d=(ModelDoc2)sw.NewDocument(proofTemplate,s.Paper,s.W/1000,s.H/1000);if(d==null)throw new Exception("Cannot instantiate saved template");d.Visible=true;
    d.ForceRebuild3(false);d.ViewZoomtofit2();
    DrawingDoc check=(DrawingDoc)d;Sheet sheet=(Sheet)check.GetCurrentSheet();
    int reloadResult=sheet.ReloadTemplate(false);
    if(reloadResult!=0)throw new Exception("External sheet-format reload failed: "+s.Name+" code="+reloadResult);
    double[] props=(double[])sheet.GetProperties();
    if(Math.Abs(props[5]-s.W/1000)>1e-7||Math.Abs(props[6]-s.H/1000)>1e-7)throw new Exception("Wrong paper dimensions");
    Save(d,Path.Combine(output,"Preview",s.Name+".pdf"));
    Demo(d);d.ForceRebuild3(false);d.GraphicsRedraw2();
    var notes=new List<object>();View view=(View)check.GetFirstView();object[] annotations=(object[])view.GetAnnotations();
    if(annotations!=null)foreach(Annotation annotation in annotations){
     Note note=annotation.GetSpecificAnnotation() as Note;if(note==null)continue;
     notes.Add(new {name=note.GetName(),text=note.GetText(),link=note.PropertyLinkedText,angle=note.Angle,extent=note.GetExtent()});
    }
    Save(d,Path.Combine(output,"Verification",s.Name+"_filled.pdf"));
    Save(d,Path.Combine(output,"Verification",s.Name+"_filled.slddrw"));
    records.Add(new{name=s.Name,template_file=templateFile,width_mm=s.W,height_mm=s.H,formatPath=sheet.GetTemplateName(),formatReloadResult=reloadResult,notes=notes,cells=Cells(s),expectedNotes=Cells(s).Count,template_bytes=new FileInfo(template).Length,format_bytes=new FileInfo(format).Length});
    sw.CloseDoc(d.GetTitle());d=null;Console.WriteLine("OK "+s.Name+" "+s.W+"x"+s.H+" mm");
   }
   var serializer=new JavaScriptSerializer();serializer.MaxJsonLength=4000000;
   File.WriteAllText(Path.Combine(output,"Verification","native-report.json"),serializer.Serialize(records),System.Text.Encoding.UTF8);
  }finally{
   if(d!=null)sw.CloseDoc(d.GetTitle());sw.SetUserPreferenceToggle(324,oldEmbed);sw.SetUserPreferenceToggle(323,oldColor);
   if(previousPath!=""){int e=0;sw.ActivateDoc3(previousPath,false,0,ref e);}
   if(owned)sw.ExitApp();Marshal.ReleaseComObject(sw);
  }
 }
}

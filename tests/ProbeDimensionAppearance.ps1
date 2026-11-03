param([Parameter(Mandatory=$true)][string]$DrawingPath,
      [Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestPaths.ps1')
$output = Get-ExternalTestPath $OutputDirectory
New-Item -ItemType Directory -Path $output -Force | Out-Null
$drawing = (Resolve-Path -LiteralPath $DrawingPath).Path
$api = 'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist\SolidWorks.Interop.sldworks.dll'
Add-Type -Path $api
$source = @'
using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
public static class DimensionAppearanceProbe {
    const int DrawingDocument=3, SilentReadOnly=3, SaveSilent=1;
    public static void Run(string path,string output) {
        var prior=new HashSet<int>();
        foreach(var p in Process.GetProcessesByName("SLDWORKS")) prior.Add(p.Id);
        var sw=(SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
        if(prior.Contains(sw.GetProcessID())) throw new Exception("A separate test process is required");
        bool oldPrint=sw.GetUserPreferenceToggle(327),oldColor=sw.GetUserPreferenceToggle(323);
        try {
            sw.UserControlBackground=true;
            sw.Visible=false;
            int error=0,warning=0;
            var doc=(ModelDoc2)sw.OpenDoc6(path,DrawingDocument,SilentReadOnly,"",ref error,ref warning);
            if(doc==null) throw new Exception("Cannot open test drawing: "+error);
            var drawing=(DrawingDoc)doc;
            var ext=(ModelDocExtension)doc.Extension;
            Console.WriteLine("Document dimension thickness="+ext.GetUserPreferenceInteger(63,0)+"; custom="+ext.GetUserPreferenceDouble(94,0));
            Console.WriteLine("Printer normal="+ext.GetUserPreferenceDouble(49,0)+"; use printer weights="+oldPrint);
            for(int variant=0;variant<4;variant++) {
                sw.SetUserPreferenceToggle(323,true);
                sw.SetUserPreferenceToggle(327,false);
                foreach(string name in (string[])drawing.GetSheetNames()) {
                    drawing.ActivateSheet(name);
                    var view=(View)((View)drawing.GetFirstView()).GetNextView();
                    while(view!=null) {
                        var dim=(DisplayDimension)view.GetFirstDisplayDimension5();
                        while(dim!=null) {
                            if(variant==1) {
                                dim.ExtensionLineUseDocumentDisplay=false;
                                Console.WriteLine("Explicit normal: "+dim.SetLineFontDimensionThickness(false,1)+" / "+dim.SetLineFontExtensionThickness(false,1));
                            } else if(variant==2) {
                                dim.ExtensionLineUseDocumentDisplay=true;
                                Console.WriteLine("Document custom: "+dim.SetLineFontDimensionThickness(true,10)+" / "+dim.SetLineFontExtensionThickness(true,10));
                            } else if(variant==3) {
                                dim.ExtensionLineUseDocumentDisplay=false;
                                Console.WriteLine("Explicit thick: "+dim.SetLineFontDimensionThickness(false,2)+" / "+dim.SetLineFontExtensionThickness(false,2));
                            }
                            dim=(DisplayDimension)dim.GetNext5();
                        }
                        view=(View)view.GetNextView();
                    }
                }
                doc.ForceRebuild3(false);
                doc.GraphicsRedraw2();
                var pdf=(ExportPdfData)sw.GetExportFileData(1);
                pdf.SetSheets(1,null);
                pdf.ViewPdfAfterSaving=false;
                string target=Path.Combine(output,"variant-"+variant+".pdf");
                if(!ext.SaveAs(target,0,SaveSilent,pdf,ref error,ref warning)) throw new Exception("PDF export failed: "+error);
                Console.WriteLine(target);
            }
        } finally {
            sw.SetUserPreferenceToggle(327,oldPrint);
            sw.SetUserPreferenceToggle(323,oldColor);
            sw.CloseAllDocuments(true);
            sw.ExitApp();
        }
    }
}
'@
Add-Type -TypeDefinition $source -ReferencedAssemblies $api
[DimensionAppearanceProbe]::Run($drawing,$output)

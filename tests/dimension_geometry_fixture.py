"""Run the actual VBA layout rules against externally captured native geometry."""
import json
import os
from pathlib import Path
import subprocess


def check_layout(test, fixture, root):
    geo=json.loads(fixture.read_text('utf8'))['geometry'][-1]
    points=sorted(set(tuple(round(p[a]/geo['scale'],8) for a in (0,1)) for p in geo['vertices']))
    minimum=[min(p[a] for p in points) for a in (0,1)]
    points=[tuple(p[a]-minimum[a] for a in (0,1)) for p in points]
    body=r'''
    Sub AddPoint(ByVal coordinates, ByVal value, ByVal perpendicular, ByVal factor)
        Dim key, record, positions
        key=CStr(Round(value/factor,5))
        If coordinates.Exists(key) Then
            Set record=coordinates(key)
            positions=record("candidatePositions")
            ReDim Preserve positions(UBound(positions)+1)
            positions(UBound(positions))=perpendicular
            record("candidatePositions")=positions
        Else
            Set record=CreateObject("Scripting.Dictionary")
            record.Add "coordinate",value
            record.Add "candidatePositions",Array(perpendicular)
            coordinates.Add key,record
        End If
    End Sub
    Function GeometryPlans(ByVal coordinates, ByVal factor, ByVal lowEdge, ByVal highEdge)
        Dim pairs,pair,plan,result,first,last,pa,pb,best,a,b,span,text,keys,i,j,temp,sorted
        Set result=CreateObject("Scripting.Dictionary")
        Set pairs=LocalDatumPairs(coordinates)
        For Each pair In pairs.Items
            Set first=pair("firstRecord")
            Set last=pair("lastRecord")
            best=1E+12
            For Each pa In first("candidatePositions")
                For Each pb In last("candidatePositions")
                    If Abs(pa-pb)<best Then
                        best=Abs(pa-pb)
                        a=pa:b=pb
                    End If
                Next
            Next
            span=last("coordinate")-first("coordinate")
            text=FormatNumber(span/factor*1000,2,-1,0,0)
            Do While Right(text,1)="0"
                text=Left(text,Len(text)-1)
            Loop
            If Right(text,1)="," Or Right(text,1)="." Then
                text=Left(text,Len(text)-1)
            End If
            Set plan=CreateObject("Scripting.Dictionary")
            plan.Add "firstCoordinate",first("coordinate")
            plan.Add "lastCoordinate",last("coordinate")
            plan.Add "span",span
            plan.Add "width",Len(text)*COORDINATE_GLYPH_WIDTH_M+COORDINATE_LABEL_PADDING_M
            plan.Add "geometricCenter",(first("coordinate")+last("coordinate"))/2
            plan.Add "desired",plan("geometricCenter")
            plan.Add "lane",-1
            plan.Add "side",0
            plan.Add "lowCost",a+b-lowEdge*2
            plan.Add "highCost",highEdge*2-a-b
            plan.Add "lineStart",first("coordinate")-CoordinateMarkerPadding(span,plan("width"))
            plan.Add "lineEnd",last("coordinate")+CoordinateMarkerPadding(span,plan("width"))
            result.Add result.Count,plan
        Next
        keys=result.Keys
        For i=1 To UBound(keys)
            temp=keys(i)
            j=i-1
            Do While j>=0
                If result(keys(j))("lineStart")<=result(temp)("lineStart") Then
                    Exit Do
                End If
                keys(j+1)=keys(j)
                j=j-1
            Loop
            keys(j+1)=temp
        Next
        Set sorted=CreateObject("Scripting.Dictionary")
        For Each i In keys
            sorted.Add sorted.Count,result(i)
        Next
        Set GeometryPlans=sorted
    End Function
    Sub GeometryMain()
        Dim factor,xs,ys,points,p,originX,originY,xmin,xmax,ymin,ymax,plans,low,high,number,desc,plan,rmin,rmax,lowUsed,highUsed,placed,reported
        points=Array(POINTS)
        For Each factor In Array(.1,.05,.04,.025,.02,.01)
            Set xs=CreateObject("Scripting.Dictionary")
            Set ys=CreateObject("Scripting.Dictionary")
            xmin=.1125-1.3*factor/2-.0042:xmax=.1125+1.3*factor/2+.0042
            ymin=.167-.348*factor/2-.0042:ymax=.167+.348*factor/2+.0042
            originX=xmin+.0042:originY=ymin+.0042
            For Each p In points
                AddPoint xs,originX+p(0)*factor,originY+p(1)*factor,factor
                AddPoint ys,originY+p(1)*factor,originX+p(0)*factor,factor
            Next
            Set plans=GeometryPlans(ys,factor,xmin,xmax)
            low=Int((xmin-.012-.026)/.007)+1
            high=Int((.198-xmax-.012)/.007)+1
            If low<0 Then
                low=0
            End If
            If high<0 Then
                high=0
            End If
            WScript.Echo "Scale=" & factor & " y bands=" & low & "+" & high
            On Error Resume Next
            BalanceCoordinateSides plans,.082,.252,low,high
            number=Err.Number:desc=Err.Description:Err.Clear
            On Error GoTo 0
            If number=0 Then
                lowUsed=0:highUsed=0
                For Each plan In plans.Items
                    If plan("side")=0 Then
                        If plan("lane")+1>lowUsed Then
                            lowUsed=plan("lane")+1
                        End If
                    Else
                        If plan("lane")+1>highUsed Then
                            highUsed=plan("lane")+1
                        End If
                    End If
                Next
                rmin=.026:rmax=.198
                If lowUsed>0 Then
                    rmin=.026+(lowUsed-1)*.007+.0055
                End If
                If highUsed>0 Then
                    rmax=.198-(highUsed-1)*.007-.0055
                End If
                WScript.Echo "Y fit rows=" & lowUsed & "+" & highUsed & "; x band=" & rmin & ".." & rmax
                Set plans=GeometryPlans(xs,factor,ymin,ymax)
                low=Int((ymin-.012-.082)/.007)+1
                high=Int((.252-ymax-.012)/.007)+1
                On Error Resume Next
                BalanceCoordinateSides plans,rmin,rmax,low,high
                number=Err.Number:desc=Err.Description:Err.Clear
                On Error GoTo 0
            End If
            WScript.Echo "Result=" & number & " " & desc
            If number<>0 Then
                placed=0:reported=False
                For Each plan In plans.Items
                    If plan("lane")>=0 Then
                        placed=placed+1
                    ElseIf Not reported Then
                        WScript.Echo "First remaining interval " & (plan("firstCoordinate")-originY)/factor*1000 & ".." & (plan("lastCoordinate")-originY)/factor*1000 & "; width=" & plan("width")*1000
                        reported=True
                    End If
                Next
                WScript.Echo "Placed=" & placed & "/" & plans.Count
            End If
        Next
    End Sub
    GeometryMain
    '''
    data=','.join('Array('+','.join(f'{v:.10f}' for v in p)+')' for p in points)
    body=body.replace('POINTS',data)
    script=root/'simulate-solver.vbs'
    script.write_text('Option Explicit\n'+'\n'.join(test['constants']+test['routines'])+'\n'+body,'ascii')
    run=subprocess.run([str(Path(os.environ['WINDIR'])/'System32/cscript.exe'),'//nologo',str(script)],capture_output=True)
    output = run.stdout.decode('ascii', errors='replace')
    print(output)
    assert run.returncode == 0 and not run.stderr, (run.returncode, output, run.stderr)
    report = {'fits_without_detail_views': 'Result=0' in output, 'diagnostics': output}
    (root / 'geometry-capacity.json').write_text(json.dumps(report, indent=2), 'utf8')
    if not report['fits_without_detail_views']:
        print('KNOWN LIMITATION: this captured geometry needs a different layout; no GOST compliance is asserted.')

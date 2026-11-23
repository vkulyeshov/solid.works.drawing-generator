"""Exercise production wall-clock guards and watchdog recovery without CAD."""
from pathlib import Path
import os
import re
import subprocess
import sys
import tempfile
import unittest

sys.dont_write_bytecode = True
BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE / 'tools'))
from project_paths import work_path
from vba_sources import read_source

HOST = str(Path(os.environ['WINDIR']) / 'System32/cscript.exe')


class ExecutionBudgetTests(unittest.TestCase):
    def test_native_clock_guards_and_timeout_markers(self):
        source = read_source()
        module = (BASE / 'modules/execution_budget.bas').read_text('utf8')
        module = re.sub(r'^Attribute .*\n|^Option Explicit\n', '', module, flags=re.M)
        module = re.sub(r' As (?:String|Double|Long|Boolean|Date|Object|Variant)\b', '', module)
        module = re.sub(r'(?m)^Private (Sub|Function)', r'\1', module)
        module = re.sub(r'(?m)^Private ', 'Dim ', module)
        module = re.sub(r'(?m)^(\s*)Next \w+\s*$', r'\1Next', module)
        module = module.replace('Left$(', 'Left(')
        constants = ['Const ' + name + ' = ' + value for name, value in re.findall(
            r'^Private Const (\w+) As \w+ = (.+)$', source, re.M)
            if name in set(re.findall(r'\b[A-Z][A-Z_]+\b', module))]
        harness = '''
Dim settings, fs, runDir, jobs, interruptedJobNumber
Set settings=CreateObject("Scripting.Dictionary")
Set fs=CreateObject("Scripting.FileSystemObject")
runDir=WScript.Arguments(0)
Function Setting(ByVal key, ByVal fallback)
    Setting=fallback
    If settings.Exists(key) Then
        Setting=settings(key)
    End If
End Function
Sub Check(ByVal condition, ByVal message)
    If Not condition Then
        Err.Raise 1001,,message
    End If
End Sub
ConfigureExecutionBudgets
Check jobTimeoutSeconds=300 And layoutTimeoutSeconds=60,"Unexpected default budgets"
Dim invalid, number, index, job, stream
For Each invalid In Array("-1","1.5","86401","invalid")
    settings("JobTimeoutSeconds")=invalid
    On Error Resume Next
    ConfigureExecutionBudgets
    number=Err.Number: Err.Clear
    On Error GoTo 0
    Check number<>0,"Invalid timeout accepted"
Next
settings("JobTimeoutSeconds")="300"
ConfigureExecutionBudgets
BeginJobBudget "fixture"
BeginStrictSearch
executionStrictStarted=DateAdd("s",-MEMORY_STRICT_SEARCH_SECONDS,Now)
For index=1 To EXECUTION_POLL_INTERVAL
    PollExecutionBudget
Next
Check MemoryStrictSearchExpired(),"The strict phase exceeded its separate clock budget"
EndStrictSearch
Check Not MemoryStrictSearchExpired(),"The strict deadline leaked into the permitted-crossing search"
BeginLayoutBudget "sheet"
executionLayoutStarted=DateAdd("s",-60,Now)
For index=1 To EXECUTION_POLL_INTERVAL-1
    PollExecutionBudget
Next
On Error Resume Next
PollExecutionBudget
number=Err.Number: Err.Clear
On Error GoTo 0
Check number=vbObjectError+ERR_EXECUTION_TIMEOUT,"An inner-loop poll did not stop the expired sheet"
Check InStr(executionTimeoutMessage,"sheet layout")>0,"The timeout lost its scope"
On Error Resume Next
EndLayoutBudget
number=Err.Number: Err.Clear
On Error GoTo 0
Check number<>0,"A caught timeout disappeared during layout cleanup"
EndJobBudget
settings("JobTimeoutSeconds")="0": settings("LayoutTimeoutSeconds")="0"
ConfigureExecutionBudgets
BeginJobBudget "disabled"
BeginLayoutBudget "disabled"
executionJobStarted=DateAdd("d",-1,Now)
executionLayoutStarted=executionJobStarted
CheckExecutionBudget
EndJobBudget
settings("JobTimeoutSeconds")="300"
ConfigureExecutionBudgets
BeginJobBudget "midnight"
executionJobStarted=DateAdd("s",-300,Now)
On Error Resume Next
CheckExecutionBudget
number=Err.Number: Err.Clear
On Error GoTo 0
Check number=vbObjectError+ERR_EXECUTION_TIMEOUT,"The job wall-clock deadline was not enforced"
EndJobBudget
Set jobs=CreateObject("Scripting.Dictionary")
For index=1 To 2
    Set job=CreateObject("Scripting.Dictionary")
    job.Add "number",index: job.Add "out","fixture" & index
    job.Add "status",STATE_FAILED: job.Add "stage","Interrupted"
    If index=2 Then
        job("status")=STATE_DONE
    End If
    jobs.Add CStr(index),job
Next
interruptedJobNumber=1
Set stream=fs.CreateTextFile(fs.BuildPath(runDir,TIMEOUT_JOBS_FILENAME),True,True)
stream.WriteLine "fixture1" & vbTab & EXECUTION_TIMEOUT_PREFIX & "job"
stream.WriteLine "fixture2" & vbTab & EXECUTION_TIMEOUT_PREFIX & "job"
stream.Close
ApplyJobTimeoutMarkers
Check JobTimedOut(jobs("1")),"A killed job was not marked on checkpoint recovery"
Check interruptedJobNumber=0,"A timed-out job would run again immediately"
Check jobs("2")("status")=STATE_DONE And Not JobTimedOut(jobs("2")),"Recovery modified a completed drawing"
settings("RetryTimedOutJobs")="1"
Check Not JobTimedOut(jobs("1")),"An explicit timeout retry was ignored"
WScript.Echo "PASS: settings, job/sheet clocks, inner-loop polling, sticky errors, disabled limits and durable recovery"
'''
        root = work_path('execution_budget/unit')
        root.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=root) as temporary:
            script = Path(temporary) / 'budget.vbs'
            script.write_text('Option Explicit\n' + '\n'.join(constants) + '\n' + module + harness, 'utf16')
            result = subprocess.run([HOST, '//nologo', str(script), temporary], capture_output=True, timeout=15)
            self.assertEqual(result.returncode, 0, (result.stdout, result.stderr))
            self.assertIn(b'PASS:', result.stdout)

    def test_controller_stops_only_owned_runtime_and_continues(self):
        source = (BASE / 'launcher.vbs').read_text('utf8')
        worker = '''Sub WorkerMain()
    Dim context, count, stream, row
    Set context=ReadKeyValues(fso.BuildPath(sessionDir,CONTEXT_FILENAME))
    outputDir=context("Output")
    progressFile=fso.BuildPath(sessionDir,PROGRESS_FILENAME)
    count=0
    If fso.FileExists(fso.BuildPath(outputDir,"calls.txt")) Then
        count=CLng(ReadUnicode(fso.BuildPath(outputDir,"calls.txt")))
    End If
    count=count+1
    WriteText fso.BuildPath(outputDir,"calls.txt"),CStr(count)
    If count=1 Then
        row=Array("JOB","1","unused","Default","body","","1","False","False","1","RUNNING","001_fixture","Opening model","0","0","1")
        WriteText fso.BuildPath(outputDir,CHECKPOINT_FILENAME),"STATE" & vbTab & "3" & vbCrLf & Join(row,vbTab) & vbCrLf
        WriteText progressFile,Time & vbTab & STATE_RUNNING & vbTab & "0" & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & outputDir & vbTab & "JOB 001_fixture" & vbCrLf
        WScript.Sleep 20000
    Else
        If Not fso.FileExists(fso.BuildPath(outputDir,TIMEOUT_JOBS_FILENAME)) Then
            Err.Raise 1001,,"Missing timeout marker"
        End If
        WriteText progressFile,Time & vbTab & STATE_COMPLETE & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & outputDir & vbTab & "Injected completion" & vbCrLf
        WriteWorkerResult EXIT_SUCCESS,"Injected completion"
    End If
End Sub'''
        source = re.sub(r'(?ms)^Sub WorkerMain\(\).*?^End Sub', lambda _: worker, source)
        source = re.sub(r'(?ms)^Function TerminateTimedOutBackground\(\).*?^End Function',
                        'Function TerminateTimedOutBackground()\n    TerminateTimedOutBackground=True\nEnd Function', source)
        source = source.replace('Const SHUTDOWN_GRACE_SECONDS = 10', 'Const SHUTDOWN_GRACE_SECONDS = 0')
        root = work_path('execution_budget/controller')
        root.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=root) as temporary:
            folder = Path(temporary)
            (folder / 'launcher.vbs').write_text(source, 'utf8')
            (folder / 'generator.swp').write_bytes(b'Unused test macro')
            model = folder / 'input.sldprt'
            model.write_bytes(b'Unused test model')
            ini = folder / 'generator.ini'
            ini.write_text('JobTimeoutSeconds=1\nDrawingTemplate=' + str(BASE / 'templates/A4.portrait.drwdot'), 'utf16')
            output = folder / 'output'
            result = subprocess.run([HOST, '//nologo', str(folder / 'launcher.vbs'), '/input:' + str(model),
                                     '/out:' + str(output)], capture_output=True, timeout=20)
            self.assertEqual(result.returncode, 0, (result.stdout, result.stderr))
            self.assertEqual((output / 'calls.txt').read_text('utf16'), '2')
            self.assertIn('TIMEOUT job', (output / 'timed-out-jobs.tsv').read_text('utf16'))


if __name__ == '__main__':
    unittest.main(verbosity=2)

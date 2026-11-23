"""Prepare an instrumented external fixture or summarize its native timings.

The production VBA modules are never edited. The generated probe measures
layout phases and overlapping logging cost in a disposable native macro.
"""
import argparse
from collections import defaultdict
from datetime import datetime
import json
from pathlib import Path
import re
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import external_path
from rebuild_native import rebuild
from vba_sources import read_source, source_digest
from prepare_model_test import prepare


PROBE_DECLARATIONS = '''
Private Const BENCHMARK_LAYOUT_FILE As String = "benchmark-layout.tsv"
Private Const BENCHMARK_MACRO_FILE As String = "benchmark-macro.tsv"
Private benchmarkLoggingSeconds As Double
Private benchmarkMacroStarted As Double
'''

PROBE_ROUTINES = '''
Private Function BenchmarkElapsed(ByVal started As Double) As Double
    BenchmarkElapsed = CDbl(Timer) - started
    If BenchmarkElapsed < 0 Then
        BenchmarkElapsed = BenchmarkElapsed + SECONDS_PER_DAY
    End If
End Function

Private Sub BenchmarkMark(ByVal phase As String)
    Static previous As Double, samples As Collection, callNumber As Long
    Dim elapsed As Double, stream As Object, row As Variant
    If phase = "start" Then
        Set samples = New Collection
        callNumber = callNumber + 1
        previous = CDbl(Timer)
        Exit Sub
    End If
    elapsed = BenchmarkElapsed(previous)
    samples.Add Array(phase, elapsed)
    previous = CDbl(Timer)
    If phase = "finish" Then
        Set stream = fs.OpenTextFile(fs.BuildPath(runDir, BENCHMARK_LAYOUT_FILE), FILE_FOR_APPENDING, True, FILE_UNICODE)
        For Each row In samples
            stream.WriteLine callNumber & vbTab & row(0) & vbTab & CStr(row(1))
        Next row
        stream.Close
    End If
End Sub

Private Sub BenchmarkComplete()
    Dim stream As Object
    Set stream = fs.OpenTextFile(fs.BuildPath(runDir, BENCHMARK_MACRO_FILE), FILE_FOR_APPENDING, True, FILE_UNICODE)
    stream.WriteLine "macro-body" & vbTab & CStr(BenchmarkElapsed(benchmarkMacroStarted))
    stream.WriteLine "logging-overlap" & vbTab & CStr(benchmarkLoggingSeconds)
    stream.Close
End Sub
'''


def replace_once(source, old, new):
    if source.count(old) != 1:
        raise ValueError(f'Benchmark insertion must be unambiguous: {old!r}')
    return source.replace(old, new, 1)


def instrument(source):
    source = replace_once(source, 'Option Explicit\n', 'Option Explicit\n' + PROBE_DECLARATIONS)
    source = replace_once(source, '    On Error GoTo Fatal\n',
                          '    On Error GoTo Fatal\n    benchmarkMacroStarted = CDbl(Timer)\n')
    source = replace_once(source, '    Debug.Print "Drawing pack: "',
                          '    BenchmarkComplete\n    Debug.Print "Drawing pack: "')
    source = replace_once(source, 'Private Sub Log(ByVal text As String)\n',
                          'Private Sub Log(ByVal text As String)\n    Dim benchmarkLogStarted As Double\n    benchmarkLogStarted = CDbl(Timer)\n')
    source = replace_once(source, '    Debug.Print text\nEnd Sub',
                          '    Debug.Print text\n    benchmarkLoggingSeconds = benchmarkLoggingSeconds + BenchmarkElapsed(benchmarkLogStarted)\nEnd Sub')
    match = re.search(r'(?ms)^Private Sub FinishMemoryDimensionLayout\(.*?^End Sub', source)
    body = match[0]
    replacements = [
        ('    started = Timer\n', '    started = Timer\n    BenchmarkMark "start"\n'),
        ('    sheetNames = AsDrawing(d).GetSheetNames\n',
         '    BenchmarkMark "ownership and callouts"\n    sheetNames = AsDrawing(d).GetSheetNames\n'),
        ('        For pass = 0 To MEMORY_LAYOUT_REPAIR_LIMIT\n',
         '        BenchmarkMark "native snapshot and workspace"\n        For pass = 0 To MEMORY_LAYOUT_REPAIR_LIMIT\n'),
        ('            planarStrictCrossings = True\n',
         '            BenchmarkMark "candidate generation"\n            planarStrictCrossings = True\n'),
        ('            For Each candidate In chosen\n',
         '            BenchmarkMark "solve and improve"\n            For Each candidate In chosen\n'),
        ('            Set actual = New Collection\n',
         '            BenchmarkMark "apply and redraw"\n            Set actual = New Collection\n'),
        ('            Log "MEMORY LAYOUT sheet="',
         '            BenchmarkMark "native verification"\n            Log "MEMORY LAYOUT sheet="'),
        ('NextMemorySheet:\n', '        BenchmarkMark "markers and audit"\nNextMemorySheet:\n'),
        ('    Log "MEMORY COMPLETE dimensions="',
         '    BenchmarkMark "finish"\n    Log "MEMORY COMPLETE dimensions="'),
    ]
    for old, new in replacements:
        body = replace_once(body, old, new)
    return source[:match.start()] + body + source[match.end():] + '\n' + PROBE_ROUTINES


def prepare_benchmark(model, directory):
    directory = external_path(directory)
    prepare(model, directory, overrides=['DimensionLayoutAudit=0'])
    source = directory / 'benchmark-generator.bas'
    source.write_text(instrument(read_source()), encoding='utf8')
    rebuild(source, directory / 'generator.swp')
    (directory / 'benchmark-source.json').write_text(json.dumps({
        'canonical_source_sha256': source_digest(), 'model': str(model.resolve()),
        'audit_enabled': False,
    }, indent=2) + '\n', encoding='ascii')


def report(directory):
    directory = external_path(directory)
    def records(path):
        return [line.split('\t') for line in path.read_text('utf-16').splitlines() if line]
    def number(value):
        return float(value.replace(',', '.'))
    execution = {row[0]: number(row[1]) for row in records(directory / 'execution-timings.tsv')[1:]}
    required = {'test-total', 'test-inspection', 'test-shutdown', 'macro'}
    if required - execution.keys():
        raise ValueError('Wait for the complete native harness run before reporting timings')
    log = (directory / 'output/run.log').read_text('utf-16')
    assert re.findall(r'BATCH END drawings=(\d+) failed=(\d+)', log)[-1] == ('1', '0')
    events = [(datetime.strptime(line[:8], '%H:%M:%S'), line[9:]) for line in log.splitlines()]
    start = next(i for i, (_, message) in enumerate(events) if message.startswith('JOB '))
    finish = next(i for i in range(start, len(events)) if events[i][1].startswith('OK '))
    marks = [(events[start][0], 'Open snapshot model')]
    marks += [(time, message.split(' | ', 1)[1]) for time, message in events[start:finish]
              if message.startswith('STAGE ')]
    marks += [(events[finish][0], 'Complete')]
    stages = [{'stage': a[1], 'seconds': (b[0]-a[0]).total_seconds() % 86400}
              for a, b in zip(marks, marks[1:])]
    phases = defaultdict(float)
    for _, phase, elapsed in records(directory / 'output/benchmark-layout.tsv'):
        phases[phase] += number(elapsed)
    macro = {row[0]: number(row[1]) for row in records(directory / 'output/benchmark-macro.tsv')}
    generation = execution['test-total'] - execution['test-inspection'] - execution['test-shutdown']
    result = {'execution_seconds': execution, 'cold_generation_seconds': generation,
              'job_seconds': sum(row['seconds'] for row in stages),
              'job_stages': stages, 'layout_phases_seconds': dict(phases),
              'macro_probe_seconds': macro,
              'notes': ['One isolated-process run; not an average.',
                        'Test inspection and test shutdown are excluded from generation.',
                        'Cold generation starts at the native application startup request; harness compilation is excluded.',
                        'Job stages use log timestamps with one-second resolution.',
                        'Logging overlap is included in other measured phases.']}
    (directory / 'benchmark-report.json').write_text(json.dumps(result, indent=2) + '\n', 'ascii')
    lines = ['# Single-part generation benchmark', '',
             f'Cold generation: **{generation:.2f} s**.',
             f'Macro: **{execution["macro"]:.2f} s**. Drawing job: **{result["job_seconds"]:.0f} s**.',
             '', '| Job stage | Seconds |', '| --- | ---: |']
    lines += [f'| {item["stage"]} | {item["seconds"]:.0f} |' for item in stages]
    lines += ['', '| Layout phase | Seconds |', '| --- | ---: |']
    lines += [f'| {name} | {seconds:.3f} |' for name, seconds in phases.items()]
    lines += ['', f'Overlapping logging time: {macro["logging-overlap"]:.3f} s.', '']
    lines += ['- ' + note for note in result['notes']]
    (directory / 'benchmark-report.md').write_text('\n'.join(lines) + '\n', 'ascii')
    print(json.dumps(result, indent=2))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    preparation = commands.add_parser('prepare')
    preparation.add_argument('--model', type=Path, required=True)
    preparation.add_argument('--work-directory', type=Path, required=True)
    summary = commands.add_parser('report')
    summary.add_argument('--work-directory', type=Path, required=True)
    arguments = parser.parse_args()
    if arguments.command == 'prepare':
        prepare_benchmark(arguments.model, arguments.work_directory)
    else:
        report(arguments.work_directory)

"""Execute compact crop transforms and baseline-safe text clearance in VBA logic."""
from pathlib import Path
import re
import subprocess
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import work_path
from vba_sources import read_source

source = read_source()
functions = []
for name in ['MinD', 'ProfilePageUnion', 'CompactRotatePoint', 'CompactRotateBounds',
             'CompactExpandTextPolygon']:
    body = re.search(r'(?ms)^Private Function ' + name + r'\(.*?^End Function', source)[0]
    body = body.replace('Private ', '', 1)
    body = re.sub(r' As (?:Object|Collection|Variant|Long|Double|Boolean|String)\b', '', body)
    body = re.sub(r'(?m)^(\s*)Next \w+\s*$', r'\1Next', body)
    functions.append(body)
constants = ['Const ' + match[1] + ' = ' + re.sub(r'(?<=\d)#', '', match[2])
             for match in re.finditer(r'^Private Const (\w+) As \w+ = (.+)$', source, re.M)]
harness = '''
Sub Check(ByVal condition, ByVal message)
    If Not condition Then
        Err.Raise 513, , message
    End If
End Sub
Sub Near(ByVal actual, ByVal expected, ByVal message)
    Check Abs(actual - expected) < 0.000000001, message
End Sub
Dim bounds, point, rotated, result, polygon
bounds = Array(-0.66, -0.01, -0.64, 0.31)
rotated = CompactRotateBounds(bounds, -PI / 2)
Near rotated(0), -0.01, "The quarter-turn crop lost its first end"
Near rotated(2), 0.31, "The quarter-turn crop lost its second end"
Near rotated(1), 0.64, "The crop normal changed sign"
Near rotated(3), 0.66, "The crop depth changed"
For Each point In Array(Array(-0.65, 0), Array(-0.65, 0.3))
    result = CompactRotatePoint(point, -PI / 2)
    Check result(0) >= rotated(0) And result(0) <= rotated(2), "A rotated native reference left the crop"
    Check result(1) >= rotated(1) And result(1) <= rotated(3), "A rotated native reference left the crop"
Next
polygon = Array(Array(0, 0), Array(0.01, 0), Array(0.01, 0.0025), Array(0, 0.0025))
result = CompactExpandTextPolygon(polygon, 0)
Near result(0)(0), -COMPACT_TEXT_CLEARANCE_M, "Missing leading-glyph clearance"
Near result(1)(0), 0.01 + COMPACT_TEXT_CLEARANCE_M, "Missing final-glyph clearance"
Near result(0)(1), 0, "Text clearance crossed its own dimension baseline"
Near result(2)(1), 0.0025, "Text height changed"
polygon = Array(Array(0, 0), Array(0, 0.01), Array(-0.0025, 0.01), Array(-0.0025, 0))
result = CompactExpandTextPolygon(polygon, PI / 2)
Near result(0)(0), 0, "Vertical text clearance crossed its baseline"
Near result(1)(1), 0.01 + COMPACT_TEXT_CLEARANCE_M, "Vertical final glyph has no clearance"
WScript.Echo "PASS: rotated crop references and baseline-safe text clearance"
'''
directory = work_path('compact_profiles/unit')
directory.mkdir(parents=True, exist_ok=True)
script = directory / 'compact-profiles.vbs'
script.write_text('Option Explicit\n' + '\n'.join(constants + functions) + harness, encoding='utf-16')
result = subprocess.run(['cscript.exe', '//nologo', str(script)], capture_output=True)
output = result.stdout.decode('oem', errors='replace') + result.stderr.decode('oem', errors='replace')
print(output)
assert result.returncode == 0 and 'PASS:' in output, result.returncode

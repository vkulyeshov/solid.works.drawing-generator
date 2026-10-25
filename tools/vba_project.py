"""Empty SolidWorks 2025 VBA host metadata; contains no drawing-generator code.

The host requires project records, library references, and an empty ThisLibrary
document module. The executable Macro11 stream is empty until generator.bas is
inserted by rebuild_native.py. This is a project template, not a second source.
"""
import base64
import json
import zlib
from build_swp import Compound

_EMPTY_PROJECT = (
    "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"
    "xneBM8xlOtfJf50SzoJxObJkcb/9+edf1kHH5Q=="
)


def empty_project():
    """Create a blank host container entirely from tracked text metadata."""
    data = json.loads(zlib.decompress(base64.b64decode(_EMPTY_PROJECT)))
    project = Compound.__new__(Compound)
    project.raw = bytes.fromhex(data['header'])
    project.entries = [bytearray.fromhex(entry) for entry in data['entries']]
    project.streams = {int(i):bytes.fromhex(value) for i,value in data['streams'].items()}
    return project

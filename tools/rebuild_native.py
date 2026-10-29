"""Rebuild one native VBA container from ordered source modules or a test BAS."""
from pathlib import Path
import argparse
import hashlib
import json
import re
import struct
import sys
sys.dont_write_bytecode = True
from build_swp import Compound, decompress, compress
from project_paths import build_path, work_path
from vba_project import empty_project
from vba_sources import DEFAULT_SOURCE, read_source, source_digest, source_files

# Source identity and VBA project format constants.
NATIVE_ENCODING = 'cp1251'
PROJECT_CODE_PAGE = 1251
SOURCE_MODULE_ATTRIBUTE = 'Attribute VB_Name = "generator"'
NATIVE_MODULE_ATTRIBUTE = 'Attribute VB_Name = "Macro11"'
ENTRY_POINT_PATTERN = r'^Public Sub main\(\)'
DIRECTORY_STREAM = 'dir'
NATIVE_MODULE_STREAM = 'Macro11'
PROJECT_STREAM = '_VBA_PROJECT'
MODULE_OFFSET_RECORD = bytes.fromhex('310004000000')
CODE_PAGE_RECORD = bytes.fromhex('030002000000')
RECORD_PAYLOAD_OFFSET = 6
MODULE_OFFSET_FORMAT = '<I'
CODE_PAGE_FORMAT = '<H'
MODULE_SOURCE_OFFSET = 0
CACHE_VERSION_START = 2
CACHE_VERSION_END = 4
FORCE_RECOMPILE_VERSION = b'\xff\xff'
BUILD_REPORT_SUFFIX = '.build.json'
BUILD_REPORT_ENCODING = 'ascii'
DEFAULT_BUILD_PATH = 'build/generator.swp'


def module_source(source):
    """Read UTF-8 source and encode it for the VBA project's Windows-1251 code page."""
    text = read_source(source)
    if text.count(SOURCE_MODULE_ATTRIBUTE) != 1:
        raise ValueError('Expected one generator module attribute in the BAS source')
    if not re.search(ENTRY_POINT_PATTERN, text, re.MULTILINE):
        raise ValueError('The generator source must expose Public Sub main()')
    return text.replace(SOURCE_MODULE_ATTRIBUTE, NATIVE_MODULE_ATTRIBUTE).replace('\r\n','\n').replace('\n','\r\n').encode(NATIVE_ENCODING)


def rebuild(source, output):
    source = Path(source)
    output = build_path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    container = empty_project()
    directory = next(i for i in container.streams if container.name(i) == DIRECTORY_STREAM)
    data = bytearray(decompress(container.streams[directory]))
    module = next(i for i in container.streams if container.name(i) == NATIVE_MODULE_STREAM)
    struct.pack_into(MODULE_OFFSET_FORMAT, data, data.rfind(MODULE_OFFSET_RECORD)+RECORD_PAYLOAD_OFFSET, MODULE_SOURCE_OFFSET)
    struct.pack_into(CODE_PAGE_FORMAT, data, data.find(CODE_PAGE_RECORD)+RECORD_PAYLOAD_OFFSET, PROJECT_CODE_PAGE)
    container.streams[directory] = compress(data)
    encoded = module_source(source)
    container.streams[module] = compress(encoded)
    project = next(i for i in container.streams if container.name(i) == PROJECT_STREAM)
    cache = bytearray(container.streams[project])
    cache[CACHE_VERSION_START:CACHE_VERSION_END] = FORCE_RECOMPILE_VERSION
    container.streams[project] = bytes(cache)
    container.write(output)
    check = Compound(output)
    assert decompress(check.streams[module]) == encoded, 'Native source verification failed'
    info = {'source':str(source.resolve()), 'source_sha256':source_digest(source),
            'source_files':[{'path':str(path.resolve()),
                             'sha256':hashlib.sha256(path.read_bytes()).hexdigest()}
                            for path in source_files(source)],
            'macro_sha256':hashlib.sha256(output.read_bytes()).hexdigest()}
    output.with_suffix(BUILD_REPORT_SUFFIX).write_text(json.dumps(info,indent=2),encoding=BUILD_REPORT_ENCODING)
    print(f'Native container rebuilt: {output}')
    return output


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', type=Path, default=DEFAULT_SOURCE,
                        help='Source-module directory, or a standalone test BAS')
    parser.add_argument('--output', type=Path, default=work_path(DEFAULT_BUILD_PATH))
    args = parser.parse_args()
    rebuild(args.source, args.output)

"""Compose ordered VBA source modules into the generator's private host module.

The unity build preserves VBA Private visibility and the existing Macro11.main
entry point. No generated BAS file is written into the source checkout.
"""
from pathlib import Path
import hashlib
import json
import re

from project_paths import REPOSITORY

DEFAULT_SOURCE = REPOSITORY / 'modules'
MANIFEST_FILENAME = 'project.json'
SOURCE_ENCODING = 'utf-8-sig'
MODULE_HEADER = 'Attribute VB_Name = "generator"\nOption Explicit\n'
MODULE_NAME = re.compile(r'[A-Za-z][A-Za-z0-9_]*\Z')
PROCEDURE = re.compile(r'^(?:Private |Public )?(?:Sub|Function) (\w+)\(', re.MULTILINE)


def source_files(source=DEFAULT_SOURCE):
    """Return a validated manifest and its modules in explicit build order."""
    source = Path(source).resolve()
    if source.is_file():
        return [source]
    manifest = source / MANIFEST_FILENAME
    project = json.loads(manifest.read_text(SOURCE_ENCODING))
    if project.get('version') != 1:
        raise ValueError('Unsupported VBA source manifest version')
    names = project.get('modules')
    if not isinstance(names, list) or not names:
        raise ValueError('The VBA source manifest needs a non-empty modules list')
    paths = []
    for name in names:
        if not isinstance(name, str) or not name.endswith('.bas') or not MODULE_NAME.fullmatch(name[:-4]):
            raise ValueError(f'Invalid VBA module filename: {name!r}')
        path = source / name
        if path.resolve().parent != source or not path.is_file():
            raise ValueError(f'Missing or external VBA source module: {name}')
        if path.name.casefold() in {item.name.casefold() for item in paths}:
            raise ValueError(f'Duplicate VBA source module: {name}')
        paths.append(path)
    unlisted = {path.name for path in source.glob('*.bas')} - set(names)
    if unlisted:
        raise ValueError(f'Unlisted VBA source modules: {sorted(unlisted)}')
    return [manifest, *paths]


def read_source(source=DEFAULT_SOURCE):
    """Compose UTF-8 modules, or read a standalone BAS used by test harnesses."""
    paths = source_files(source)
    if len(paths) == 1:
        return paths[0].read_text(SOURCE_ENCODING)
    bodies = []
    procedures = set()
    for path in paths[1:]:
        text = path.read_text(SOURCE_ENCODING)
        header = f'Attribute VB_Name = "{path.stem}"\nOption Explicit\n'
        if not text.startswith(header):
            raise ValueError(f'Expected a matching module attribute and Option Explicit: {path.name}')
        body = text[len(header):]
        if re.search(r'^(?:Attribute |Option Explicit\b)', body, re.MULTILINE):
            raise ValueError(f'Unexpected module header inside {path.name}')
        if not body.endswith('\n'):
            body += '\n'
        for procedure in PROCEDURE.findall(body):
            key = procedure.casefold()
            if key in procedures:
                raise ValueError(f'Duplicate VBA procedure: {procedure}')
            procedures.add(key)
        bodies.append(body)
    text = MODULE_HEADER + ''.join(bodies)
    if len(re.findall(r'^Public Sub main\(\)', text, re.MULTILINE)) != 1:
        raise ValueError('The VBA modules must expose exactly one Public Sub main()')
    first_procedure = PROCEDURE.search(text)
    if first_procedure and re.search(
            r'^Private (?:Const |Type |Declare |\w+ As )',
            text[first_procedure.start():], re.MULTILINE):
        raise ValueError('VBA declarations must precede procedures in the module manifest')
    return text


def source_digest(source=DEFAULT_SOURCE):
    """Hash the composed UTF-8 source used by the native builder."""
    return hashlib.sha256(read_source(source).encode('utf-8')).hexdigest()

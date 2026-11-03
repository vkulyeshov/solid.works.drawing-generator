"""Check a native saved-reference inventory and optional unchanged snapshot bytes."""
from pathlib import Path
import argparse
import hashlib
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import external_path

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--directory', type=Path, required=True)
parser.add_argument('--source-models', type=Path)
parser.add_argument('--prefix', required=True)
parser.add_argument('--report', default='inventory.tsv')
parser.add_argument('--expect', action='append', default=[], metavar='MODEL=QUANTITY')
args = parser.parse_args()
directory = external_path(args.directory)
models = directory / 'Models'
events = (directory / 'progress.tsv').read_text('utf-16')
assert '\tCOMPLETE\t' in events, 'The latest native inventory did not complete'
rows = [line.split('\t') for line in (directory / args.report).read_text('utf-16').splitlines()]
assert rows[0] == ['Path', 'Configuration', 'Quantity']
assert len(rows) > 1
seen = set()
totals = {}
for row in rows[1:]:
    assert len(row) == 3, row
    path = Path(row[0]).resolve(strict=True)
    assert models.resolve() in path.parents, ('A source outside the copied snapshot entered inventory', row)
    assert row[1], ('An empty referenced configuration entered inventory', row)
    assert int(row[2]) > 0, row
    key = (str(path).casefold(), row[1].casefold())
    assert key not in seen, ('Duplicate model/configuration entry', row)
    seen.add(key)
    name = path.name
    assert name.casefold().startswith(args.prefix.casefold()), row
    name = name[len(args.prefix):].casefold()
    totals[name] = totals.get(name, 0) + int(row[2])
for expectation in args.expect:
    name, quantity = expectation.rsplit('=', 1)
    assert totals.get(name.casefold()) == int(quantity), (expectation, totals.get(name.casefold()))
hash_file = directory / 'inventory-input-hashes.tsv'
if hash_file.is_file():
    hashes = [line.split('\t') for line in hash_file.read_text('utf-16').splitlines()]
    for name, expected in hashes:
        assert hashlib.sha256((models / name).read_bytes()).hexdigest() == expected, ('Snapshot bytes changed during inventory', name)
    print('PASS: unchanged remapped snapshot files:', len(hashes))
elif args.source_models:
    source = args.source_models.resolve(strict=True)
    source_files = [path for path in source.iterdir() if path.is_file() and path.name.casefold().startswith(args.prefix.casefold())]
    assert source_files
    for original in source_files:
        copy = models / original.name
        assert copy.is_file(), ('Missing copied snapshot file', original.name)
        assert hashlib.sha256(copy.read_bytes()).digest() == hashlib.sha256(original.read_bytes()).digest(), ('Snapshot bytes changed during inventory', original.name)
    print('PASS: unchanged snapshot files:', len(source_files))
print('PASS: native model/configuration entries:', len(seen), '; occurrences:', sum(totals.values()))

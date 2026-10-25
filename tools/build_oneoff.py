"""Build a test macro without modifying the production VBA files."""
from pathlib import Path
import argparse
import sys
sys.dont_write_bytecode = True
from project_paths import work_path
from rebuild_native import rebuild

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('source', type=Path)
parser.add_argument('--output', type=Path, default=work_path('build/OneOff.swp'))
args = parser.parse_args()
rebuild(args.source, args.output)

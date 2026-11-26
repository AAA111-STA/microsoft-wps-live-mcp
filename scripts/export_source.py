"""Export only reviewed source directories; never walk local runtime or document data."""
from pathlib import Path
import zipfile

ROOT = Path(__file__).resolve().parent.parent
ROOT_FILES = ('SKILL.md', 'README.md', 'README_CN.md', 'pyproject.toml', '.gitignore', '.gitattributes')
SOURCE_DIRS = {'scripts': {'.py', '.ps1', '.sh'}, 'assets': {'.html', '.js', '.css', '.png', '.jpg', '.svg'},
               'tests': {'.py'}, 'docs': {'.md', '.json'}, 'examples': {'.json', '.md'},
               'skills/ppt-architecture-diagram': {'.md'},
               'skills/ppt-architecture-diagram/references': {'.md'},
               'src/microsoft_wps_live_mcp': {'.py'}, '.github/workflows': {'.yml'}}


def source_files(root):
    for name in ROOT_FILES:
        path = root / name
        if path.is_file() and not path.is_symlink():
            yield path
    for directory, extensions in SOURCE_DIRS.items():
        folder = root / directory
        if not folder.is_dir() or folder.is_symlink():
            continue
        # Source is intentionally flat: do not traverse hidden caches or junctions.
        for path in sorted(folder.iterdir()):
            if (path.is_file() and not path.is_symlink() and path.suffix in extensions
                    and path.resolve().parent == folder.resolve()):
                yield path


def export(root=ROOT):
    destination = root / 'dist' / 'microsoft-wps-live-mcp-source.zip'
    destination.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(destination, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        for path in source_files(root):
            archive.write(path, 'microsoft-wps-live-mcp/' + path.relative_to(root).as_posix())
    return destination


if __name__ == '__main__':
    print(export())

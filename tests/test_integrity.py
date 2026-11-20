"""Checks for callback authentication, staging boundaries, and original-file conflicts."""
import importlib.util
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch, MagicMock
import io
import zipfile
import hashlib


def module(name, filename):
    path = Path(__file__).resolve().parents[1] / 'scripts' / filename
    spec = importlib.util.spec_from_file_location(name, path)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


server = module('live_server', 'live_server.py')
client = module('office_live', 'office_live.py')
exporter = module('export_source', 'export_source.py')


class IntegrityTests(unittest.TestCase):
    @unittest.skipUnless(client.os.name == 'nt', 'Windows forwarding fallback')
    def test_fallback_only_binds_the_unavailable_service(self):
        for working, missing in ((18820, 18821), (18821, 18820)):
            with self.subTest(working=working), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                (root / 'logs').mkdir()
                with patch.object(client, 'CONFIG', root), patch.object(client, 'ROOT', root), \
                     patch.object(client, 'reachable_ports', side_effect=[{working}, {18820, 18821}]), \
                     patch.object(client.socket, 'create_connection', side_effect=OSError('missing')), \
                     patch.object(client.subprocess, 'Popen') as spawn:
                    client.ensure_windows_proxy()
                    command = spawn.call_args.args[0]
                    self.assertEqual(command[command.index('--ports') + 1:], [str(missing)])

    def test_both_service_ports_are_probed(self):
        with patch.object(client, 'urlopen') as opening:
            response = MagicMock()
            response.__enter__.return_value = io.BytesIO(b'{"ok":true}')
            opening.side_effect = [response, OSError('Docs unavailable')]
            self.assertEqual(client.reachable_ports(), {18820})
            self.assertIn(':18821/healthcheck', opening.call_args_list[1].args[0])

    def test_source_export_excludes_machine_data(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'SKILL.md').write_text('public skill')
            for name in ('config', 'data', 'wsl', 'scripts'):
                (root / name).mkdir()
            (root / 'config' / 'settings.json').write_text('secret')
            (root / 'data' / 'personal.docx').write_bytes(b'private')
            (root / 'wsl' / 'ext4.vhdx').write_bytes(b'runtime')
            (root / 'scripts' / 'office_live.py').write_text('public source')
            with zipfile.ZipFile(exporter.export(root)) as archive:
                self.assertEqual(set(archive.namelist()), {'microsoft-wps-live-mcp/SKILL.md',
                                                         'microsoft-wps-live-mcp/scripts/office_live.py'})

    def test_older_callback_cannot_replace_new_generation(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            server.write_json(root / 'revision.json', {'number': 2, 'generation': 5})
            with patch.object(server, 'verify', return_value={'status': 6, 'key': 'sample', 'userdata': '4'}), \
                 patch.object(server, 'load_session', return_value={'key': 'sample'}), \
                 patch.object(server, 'session_path', return_value=root), patch.object(server, 'build_opener') as fetch:
                self.assertEqual(server.store_callback_serial('sample', {}, 'signed'), {'error': 0})
                fetch.assert_not_called()
                self.assertEqual(server.read_json(root / 'revision.json')['generation'], 5)

    def test_identical_content_advances_disk_save_confirmation(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with patch.object(client, 'DATA', root / 'data'):
                folder = client.DATA / 'sessions' / 'sample'
                folder.mkdir(parents=True)
                original = root / 'document.docx'
                original.write_bytes(b'unchanged')
                staged = folder / 'revision-000001.docx'
                staged.write_bytes(original.read_bytes())
                digest = hashlib.sha256(b'unchanged').hexdigest()
                binding = {'source': str(original), 'sha256': digest, 'revision': 1, 'generation': 1}
                client.write(folder / 'revision.json', {'number': 1, 'generation': 2,
                             'file': staged.name, 'sha256': digest})
                client.save_revision('sample', binding)
                self.assertEqual(client.read(folder / 'save-status.json')['generation'], 2)
                self.assertEqual(original.read_bytes(), b'unchanged')

    def test_signed_callback_cannot_be_modified(self):
        server.JWT_SECRET = 'local-test-secret'
        token = server.sign({'status': 6, 'key': 'sample'})
        self.assertEqual(server.verify(token)['status'], 6)
        with self.assertRaises(ValueError):
            server.verify(token[:-1] + ('x' if token[-1] != 'x' else 'y'))

    def test_malformed_session_rejected(self):
        with self.assertRaises(ValueError):
            server.session_path('../../outside')

    def test_command_kind_must_match_file(self):
        with self.assertRaises(ValueError):
            server.validate_command('word', {'op': 'sheet.set', 'range': 'A1', 'values': 2})

    def test_invalid_range_rejected(self):
        with self.assertRaises(ValueError):
            server.validate_command('cell', {'op': 'sheet.read', 'range': '../../A1'})

    def test_external_edit_is_not_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            client.DATA = root / 'data'
            folder = client.DATA / 'sessions' / 'sample'
            folder.mkdir(parents=True)
            original = root / 'document.docx'
            original.write_bytes(b'original')
            binding = {'source': str(original), 'sha256': client.sha(original), 'revision': 0}
            original.write_bytes(b'external modification')
            client.write(folder / 'revision.json', {'number': 1, 'file': 'revision-000001.docx'})
            client.save_revision('sample', binding)
            self.assertEqual(original.read_bytes(), b'external modification')
            self.assertEqual(client.read(folder / 'save-status.json')['state'], 'conflict')

    def test_revision_path_cannot_escape_session(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            client.DATA = root / 'data'
            folder = client.DATA / 'sessions' / 'sample'
            folder.mkdir(parents=True)
            original = root / 'document.docx'
            original.write_bytes(b'original')
            binding = {'source': str(original), 'sha256': client.sha(original), 'revision': 0}
            client.write(folder / 'revision.json', {'number': 1, 'file': '../../outside.docx'})
            with self.assertRaises(ValueError):
                client.save_revision('sample', binding)
            self.assertEqual(original.read_bytes(), b'original')


if __name__ == '__main__':
    unittest.main()

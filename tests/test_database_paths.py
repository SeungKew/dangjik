"""data 경로·구경로 호환·임시DB 격리·초기화 검증."""
import contextlib
import hashlib
import io
from pathlib import Path
import sqlite3
import subprocess
import tempfile
import unittest
from unittest.mock import patch

import dangjik_basic1 as dy


class DatabasePathsTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.data = self.root / 'data'
        self.default = str(self.data / 'duty_system_v3.db')
        self.patcher = patch.multiple(dy, PROJECT_DIR=self.root, DATA_DIR=self.data,
                                     DEFAULT_DB_NAME=self.default, DB_NAME=self.default,
                                     V2_DB_NAME=str(self.data / 'duty_system_v2.db'))
        self.patcher.start()

    def tearDown(self) -> None:
        self.patcher.stop()
        self.temp.cleanup()

    def legacy(self, filename: str, label: str) -> Path:
        path = self.root / filename
        with sqlite3.connect(path) as conn:
            conn.execute('CREATE TABLE marker(value TEXT)')
            conn.execute('INSERT INTO marker VALUES (?)', (label,))
        return path

    def test_default_paths_are_absolute_from_another_working_directory(self) -> None:
        module_dir = Path(dy.__file__).resolve().parent
        code = 'import sys;sys.path.insert(0,sys.argv[1]);import dangjik_basic1 as d;print(d.DB_NAME);print(d.V2_DB_NAME)'
        result = subprocess.check_output([str(module_dir / '.venv/bin/python'), '-c', code, str(module_dir)],
                                         cwd=self.root, text=True).splitlines()
        self.assertEqual(result, [str(module_dir / 'data/duty_system_v3.db'),
                                  str(module_dir / 'data/duty_system_v2.db')])

    def test_legacy_v3_is_used_before_v2_and_never_overwrites_data(self) -> None:
        legacy = self.legacy('duty_system_v3.db', 'latest-v3')
        self.legacy('duty_system_v2.db', 'older-v2')
        fingerprint = hashlib.sha256(legacy.read_bytes()).hexdigest()
        dy.migrate_v2_to_v3()
        with dy.get_conn() as conn:
            self.assertEqual(conn.execute('SELECT value FROM marker').fetchone()[0], 'latest-v3')
            conn.execute("UPDATE marker SET value='new-data'")
        dy.migrate_v2_to_v3()
        with dy.get_conn() as conn:
            self.assertEqual(conn.execute('SELECT value FROM marker').fetchone()[0], 'new-data')
        self.assertEqual(hashlib.sha256(legacy.read_bytes()).hexdigest(), fingerprint)
        self.assertTrue(Path(dy.V2_DB_NAME).exists())

    def test_v2_fallback_and_reset_do_not_restore_old_data(self) -> None:
        self.legacy('duty_system_v2.db', 'v2')
        with contextlib.redirect_stdout(io.StringIO()):
            dy.migrate_v2_to_v3()
        dy.register_employee('00000001', '직원', '부서', dy.GROUP_SUK)
        # 구루트에 v3가 남아 있어도 초기화 후 복원되면 안 된다.
        self.legacy('duty_system_v3.db', 'legacy-v3')
        with contextlib.redirect_stdout(io.StringIO()):
            backup = dy.reset_database()
            dy.migrate_v2_to_v3()
        self.assertEqual(Path(backup).parent, self.data)
        with dy.get_conn() as conn:
            self.assertEqual(conn.execute('SELECT count(*) FROM employees').fetchone()[0], 0)
            self.assertIsNone(conn.execute("SELECT name FROM sqlite_master WHERE name='marker'").fetchone())

    def test_temporary_override_does_not_touch_production_or_legacy(self) -> None:
        legacy = self.legacy('duty_system_v3.db', 'production')
        fingerprint = hashlib.sha256(legacy.read_bytes()).hexdigest()
        with patch.object(dy, 'DB_NAME', str(self.root / 'isolated/test.db')):
            dy.init_db()
            self.assertEqual(dy.list_employees(), [])
        self.assertFalse(self.data.exists())
        self.assertEqual(hashlib.sha256(legacy.read_bytes()).hexdigest(), fingerprint)

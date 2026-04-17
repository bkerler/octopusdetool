import shutil
import unittest
from contextlib import contextmanager
from pathlib import Path
import sys
import uuid
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

TEST_TEMP_ROOT = PROJECT_ROOT / "tests_runtime"
TEST_TEMP_ROOT.mkdir(parents=True, exist_ok=True)

from octopusdetool.octopusdetool import (
    TARIFF_INTELLIGENT_GO,
    TARIFF_INTELLIGENT_HEAT,
    _get_bundled_excel_template_resource,
    detect_excel_template_type,
    ensure_excel_template,
)


class EnsureExcelTemplateWindowsTests(unittest.TestCase):
    """Regression coverage for Windows startup template creation."""

    def _make_scratch_dir(self) -> Path:
        scratch_dir = TEST_TEMP_ROOT / f"template-test-{uuid.uuid4().hex}"
        scratch_dir.mkdir(parents=True, exist_ok=True)
        return scratch_dir

    def test_ensure_excel_template_creates_go_template_in_data_dir(self) -> None:
        scratch_dir = self._make_scratch_dir()
        try:
            smartmeter_dir = scratch_dir / "Documents" / "smartmeter_data"
            smartmeter_dir.mkdir(parents=True, exist_ok=True)

            with patch(
                "octopusdetool.octopusdetool.ensure_smartmeter_data_folder",
                return_value=smartmeter_dir,
            ):
                template_path = ensure_excel_template(TARIFF_INTELLIGENT_GO)

            self.assertEqual(template_path, smartmeter_dir / "smartmeter_daten.xlsx")
            self.assertTrue(template_path.exists())
            self.assertEqual(
                detect_excel_template_type(template_path),
                TARIFF_INTELLIGENT_GO,
            )
        finally:
            shutil.rmtree(scratch_dir, ignore_errors=True)

    def test_ensure_excel_template_creates_heat_template_in_data_dir(self) -> None:
        scratch_dir = self._make_scratch_dir()
        try:
            smartmeter_dir = scratch_dir / "Documents" / "smartmeter_data"
            smartmeter_dir.mkdir(parents=True, exist_ok=True)

            with patch(
                "octopusdetool.octopusdetool.ensure_smartmeter_data_folder",
                return_value=smartmeter_dir,
            ):
                template_path = ensure_excel_template(TARIFF_INTELLIGENT_HEAT)

            self.assertEqual(
                template_path,
                smartmeter_dir / "smartmeter_heat_daten.xlsx",
            )
            self.assertTrue(template_path.exists())
            self.assertEqual(
                detect_excel_template_type(template_path),
                TARIFF_INTELLIGENT_HEAT,
            )
        finally:
            shutil.rmtree(scratch_dir, ignore_errors=True)

    def test_ensure_excel_template_can_build_heat_template_without_heat_resource(self) -> None:
        scratch_dir = self._make_scratch_dir()
        try:
            smartmeter_dir = scratch_dir / "Documents" / "smartmeter_data"
            smartmeter_dir.mkdir(parents=True, exist_ok=True)
            stock_resource = _get_bundled_excel_template_resource(TARIFF_INTELLIGENT_GO)

            def fake_resource_lookup(tariff_type: str):
                if tariff_type == TARIFF_INTELLIGENT_HEAT:
                    class MissingResource:
                        def is_file(self) -> bool:
                            return False

                    return MissingResource()
                return stock_resource

            with patch(
                "octopusdetool.octopusdetool.ensure_smartmeter_data_folder",
                return_value=smartmeter_dir,
            ), patch(
                "octopusdetool.octopusdetool._get_bundled_excel_template_resource",
                side_effect=fake_resource_lookup,
            ):
                template_path = ensure_excel_template(TARIFF_INTELLIGENT_HEAT)

            self.assertEqual(
                template_path,
                smartmeter_dir / "smartmeter_heat_daten.xlsx",
            )
            self.assertTrue(template_path.exists())
            self.assertEqual(
                detect_excel_template_type(template_path),
                TARIFF_INTELLIGENT_HEAT,
            )
        finally:
            shutil.rmtree(scratch_dir, ignore_errors=True)

    def test_ensure_excel_template_falls_back_to_bundled_go_template_when_resource_copy_fails(self) -> None:
        scratch_dir = self._make_scratch_dir()
        try:
            smartmeter_dir = scratch_dir / "Documents" / "smartmeter_data"
            smartmeter_dir.mkdir(parents=True, exist_ok=True)

            @contextmanager
            def missing_resource_file():
                yield smartmeter_dir / "missing-smartmeter_daten.xlsx"

            with patch(
                "octopusdetool.octopusdetool.ensure_smartmeter_data_folder",
                return_value=smartmeter_dir,
            ), patch(
                "octopusdetool.octopusdetool.package_resources.as_file",
                side_effect=lambda resource: missing_resource_file(),
            ):
                template_path = ensure_excel_template(TARIFF_INTELLIGENT_GO)

            self.assertEqual(
                template_path,
                PROJECT_ROOT / "octopusdetool" / "smartmeter_daten.xlsx",
            )
            self.assertTrue(template_path.exists())
            self.assertEqual(
                detect_excel_template_type(template_path),
                TARIFF_INTELLIGENT_GO,
            )
        finally:
            shutil.rmtree(scratch_dir, ignore_errors=True)

    def test_ensure_excel_template_falls_back_to_stock_heat_generation_when_resource_copy_fails(self) -> None:
        scratch_dir = self._make_scratch_dir()
        try:
            smartmeter_dir = scratch_dir / "Documents" / "smartmeter_data"
            smartmeter_dir.mkdir(parents=True, exist_ok=True)

            @contextmanager
            def missing_resource_file():
                yield smartmeter_dir / "missing-smartmeter_heat_daten.xlsx"

            with patch(
                "octopusdetool.octopusdetool.ensure_smartmeter_data_folder",
                return_value=smartmeter_dir,
            ), patch(
                "octopusdetool.octopusdetool.package_resources.as_file",
                side_effect=lambda resource: missing_resource_file(),
            ):
                template_path = ensure_excel_template(TARIFF_INTELLIGENT_HEAT)

            self.assertEqual(
                template_path,
                smartmeter_dir / "smartmeter_heat_daten.xlsx",
            )
            self.assertTrue(template_path.exists())
            self.assertEqual(
                detect_excel_template_type(template_path),
                TARIFF_INTELLIGENT_HEAT,
            )
        finally:
            shutil.rmtree(scratch_dir, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()

import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import cv2
import numpy as np
from google.api_core.exceptions import PermissionDenied
from googleapiclient.errors import HttpError

import main
from main import (
    Movement,
    create_drive_service,
    download_drive_image,
    drive_literal,
    load_config,
    load_google_credentials,
    locate_drive_image,
    parse_money,
    preprocess,
    save_outputs,
    valid_date,
    vision_request,
)


class CoreValidationTests(unittest.TestCase):
    def setUp(self):
        self.cfg = {"year": 2026, "allowed_months": [6, 7]}

    def test_dates(self):
        self.assertTrue(valid_date("26/06", self.cfg))
        self.assertTrue(valid_date("01/07", self.cfg))
        self.assertFalse(valid_date("2606", self.cfg))
        self.assertFalse(valid_date("24/87", self.cfg))
        self.assertFalse(valid_date("31/06", self.cfg))

    def test_money(self):
        self.assertEqual(parse_money("35,15"), 35.15)
        self.assertEqual(parse_money("1.234,56"), 1234.56)
        self.assertEqual(parse_money("0.04"), 0.04)
        self.assertIsNone(parse_money("8102"))
        self.assertIsNone(parse_money("$ 1,20 1,45"))

    def test_drive_literal(self):
        self.assertEqual(drive_literal("Clayton's\\Drive"), "Clayton\\'s\\\\Drive")


class ConfigurationTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.dir_path = Path(self.temp_dir.name)

    def test_load_config_valid(self):
        cfg_file = self.dir_path / "config.yaml"
        cfg_file.write_text("google:\n  project_id: test\n", encoding="utf-8")
        loaded = load_config(cfg_file)
        self.assertEqual(loaded.get("google", {}).get("project_id"), "test")

    def test_load_config_missing_file(self):
        cfg_file = self.dir_path / "non_existent.yaml"
        with self.assertRaises(FileNotFoundError):
            load_config(cfg_file)


class ServiceAccountAuthTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.dir_path = Path(self.temp_dir.name)

    @patch("main.service_account.Credentials.from_service_account_file")
    def test_load_credentials_valid(self, mock_from_file):
        mock_creds = MagicMock()
        mock_creds.service_account_email = "test@project.iam.gserviceaccount.com"
        mock_from_file.return_value = mock_creds

        sa_file = self.dir_path / "valid_sa.json"
        sa_file.write_text(
            json.dumps({"type": "service_account", "client_email": "test@project.iam.gserviceaccount.com"}),
            encoding="utf-8",
        )

        cfg = {"google": {"service_account_file": str(sa_file)}}
        creds = load_google_credentials(cfg, ["scope1"])
        self.assertEqual(creds, mock_creds)
        mock_from_file.assert_called_once_with(str(sa_file), scopes=["scope1"])

    @patch("main.service_account.Credentials.from_service_account_file")
    def test_load_credentials_env_override(self, mock_from_file):
        mock_creds = MagicMock()
        mock_from_file.return_value = mock_creds

        sa_file = self.dir_path / "env_sa.json"
        sa_file.write_text(
            json.dumps({"type": "service_account", "client_email": "env@project.iam.gserviceaccount.com"}),
            encoding="utf-8",
        )

        with patch.dict(os.environ, {"SOMA_GOOGLE_CREDENTIALS": str(sa_file)}):
            cfg = {"google": {"service_account_file": "other.json"}}
            creds = load_google_credentials(cfg, ["scope1"])
            self.assertEqual(creds, mock_creds)
            mock_from_file.assert_called_once_with(str(sa_file), scopes=["scope1"])

    def test_load_credentials_missing_file(self):
        cfg = {"google": {"service_account_file": str(self.dir_path / "non_existent.json")}}
        with self.assertRaises(FileNotFoundError) as ctx:
            load_google_credentials(cfg, ["scope1"])
        self.assertIn("não encontrado", str(ctx.exception))

    def test_load_credentials_invalid_json(self):
        sa_file = self.dir_path / "bad.json"
        sa_file.write_text("NOT_JSON", encoding="utf-8")

        cfg = {"google": {"service_account_file": str(sa_file)}}
        with self.assertRaises(ValueError) as ctx:
            load_google_credentials(cfg, ["scope1"])
        self.assertIn("inválido", str(ctx.exception))

    def test_load_credentials_not_service_account(self):
        sa_file = self.dir_path / "user.json"
        sa_file.write_text(json.dumps({"type": "authorized_user"}), encoding="utf-8")

        cfg = {"google": {"service_account_file": str(sa_file)}}
        with self.assertRaises(ValueError) as ctx:
            load_google_credentials(cfg, ["scope1"])
        self.assertIn("não é uma chave de conta de serviço", str(ctx.exception))


class DriveIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.drive_service = MagicMock()
        self.drive_cfg = {
            "folder_id": "1mwJmtDnPOMrlKkJuUL91iCkYxHI0Noll",
            "folder_name": "EXTRATO_CARTÃO_Images",
            "filename": "07-2026.IMAGEM.051625.jpg",
        }
        self.sa_email = "sa@project.iam.gserviceaccount.com"

    @patch("main.build")
    def test_create_drive_service_explicit_credentials(self, mock_build):
        creds = MagicMock()
        create_drive_service(creds)
        mock_build.assert_called_once_with("drive", "v3", credentials=creds, cache_discovery=False)

    def test_locate_drive_image_success(self):
        files_mock = self.drive_service.files.return_value
        list_mock = files_mock.list.return_value
        list_mock.execute.return_value = {
            "files": [
                {
                    "id": "file123",
                    "name": "07-2026.IMAGEM.051625.jpg",
                    "mimeType": "image/jpeg",
                    "modifiedTime": "2026-08-01T00:00:00Z",
                }
            ]
        }

        info = locate_drive_image(self.drive_service, self.drive_cfg, self.sa_email)
        self.assertEqual(info["id"], "file123")
        self.assertEqual(info["name"], "07-2026.IMAGEM.051625.jpg")

    def test_locate_drive_image_not_found(self):
        files_mock = self.drive_service.files.return_value
        list_mock = files_mock.list.return_value
        list_mock.execute.return_value = {"files": []}

        with self.assertRaises(FileNotFoundError) as ctx:
            locate_drive_image(self.drive_service, self.drive_cfg, self.sa_email)
        self.assertIn("não foi encontrada", str(ctx.exception))

    def test_locate_drive_image_duplicates(self):
        files_mock = self.drive_service.files.return_value
        list_mock = files_mock.list.return_value
        list_mock.execute.return_value = {
            "files": [
                {"id": "file1", "name": "07-2026.IMAGEM.051625.jpg", "mimeType": "image/jpeg", "modifiedTime": "2026-08-01"},
                {"id": "file2", "name": "07-2026.IMAGEM.051625.jpg", "mimeType": "image/jpeg", "modifiedTime": "2026-08-02"},
            ]
        }

        with self.assertRaises(RuntimeError) as ctx:
            locate_drive_image(self.drive_service, self.drive_cfg, self.sa_email)
        self.assertIn("duplicidade", str(ctx.exception))
        self.assertIn("file1", str(ctx.exception))
        self.assertIn("file2", str(ctx.exception))

    def test_locate_drive_image_native_google_file(self):
        files_mock = self.drive_service.files.return_value
        list_mock = files_mock.list.return_value
        list_mock.execute.return_value = {
            "files": [
                {
                    "id": "doc123",
                    "name": "07-2026.IMAGEM.051625.jpg",
                    "mimeType": "application/vnd.google-apps.document",
                }
            ]
        }

        with self.assertRaises(ValueError) as ctx:
            locate_drive_image(self.drive_service, self.drive_cfg, self.sa_email)
        self.assertIn("ficheiro nativo do Google", str(ctx.exception))

    def test_locate_drive_image_403_error(self):
        files_mock = self.drive_service.files.return_value
        list_mock = files_mock.list.return_value
        resp = MagicMock()
        resp.status = 403
        list_mock.execute.side_effect = HttpError(resp=resp, content=b"Forbidden")

        with self.assertRaises(RuntimeError) as ctx:
            locate_drive_image(self.drive_service, self.drive_cfg, self.sa_email)
        self.assertIn("A conta de serviço não tem permissão", str(ctx.exception))
        self.assertIn(self.sa_email, str(ctx.exception))
        self.assertIn("EXTRATO_CARTÃO_Images", str(ctx.exception))

    @patch("main.MediaIoBaseDownload")
    def test_download_drive_image_success(self, mock_downloader_cls):
        def fake_download(fh, request):
            fh.write(b"fake_image_bytes")
            inst = MagicMock()
            inst.next_chunk.return_value = (MagicMock(), True)
            return inst

        mock_downloader_cls.side_effect = fake_download

        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp_path = Path(tmp.name)

        self.addCleanup(lambda: tmp_path.unlink(missing_ok=True))
        download_drive_image(self.drive_service, "file123", tmp_path)
        self.assertTrue(tmp_path.exists())
        self.assertGreater(tmp_path.stat().st_size, 0)

    @patch("main.MediaIoBaseDownload")
    def test_download_drive_image_empty_file(self, mock_downloader_cls):
        def fake_download(fh, request):
            inst = MagicMock()
            inst.next_chunk.return_value = (MagicMock(), True)
            return inst

        mock_downloader_cls.side_effect = fake_download

        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp_path = Path(tmp.name)

        self.addCleanup(lambda: tmp_path.unlink(missing_ok=True))
        with self.assertRaises(RuntimeError) as ctx:
            download_drive_image(self.drive_service, "file123", tmp_path)
        self.assertIn("ficheiro vazio", str(ctx.exception))


class CloudVisionTests(unittest.TestCase):
    @patch("main.vision.ImageAnnotatorClient")
    def test_vision_request_explicit_credentials(self, mock_client_cls):
        mock_client = MagicMock()
        mock_client_cls.return_value = mock_client
        mock_response = MagicMock()
        mock_response.error.message = ""
        mock_client.document_text_detection.return_value = mock_response

        creds = MagicMock()
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp_path = Path(tmp.name)
            tmp.write(b"fake_img")

        self.addCleanup(lambda: tmp_path.unlink(missing_ok=True))

        cfg = {"ocr": {"language_hints": ["pt"]}}
        res = vision_request(tmp_path, cfg, creds)
        self.assertEqual(res, mock_response)
        mock_client_cls.assert_called_once_with(credentials=creds)

    @patch("main.vision.ImageAnnotatorClient")
    def test_vision_request_error_message(self, mock_client_cls):
        mock_client = MagicMock()
        mock_client_cls.return_value = mock_client
        mock_response = MagicMock()
        mock_response.error.message = "Quota exceeded"
        mock_client.document_text_detection.return_value = mock_response

        creds = MagicMock()
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp_path = Path(tmp.name)
            tmp.write(b"fake_img")

        self.addCleanup(lambda: tmp_path.unlink(missing_ok=True))

        cfg = {"ocr": {"language_hints": ["pt"]}}
        with self.assertRaises(RuntimeError) as ctx:
            vision_request(tmp_path, cfg, creds)
        self.assertIn("Quota exceeded", str(ctx.exception))

    @patch("main.vision.ImageAnnotatorClient")
    def test_vision_request_permission_denied(self, mock_client_cls):
        mock_client = MagicMock()
        mock_client_cls.return_value = mock_client
        mock_client.document_text_detection.side_effect = PermissionDenied("API disabled")

        creds = MagicMock()
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp_path = Path(tmp.name)
            tmp.write(b"fake_img")

        self.addCleanup(lambda: tmp_path.unlink(missing_ok=True))

        cfg = {"ocr": {"language_hints": ["pt"]}}
        with self.assertRaises(RuntimeError) as ctx:
            vision_request(tmp_path, cfg, creds)
        self.assertIn("Falta de permissão", str(ctx.exception))


class PipelineOutputsAndCLITests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.out_dir = Path(self.temp_dir.name)

    def test_preprocess_invalid_image(self):
        invalid_path = self.out_dir / "invalid.jpg"
        invalid_path.write_text("NOT_AN_IMAGE", encoding="utf-8")
        with self.assertRaises(ValueError) as ctx:
            preprocess(invalid_path, {}, self.out_dir)
        self.assertIn("Não foi possível abrir a imagem", str(ctx.exception))

    def test_save_outputs(self):
        movement = Movement(
            line=1,
            data_movimento="26/06",
            data_valor="26/06",
            descricao="TEST STORE",
            pais="PT",
            moeda_original="EUR",
            taxa_cambio="1.00",
            debito_eur="12,50",
            credito_eur="",
            confidence=0.95,
            status="VÁLIDO",
            motivos_revisao="",
            texto_ocr="26/06 26/06 TEST STORE 12,50",
        )
        metadata = {"imagem": "test.jpg", "linhas": 1}
        save_outputs([movement], self.out_dir, metadata)

        self.assertTrue((self.out_dir / "movimentos.csv").exists())
        self.assertTrue((self.out_dir / "resultado.json").exists())
        self.assertTrue((self.out_dir / "movimentos.xlsx").exists())

    @patch("main.sync_movements_to_sheets")
    @patch("main.create_sheets_service")
    @patch("main.annotate")
    @patch("main.save_outputs")
    @patch("main.build_movements")
    @patch("main.group_rows")
    @patch("main.extract_words")
    @patch("main.vision_request")
    @patch("main.preprocess")
    @patch("main.download_drive_image")
    @patch("main.locate_drive_image")
    @patch("main.create_drive_service")
    @patch("main.load_google_credentials")
    def test_main_drive_mode(
        self,
        mock_creds,
        mock_create_drive,
        mock_locate,
        mock_download,
        mock_preprocess,
        mock_vision,
        mock_extract,
        mock_group,
        mock_build_mov,
        mock_save,
        mock_annotate,
        mock_create_sheets,
        mock_sync_sheets,
    ):
        mock_creds.return_value = MagicMock(service_account_email="sa@test.com")
        mock_locate.return_value = {"id": "file123", "name": "test.jpg"}
        mock_sync_sheets.return_value = {"sheet_write": False}

        fake_ocr_img = self.out_dir / "01_contraste.png"
        cv2.imwrite(str(fake_ocr_img), np.zeros((10, 10, 3), dtype=np.uint8))
        mock_preprocess.return_value = [fake_ocr_img]

        mock_extract.return_value = []
        mock_group.return_value = []
        mock_build_mov.return_value = []

        with patch("sys.argv", ["main.py"]):
            ret = main.main()
            self.assertEqual(ret, 0)
            mock_locate.assert_called_once()
            mock_download.assert_called_once()

    @patch("main.sync_movements_to_sheets")
    @patch("main.create_sheets_service")
    @patch("main.annotate")
    @patch("main.save_outputs")
    @patch("main.build_movements")
    @patch("main.group_rows")
    @patch("main.extract_words")
    @patch("main.vision_request")
    @patch("main.preprocess")
    @patch("main.load_google_credentials")
    def test_main_local_image_mode(
        self,
        mock_creds,
        mock_preprocess,
        mock_vision,
        mock_extract,
        mock_group,
        mock_build_mov,
        mock_save,
        mock_annotate,
        mock_create_sheets,
        mock_sync_sheets,
    ):
        mock_creds.return_value = MagicMock(service_account_email="sa@test.com")
        mock_sync_sheets.return_value = {"sheet_write": False}
        local_img = self.out_dir / "local_test.jpg"
        cv2.imwrite(str(local_img), np.zeros((10, 10, 3), dtype=np.uint8))

        fake_ocr_img = self.out_dir / "01_contraste.png"
        cv2.imwrite(str(fake_ocr_img), np.zeros((10, 10, 3), dtype=np.uint8))
        mock_preprocess.return_value = [fake_ocr_img]

        mock_extract.return_value = []
        mock_group.return_value = []
        mock_build_mov.return_value = []

        with patch("sys.argv", ["main.py", str(local_img)]):
            ret = main.main()
            self.assertEqual(ret, 0)


class GoogleSheetsSyncTests(unittest.TestCase):
    def setUp(self):
        self.header = [
            "Data Mov.", "Data Valor", "Descrição", "País",
            "Valor Original", "Moeda Original", "Taxa de Câmbio",
            "Débito EUR (+)", "Crédito EUR (-)", "ID_INTERNO",
            "FINANCE", "STATUS", "MOTIVO_REVISÃO"
        ]
        self.cfg = {
            "google_sheets": {
                "enabled": True,
                "spreadsheet_id": "test_sheet_id",
                "worksheet": "CARTÃO",
                "write_mode": "append_idempotent",
                "dry_run": True,
                "id_prefix": "CAR",
                "id_digits": 10,
            },
            "validation": {"year": 2026},
        }

    def test_map_sheet_header_valid(self):
        mapping = main.map_sheet_header(self.header)
        self.assertEqual(mapping["data_movimento"], 0)
        self.assertEqual(mapping["data_valor"], 1)
        self.assertEqual(mapping["descricao"], 2)
        self.assertEqual(mapping["debito_eur"], 7)
        self.assertEqual(mapping["id_interno"], 9)
        self.assertEqual(mapping["motivos_revisao"], 12)

    def test_map_sheet_header_missing_column(self):
        incomplete_header = ["Data Mov.", "Data Valor", "Descrição"]
        with self.assertRaises(ValueError):
            main.map_sheet_header(incomplete_header)

    def test_generate_next_id_interno(self):
        existing = ["CAR0000000001", "CAR0000000005", "INVALID"]
        next_id = main.generate_next_id_interno(existing, prefix="CAR", digits=10)
        self.assertEqual(next_id, "CAR0000000006")

    def test_format_date_for_sheet(self):
        self.assertEqual(main.format_date_for_sheet("26/06", 2026), "26/06/2026")
        self.assertEqual(main.format_date_for_sheet("invalid_date", 2026), "invalid_date")

    def test_sync_dry_run_no_writes(self):
        mock_service = MagicMock()
        mock_get = mock_service.spreadsheets().values().get().execute
        mock_get.return_value = {"values": [self.header]}

        mov_valid = Movement(1, "26/06", "25/06", "MERCADONA", "", "", "", "35,15", "", 0.9, "VÁLIDO", "", "26/06 25/06 MERCADONA 35,15")
        mov_rev = Movement(2, "27/06", "27/06", "COMISSAD", "", "", "", "1,02", "", 0.6, "REVISÃO", "confiança OCR baixa", "27/06 27/06 COMISSAD 1,02")

        stats = main.sync_movements_to_sheets(mock_service, self.cfg, [mov_valid, mov_rev], None, "sa@test.com")

        self.assertFalse(stats["sheet_write"])
        self.assertTrue(stats["dry_run"])
        self.assertEqual(stats["linhas_novas"], 2)
        self.assertEqual(stats["validos_inseridos"], 0)
        # Garantir que append não foi chamado em dry_run
        mock_service.spreadsheets().values().append.assert_not_called()

    def test_sync_real_write(self):
        cfg_real = dict(self.cfg)
        cfg_real["google_sheets"] = dict(self.cfg["google_sheets"])
        cfg_real["google_sheets"]["dry_run"] = False

        mock_service = MagicMock()
        mock_get = mock_service.spreadsheets().values().get().execute
        mock_get.return_value = {"values": [self.header]}

        mov_valid = Movement(1, "26/06", "25/06", "MERCADONA", "", "", "", "35,15", "", 0.9, "VÁLIDO", "", "26/06 25/06 MERCADONA 35,15")
        mov_rev = Movement(2, "27/06", "27/06", "COMISSAD", "", "", "", "1,02", "", 0.6, "REVISÃO", "confiança OCR baixa", "27/06 27/06 COMISSAD 1,02")

        stats = main.sync_movements_to_sheets(mock_service, cfg_real, [mov_valid, mov_rev], None, "sa@test.com")

        self.assertTrue(stats["sheet_write"])
        self.assertFalse(stats["dry_run"])
        self.assertEqual(stats["linhas_inseridas"], 2)
        self.assertGreaterEqual(mock_service.spreadsheets().values().update.call_count, 1)

        # Verificar dados passados para update do CARTÃO
        update_call = mock_service.spreadsheets().values().update.call_args_list[0]
        written_body = update_call[1]["body"]["values"]
        self.assertEqual(len(written_body), 2)

        # Validação da linha VÁLIDA
        row0 = written_body[0]
        self.assertEqual(row0[0], "26/06/2026")  # Data Mov.
        self.assertEqual(row0[1], "25/06/2026")  # Data Valor
        self.assertEqual(row0[2], "MERCADONA")  # Descrição
        self.assertEqual(row0[7], 35.15)  # Débito EUR (+)
        self.assertEqual(row0[8], "")  # Crédito EUR (-)
        self.assertEqual(row0[9], "CAR0000000001")  # ID_INTERNO
        self.assertEqual(row0[10], "")  # FINANCE
        self.assertEqual(row0[11], "VÁLIDO")  # STATUS
        self.assertEqual(row0[12], "")  # MOTIVO_REVISÃO

        # Validação da linha REVISÃO
        row1 = written_body[1]
        self.assertEqual(row1[11], "REVISÃO")
        self.assertEqual(row1[12], "confiança OCR baixa")
        self.assertEqual(row1[9], "CAR0000000002")

    def test_idempotency_duplicate_prevention(self):
        mock_service = MagicMock()
        mock_get = mock_service.spreadsheets().values().get().execute
        # Simular que a sheet já possui a linha do MERCADONA
        existing_row = ["2026-06-26", "2026-06-25", "MERCADONA", "", "", "", "", "35.15", "", "CAR0000000001", "", "VÁLIDO", ""]
        mock_get.return_value = {"values": [self.header, existing_row]}

        mov_valid = Movement(1, "26/06", "25/06", "MERCADONA", "", "", "", "35,15", "", 0.9, "VÁLIDO", "", "26/06 25/06 MERCADONA 35,15")

        stats = main.sync_movements_to_sheets(mock_service, self.cfg, [mov_valid], None, "sa@test.com")

        self.assertEqual(stats["linhas_novas"], 0)
        self.assertEqual(stats["duplicados_ignorados"], 1)
        self.assertEqual(stats["ids_reutilizados"], ["CAR0000000001"])

    def test_conflict_handling(self):
        mock_service = MagicMock()
        mock_get = mock_service.spreadsheets().values().get().execute
        # Linha existente com mesmo movimento mas montante diferente
        existing_row = ["2026-06-26", "2026-06-25", "MERCADONA", "", "", "", "", "50.0", "", "CAR0000000001", "", "VÁLIDO", ""]
        mock_get.return_value = {"values": [self.header, existing_row]}

        mov_conflict = Movement(1, "26/06", "25/06", "MERCADONA", "", "", "", "35,15", "", 0.9, "VÁLIDO", "", "26/06 25/06 MERCADONA 35,15")

        stats = main.sync_movements_to_sheets(mock_service, self.cfg, [mov_conflict], None, "sa@test.com")

        self.assertEqual(stats["linhas_novas"], 0)
        self.assertEqual(stats["conflitos"], 1)

    def test_http_403_error_message(self):
        mock_service = MagicMock()
        resp_mock = MagicMock(status=403)
        mock_service.spreadsheets().values().get().execute.side_effect = HttpError(resp_mock, b"Forbidden")

        mov_valid = Movement(1, "26/06", "25/06", "MERCADONA", "", "", "", "35,15", "", 0.9, "VÁLIDO", "", "26/06 25/06 MERCADONA 35,15")

        with self.assertRaises(RuntimeError) as ctx:
            main.sync_movements_to_sheets(mock_service, self.cfg, [mov_valid], None, "sa@test.com")
        
        self.assertIn("Partilhe o ficheiro AppTesouraria com:", str(ctx.exception))
        self.assertIn("sa@test.com", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()


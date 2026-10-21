"""Integración de la interfaz con datos sintéticos, sin conexiones externas."""
from io import BytesIO
from unittest import TestCase
from unittest.mock import patch
from pathlib import Path

import streamlit as st
from streamlit.testing.v1 import AppTest

from test_model import workbook_bytes


class Upload(BytesIO):
    name='mercados_sinteticos.xlsx'

    @property
    def size(self):return len(self.getvalue())


class AppTests(TestCase):
    def test_upload_calculation_and_stale_downloads(self):
        uploaded=Upload(workbook_bytes())
        with patch.object(st,'file_uploader',side_effect=lambda label,**kwargs:uploaded if label=='Excel actualizado en Bloomberg' else None):
            app=AppTest.from_file(str(Path(__file__).resolve().parents[1]/'app.py')).run(timeout=30)
            self.assertFalse(app.exception)
            next(b for b in app.button if b.label=='Calcular asignación').click()
            app.run(timeout=30)
            self.assertFalse(app.exception)
            self.assertEqual(len(app.get('download_button')),3)
            next(x for x in app.number_input if x.label=='Renta variable estratégica (%)').set_value(41.)
            app.run(timeout=30)
            self.assertTrue(app.error)
            self.assertEqual(len(app.get('download_button')),0)

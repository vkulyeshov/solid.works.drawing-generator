"""Reject unrelated PDF numbers when reopened Toolbox text is incomplete."""
import math
import unittest
from test_memory_layout_native import pdf_measurement_near_graphics


class Page:
    height = 842

    def __init__(self, text, x=0.04, y=0.10):
        unit = 72 / 0.0254
        self.word = {'text': text, 'x0': x * unit, 'x1': (x + 0.004) * unit,
                     'top': self.height - (y + 0.003) * unit,
                     'bottom': self.height - y * unit}

    def extract_words(self):
        return [self.word]


class PdfMeasurementTests(unittest.TestCase):
    box = [0.039, 0.099, 0.05, 0.11]

    def test_prefix_only_com_text_has_a_real_pdf_value(self):
        self.assertTrue(pdf_measurement_near_graphics(Page('R1'), self.box, 0.001, 5))
        self.assertTrue(pdf_measurement_near_graphics(Page('5'), self.box, 0.005, 2))

    def test_unrelated_values_do_not_validate_a_missing_dimension(self):
        self.assertFalse(pdf_measurement_near_graphics(Page('15'), self.box, 0.005, 2))
        self.assertFalse(pdf_measurement_near_graphics(Page('5', x=0.15), self.box, 0.005, 2))
        self.assertFalse(pdf_measurement_near_graphics(Page('R'), self.box, 0.001, 5))

    def test_units_angles_and_printed_rounding(self):
        self.assertTrue(pdf_measurement_near_graphics(Page('4,01'), self.box, 0.0040125, 2))
        self.assertTrue(pdf_measurement_near_graphics(Page('45°'), self.box, math.pi / 4, 3))


if __name__ == '__main__':
    unittest.main(verbosity=2)

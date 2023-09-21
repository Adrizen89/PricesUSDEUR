import unittest
from unittest.mock import patch, Mock
from datetime import datetime
# Assurez-vous d'importer votre fonction EURX_scrap ici
from functions_scrapping import EURX_scrap

class TestEURXScrap(unittest.TestCase):

    @patch('requests.get')
    def test_EURX_scrap(self, mock_get):
        # Simuler une réponse HTTP
        mock_response = Mock()
        mock_response.text = '''<table>
            <tr><td>Date</td><td>Valeur</td></tr>
            <tr><td>01. January 2023</td><td>1.2345</td></tr>
            <tr><td>02. January 2023</td><td>1.6789</td></tr>
        </table>'''
        mock_response.raise_for_status = Mock()
        mock_get.return_value = mock_response

        # Appeler la fonction avec une plage de dates
        start_date = datetime.strptime("01.01.2023", "%d.%m.%Y").date()
        end_date = datetime.strptime("02.01.2023", "%d.%m.%Y").date()
        result = EURX_scrap(start_date, end_date)

        # Vérifier le résultat
        expected_result = [('02.01.2023', '1,6789'), ('01.01.2023', '1,2345')]
        self.assertEqual(result, expected_result)

        # Appeler la fonction sans plage de dates
        result_without_dates = EURX_scrap()
        expected_result_without_dates = [('01.01.2023', '1,2345')]
        self.assertEqual(result_without_dates, expected_result_without_dates)

if __name__ == '__main__':
    unittest.main()

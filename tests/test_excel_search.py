import io
from pathlib import Path
import tempfile
import unittest
from openpyxl import Workbook
from excel_search.app import create_app
from excel_search.search import search_files


class SearchTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.path = self.root / 'people.xlsx'
        book = Workbook()
        sheet = book.active
        sheet.title = 'People'
        sheet.append(['Name', 'Department', 'Reference'])
        sheet.append(['Ahmed Ali', 'Sales', 42])
        sheet.append(['Sara Mohamed', 'Operations', '=1+1'])
        sheet.append(['AHMED  ALI', 'Office', None])
        book.create_sheet('Other').append(['Ahmed Ali'])
        book.save(self.path)
        book.close()
        self.client = create_app(self.root).test_client()

    def tearDown(self):
        self.temp.cleanup()

    def test_partial_and_metadata(self):
        result = search_files([(self.path, 'people.xlsx')], ['ahmed'])
        self.assertEqual(len(result['results']), 2)
        self.assertEqual(result['results'][0]['row'], 2)
        self.assertEqual(result['results'][0]['cells'][1]['label'], 'Department')
        self.assertEqual(result['results'][0]['matched_columns'], [0])

    def test_exact_normalized_and_multiple_names(self):
        result = search_files([(self.path, 'people.xlsx')], ['ahmed ali', 'Sara'], exact=True)
        self.assertEqual(len(result['results']), 2)
        self.assertEqual(search_files([(self.path, 'p')], ['ahmed'], exact=True)['results'], [])

    def test_no_header_and_limit(self):
        result = search_files([(self.path, 'p')], ['Ahmed'], header_row=0)
        self.assertEqual(len(result['results']), 3)
        result = search_files([(self.path, 'p')], ['Ahmed'], limit=1)
        self.assertTrue(result['truncated'])
        self.assertEqual(len(result['results']), 1)

    def test_corrupt_file_does_not_stop_search(self):
        bad = self.root / 'bad.xlsx'
        bad.write_text('not excel')
        result = search_files([(bad, 'bad'), (self.path, 'good')], ['Sara'])
        self.assertEqual(len(result['errors']), 1)
        self.assertEqual(len(result['results']), 1)

    def test_api_folder_upload_and_validation(self):
        self.assertEqual(self.client.get('/').status_code, 200)
        self.assertEqual(self.client.get('/api/files').json['files'], ['people.xlsx'])
        response = self.client.post('/api/search', data={'names': 'Sara'})
        self.assertEqual(len(response.json['results']), 1)
        data = {'names': 'Ahmed', 'source': 'upload', 'files': (io.BytesIO(self.path.read_bytes()), '../../people.xlsx')}
        response = self.client.post('/api/search', data=data)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json['results'][0]['file'], 'people.xlsx')
        self.assertEqual(self.client.post('/api/search', data={'names': ''}).status_code, 400)
        self.assertEqual(self.client.post('/api/search', data={'names': 'Sara', 'header_row': '-1'}).status_code, 400)
        self.assertEqual(self.client.post('/api/search', data={'names': 'Sara', 'source': 'upload'}).status_code, 400)


if __name__ == '__main__':
    unittest.main()

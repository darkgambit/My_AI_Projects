import argparse
from pathlib import Path
import tempfile

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.exceptions import RequestEntityTooLarge

from .search import SUPPORTED, search_files


def create_app(folder=None):
    app = Flask(__name__)
    app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024
    root = Path(folder or Path.cwd()).resolve()

    def folder_files():
        return sorted(((p, str(p.relative_to(root))) for p in root.rglob('*')
                       if p.is_file() and p.suffix.lower() in SUPPORTED
                       and not p.name.startswith('~$')
                       and not any(part.startswith('.') for part in p.relative_to(root).parts)),
                      key=lambda item: item[1].lower())

    @app.get('/')
    def index():
        return render_template('index.html')

    @app.get('/download')
    def download():
        import io
        from zipfile import ZipFile, ZIP_DEFLATED
        project = Path(__file__).resolve().parent.parent
        archive = io.BytesIO()
        with ZipFile(archive, 'w', ZIP_DEFLATED) as bundle:
            for name in ('README.md', 'requirements.txt'):
                bundle.write(project / name, f'SheetScout/{name}')
            for directory in ('excel_search', 'tests'):
                for path in (project / directory).rglob('*'):
                    if path.is_file() and '__pycache__' not in path.parts:
                        bundle.write(path, f'SheetScout/{path.relative_to(project)}')
        archive.seek(0)
        return send_file(archive, mimetype='application/zip', as_attachment=True,
                         download_name='SheetScout.zip')

    @app.get('/api/files')
    def files():
        return jsonify(files=[name for _, name in folder_files()])

    @app.errorhandler(RequestEntityTooLarge)
    def too_large(error):
        return jsonify(error='Uploads must total less than 100 MB.'), 413

    @app.post('/api/search')
    def search():
        queries = list(dict.fromkeys(q.strip() for q in request.form.get('names', '').splitlines() if q.strip()))
        if not queries:
            return jsonify(error='Enter at least one name, one per line.'), 400
        if len(queries) > 100 or any(len(q) > 200 for q in queries):
            return jsonify(error='Use up to 100 names, with at most 200 characters each.'), 400
        try:
            header = int(request.form.get('header_row', '1'))
            if not 0 <= header <= 1000:
                raise ValueError()
        except ValueError:
            return jsonify(error='Header row must be a number from 0 to 1000.'), 400
        exact = request.form.get('mode') == 'exact'
        source = request.form.get('source', 'folder')
        if source not in ('folder', 'upload'):
            return jsonify(error='Choose a valid file source.'), 400
        with tempfile.TemporaryDirectory(prefix='excel-search-') as temp:
            if source == 'upload':
                paths = []
                for i, upload in enumerate(request.files.getlist('files')):
                    name = (upload.filename or '').replace('\\', '/').split('/')[-1]
                    suffix = Path(name).suffix.lower()
                    if suffix not in SUPPORTED or name.startswith('~$'):
                        return jsonify(error=f'Unsupported file: {name}. Choose .xlsx, .xlsm or .xls.'), 400
                    path = Path(temp) / f'{i}{suffix}'
                    upload.save(path)
                    paths.append((path, name))
            else:
                paths = folder_files()
            if not paths:
                return jsonify(error='No Excel files found. Add files or choose another source.'), 400
            return jsonify(search_files(paths, queries, exact=exact, header_row=header))

    return app


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Search names across Excel workbooks.')
    parser.add_argument('--folder', default=str(Path.cwd()), help='Folder to search, including subfolders')
    parser.add_argument('--port', type=int, default=5000)
    parser.add_argument('--host', default='127.0.0.1', help='Use 0.0.0.0 for a hosted preview')
    args = parser.parse_args()
    if not Path(args.folder).is_dir():
        parser.error('--folder must be an existing directory')
    create_app(args.folder).run(host=args.host, port=args.port, debug=False)

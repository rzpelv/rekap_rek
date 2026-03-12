import os, io, json, tempfile
from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename
import rekap_rek as rr

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB

def allowed(filename):
    return filename.lower().endswith('.pdf')

def parse_pdfs(files):
    all_transactions = []
    meta_combined = {
        "accountNo": "", "companyName": "", "period": "",
        "opening": 0, "totalDebet": 0, "totalKredit": 0, "closing": 0
    }
    periods = []
    file_results = []

    with tempfile.TemporaryDirectory() as tmpdir:
        for f in files:
            if not allowed(f.filename):
                continue
            fname = secure_filename(f.filename)
            fpath = os.path.join(tmpdir, fname)
            f.save(fpath)
            try:
                meta, txs = rr.parse_pdf(fpath)
                all_transactions.extend(txs)
                file_results.append({
                    'file': fname,
                    'rekening': meta['accountNo'],
                    'nama': meta['companyName'],
                    'periode': meta['period'],
                    'jumlah': len(txs),
                    'penjualan': sum(1 for t in txs if t['kategori'] == 'Penjualan'),
                })
                if not meta_combined['accountNo'] and meta['accountNo']:
                    meta_combined['accountNo']   = meta['accountNo']
                    meta_combined['companyName'] = meta['companyName']
                meta_combined['totalDebet']  += meta['totalDebet']
                meta_combined['totalKredit'] += meta['totalKredit']
                if meta['period']:
                    periods.append(meta['period'])
                if not meta_combined['opening'] and meta['opening']:
                    meta_combined['opening'] = meta['opening']
                if meta['closing']:
                    meta_combined['closing'] = meta['closing']
            except Exception as e:
                file_results.append({'file': fname, 'error': str(e)})

    if periods:
        meta_combined['period'] = (
            f"{periods[0].split(' - ')[0]} - {periods[-1].split(' - ')[-1]}"
        )
    return all_transactions, meta_combined, file_results


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/proses', methods=['POST'])
def proses():
    files = request.files.getlist('pdfs')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'error': 'Tidak ada file yang diupload'}), 400

    all_transactions, meta, file_results = parse_pdfs(files)
    if not all_transactions:
        return jsonify({'error': 'Tidak ada transaksi berhasil dibaca'}), 400

    return jsonify({
        'meta': meta,
        'files': file_results,
        'total_tx': len(all_transactions),
        'total_penj': sum(1 for t in all_transactions if t['kategori'] == 'Penjualan'),
        'transactions': [
            {
                'no': i + 1,
                'month': t['month'],
                'date': t['date'],
                'desc': t['desc'],
                'debet': t['debet'],
                'kredit': t['kredit'],
                'balance': t['balance'],
                'kategori': t['kategori'],
            }
            for i, t in enumerate(all_transactions)
        ]
    })


@app.route('/download', methods=['POST'])
def download():
    overrides_raw = request.form.get('overrides', '[]')
    try:
        kat_map = {o['no']: o['kategori'] for o in json.loads(overrides_raw)}
    except Exception:
        kat_map = {}

    files = request.files.getlist('pdfs')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'error': 'Tidak ada file PDF'}), 400

    all_transactions, meta, _ = parse_pdfs(files)
    if not all_transactions:
        return jsonify({'error': 'Tidak ada transaksi'}), 400

    for i, tx in enumerate(all_transactions):
        if (i + 1) in kat_map:
            tx['kategori'] = kat_map[i + 1]

    acc = meta.get('accountNo', 'rekening')
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
    try:
        rr.build_excel(all_transactions, meta, tmp_path)
        buf = io.BytesIO()
        with open(tmp_path, 'rb') as fh:
            buf.write(fh.read())
        buf.seek(0)
    finally:
        os.unlink(tmp_path)

    return send_file(
        buf,
        as_attachment=True,
        download_name=f"rekap_{acc}.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)

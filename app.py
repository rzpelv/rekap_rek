import os, uuid, threading, time
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template

app = Flask(__name__)

UPLOAD_DIR = Path("/tmp/bri_uploads")
OUTPUT_DIR = Path("/tmp/bri_outputs")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

jobs = {}

def _cleanup_old_files():
    now = time.time()
    for d in [UPLOAD_DIR, OUTPUT_DIR]:
        for f in d.iterdir():
            if now - f.stat().st_mtime > 3600:
                f.unlink(missing_ok=True)

def _process_job(job_id, pdf_paths, meta_override):
    try:
        import rekap_rek as rk
        all_txs, metas = [], []
        for p in pdf_paths:
            meta, txs = rk.parse_pdf(p)
            all_txs.extend(txs)
            metas.append(meta)

        mc = metas[0].copy() if metas else {}
        for m in metas[1:]:
            mc["totalDebet"]  += m.get("totalDebet", 0)
            mc["totalKredit"] += m.get("totalKredit", 0)
            if m.get("closing"): mc["closing"] = m["closing"]
        if len(metas) > 1:
            p0 = metas[0].get("period","").split(" - ")[0]
            p1 = metas[-1].get("period","").split(" - ")[-1]
            mc["period"] = f"{p0} - {p1}"

        if meta_override.get("companyName"): mc["companyName"] = meta_override["companyName"]
        if meta_override.get("accountNo"):   mc["accountNo"]   = meta_override["accountNo"]

        out_path = OUTPUT_DIR / f"rekap_{job_id}.xlsx"
        rk.build_excel(all_txs, mc, str(out_path))

        penj = [t for t in all_txs if t["kategori"] == "Penjualan"]
        by_month = {}
        for t in all_txs:
            by_month.setdefault(t["month"], {"debet":0,"kredit":0,"penjualan":0,"count":0})
            by_month[t["month"]]["debet"]    += t["debet"]
            by_month[t["month"]]["kredit"]   += t["kredit"]
            by_month[t["month"]]["count"]    += 1
            if t["kategori"] == "Penjualan":
                by_month[t["month"]]["penjualan"] += t["kredit"]

        jobs[job_id].update({
            "status": "done",
            "filename": out_path.name,
            "stats": {
                "company":       mc.get("companyName",""),
                "account":       mc.get("accountNo",""),
                "period":        mc.get("period",""),
                "total_tx":      len(all_txs),
                "total_penj":    len(penj),
                "total_penj_rp": sum(t["kredit"] for t in penj),
                "total_debet":   sum(t["debet"]  for t in all_txs),
                "total_kredit":  sum(t["kredit"] for t in all_txs),
                "by_month":      by_month,
            }
        })
    except Exception as e:
        import traceback
        jobs[job_id].update({"status":"error","message":str(e),"trace":traceback.format_exc()})
    finally:
        for p in pdf_paths:
            try: Path(p).unlink()
            except: pass

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    _cleanup_old_files()
    files = request.files.getlist("pdfs")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "Tidak ada file PDF"}), 400
    job_id, pdf_paths = uuid.uuid4().hex[:12], []
    for f in files:
        if not f.filename.lower().endswith(".pdf"):
            return jsonify({"error": f"{f.filename} bukan PDF"}), 400
        dest = UPLOAD_DIR / f"{job_id}_{uuid.uuid4().hex[:6]}.pdf"
        f.save(dest)
        pdf_paths.append(str(dest))
    meta_override = {
        "companyName": request.form.get("companyName","").strip(),
        "accountNo":   request.form.get("accountNo","").strip(),
    }
    jobs[job_id] = {"status": "processing"}
    threading.Thread(target=_process_job, args=(job_id, pdf_paths, meta_override), daemon=True).start()
    return jsonify({"job_id": job_id})

@app.route("/status/<job_id>")
def status(job_id):
    job = jobs.get(job_id)
    if not job: return jsonify({"error": "Job tidak ditemukan"}), 404
    return jsonify(job)

@app.route("/download/<job_id>")
def download(job_id):
    job = jobs.get(job_id)
    if not job or job.get("status") != "done":
        return jsonify({"error": "File belum siap"}), 404
    path = OUTPUT_DIR / job["filename"]
    if not path.exists(): return jsonify({"error": "File tidak ditemukan"}), 404
    company = job["stats"].get("company","rekap").replace(" ","_")[:20]
    return send_file(path, as_attachment=True,
        download_name=f"rekap_bri_{company}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT",5000)), debug=False)

"""Flask + Vercel entrypoint. Uploads processed only in memory."""
import os
from datetime import datetime
from io import BytesIO
from pathlib import Path
from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.exceptions import HTTPException, RequestEntityTooLarge
from engine import raw_parse, master_load, preview_data, part1_book, create_master, make_zip

app=Flask(__name__)
app.config['MAX_CONTENT_LENGTH']=int(os.environ.get('MAX_UPLOAD_BYTES','4000000'))
app.config['SEND_FILE_MAX_AGE_DEFAULT']=0

@app.after_request
def headers(response):
    response.headers['Cache-Control']='no-store, private'
    response.headers['X-Content-Type-Options']='nosniff'
    response.headers['Referrer-Policy']='no-referrer'
    response.headers['X-Frame-Options']='DENY'
    response.headers['Content-Security-Policy']="default-src 'self'; script-src 'self' https://cdn.tailwindcss.com; style-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net https://fonts.googleapis.com; font-src 'self' data: https://fonts.gstatic.com; img-src 'self' data:; connect-src 'self'; object-src 'none'; base-uri 'self'; frame-ancestors 'none'"
    return response

@app.errorhandler(Exception)
def fail(exc):
    if isinstance(exc,RequestEntityTooLarge):return jsonify(error='Combined upload is too large for this deployment. Stay under 4 MB including files.'),413
    if isinstance(exc,HTTPException):return jsonify(error=exc.description),exc.code
    if isinstance(exc,(ValueError,TypeError)):return jsonify(error=str(exc)),400
    app.logger.exception('Processing failed')
    return jsonify(error='File processing failed. Check source workbooks and try again.'),500

@app.get('/')
def index():return render_template('index.html',protected=bool(os.environ.get('APP_PASSWORD')))

@app.get('/api/health')
def health():return jsonify(status='ok',service='ECHS NMI Report Studio')

def intake():
    password=os.environ.get('APP_PASSWORD')
    if password:
        import hmac
        if not hmac.compare_digest(request.headers.get('X-App-Password',''),password):
            from werkzeug.exceptions import Unauthorized
            raise Unauthorized('Incorrect access password.')
    raw=request.files.getlist('raw_files')
    if not raw or any(not f.filename for f in raw):raise ValueError('Select your Raw NMI files.')
    if len(raw)>40:raise ValueError('Maximum 40 Raw files per request.')
    files=[(Path(f.filename).name,f.read()) for f in raw]
    master=request.files.get('master_file')
    if master and master.filename: old=master.read();filename=Path(master.filename).name
    else:old=None;filename=None
    inp=(request.form.get('report_date') or datetime.now().strftime('%d-%m-%Y')).strip()
    try:asof=datetime.strptime(inp,'%d-%m-%Y')
    except ValueError:raise ValueError('Report date must be DD-MM-YYYY.')
    rows,unique,stats=raw_parse(files)
    wb,ws,hr=master_load(old,filename)
    preview=preview_data(rows,unique,stats,wb,ws,hr,asof)
    return rows,unique,stats,wb,ws,hr,asof,preview,filename

@app.post('/api/preview')
def preview():
    *_,data,filename=intake()
    return jsonify(data)

@app.post('/api/generate')
def generate():
    rows,unique,stats,wb,ws,hr,asof,data,filename=intake()
    if data['pending_to_done'] and request.form.get('confirm_absent')!='true':
        return jsonify(error='Confirmation required before marking previous Pending records Done.'),409
    rawbook=part1_book(rows,stats)
    masterbook=create_master(wb,ws,hr,unique,stats,asof)
    packed,outputname=make_zip(rawbook,masterbook,asof,'xlsm' if filename and filename.lower().endswith('.xlsm') else 'xlsx')
    if len(packed)>4200000:return jsonify(error='Generated report exceeds Vercel response limit (4.5 MB). Run this batch locally/Colab or use external storage.'),413
    return send_file(BytesIO(packed),mimetype='application/zip',as_attachment=True,download_name=outputname)

if __name__=='__main__':app.run(debug=os.environ.get('FLASK_DEBUG')=='1',port=int(os.environ.get('PORT',5000)))

import os
import re
import uuid
import zipfile
import subprocess
import logging
import json
from flask import Flask, request, jsonify, send_file, render_template
import openpyxl
try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'

ALLOWED_EXTENSIONS = {'doc', 'docx', 'xls', 'xlsx'}

FIELD_DEFS = [
    {'key': 'seat_count',    'label': '座位数',             'example': '45'},
    {'key': 'car_type',      'label': '车型',               'example': '大巴'},
    {'key': 'plate',         'label': '车牌号（苏K-后面部分）', 'example': '12345'},
    {'key': 'driver1_name',  'label': '驾驶员1姓名',         'example': '张三'},
    {'key': 'driver1_phone', 'label': '驾驶员1电话',         'example': '13800000001'},
    {'key': 'driver1_cert',  'label': '驾驶员1从业资格证号',  'example': 'JS0001'},
    {'key': 'driver2_name',  'label': '驾驶员2姓名',         'example': ''},
    {'key': 'driver2_phone', 'label': '驾驶员2电话',         'example': ''},
    {'key': 'driver2_cert',  'label': '驾驶员2从业资格证号',  'example': ''},
    {'key': 'transport_cert','label': '道路运输证号',        'example': 'YZ001'},
    {'key': 'start_date',    'label': '用车开始日期',        'example': '2026-06-15', 'type': 'date'},
    {'key': 'end_date',      'label': '用车结束日期',        'example': '2026-06-16', 'type': 'date'},
    {'key': 'dest',          'label': '目的地',             'example': '南京'},
    {'key': 'route',         'label': '途经',               'example': '宁沪高速'},
    {'key': 'fee',           'label': '运输费用（元）',      'example': '2000'},
    {'key': 'sign_date',     'label': '签订日期',           'example': '2026-06-10', 'type': 'date'},
]

NANJING_BLANK_MAP = [
    (4,  0, 'seat_count'),
    (4,  1, 'car_type'),
    (4,  2, 'plate'),
    (5,  0, 'driver1_name'),
    (5,  1, 'driver1_phone'),
    (6,  0, 'driver1_cert'),
    (7,  0, 'driver2_name'),
    (7,  1, 'driver2_phone'),
    (8,  0, 'driver2_cert'),
    (9,  1, 'transport_cert'),
    (9,  2, 'start_year'),
    (9,  3, 'start_month'),
    (9,  4, 'start_day'),
    (9,  5, 'end_year'),
    (9,  6, 'end_month'),
    (9,  7, 'end_day'),
    (9,  8, 'dest'),
    (10, 0, 'route'),
    (11, 2, 'fee'),
    (28, 0, 'sign_year'),
    (28, 1, 'sign_month'),
    (28, 2, 'sign_day'),
]


def read_excel(path):
    ext = path.rsplit('.', 1)[-1].lower()
    if ext == 'xls':
        if not HAS_XLRD:
            raise RuntimeError('服务器缺少xlrd库，无法读取.xls文件')
        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_index(0)
        all_rows = [ws.row_values(i) for i in range(ws.nrows)]
    else:
        wb = openpyxl.load_workbook(path, read_only=True)
        ws = wb.active
        all_rows = [list(r) for r in ws.iter_rows(values_only=True)]

    if not all_rows:
        raise RuntimeError('Excel为空')

    headers = [str(h).strip() if h is not None else '' for h in all_rows[0]]
    records = []
    for row in all_rows[1:]:
        if any(v is not None and str(v).strip() != '' for v in row):
            first = str(row[0]) if row[0] is not None else ''
            if '示例' in first or '↓' in first:
                continue
            records.append({headers[i]: (str(v) if v is not None else '') for i, v in enumerate(row)})
    return headers, records


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def normalize_plate(val):
    """车牌号去掉苏K-前缀"""
    if val and val.upper().startswith('苏K-'):
        return val[3:]
    return val


def fill_doc_template(template_path, data, output_path):
    import re as _re
    from docx import Document
    from collections import defaultdict

    doc = Document(template_path)
    all_text = ' '.join(p.text for p in doc.paragraphs)
    use_placeholder = '{{' in all_text

    if use_placeholder:
        def replace_in_para(para):
            full = ''.join(r.text for r in para.runs)
            for key, val in data.items():
                full = full.replace('{{' + key + '}}', str(val) if val else '')
            if para.runs:
                para.runs[0].text = full
                for r in para.runs[1:]:
                    r.text = ''
        for para in doc.paragraphs:
            replace_in_para(para)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        replace_in_para(para)
    else:
        def replace_blanks_in_para(para, replacements):
            full = ''.join(r.text for r in para.runs)
            gaps = list(_re.finditer(r' {2,}', full))
            for gap_i, val in sorted(replacements.items(), reverse=True):
                if gap_i < len(gaps):
                    m = gaps[gap_i]
                    full = full[:m.start()] + str(val) + full[m.end():]
            if para.runs:
                para.runs[0].text = full
                for r in para.runs[1:]:
                    r.text = ''

        para_replacements = defaultdict(dict)
        for para_i, gap_i, key in NANJING_BLANK_MAP:
            val = data.get(key, '')
            if val:
                para_replacements[para_i][gap_i] = val

        for para_i, replacements in para_replacements.items():
            if para_i < len(doc.paragraphs):
                replace_blanks_in_para(doc.paragraphs[para_i], replacements)

    _fix_to_single_page(doc)
    doc.save(output_path)


def _fix_to_single_page(doc):
    """调整格式使所有内容适合单页A4输出"""
    from docx.shared import Pt
    from docx.enum.text import WD_LINE_SPACING
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    # 1. 修正页面：去landscape，竖向A4，收紧页边距
    for sectPr in doc.element.body.iter('{%s}sectPr' % ns):
        pgSz = sectPr.find('{%s}pgSz' % ns)
        if pgSz is not None:
            orient_key = '{%s}orient' % ns
            if pgSz.get(orient_key) == 'landscape':
                del pgSz.attrib[orient_key]
            pgSz.set('{%s}w' % ns, '11906')
            pgSz.set('{%s}h' % ns, '16838')
        pgMar = sectPr.find('{%s}pgMar' % ns)
        if pgMar is not None:
            pgMar.set('{%s}top' % ns, '200')
            pgMar.set('{%s}bottom' % ns, '100')
            pgMar.set('{%s}left' % ns, '700')
            pgMar.set('{%s}right' % ns, '700')

    # 2. 所有run字体统一缩小到10.5pt（标题保持原大小）
    for i, p in enumerate(doc.paragraphs):
        for r in p.runs:
            if i != 0:  # 跳过标题
                if not r.font.size or r.font.size.pt > 10.5:
                    r.font.size = Pt(10.5)
        # 3. 行距：正文11pt固定，空行压为1pt
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        if p.text.strip():
            pf.line_spacing = Pt(11)
        else:
            pf.line_spacing = Pt(1)


def run_cmd(cmd, timeout=60):
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    try:
        stdout, stderr = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        proc.kill()
        proc.communicate()
        raise RuntimeError('命令超时: ' + ' '.join(cmd))
    return proc.returncode, stdout.decode('utf-8', errors='replace'), stderr.decode('utf-8', errors='replace')


def docx_to_single_image(docx_path, output_dir, out_name='contract'):
    """docx -> PDF -> 所有页拼成一张PNG"""
    code, out, err = run_cmd(
        ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, docx_path]
    )
    if code != 0:
        raise RuntimeError('LibreOffice转换失败: ' + err)

    pdf_name = os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
    pdf_path = os.path.join(output_dir, pdf_name)
    if not os.path.exists(pdf_path):
        raise RuntimeError('PDF文件未生成')

    img_prefix = os.path.join(output_dir, 'page')
    # 只取第1页（文档内容实际在第1页，LibreOffice渲染差异导致部分内容溢出第2页）
    code2, out2, err2 = run_cmd(['pdftoppm', '-r', '200', '-png', '-f', '1', '-l', '1', pdf_path, img_prefix])
    if code2 != 0:
        raise RuntimeError('pdftoppm转换失败: ' + err2)

    pages = sorted([
        os.path.join(output_dir, f)
        for f in os.listdir(output_dir)
        if f.startswith('page') and f.endswith('.png')
    ])
    if not pages:
        raise RuntimeError('未找到生成的图片')

    final_path = os.path.join(output_dir, out_name + '.png')

    if len(pages) == 1:
        os.rename(pages[0], final_path)
    else:
        from PIL import Image, ImageChops
        def trim_white(img):
            """裁掉图片四周纯白边框"""
            bg = Image.new(img.mode, img.size, (255, 255, 255))
            diff = ImageChops.difference(img, bg)
            bbox = diff.getbbox()
            if bbox:
                # 只裁上下，保留左右原始宽度（避免不同页宽度不一致）
                return img.crop((0, bbox[1], img.width, bbox[3]))
            return img

        imgs = [Image.open(p) for p in pages]
        # 裁掉每页上下空白边
        trimmed = [trim_white(img) for img in imgs]
        w = max(i.width for i in trimmed)
        h = sum(i.height for i in trimmed)
        merged = Image.new('RGB', (w, h), (255, 255, 255))
        y = 0
        for img in trimmed:
            merged.paste(img, (0, y))
            y += img.height
        merged.save(final_path)

    return final_path


def split_date(data, date_key, year_key, month_key, day_key):
    """把 yyyy-mm-dd 日期字段拆分为年月日三个字段"""
    val = data.pop(date_key, '') or ''
    val = val.strip()
    if val and '-' in val:
        parts = val.split('-')
        if len(parts) == 3:
            data[year_key]  = parts[0].lstrip('0') or parts[0]
            data[month_key] = parts[1].lstrip('0') or parts[1]
            data[day_key]   = parts[2].lstrip('0') or parts[2]
            return
    # 没填或格式不对，留空
    data[year_key] = data[month_key] = data[day_key] = ''


def build_fill_data(source_data):
    """处理填充数据：拆分日期、归一化车牌"""
    data = dict(source_data)
    data['plate'] = normalize_plate(data.get('plate', ''))
    split_date(data, 'start_date', 'start_year', 'start_month', 'start_day')
    split_date(data, 'end_date',   'end_year',   'end_month',   'end_day')
    split_date(data, 'sign_date',  'sign_year',  'sign_month',  'sign_day')
    return data


@app.route('/')
def index():
    return render_template('index.html', fields=FIELD_DEFS)


@app.route('/api/fields')
def get_fields():
    return jsonify(FIELD_DEFS)


@app.route('/api/parse-excel', methods=['POST'])
def parse_excel():
    if 'excel' not in request.files:
        return jsonify({'error': '请上传Excel文件'}), 400
    f = request.files['excel']
    if not allowed_file(f.filename):
        return jsonify({'error': '不支持的文件格式'}), 400
    ext = f.filename.rsplit('.', 1)[-1].lower() if '.' in f.filename else ''
    path = os.path.join(app.config['UPLOAD_FOLDER'], str(uuid.uuid4())[:8] + '.' + ext)
    f.save(path)
    try:
        headers, data = read_excel(path)
        return jsonify({'headers': headers, 'preview': data[:5], 'total': len(data)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/generate-manual', methods=['POST'])
def generate_manual():
    """手动输入单条数据，生成一张合同图片并直接返回"""
    if 'template' not in request.files:
        return jsonify({'error': '请上传Word模板'}), 400

    tmpl_file = request.files['template']
    raw_data = request.form.get('data')
    if not raw_data:
        return jsonify({'error': '请提供字段数据'}), 400

    try:
        fill_data = build_fill_data(json.loads(raw_data))
    except Exception as e:
        return jsonify({'error': '数据格式错误: ' + str(e)}), 400

    task_id = str(uuid.uuid4())[:8]
    task_dir = os.path.join(app.config['OUTPUT_FOLDER'], task_id)
    os.makedirs(task_dir)

    tmpl_ext = tmpl_file.filename.rsplit('.', 1)[1].lower() if '.' in tmpl_file.filename else 'docx'
    tmpl_path = os.path.join(app.config['UPLOAD_FOLDER'], task_id + '_template.' + tmpl_ext)
    tmpl_file.save(tmpl_path)

    if tmpl_ext == 'doc':
        code, out, err = run_cmd(
            ['libreoffice', '--headless', '--convert-to', 'docx',
             '--outdir', app.config['UPLOAD_FOLDER'], tmpl_path]
        )
        tmpl_docx = tmpl_path.replace('.doc', '.docx')
        if not os.path.exists(tmpl_docx):
            return jsonify({'error': 'doc转docx失败: ' + err}), 500
        tmpl_path = tmpl_docx

    try:
        filled_docx = os.path.join(task_dir, 'contract.docx')
        fill_doc_template(tmpl_path, fill_data, filled_docx)
        img_path = docx_to_single_image(filled_docx, task_dir, 'contract')
        return send_file(img_path, mimetype='image/png',
                         as_attachment=True, download_name='合同.png')
    except Exception as e:
        logger.error('generate_manual failed: %s', e)
        return jsonify({'error': str(e)}), 500


@app.route('/api/generate', methods=['POST'])
def generate():
    """批量Excel生成，多份合同打包ZIP（每份一张图片）"""
    if 'template' not in request.files or 'excel' not in request.files:
        return jsonify({'error': '请上传模板和Excel'}), 400

    tmpl_file = request.files['template']
    excel_file = request.files['excel']
    mapping_raw = request.form.get('mapping')
    if not mapping_raw:
        return jsonify({'error': '请提供字段映射'}), 400
    mapping = json.loads(mapping_raw)

    task_id = str(uuid.uuid4())[:8]
    task_dir = os.path.join(app.config['OUTPUT_FOLDER'], task_id)
    os.makedirs(task_dir)

    tmpl_ext = tmpl_file.filename.rsplit('.', 1)[1].lower()
    tmpl_path = os.path.join(app.config['UPLOAD_FOLDER'], task_id + '_template.' + tmpl_ext)
    tmpl_file.save(tmpl_path)

    excel_ext = excel_file.filename.rsplit('.', 1)[-1].lower() if '.' in excel_file.filename else 'xlsx'
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], task_id + '_data.' + excel_ext)
    excel_file.save(excel_path)

    if tmpl_ext == 'doc':
        code, out, err = run_cmd(
            ['libreoffice', '--headless', '--convert-to', 'docx',
             '--outdir', app.config['UPLOAD_FOLDER'], tmpl_path]
        )
        tmpl_docx = tmpl_path.replace('.doc', '.docx')
        if not os.path.exists(tmpl_docx):
            return jsonify({'error': 'doc转docx失败: ' + err}), 500
        tmpl_path = tmpl_docx

    try:
        headers, records = read_excel(excel_path)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    if not records:
        return jsonify({'error': 'Excel无数据行'}), 400

    # 单条直接返回图片，多条打ZIP
    if len(records) == 1:
        record = records[0]
        raw = {field_key: record.get(col_name, '') for field_key, col_name in mapping.items()}
        fill_data = build_fill_data(raw)
        row_dir = os.path.join(task_dir, 'row_1')
        os.makedirs(row_dir)
        try:
            filled_docx = os.path.join(row_dir, 'contract.docx')
            fill_doc_template(tmpl_path, fill_data, filled_docx)
            img_path = docx_to_single_image(filled_docx, row_dir, 'contract')
            return send_file(img_path, mimetype='image/png',
                             as_attachment=True, download_name='合同.png')
        except Exception as e:
            return jsonify({'error': str(e)}), 500

    zip_path = os.path.join(app.config['OUTPUT_FOLDER'], task_id + '.zip')
    errors = []
    generated = 0

    with zipfile.ZipFile(zip_path, 'w') as zf:
        for idx, record in enumerate(records):
            raw = {field_key: record.get(col_name, '') for field_key, col_name in mapping.items()}
            fill_data = build_fill_data(raw)
            row_dir = os.path.join(task_dir, 'row_%d' % (idx + 1))
            os.makedirs(row_dir)
            try:
                filled_docx = os.path.join(row_dir, 'contract.docx')
                fill_doc_template(tmpl_path, fill_data, filled_docx)
                img_path = docx_to_single_image(filled_docx, row_dir, 'contract')
                zf.write(img_path, '合同_%d.png' % (idx + 1))
                generated += 1
            except Exception as e:
                errors.append('第%d行: %s' % (idx + 1, str(e)))
                logger.error('Row %d failed: %s', idx + 1, e)

    if generated == 0:
        return jsonify({'error': '所有记录生成失败', 'details': errors}), 500

    return jsonify({
        'task_id': task_id,
        'generated': generated,
        'errors': errors,
        'download_url': '/api/download/' + task_id
    })


@app.route('/api/download/<task_id>')
def download(task_id):
    if not re.match(r'^[a-f0-9]{8}$', task_id):
        return jsonify({'error': '无效task_id'}), 400
    zip_path = os.path.join(app.config['OUTPUT_FOLDER'], task_id + '.zip')
    if not os.path.exists(zip_path):
        return jsonify({'error': '文件不存在'}), 404
    return send_file(zip_path, as_attachment=True, download_name='合同图片.zip')


if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('output', exist_ok=True)
    app.run(host='127.0.0.1', port=5000, debug=False)

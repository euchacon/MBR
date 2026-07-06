from flask import Flask, request, jsonify, send_file, render_template
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import copy, io, os, re

app = Flask(__name__)
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.pptx')

# ── helpers ──────────────────────────────────────────────────────────────────

def set_text_frame(tf, lines, preserve_format=True):
    """Replace all paragraphs in a text frame with new lines."""
    if not lines:
        lines = ['']
    # Grab format from first run of first paragraph
    ref_para = tf.paragraphs[0] if tf.paragraphs else None
    ref_run = ref_para.runs[0] if (ref_para and ref_para.runs) else None

    # Clear existing paragraphs (keep first, remove rest)
    from pptx.oxml.ns import qn
    from lxml import etree
    txBody = tf._txBody
    existing = txBody.findall(qn('a:p'))
    for p in existing[1:]:
        txBody.remove(p)

    first_p = existing[0]
    # Clear runs from first para
    for r in first_p.findall(qn('a:r')):
        first_p.remove(r)
    # Set first line
    r_elem = copy.deepcopy(ref_run._r) if ref_run else etree.SubElement(first_p, qn('a:r'))
    rPr = r_elem.find(qn('a:rPr'))
    t_elem = r_elem.find(qn('a:t'))
    if t_elem is None:
        t_elem = etree.SubElement(r_elem, qn('a:t'))
    t_elem.text = lines[0]
    first_p.append(r_elem)

    # Add remaining lines as new paragraphs
    for line in lines[1:]:
        new_p = copy.deepcopy(first_p)
        for r in new_p.findall(qn('a:r')):
            new_p.remove(r)
        new_r = copy.deepcopy(r_elem)
        new_t = new_r.find(qn('a:t'))
        if new_t is None:
            new_t = etree.SubElement(new_r, qn('a:t'))
        new_t.text = line
        new_p.append(new_r)
        txBody.append(new_p)


def replace_shape_text(slide, shape_name, new_text):
    """Find shape by name and replace its text frame."""
    for shape in slide.shapes:
        if shape.name == shape_name and shape.has_text_frame:
            lines = new_text.split('\n') if new_text else ['']
            set_text_frame(shape.text_frame, lines)
            return True
    return False


def replace_table_cell(slide, shape_name, row, col, text):
    for shape in slide.shapes:
        if shape.name == shape_name and shape.has_table:
            try:
                cell = shape.table.cell(row, col)
                tf = cell.text_frame
                from pptx.oxml.ns import qn
                txBody = tf._txBody
                paras = txBody.findall(qn('a:p'))
                for p in paras[1:]:
                    txBody.remove(p)
                first_p = paras[0]
                for r in first_p.findall(qn('a:r')):
                    first_p.remove(r)
                from lxml import etree
                r_elem = etree.SubElement(first_p, qn('a:r'))
                t_elem = etree.SubElement(r_elem, qn('a:t'))
                t_elem.text = str(text) if text else ''
            except:
                pass


def fill_table(slide, table_shape_name, data_rows, start_row=1):
    """Fill table rows starting from start_row with data_rows (list of lists)."""
    for shape in slide.shapes:
        if shape.name == table_shape_name and shape.has_table:
            tbl = shape.table
            for ri, row_data in enumerate(data_rows):
                tbl_row = ri + start_row
                if tbl_row >= len(tbl.rows):
                    break
                for ci, val in enumerate(row_data):
                    if ci < len(tbl.columns):
                        cell = tbl.cell(tbl_row, ci)
                        tf = cell.text_frame
                        from pptx.oxml.ns import qn
                        from lxml import etree
                        txBody = tf._txBody
                        paras = txBody.findall(qn('a:p'))
                        for p in paras[1:]:
                            txBody.remove(p)
                        first_p = paras[0]
                        for r in first_p.findall(qn('a:r')):
                            first_p.remove(r)
                        r_elem = etree.SubElement(first_p, qn('a:r'))
                        t_elem = etree.SubElement(r_elem, qn('a:t'))
                        t_elem.text = str(val) if val else ''


# ── main route ────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    d = request.json
    prs = Presentation(TEMPLATE_PATH)
    slides = prs.slides

    # ── SLIDE 1: COVER ────────────────────────────────────────────────────────
    s1 = slides[0]
    replace_shape_text(s1, 'Text 6', d.get('dealer', ''))
    replace_shape_text(s1, 'Text 8', d.get('country', ''))
    replace_shape_text(s1, 'Text 10', d.get('period', ''))
    replace_shape_text(s1, 'Text 12', d.get('prepared_by', 'Eugenio Chacon'))

    # ── SLIDE 2: SNAPSHOT ─────────────────────────────────────────────────────
    s2 = slides[1]
    replace_shape_text(s2, 'Text 4', d.get('active_accounts', '[#]'))
    replace_shape_text(s2, 'Text 8', d.get('total_backlog', '$[X,XXX]'))
    replace_shape_text(s2, 'Text 12', d.get('units_invoiced', '[#]'))
    replace_shape_text(s2, 'Text 16', d.get('revenue_invoiced', '$[X,XXX]'))
    # Key movements
    movements = d.get('key_movements', '')
    replace_shape_text(s2, 'Text 22', movements)
    # Flags
    replace_shape_text(s2, 'Text 27', d.get('lzb_needs', ''))
    replace_shape_text(s2, 'Text 30', d.get('blockers', ''))

    # ── SLIDE 3: ACCOUNT PERFORMANCE ─────────────────────────────────────────
    s3 = slides[2]
    # Table — find table shape
    acc_rows = d.get('accounts', [])
    for shape in s3.shapes:
        if shape.has_table:
            tbl = shape.table
            # rows 1+ are data (row 0 = header)
            from pptx.oxml.ns import qn
            from lxml import etree
            for ri, row_data in enumerate(acc_rows):
                tbl_row = ri + 1
                if tbl_row >= len(tbl.rows):
                    break
                cols = ['account','doors','slots','units_sold','revenue','floor_inv','backlog','rotation','vs_target','status']
                for ci, col_key in enumerate(cols):
                    if ci < len(tbl.columns):
                        cell = tbl.cell(tbl_row, ci)
                        tf = cell.text_frame
                        txBody = tf._txBody
                        paras = txBody.findall(qn('a:p'))
                        for p in paras[1:]:
                            txBody.remove(p)
                        first_p = paras[0]
                        for r in first_p.findall(qn('a:r')):
                            first_p.remove(r)
                        r_elem = etree.SubElement(first_p, qn('a:r'))
                        t_elem = etree.SubElement(r_elem, qn('a:t'))
                        t_elem.text = str(row_data.get(col_key, ''))
            break

    # Performance notes
    perf_notes = d.get('perf_notes', '')
    replace_shape_text(s3, 'Text 148', perf_notes)

    # ── SLIDE 4: INVENTORY ────────────────────────────────────────────────────
    s4 = slides[3]
    replace_shape_text(s4, 'Text 4',  d.get('units_on_floor', '[#]'))
    replace_shape_text(s4, 'Text 8',  d.get('units_in_transit', '[#]'))
    replace_shape_text(s4, 'Text 12', d.get('backlog_value', '$[X,XXX]'))
    replace_shape_text(s4, 'Text 16', d.get('weeks_cover', '[#]') + ' wks')
    # Inventory table
    inv_rows = d.get('inventory', [])
    tbl_count = 0
    for shape in s4.shapes:
        if shape.has_table:
            tbl_count += 1
            if tbl_count == 1:
                tbl = shape.table
                from pptx.oxml.ns import qn
                from lxml import etree
                cols = ['account','sku','on_floor','sold_mtd','in_transit','reorder','next_action','target_stock','days_cover']
                for ri, row_data in enumerate(inv_rows):
                    tbl_row = ri + 1
                    if tbl_row >= len(tbl.rows):
                        break
                    for ci, col_key in enumerate(cols):
                        if ci < len(tbl.columns):
                            cell = tbl.cell(tbl_row, ci)
                            tf = cell.text_frame
                            txBody = tf._txBody
                            paras = txBody.findall(qn('a:p'))
                            for p in paras[1:]:
                                txBody.remove(p)
                            first_p = paras[0]
                            for r in first_p.findall(qn('a:r')):
                                first_p.remove(r)
                            r_elem = etree.SubElement(first_p, qn('a:r'))
                            t_elem = etree.SubElement(r_elem, qn('a:t'))
                            t_elem.text = str(row_data.get(col_key, ''))
    replace_shape_text(s4, 'Text 114', d.get('inv_issues', ''))

    # ── SLIDE 5: COMMERCIAL NARRATIVE ─────────────────────────────────────────
    s5 = slides[4]
    replace_shape_text(s5, 'Text 5',  d.get('what_worked', ''))
    replace_shape_text(s5, 'Text 9',  d.get('what_didnt', ''))
    replace_shape_text(s5, 'Text 13', d.get('cust_feedback', ''))
    replace_shape_text(s5, 'Text 17', d.get('market_obs', ''))

    # ── SLIDE 6: PIPELINE ─────────────────────────────────────────────────────
    s6 = slides[5]
    # Pipeline table
    pipe_rows = d.get('pipeline', [])
    for shape in s6.shapes:
        if shape.has_table:
            tbl = shape.table
            from pptx.oxml.ns import qn
            from lxml import etree
            cols = ['account','opportunity','est_units','est_usd','probability','expected_close','next_action','owner']
            for ri, row_data in enumerate(pipe_rows):
                tbl_row = ri + 1
                if tbl_row >= len(tbl.rows):
                    break
                for ci, col_key in enumerate(cols):
                    if ci < len(tbl.columns):
                        cell = tbl.cell(tbl_row, ci)
                        tf = cell.text_frame
                        txBody = tf._txBody
                        paras = txBody.findall(qn('a:p'))
                        for p in paras[1:]:
                            txBody.remove(p)
                        first_p = paras[0]
                        for r in first_p.findall(qn('a:r')):
                            first_p.remove(r)
                        r_elem = etree.SubElement(first_p, qn('a:r'))
                        t_elem = etree.SubElement(r_elem, qn('a:t'))
                        t_elem.text = str(row_data.get(col_key, ''))
            break

    replace_shape_text(s6, 'Text 88', d.get('orders_to_place', ''))
    replace_shape_text(s6, 'Text 91', d.get('activities_planned', ''))

    # ── SLIDE 7: NEW CUSTOMER DEPLOYMENT ──────────────────────────────────────
    s7 = slides[6]
    nc_rows = d.get('new_customers', [])
    for shape in s7.shapes:
        if shape.has_table:
            tbl = shape.table
            from pptx.oxml.ns import qn
            from lxml import etree
            cols = ['market','dealer','customer','doors','total_slots','launch_date','status','next_step']
            for ri, row_data in enumerate(nc_rows):
                tbl_row = ri + 1
                if tbl_row >= len(tbl.rows):
                    break
                for ci, col_key in enumerate(cols):
                    if ci < len(tbl.columns):
                        cell = tbl.cell(tbl_row, ci)
                        tf = cell.text_frame
                        txBody = tf._txBody
                        paras = txBody.findall(qn('a:p'))
                        for p in paras[1:]:
                            txBody.remove(p)
                        first_p = paras[0]
                        for r in first_p.findall(qn('a:r')):
                            first_p.remove(r)
                        r_elem = etree.SubElement(first_p, qn('a:r'))
                        t_elem = etree.SubElement(r_elem, qn('a:t'))
                        t_elem.text = str(row_data.get(col_key, ''))
            break

    replace_shape_text(s7, 'Text 71', d.get('onboarding_notes', ''))

    # ── SLIDE 8: NEEDS & ESCALATIONS ──────────────────────────────────────────
    s8 = slides[7]
    replace_shape_text(s8, 'Text 5',  d.get('need_marketing', ''))
    replace_shape_text(s8, 'Text 9',  d.get('need_pricing', ''))
    replace_shape_text(s8, 'Text 13', d.get('need_product', ''))
    replace_shape_text(s8, 'Text 17', d.get('need_training', ''))
    replace_shape_text(s8, 'Text 21', d.get('need_logistics', ''))
    replace_shape_text(s8, 'Text 25', d.get('need_strategic', ''))

    # ── SLIDE 9: MARKETING ────────────────────────────────────────────────────
    s9 = slides[8]
    replace_shape_text(s9, 'Text 4',  d.get('mkt_events', ''))
    replace_shape_text(s9, 'Text 7',  d.get('mkt_digital', ''))
    replace_shape_text(s9, 'Text 10', d.get('mkt_instore', ''))
    replace_shape_text(s9, 'Text 15', d.get('mkt_asks', ''))

    # ── SLIDE 10: CLOSING ─────────────────────────────────────────────────────
    s10 = slides[9]
    next_review = d.get('next_review', '[Date]')
    next_format = d.get('next_format', 'video call')
    replace_shape_text(s10, 'Text 2', f'Next review:  {next_review}  —  {next_format}')
    dealer_email = d.get('dealer_email', '')
    jeff_email   = d.get('jeff_email', 'jeff.lillich@la-z-boy.com')
    latam_email  = d.get('latam_email', 'eugenio.chacon@gmail.com')
    replace_shape_text(s10, 'Text 3', f'Questions or follow-ups:  {dealer_email}  |  Jeff:  {jeff_email}  |  LATAM Team:  {latam_email}')
    replace_shape_text(s10, 'Text 4', f'La-Z-Boy LATAM  •  Confidential. Internal Use Only.  •  {d.get("period","[Month YYYY]")}')

    # ── OUTPUT ────────────────────────────────────────────────────────────────
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    period_safe = re.sub(r'[^a-zA-Z0-9_-]', '_', d.get('period', 'MBR'))
    filename = f'LZB_LATAM_MBR_{period_safe}.pptx'
    return send_file(buf, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')


if __name__ == '__main__':
    app.run(debug=True, port=5050)

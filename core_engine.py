# -*- coding: utf-8 -*-
"""
密评报告对比核心引擎
- 解析docx文档的XML结构
- 提取附录D层级树 (D.x → 子类型 → 测评对象 → 符合程度表格)
- 对比两份文档的符合程度差异

依赖: python-docx, lxml (用于XML解析)
"""

import os, re, zipfile, xml.etree.ElementTree as ET


# ── XML Namespace ───────────────────────
WNS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WT  = WNS + 't'
WTR = WNS + 'tr'
WTC = WNS + 'tc'
WP  = WNS + 'p'
WPPr = WNS + 'pPr'
WPStyle = WNS + 'pStyle'
WNumPr = WNS + 'numPr'
WNumId = WNS + 'numId'
WTbl = WNS + 'tbl'
WBody = WNS + 'body'
Wval = WNS + 'val'


# ── Known Section Names ─────────────────
KNOWN_D_SECTIONS = [
    '安全物理环境', '安全通信网络', '安全区域边界', '安全计算环境',
    '安全管理中心', '安全管理制度', '安全管理机构', '安全管理人员',
    '安全建设管理', '安全运维管理',
]


# ==================== XML Parsing ====================

def get_text_from_cell(cell):
    texts = []
    for t in cell.iter(WT):
        if t.text:
            texts.append(t.text)
    return ''.join(texts).strip()


def parse_table_rows(table):
    rows = []
    for tr in table.iter(WTR):
        cells = []
        for tc in tr.iter(WTC):
            cells.append(get_text_from_cell(tc))
        if cells:
            rows.append(cells)
    return rows


def extract_docx_elements(docx_path):
    """
    Extract all paragraphs and tables from docx in document order.
    
    Returns: list of element dicts:
      - type='p': {text, style, numId}
      - type='tbl': {hdr, data, nr, comply}  (comply=True if header has '符合程度')
    """
    elements = []
    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            body = root.find(WBody)

            global_idx = 0
            for elem in body:
                tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

                # ---- Paragraph ----
                if tag == 'p':
                    texts = []
                    pPr = elem.find(WPPr)
                    style_name = ''
                    numId = ''
                    if pPr is not None:
                        pStyle = pPr.find(WPStyle)
                        if pStyle is not None:
                            style_name = pStyle.get(Wval, '')
                        numPr = pPr.find(WNumPr)
                        if numPr is not None:
                            numIdEl = numPr.find(WNumId)
                            if numIdEl is not None:
                                numId = numIdEl.get(Wval, '')

                    for t in elem.iter(WT):
                        if t.text:
                            texts.append(t.text)
                    full_text = ''.join(texts).strip()

                    if full_text:
                        elements.append({
                            'type': 'p',
                            'text': full_text,
                            'style': style_name,
                            'numId': numId,
                        })
                        global_idx += 1

                # ---- Table ----
                elif tag == 'tbl':
                    rows = parse_table_rows(elem)
                    if not rows:
                        continue
                    
                    header_str = '|'.join(rows[0]) if rows else ''
                    
                    elements.append({
                        'type': 'tbl',
                        'nr': len(rows),
                        'hdr': [c.strip() for c in rows[0]] if rows else [],
                        'data': [r for r in rows[1:]],  # data rows only
                        'comply': '符合程度' in header_str,
                    })
                    global_idx += 1

    return elements


# ==================== Locate Appendix D ====================

def find_appendix_d(elems):
    """
    Find the index of "单项测评结果记录" paragraph that starts Appendix D.
    Returns -1 if not found.
    """
    best = -1
    for i, el in enumerate(elems):
        if el.get('text') == '单项测评结果记录' or el.get('text') == '单项测评结果':
            # Prefer later occurrences (skip TOC)
            if i > len(elems) * 0.3:
                best = i
    return best


# ==================== Hierarchy Builder ====================

def build_d_hierarchy(elems, d_start_idx):
    """
    Build the D.x.y.z hierarchy tree from elements starting at d_start_idx.
    
    Structure:
      D.x 大章节 (Section)
        ├── 子类型 Sub (安全通用要求/云计算扩展要求)
        │     └── 测评对象 Object (a) 云平台机房 / b) 堡垒机 ...
        │           └── 表格 Table (符合程度表)
        │                 └── Items (控制点|指标|记录|符合程度)
    """
    sections = []
    sec = None
    sub = None
    obj = None
    
    sec_names = set(KNOWN_D_SECTIONS)
    i = d_start_idx
    
    while i < len(elems):
        el = elems[i]
        
        # ===== Handle Tables =====
        if el['type'] == 'tbl' and el['comply']:
            tbl_info = {
                'ei': i,
                'hdr': el['hdr'],
                'items': [],
                'nr_data': len(el['data']),
            }
            
            for ri, row in enumerate(el['data']):
                ncol = len(row)
                item = {
                    'ctrl': row[0].strip() if ncol > 0 else '',
                    'indi': row[1].strip() if ncol > 1 else '',
                    'rec': row[2].strip() if ncol > 2 else '',
                    'cpl': row[3].strip() if ncol > 3 else (row[2].strip() if ncol > 2 else ''),
                }
                
                # Clean up control point text (remove leading numbers like "a)")
                if item['ctrl']:
                    m = re.match(r'^([a-z][\)\.\．]|（[a-z]）|\(\d+\)|[\u2460-\u2473])\s*(.*)', item['ctrl'])
                    if m and m.group(2).strip():
                        item['ctrl'] = m.group(2).strip()
                
                tbl_info['items'].append(item)
            
            # Attach to nearest parent
            if obj is not None:
                obj.setdefault('tables', []).append(tbl_info)
            elif sub is not None:
                sub.setdefault('tables', []).append(tbl_info)
            elif sec is not None:
                sec.setdefault('tables', []).append(tbl_info)
            
            i += 1
            continue
        
        # Skip non-paragraphs
        if el['type'] != 'p':
            i += 1
            continue
        
        t = el['text']
        
        # Stop condition
        if re.match(r'^附录[E-Z]', t) or re.match(r'^附\s*录\s*[E-Z]', t):
            break
        
        # Skip "单项测评结果记录"
        if t in ('单项测评结果记录', '单项测评结果'):
            i += 1
            continue
        
        # ===== Section Title =====
        m_d = re.match(r'^([Dd]\.\d+(?:\.\d+)?)\s+(.{3,})$', t)
        is_known = t.strip() in sec_names
        
        if m_d or is_known:
            if m_d:
                sid = m_d.group(1).upper().rstrip('.')
                sname = m_d.group(2).strip()
                if re.match(r'^\d+$', sname):
                    i += 1; continue
            else:
                sname = t.strip()
                sid = f"D.{len(sections)+1}"
            
            sec = {'id': sid, 'name': sname, 'subs': [], 'objects': [], 'tables': []}
            sections.append(sec)
            sub = None; obj = None
            i += 1; continue
        
        # ===== Sub-type title =====
        sub_kws = ['安全通用要求部分', '云计算安全扩展要求部分', 
                   '安全扩展要求部分', '其它安全要求部分', '其他安全要求部分']
        
        if sec is not None and any(kw in t for kw in sub_kws):
            sub_id_base = f"{sec['id']}"
            if '通用' in t and '扩展' not in t:
                sub_id = f"{sub_id_base}.1"
            elif '其它' in t or '其他' in t:
                sub_id = f"{sub_id_base}.3"
            else:
                sub_id = f"{sub_id_base}.2"
            
            sub = {'id': sub_id, 'name': t.strip(), 'objects': [], 'tables': []}
            sec['subs'].append(sub)
            obj = None
            i += 1; continue
        
        # ===== Object title: a), b), (1), etc. =====
        m_obj = re.match(r'^([a-z][\)\.\．]|（[a-z]）|\(\d+\)|[\u2460-\u2473])\s*(.{2,})$', t)
        if m_obj and sec is not None:
            prefix = m_obj.group(1).strip()
            oname = m_obj.group(2).strip()
            
            # Build object ID like D.4.1.a or D.4.1.1
            if sub:
                oid = f"{sub['id']}.{prefix.rstrip(')）.')}"
            else:
                oid = f"{sec['id']}.{prefix.rstrip(')）。')}"
            
            obj = {'name': oname, 'prefix': prefix, 'id': oid, 'tables': []}
            sec['objects'].append(obj)
            i += 1; continue
        
        i += 1
    
    return sections


# ==================== Flatten to items ====================

def flatten_section_items(section):
    """Extract all items from a section into flat list."""
    items = []
    
    # Items from objects' tables
    for obj in section.get('objects', []):
        for tbl in obj.get('tables', []):
            for item in tbl.get('items', []):
                item['_obj_name'] = obj.get('name', '')
                items.append(item)
    
    # Items directly under section (no object grouping)
    for tbl in section.get('tables', []):
        for item in tbl.get('items', []):
            item['_obj_name'] = section.get('name', '')
            items.append(item)
    
    return items


# ==================== Comparison Engine ====================

COMPLY_ORDER = ['不符合', '部分符合', '符合', '不适用']

def classify_comply_change(cb_val, zg_val):
    """Classify the direction of compliance change."""
    cb = (cb_val or '').strip().replace('\n', ' ')
    zg = (zg_val or '').strip().replace('\n', ' ')
    
    # One side empty
    if not zg: return 'deleted'
    if not cb: return 'added'
    
    if cb == zg: return None
    
    # Special cases
    if '不适用' in zg and '不适用' not in cb: return 'to_na'
    if '不适用' in cb and '不适用' not in zg: return 'from_na'
    
    # Order-based comparison
    ci = next((i for i, x in enumerate(COMPLY_ORDER) if x in cb), -1)
    zi = next((i for i, x in enumerate(COMPLY_ORDER) if x in zg), -1)
    
    if ci >= 0 and zi >= 0:
        if zi < ci: return 'downgrade'
        if zi > ci: return 'upgrade'
    
    # Fallback: string length heuristic (longer usually means worse)
    if len(zg) > len(cb) + 4: return 'downgrade'
    return 'upgrade'


def compare_hierarchies(hier_a, hier_b):
    """
    Compare two hierarchies and produce diff-ready structure for frontend.
    
    Returns dict with:
      - stats: summary statistics
      - sections: list of comparison sections with flattened items
    """
    stats = {
        'total_items': 0,
        'comply_changed': 0,
        'cb_only': 0,
        'zg_only': 0,
        'downgrades': 0,
        'upgrades': 0,
        'to_na_count': 0,
        'from_na_count': 0,
    }
    
    sections_out = []
    
    # Build lookup by section ID for both
    def build_sec_map(hier):
        m = {}
        for s in hier:
            m[s['id']] = s
        return m
    
    map_a = build_sec_map(hier_a)
    map_b = build_sec_map(hier_b)
    
    # All unique section IDs (union)
    all_sec_ids = set(list(map_a.keys()) + list(map_b.keys()))
    
    for sid in sorted(all_sec_ids, key=sec_sort_key):
        sec_a = map_a.get(sid)
        sec_b = map_b.get(sid)
        
        # Get name
        name = sec_a['name'] if sec_a else (sec_b['name'] if sec_b else sid)
        
        # Flatten items
        items_a = flatten_section_items(sec_a) if sec_a else []
        items_b = flatten_section_items(sec_b) if sec_b else []
        
        # Build B lookup by (control_point, indicator)
        b_lookup = {}
        for idx, it in enumerate(items_b):
            key = (it['ctrl'][:60], it['indi'][:80])
            if key not in b_lookup:
                b_lookup[key] = []
            b_lookup[key].append(idx)
        
        # Match and compare
        compared = []
        used_b_indices = set()
        
        for ia_idx, item_a in enumerate(items_a):
            key = (item_a['ctrl'][:60], item_a['indi'][:80])
            
            b_match = None
            b_idx = -1
            
            if key in b_lookup:
                for bi in b_lookup[key]:
                    if bi not in used_b_indices:
                        b_idx = bi
                        b_match = items_b[bi]
                        used_b_indices.add(bi)
                        break
            
            if b_match:
                diff_type = classify_comply_change(item_a['cpl'], b_match['cpl'])
                has_diff = diff_type is not None
                
                entry = {
                    'ctrl': item_a['ctrl'],
                    'indi': item_a['indi'],
                    'cb_rec': item_a['rec'],
                    'cb_comply': item_a['cpl'],
                    'zg_comply': b_match['cpl'],
                    'diff_type': diff_type,
                    'has_diff': has_diff,
                    '_obj_name': item_a.get('_obj_name', ''),
                }
                
                # Update stats
                stats['total_items'] += 1
                if has_diff:
                    stats['comply_changed'] += 1
                    if diff_type == 'downgrade': stats['downgrades'] += 1
                    elif diff_type == 'upgrade': stats['upgrades'] += 1
                    elif diff_type == 'to_na': stats['to_na_count'] += 1
                    elif diff_type == 'from_na': stats['from_na_count'] += 1
                
                compared.append(entry)
            else:
                # Only in A (deleted in B)
                stats['total_items'] += 1
                stats['cb_only'] += 1
                _a_oname = item_a.get('_obj_name', '')
                compared.append({
                    'ctrl': item_a['ctrl'], 'indi': item_a['indi'],
                    'cb_rec': item_a['rec'], 'cb_comply': item_a['cpl'],
                    'zg_comply': '', 'diff_type': 'deleted', 'has_diff': True,
                    '_obj_name': _a_oname,
                })
                _ = _obj_name  # suppress warning
        
        # Items only in B (added)
        for ib_idx, item_b in enumerate(items_b):
            if ib_idx not in used_b_indices:
                stats['total_items'] += 1
                stats['zg_only'] += 1
                b_oname = item_b.get('_obj_name', '')
                compared.append({
                    'ctrl': item_b['ctrl'], 'indi': item_b['indi'],
                    'cb_rec': '', 'cb_comply': '',
                    'zg_comply': item_b['cpl'], 'diff_type': 'added', 'has_diff': True,
                    '_obj_name': b_oname,
                })
        
        # Group compared items by object name for display
        obj_groups = {}
        for item in compared:
            oname = item.pop('_obj_name', '') or '(未分类)'
            obj_groups.setdefault(oname, []).append(item)
        
        sec_objects = [
            {'name': k, 'items': v} for k, v in obj_groups.items()
        ]
        
        cb_total = len(items_a)
        zg_total = len(items_b)
        
        sec_out = {
            'id': sid,
            'name': name,
            'objects': sec_objects,
            'cb_count': cb_total,
            'zg_count':zg_total,
        }
        sections_out.append(sec_out)
    
    stats['sections_count'] = len(sections_out)
    
    return {
        'stats': stats,
        'sections': sections_out,
    }


def sec_sort_key(sid):
    """Sort section IDs like D.1, D.2 ... D.10"""
    parts = sid.replace('.', '').lstrip('D').lstrip('d')
    try:
        return int(parts) if parts else 999
    except ValueError:
        return 999

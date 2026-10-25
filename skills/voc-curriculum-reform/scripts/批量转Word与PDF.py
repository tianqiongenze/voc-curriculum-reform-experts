#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量转 Word 与 PDF（院校体例）
=============================
输入：材料包目录（含 .md / .html）
输出：同名 .docx 与 .pdf，与院校目录惯例一致（docx 与 pdf 同名并存）

转换链：
  .md   --pandoc(院校参考模板)--> .docx --后处理--> .docx --LibreOffice--> .pdf
  .html --pandoc(院校参考模板)--> .docx --后处理--> .docx --LibreOffice--> .pdf

后处理三项（对齐院校 docx 实测样式）：
  1. 引用块  pandoc 的 "Block Text" → 院校自定义样式 "Blockquote"（楷体灰字）
  2. 分隔线  pandoc 3.x 的 VML <v:rect> → HorizontalRule 样式 + CCCCCC 底边框
  3. 表头    每个表格首行加粗 + <w:tblHeader/>（跨页重复表头；院校原文件未设，属改进项）

用法：
  python3 批量转Word与PDF.py <材料包目录> [--only 文件名子串]
依赖：pandoc ≥3、LibreOffice（soffice）、院校Word参考模板.docx
"""
import os, re, shutil, subprocess, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REF = os.path.normpath(os.path.join(HERE, '..', '..', 'templates', '院校Word参考模板.docx'))
LO = '/Applications/LibreOffice.app/Contents/MacOS/soffice'
LO_PROFILE = 'file:///tmp/lo_profile_voc'
PY = '/Users/joelgong/.workbuddy/binaries/python/versions/3.13.12/bin/python3'

def run(cmd, **kw):
    r = subprocess.run(cmd, capture_output=True, text=True, **kw)
    if r.returncode != 0:
        print('  ⚠ 命令退出', r.returncode, ':', ' '.join(cmd[:3]), (r.stderr or r.stdout)[:300])
    return r

# ---------------- 后处理 ----------------
HR_XML = ('<w:p><w:pPr><w:pStyle w:val="HorizontalRule"/>'
          '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="CCCCCC"/></w:pBdr>'
          '</w:pPr></w:p>')

# 页脚（"— 页码 —"居中，样文实测）；人培方案注入，供封面/目录/正文各节引用
FOOTER_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
 '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
 '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
 '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman"/></w:rPr><w:t xml:space="preserve">— </w:t></w:r>'
 '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
 '<w:r><w:instrText xml:space="preserve"> PAGE  \\* MERGEFORMAT </w:instrText></w:r>'
 '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
 '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman"/></w:rPr><w:t xml:space="preserve"> —</w:t></w:r>'
 '</w:p></w:ftr>')
FTR_ID = 'rIdFtrVoc1'

# 横版节（A4 横向，四边 1440）与纵版节的 sectPr
# 注意：带页脚引用的变体只给"人培方案"用（rebuild_rp_cover_toc 会注入 footer1.xml）；
# 课标等其余文件按院校惯例不设页码，用无 footer 版，否则引用不存在的部件会导致 LibreOffice 拒绝加载。
SECT_PORTRAIT = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1440" w:right="1797" w:bottom="1440" w:left="1797" '
                 'w:header="720" w:footer="720" w:gutter="0"/>'
                 '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>')
SECT_PORTRAIT_FTR = ('<w:sectPr><w:footerReference w:type="default" r:id="%s"/>'
                     '<w:pgSz w:w="11906" w:h="16838"/>'
                     '<w:pgMar w:top="1440" w:right="1797" w:bottom="1440" w:left="1797" '
                     'w:header="720" w:footer="720" w:gutter="0"/>'
                     '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>' % FTR_ID)
SECT_LANDSCAPE = ('<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
                  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
                  'w:header="720" w:footer="720" w:gutter="0"/>'
                  '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>')

def _sect_break(sect_xml):
    """承载节切换的空段（1pt 字号、无段距，视觉上近不可见）。"""
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>' + sect_xml + '</w:pPr></w:p>')

def _rpr(ea=None, sz=None, b=False):
    """构造 run 属性：eastAsia 字体 + 字号（半磅）。"""
    f = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"'
    if ea: f += ' w:eastAsia="%s"' % ea
    f += '/>'
    s = '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz) if sz else ''
    bb = '<w:b/><w:bCs/>' if b else ''
    return '<w:rPr>' + f + bb + s + '</w:rPr>'

# —— 人培方案封面/目录重构（学校标准体例，样文实测）——
RP_COVER_FIELDS = [('专业名称：', '人工智能数据工程技术'), ('专业代码：', '510217'),
                   ('专业大类：', '电子与信息大类（51）'), ('学制：', '三年制'),
                   ('专业负责人：', '（待定）'), ('企业负责人：', '（待定）'),
                   ('学院负责人：', '（待定）'), ('制订日期：', '2026 年 10 月')]

def _p(text, ea=None, sz=None, jc=None, b=False, spacing_after=None):
    ppr = '<w:pPr>'
    if jc: ppr += '<w:jc w:val="%s"/>' % jc
    if spacing_after is not None: ppr += '<w:spacing w:after="%d"/>' % spacing_after
    ppr += _rpr(ea, sz, b) + '</w:pPr>'
    if not text: return '<w:p>' + ppr + '</w:p>'
    runs = ''
    for i, seg in enumerate(text.split('\n')):
        if i: runs += '<w:r>' + _rpr(ea, sz, b) + '<w:br/></w:r>'
        runs += '<w:r>' + _rpr(ea, sz, b) + '<w:t xml:space="preserve">' + seg + '</w:t></w:r>'
    return '<w:p>' + ppr + runs + '</w:p>'

def _toc_entry(title, level, page_ph):
    """一条目录项：标题 + 右制表位点线 + 页码占位（二级回填）。章级黑体、节级楷体、15pt。"""
    ea = '方正黑体_GBK' if level == 1 else '方正楷体_GBK'
    tabs = ('<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8306"/></w:tabs>')
    ppr = ('<w:pPr>' + tabs + '<w:spacing w:after="60" w:line="240" w:lineRule="auto"/>'
           + _rpr(ea, 30) + '</w:pPr>')
    runs = ('<w:r>' + _rpr(ea, 30) + '<w:t xml:space="preserve">' + title + '</w:t></w:r>'
            '<w:r>' + _rpr(ea, 30) + '<w:tab/></w:r>'
            '<w:r>' + _rpr(ea, 30) + '<w:t xml:space="preserve">PGNUM:' + title + '</w:t></w:r>')
    return '<w:p>' + ppr + runs + '</w:p>'

def rebuild_rp_cover_toc(docx_path, entries, body_title, body_sub):
    """把人培方案 docx 的封面+目录区重构为学校标准体例；注入页脚；正文 H2→H1、H3→H2。
    entries: [(title, level)]；返回占位页码字典（key=title, value='PGNUM:title' 文本）。"""
    tmp = docx_path + '.unz'
    shutil.rmtree(tmp, ignore_errors=True); os.makedirs(tmp)
    with zipfile.ZipFile(docx_path) as z: z.extractall(tmp)
    dp = os.path.join(tmp, 'word', 'document.xml')
    s = open(dp, encoding='utf-8').read()

    # 1) 定位正文起点（第一个 Heading2 且文本以"一、"开头）
    m = re.search(r'<w:p\b[^>]*>(?:(?!</w:p>).)*?<w:pStyle w:val="Heading2"/>(?:(?!</w:p>).)*?</w:p>', s, re.S)
    start = None
    for pm in re.finditer(r'<w:p\b.*?</w:p>', s, re.S):
        if 'w:val="Heading2"' in pm.group(0) and ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', pm.group(0))).startswith('一、'):
            start = pm.start(); break
    if start is None:
        shutil.rmtree(tmp, ignore_errors=True); return None

    # 2) 构造学校标准封面 + 目录 + 正文大标题
    cover = [_p('重庆现代制造职业学院', '方正黑体_GBK', 44, 'center', spacing_after=0),
             _p('', spacing_after=0), _p('', spacing_after=0),
             _p(body_title, '方正小标宋_GBK', 52, 'center', spacing_after=0),
             _p(body_sub, '楷体', 32, 'center', spacing_after=0),
             _p('', spacing_after=0), _p('', spacing_after=0), _p('', spacing_after=0)]
    # 封面信息表：2 列 2667/4815，8 行，每格下边线，标签黑体右对齐/值仿宋居中，15pt
    rows_xml = ''
    for label, value in RP_COVER_FIELDS:
        rows_xml += ('<w:tr><w:tc><w:tcPr><w:tcW w:w="2667" w:type="dxa"/>'
                     '<w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>'
                     '</w:tcPr>' + _p(label, '方正黑体_GBK', 30, 'right', spacing_after=0) + '</w:tc>'
                     '<w:tc><w:tcPr><w:tcW w:w="4815" w:type="dxa"/>'
                     '<w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>'
                     '</w:tcPr>' + _p(value, '方正仿宋_GBK', 30, 'center', spacing_after=0) + '</w:tc></w:tr>')
    cover_tbl = ('<w:tbl><w:tblPr><w:tblW w:w="7482" w:type="dxa"/><w:jc w:val="center"/>'
                 '<w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" '
                 'w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
                 '<w:tblGrid><w:gridCol w:w="2667"/><w:gridCol w:w="4815"/></w:tblGrid>'
                 + rows_xml + '</w:tbl>')
    # 目录：居中标题 + 真 TOC 域（begin/instr/separate 在首条，end 在末条，缓存为点线条目）
    toc = [_p('目　录', '方正小标宋_GBK', 36, 'center', spacing_after=0), _p('', spacing_after=0)]
    n = len(entries)
    for i, (title, level) in enumerate(entries):
        entry = _toc_entry(title, level, None)
        if i == 0:
            entry = entry.replace('</w:pPr>',
                '</w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
                '<w:r><w:fldChar w:fldCharType="separate"/></w:r>', 1)
        if i == n - 1:
            entry = entry.replace('</w:p>', '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
        toc.append(entry)
    page_break = ('<w:p><w:pPr><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>'
                  '<w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>')
    body_head = [_p(body_title, '方正小标宋_GBK', 44, 'center', spacing_after=0),
                 _p(body_sub, '方正仿宋_GBK', 32, 'center', spacing_after=0)]
    new_head = ''.join(cover) + cover_tbl + _sect_break(SECT_PORTRAIT_FTR) + ''.join(toc) \
               + page_break + ''.join(body_head)
    # 保留 XML 声明＋<w:document>根＋<w:body>开标签，只替换 body 内的旧封面/目录区
    body_open = s.find('<w:body>')
    if body_open < 0 or start <= body_open:
        shutil.rmtree(tmp, ignore_errors=True); return None
    body_open += len('<w:body>')
    s = s[:body_open] + new_head + s[start:]

    # 3) 正文标题层级：H2→Heading1、H3→Heading2（保证 TOC 域更新后层级正确）
    s = s.replace('<w:pStyle w:val="Heading2"/>', '<w:pStyle w:val="Heading1"/>')
    s = s.replace('<w:pStyle w:val="Heading3"/>', '<w:pStyle w:val="Heading2"/>')

    # 4) body 末尾 sectPr 挂页脚引用
    s = re.sub(r'(<w:sectPr(?:(?!</w:sectPr>).)*?)(<w:pgSz)', r'\1<w:footerReference w:type="default" r:id="%s"/>\2' % FTR_ID, s, count=1, flags=re.S)
    open(dp, 'w', encoding='utf-8').write(s)

    # 5) 注入 footer part + rels + content types
    ftp = os.path.join(tmp, 'word', 'footer1.xml')
    open(ftp, 'w', encoding='utf-8').write(FOOTER_XML)
    rp = os.path.join(tmp, 'word', '_rels', 'document.xml.rels')
    rels = open(rp, encoding='utf-8').read()
    if FTR_ID not in rels:
        rels = rels.replace('</Relationships>',
            '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>' % FTR_ID)
        open(rp, 'w', encoding='utf-8').write(rels)
    ctp = os.path.join(tmp, '[Content_Types].xml')
    ct = open(ctp, encoding='utf-8').read()
    if 'footer1.xml' not in ct:
        ct = ct.replace('</Types>',
            '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
        open(ctp, 'w', encoding='utf-8').write(ct)

    # 6) 重包
    os.remove(docx_path)
    zf = zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED)
    ct2 = os.path.join(tmp, '[Content_Types].xml')
    zf.write(ct2, '[Content_Types].xml')
    for root, _, files in os.walk(tmp):
        for f in files:
            full = os.path.join(root, f); rel = os.path.relpath(full, tmp)
            if rel != '[Content_Types].xml': zf.write(full, rel)
    zf.close(); shutil.rmtree(tmp, ignore_errors=True)
    return {t: 'PGNUM:' + t for t, _ in entries}

def fill_toc_pages(docx_path, page_map):
    """把目录缓存里的 PGNUM:标题 占位替换为真实页码。page_map: {标题: 页码}"""
    tmp = docx_path + '.unz'
    shutil.rmtree(tmp, ignore_errors=True); os.makedirs(tmp)
    with zipfile.ZipFile(docx_path) as z: z.extractall(tmp)
    dp = os.path.join(tmp, 'word', 'document.xml')
    s = open(dp, encoding='utf-8').read()
    n = 0
    for title, page in page_map.items():
        old = '<w:t xml:space="preserve">PGNUM:%s</w:t>' % title
        if old in s:
            s = s.replace(old, '<w:t xml:space="preserve">%s</w:t>' % page); n += 1
    open(dp, 'w', encoding='utf-8').write(s)
    os.remove(docx_path)
    zf = zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED)
    zf.write(os.path.join(tmp, '[Content_Types].xml'), '[Content_Types].xml')
    for root, _, files in os.walk(tmp):
        for f in files:
            full = os.path.join(root, f); rel = os.path.relpath(full, tmp)
            if rel != '[Content_Types].xml': zf.write(full, rel)
    zf.close(); shutil.rmtree(tmp, ignore_errors=True)
    return n

def extract_toc_pages(pdf_path, entries, offset=None):
    """从 PDF 提取每个目录标题的页码。动态测定正文起始物理页：跳过封面页与
    含点线（目录条目）的目录页——目录可能占多页，不能写死 offset。"""
    import subprocess as sp
    info = sp.run(['pdfinfo', pdf_path], capture_output=True, text=True).stdout
    total = int(re.search(r'Pages:\s+(\d+)', info).group(1))
    norms = {}
    for p in range(1, total + 1):
        txt = sp.run(['pdftotext', '-f', str(p), '-l', str(p), pdf_path, '-'],
                     capture_output=True, text=True).stdout
        norms[p] = re.sub(r'\s+', '', txt)
    body_start = None
    for p in range(2, total + 1):
        if '....' not in norms[p] and '目录' not in norms[p]:
            body_start = p; break
    if body_start is None: body_start = 3
    off = body_start - 1  # 正文第 1 页记为页码 1
    pages = {}
    for p in range(body_start, total + 1):
        for title, _ in entries:
            if title not in pages and re.sub(r'\s+', '', title) in norms[p]:
                pages[title] = str(p - off)
    return pages

def _set_run_size(run_xml, half_pt):
    """把单个 run 的字号改为 half_pt（半磅值）。"""
    if '<w:sz ' in run_xml:
        run_xml = re.sub(r'<w:sz w:val="\d+"/>', '<w:sz w:val="%d"/>' % half_pt, run_xml)
        run_xml = re.sub(r'<w:szCs w:val="\d+"/>', '<w:szCs w:val="%d"/>' % half_pt, run_xml)
        if '<w:sz ' not in run_xml:
            run_xml = run_xml.replace('</w:rPr>', '<w:sz w:val="%d"/></w:rPr>' % half_pt, 1)
        if '<w:szCs ' not in run_xml:
            run_xml = run_xml.replace('</w:rPr>', '<w:szCs w:val="%d"/></w:rPr>' % half_pt, 1)
        return run_xml
    rpr = re.search(r'<w:rPr>.*?</w:rPr>|<w:rPr/>', run_xml, re.S)
    ins = '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (half_pt, half_pt)
    if rpr:
        r = rpr.group(0)
        if r == '<w:rPr/>':
            return run_xml.replace('<w:rPr/>', '<w:rPr>' + ins + '</w:rPr>', 1)
        return run_xml.replace(r, r.replace('</w:rPr>', ins + '</w:rPr>', 1), 1)
    return run_xml.replace('<w:r>', '<w:r><w:rPr>' + ins + '</w:rPr>', 1)

def postprocess(docx_path):
    """解包→改 document.xml→重包。返回改动计数。"""
    tmp = docx_path + '.unz'
    shutil.rmtree(tmp, ignore_errors=True)
    os.makedirs(tmp)
    with zipfile.ZipFile(docx_path) as z:
        z.extractall(tmp)
    p = os.path.join(tmp, 'word', 'document.xml')
    s = open(p, encoding='utf-8').read()
    n0 = len(s)

    # 1) Block Text → Blockquote（注意 pandoc 输出 <w:pStyle w:val="BlockText" />，斜杠前有空格）
    n_bq = len(re.findall(r'<w:pStyle w:val="BlockText"\s*/>', s))
    s = re.sub(r'<w:pStyle w:val="BlockText"\s*/>', '<w:pStyle w:val="Blockquote"/>', s)

    # 2) VML 分隔线 → HorizontalRule 样式段
    n_hr = len(re.findall(r'<w:p><w:r><w:pict><v:rect[^>]*/></w:pict></w:r></w:p>', s))
    s = re.sub(r'<w:p><w:r><w:pict><v:rect[^>]*/></w:pict></w:r></w:p>', HR_XML, s)

    # 3) Compact 残留清理（参考模板已删该样式，pandoc 仍会引用；去掉即回退 Normal，同院校文件）
    n_cp = len(re.findall(r'<w:pStyle w:val="Compact"\s*/>', s))
    s = re.sub(r'<w:p>\s*<w:pPr>\s*<w:pStyle w:val="Compact"\s*/>\s*</w:pPr>', '<w:p><w:pPr/>', s)
    s = re.sub(r'<w:pStyle w:val="Compact"\s*/>', '', s)

    # 4) 表格处理：首行加粗+跨页重复；≥9 列整体横版并降字号；列宽等比撑满所在节文本宽
    #    纵版文本宽 8312 twips（11906-1797×2）；横版 13958 twips（16838-1440×2）
    n_tbl = 0
    n_land = 0
    out, pos = [], 0
    for m in re.finditer(r'<w:tbl>.*?</w:tbl>', s, re.S):
        tbl = m.group(0)
        cols = len(re.findall(r'<w:gridCol', tbl))
        wide = cols >= 9
        # 列宽等比撑满
        gw = [int(x) for x in re.findall(r'<w:gridCol w:w="(\d+)"', tbl)]
        if gw:
            target = 13958 if wide else 8312
            f = target / float(sum(gw))
            if abs(f - 1) > 0.01:
                def scale_gc(mg):
                    return '<w:gridCol w:w="%d" />' % max(60, int(int(mg.group(1)) * f))
                tbl = re.sub(r'<w:gridCol w:w="(\d+)"\s*/>', scale_gc, tbl)
                def scale_tc(mc):
                    return '<w:tcW w:w="%d" w:type="dxa"/>' % max(60, int(int(mc.group(1)) * f))
                tbl = re.sub(r'<w:tcW w:w="(\d+)" w:type="dxa"\s*/>', scale_tc, tbl)
        tr = re.search(r'<w:tr\b[^>]*>', tbl)
        if tr:
            head_close = tbl.find('</w:tr>')
            head = tbl[tr.end():head_close]
            tr_open = tr.group(0)
            if '<w:trPr>' not in tbl[:tr.end() + 40]:
                new_open = tr_open + '<w:trPr><w:tblHeader/></w:trPr>'
            else:
                new_open = tr_open  # 已有 trPr 则不重复插
            def bold(mr):
                run = mr.group(0)
                if '<w:b/>' in run:
                    return run
                rpr = re.search(r'<w:rPr>.*?</w:rPr>', run, re.S)
                if rpr:
                    return run.replace(rpr.group(0), rpr.group(0).replace('<w:rPr>', '<w:rPr><w:b/>', 1), 1)
                return run.replace('<w:r>', '<w:r><w:rPr><w:b/></w:rPr>', 1)
            head = re.sub(r'<w:r\b[^>]*>.*?</w:r>', bold, head, flags=re.S)
            tbl = tbl[:tr.start()] + new_open + head + tbl[head_close:]
            n_tbl += 1
        if wide:  # 宽表：横版 + 降字号（≥13 列 8pt，9-12 列 9pt）
            half = 16 if cols >= 13 else 18
            tbl = re.sub(r'<w:r\b[^>]*>.*?</w:r>',
                         lambda mr: _set_run_size(mr.group(0), half), tbl, flags=re.S)
            out.append(s[pos:m.start()])
            out.append(_sect_break(SECT_PORTRAIT) + tbl + _sect_break(SECT_LANDSCAPE))
            pos = m.end()
            n_land += 1
            continue
        out.append(s[pos:m.start()]); out.append(tbl); pos = m.end()
    out.append(s[pos:])
    s = ''.join(out)

    open(p, 'w', encoding='utf-8').write(s)
    # 重包（[Content_Types].xml 放最前）
    os.remove(docx_path)
    zf = zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED)
    ct = os.path.join(tmp, '[Content_Types].xml')
    if os.path.exists(ct):
        zf.write(ct, '[Content_Types].xml')
    for root, _, files in os.walk(tmp):
        for f in files:
            full = os.path.join(root, f)
            rel = os.path.relpath(full, tmp)
            if rel != '[Content_Types].xml':
                zf.write(full, rel)
    zf.close()
    shutil.rmtree(tmp, ignore_errors=True)
    return {'blockquote': n_bq, 'hr': n_hr, 'table': n_tbl, 'landscape': n_land,
            'compact': n_cp, 'changed': len(s) != n0}

# ---------------- 主流程 ----------------
def rp_toc_entries(md_path):
    """从人培方案 md 提取目录条目：[(title, level)]，章=1、节=（一）=2。"""
    entries = []
    for line in open(md_path, encoding='utf-8'):
        m = re.match(r'^## ([一二三四五六七八九十]+、.+)$', line.strip())
        if m: entries.append((m.group(1), 1)); continue
        m = re.match(r'^### (（[一二三四五六七八九十]+）.+)$', line.strip())
        if m: entries.append((m.group(1), 2))
    return entries

def convert(pkg):
    if not os.path.exists(REF):
        print('❌ 缺院校参考模板，先运行 构建院校Word参考模板.py')
        sys.exit(1)
    targets = sorted(f for f in os.listdir(pkg)
                     if f.endswith(('.md', '.html')))
    docs = []
    rp_doc = None; rp_entries = None
    for f in targets:
        src = os.path.join(pkg, f)
        base = os.path.splitext(f)[0]
        dst = os.path.join(pkg, base + '.docx')
        fmt = 'markdown' if f.endswith('.md') else 'html'
        r = run(['pandoc', src, '-f', fmt, '-t', 'docx', '-o', dst,
                 '--reference-doc', REF])
        if not os.path.exists(dst):
            print('  ❌ 转换失败:', f)
            continue
        info = postprocess(dst)
        print('  ✓ %-52s → docx %6.1fKB｜引用块%d 分隔线%d 表%d（横版%d）'
              % (f[:52], os.path.getsize(dst) / 1024, info['blockquote'], info['hr'],
                 info['table'], info['landscape']))
        # 人培方案：封面/目录重构为学校标准体例（居中大标题＋信息表＋真 TOC 域＋页脚）
        if '人才培养方案' in f and f.endswith('.md'):
            entries = rp_toc_entries(src)
            if entries:
                title = '人工智能数据工程技术专业\n人才培养方案'
                sub = '（适用于2026级）'
                rebuild_rp_cover_toc(dst, entries, title, sub)
                rp_doc, rp_entries = dst, entries
                print('    └ 封面/目录已按学校标准重构：%d 条目录（章%d／节%d）'
                      % (len(entries), sum(1 for _, l in entries if l == 1),
                         sum(1 for _, l in entries if l == 2)))
        docs.append(dst)
    # LibreOffice 批量转 PDF（一次进程吃多个文件）；先删旧 PDF，防止"按存在计数"误报
    if docs:
        for d in docs:
            old_pdf = os.path.splitext(d)[0] + '.pdf'
            if os.path.exists(old_pdf): os.remove(old_pdf)
        r = run([LO, '--headless', '--norestore', '--nolockcheck',
                 '-env:UserInstallation=' + LO_PROFILE,
                 '--convert-to', 'pdf', '--outdir', pkg] + docs)
        # 按磁盘实数：md 与 html 同名时后者覆盖前者，唯一文件名可能少于源文件数
        uniq = sorted({os.path.splitext(d)[0] for d in docs})
        ok = sum(1 for b in uniq if os.path.exists(b + '.pdf'))
        print('  ✓ PDF 生成 %d/%d（唯一文件名 %d 个；md 与 html 同名时以 md 为权威源）'
              % (ok, len(uniq), len(uniq)))
        if ok < len(uniq):
            print('  ⚠ 部分失败，stderr:', (r.stderr or '')[:300])
    # 人培方案第二遍：从 PDF v1 提取各章真实页码，回填目录缓存，重转 PDF
    if rp_doc and rp_entries:
        pdf = os.path.splitext(rp_doc)[0] + '.pdf'
        if os.path.exists(pdf):
            page_map = extract_toc_pages(pdf, rp_entries)
            filled = fill_toc_pages(rp_doc, page_map)
            if os.path.exists(pdf): os.remove(pdf)  # 删旧 PDF，强制重生成
            r2 = run([LO, '--headless', '--norestore', '--nolockcheck',
                      '-env:UserInstallation=' + LO_PROFILE,
                      '--convert-to', 'pdf', '--outdir', os.path.dirname(rp_doc), rp_doc])
            ok2 = os.path.exists(pdf) and os.path.getmtime(pdf) > os.path.getmtime(rp_doc)
            miss = [t for t, _ in rp_entries if t not in page_map]
            print('  ✓ 目录页码回填 %d/%d 条（页码＝物理页−2，自封面起算）%s'
                  % (filled, len(rp_entries),
                     ('；未定位：' + '、'.join(miss[:3]) if miss else '')
                     if ok2 else '；⚠ PDF 重转失败：' + (r2.stderr or '')[:200]))
    # 汇总
    print('\n=== 产出汇总 ===')
    for f in sorted(os.listdir(pkg)):
        if f.endswith(('.docx', '.pdf')):
            print('  %-58s %8.1f KB' % (f, os.path.getsize(os.path.join(pkg, f)) / 1024))

if __name__ == '__main__':
    pkg = sys.argv[1] if len(sys.argv) > 1 else '.'
    convert(pkg)

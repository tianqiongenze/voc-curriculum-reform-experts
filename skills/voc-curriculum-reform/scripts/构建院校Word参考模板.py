#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
构建院校版 Word 参考模板（院校Word参考模板.docx）
================================================
来源：pandoc 默认 reference.docx，按院校既有 docx（LF4 课标、Table1 V80 等）实测样式打补丁。

院校版式实测值（2026-10-24 解包院校 docx 得出）：
  页面   A4 11906×16838 twips；页边距 上1440 右1797 下1440 左1797；页眉720 页脚720
  正文   宋体 11pt（sz22）黑色；段后120，行距276
  H1     加粗 14pt 色365F91，段前480；H2 加粗13pt 色4F81BD，段前200；H3 加粗11pt 色4F81BD，段前200
  Title  26pt 色17365D，底部边框4F81BD
  引用块 楷体 11pt 色444444（院校元数据行的承载样式）
  表格   TableGrid 全边框 single sz4，单元格边距左右108；表格居中
  无页眉页脚内容（院校文件本身未设页码）

用法：
  python3 构建院校Word参考模板.py [输出路径]
  默认输出：与本脚本同级 ../templates/院校Word参考模板.docx
"""
import os, re, shutil, subprocess, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.normpath(
    os.path.join(HERE, '..', '..', 'templates', '院校Word参考模板.docx'))

# ---------- 1. 取 pandoc 默认参考档并解包 ----------
base = os.path.join('/tmp', 'pandoc_ref_base.docx')
subprocess.run(['pandoc', '--print-default-data-file', 'reference.docx'],
               stdout=open(base, 'wb'), check=True)
work = '/tmp/_ref_build'
shutil.rmtree(work, ignore_errors=True)
os.makedirs(work)
with zipfile.ZipFile(base) as z:
    z.extractall(work)

S = open(os.path.join(work, 'word', 'styles.xml'), encoding='utf-8').read()

def replace_style(s, sid, new_xml):
    """整体替换某 styleId 的 <w:style> 块；不存在则插入到 </w:styles> 前。"""
    pat = re.compile(r'<w:style [^>]*w:styleId="%s"[^>]*>.*?</w:style>' % re.escape(sid), re.S)
    if pat.search(s):
        return pat.sub(lambda m: new_xml, s, count=1)
    return s.replace('</w:styles>', new_xml + '</w:styles>')

# ---------- 2. docDefaults：主题字体 + 11pt + 段后200行距276 ----------
S = re.sub(r'<w:docDefaults>.*?</w:docDefaults>',
'''<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>''',
S, count=1, flags=re.S)

# ---------- 3. Normal：宋体 11pt 黑字 ----------
S = replace_style(S, 'Normal',
'''<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体"/><w:color w:val="000000"/><w:sz w:val="22"/></w:rPr></w:style>''')

# ---------- 4. Heading 1-3：院校蓝灰体系 ----------
for sid, lvl, sz, color, before in [('Heading1', 0, '28', '365F91', '480'),
                                    ('Heading2', 1, '26', '4F81BD', '200'),
                                    ('Heading3', 2, None, '4F81BD', '200')]:
    rpr_sz = '<w:sz w:val="%s"/><w:szCs w:val="%s"/>' % (sz, sz) if sz else ''
    S = replace_style(S, sid,
'''<w:style w:type="paragraph" w:styleId="%s"><w:name w:val="heading %d"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:link w:val="%sChar"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="%s" w:after="0"/><w:outlineLvl w:val="%d"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:color w:val="%s"/>%s</w:rPr></w:style>'''
        % (sid, lvl + 1, sid, before, lvl, color, rpr_sz))

# ---------- 5. Title：26pt 深蓝 + 底部边框 ----------
S = replace_style(S, 'Title',
'''<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:link w:val="TitleChar"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="4" w:color="4F81BD" w:themeColor="accent1"/></w:pBdr><w:spacing w:after="300" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:color w:val="17365D" w:themeColor="text2" w:themeShade="BF"/><w:spacing w:val="5"/><w:kern w:val="28"/><w:sz w:val="52"/><w:szCs w:val="52"/></w:rPr></w:style>''')

# ---------- 5.5 Heading4：斜体加粗（院校 LF4 课标实测） ----------
S = replace_style(S, 'Heading4',
'''<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:link w:val="Heading4Char"/><w:uiPriority w:val="9"/><w:semiHidden/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="0"/><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:i/><w:iCs/><w:color w:val="4F81BD" w:themeColor="accent1"/></w:rPr></w:style>''')

# ---------- 5.6 HorizontalRule：灰色细线（院校自定义样式） ----------
S = replace_style(S, 'HorizontalRule',
'''<w:style w:type="paragraph" w:customStyle="1" w:styleId="HorizontalRule"><w:name w:val="HorizontalRule"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:color w:val="CCCCCC"/></w:rPr></w:style>''')

# ---------- 5.7 删除 BodyText/FirstParagraph/Compact ----------
# 院校 docx 正文段落一律无 pStyle（即 Normal）。删除这三个样式后，
# pandoc 会自动回退为 Normal，与院校文件逐段同构。
for sid in ['BodyText', 'BodyTextChar', 'FirstParagraph', 'Compact']:
    S = re.sub(r'<w:style [^>]*w:styleId="%s"[^>]*>.*?</w:style>' % re.escape(sid), '', S, count=1, flags=re.S)
# ---------- 6. Blockquote：楷体灰字（院校元数据行） ----------
S = replace_style(S, 'Blockquote',
'''<w:style w:type="paragraph" w:customStyle="1" w:styleId="Blockquote"><w:name w:val="Blockquote"/><w:basedOn w:val="Normal"/><w:rPr><w:rFonts w:ascii="楷体" w:hAnsi="楷体" w:eastAsia="楷体"/><w:color w:val="444444"/><w:sz w:val="22"/></w:rPr></w:style>''')

# ---------- 7. ListParagraph / Caption ----------
S = replace_style(S, 'ListParagraph',
'''<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>''')
S = replace_style(S, 'Caption',
'''<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="35"/><w:semiHidden/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:b/><w:bCs/><w:color w:val="4F81BD" w:themeColor="accent1"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>''')

# ---------- 8. Table 样式：TableGrid 全边框 + 居中 ----------
S = replace_style(S, 'Table',
'''<w:style w:type="table" w:default="1" w:styleId="Table"><w:name w:val="Table"/><w:basedOn w:val="TableNormal"/><w:qFormat/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>''')

open(os.path.join(work, 'word', 'styles.xml'), 'w', encoding='utf-8').write(S)

# ---------- 9. document.xml：A4 页面设置 ----------
D = open(os.path.join(work, 'word', 'document.xml'), encoding='utf-8').read()
D = re.sub(r'<w:sectPr.*?</w:sectPr>',
'''<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1797" w:bottom="1440" w:left="1797" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>''',
D, count=1, flags=re.S)
open(os.path.join(work, 'word', 'document.xml'), 'w', encoding='utf-8').write(D)

# ---------- 10. 重新打包 ----------
os.makedirs(os.path.dirname(OUT), exist_ok=True)
if os.path.exists(OUT):
    os.remove(OUT)
zf = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)
for root, _, files in os.walk(work):
    for f in files:
        full = os.path.join(root, f)
        zf.write(full, os.path.relpath(full, work))
zf.close()
print('✅ 院校Word参考模板已生成：', OUT, os.path.getsize(OUT), 'bytes')

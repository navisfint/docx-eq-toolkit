# -*- coding: utf-8 -*-
"""把 7 个编号公式段落替换为「无框线两列表格」：左列公式 / 右列编号。

只做结构重排：OLE 对象、关系 id（rId*）、嵌入流全部原样搬移，
不触碰任何公式内容，因此 MathType 对象保持可双击编辑。
"""
import zipfile, re, os, sys

BASE = r'F:\workBuddy\2026-10-13-11-10-59'
SRC = os.path.join(BASE, 'Cost-Optimal Data Annotation via Multi-Task Cost Prediction and Constrained Optimization (MathType).docx')
DST = os.path.join(BASE, 'Cost-Optimal Data Annotation via Multi-Task Cost Prediction and Constrained Optimization (MathType, \u516c\u5f0f\u8868\u683c).docx')

TEXT_W = 10342          # 版心宽 (11906 - 782 - 782) twips
RIGHT_W = 900           # 右列：编号 = 45pt，够放 �（n） 且留 15pt 余量
LEFT_W = TEXT_W - RIGHT_W
IND_LEFT = RIGHT_W      # 左列段落左侧缩进；缩进量 = 右列宽，
                        # 使左列可用区中心恰好落在版心中心（公式真正居中）

NUM_RE = re.compile(r'^\uff08([1-7])\uff09$')   # ^（1）…（7）$

z = zipfile.ZipFile(SRC)
xml = z.read('word/document.xml').decode('utf-8')

PARA = re.compile(r'<w:p\b[^>]*>.*?</w:p>', re.S)
RUN = re.compile(r'<w:r\b[^>]*>.*?</w:r>', re.S)


def para_text(s):
    return ''.join(re.findall(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', s, re.S))


def para_num(s):
    """段落里唯一的 〈（n）〉 编号，否则 None"""
    m = NUM_RE.match(para_text(s).strip())
    return m.group(1) if m else None


def split_ppr(para):
    """返回 (除 jc/wordWrap 外的 pPr 内部, rPr 片段)"""
    m = re.search(r'<w:pPr(?:\s[^>]*)?>(.*?)</w:pPr>', para, re.S)
    inner = m.group(1) if m else ''
    inner = re.sub(r'<w:wordWrap[^>]*/>', '', inner)
    inner = re.sub(r'<w:jc[^>]*/>', '', inner)
    inner = re.sub(r'<w:ind[^>]*/>', '', inner)
    mr = re.search(r'<w:rPr>.*?</w:rPr>', inner, re.S)
    rpr = mr.group(0) if mr else ''
    if rpr:
        inner = inner.replace(rpr, '')
    return inner, rpr


def find_obj_run(para):
    for rm in RUN.finditer(para):
        if '<w:object' in rm.group(0):
            return rm.group(0)
    return None


def strip_position(run):
    return re.sub(r'<w:rPr>\s*<w:position[^>]*/>\s*</w:rPr>', '', run)


def find_num_rpr(para):
    """取包含 〈（〉 那个 run 的 rPr 作为编号格式"""
    for rm in RUN.finditer(para):
        body = rm.group(0)
        if '\uff08' in body and '<w:object' not in body:
            mr = re.search(r'<w:rPr>.*?</w:rPr>', body, re.S)
            if mr:
                return mr.group(0)
    return ('<w:rPr><w:rFonts w:ascii="Times New Roman Regular" w:hAnsi="Times New Roman Regular" '
            'w:cs="Times New Roman Regular"/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/>'
            '<w:lang w:eastAsia="zh-CN"/></w:rPr>')


BORDERLESS = (
    '<w:tblBorders>'
    '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    '</w:tblBorders>'
)
CELLMAR0 = ('<w:tblCellMar>'
            '<w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
            '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>'
            '</w:tblCellMar>')


def build_table(para):
    obj_run = strip_position(find_obj_run(para))
    num = para_num(para)                   # 公式序号，如 1
    inner, rpr = split_ppr(para)
    num_rpr = find_num_rpr(para)

    left_p = ('<w:p><w:pPr>' + inner + '<w:ind w:left="%d"/>' % IND_LEFT
              + '<w:jc w:val="center"/>' + rpr + '</w:pPr>'
              + obj_run + '</w:p>')
    num_run = '<w:r>' + num_rpr + '<w:t>\uff08' + num + '\uff09</w:t></w:r>'
    right_p = ('<w:p><w:pPr>' + inner + '<w:jc w:val="right"/>' + rpr + '</w:pPr>'
               + num_run + '</w:p>')

    return (
        '<w:tbl><w:tblPr>'
        '<w:tblW w:w="%d" w:type="dxa"/>' % TEXT_W +
        '<w:jc w:val="center"/>' + BORDERLESS +
        '<w:tblLayout w:type="fixed"/>' + CELLMAR0 +
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="%d"/><w:gridCol w:w="%d"/></w:tblGrid>' % (LEFT_W, RIGHT_W) +
        '<w:tr><w:trPr><w:cantSplit/></w:trPr>'
        '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>' % LEFT_W + left_p + '</w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>' % RIGHT_W + right_p + '</w:tc>'
        '</w:tr></w:tbl>'
    )


hits = []


def repl(m):
    para = m.group(0)
    if 'OLEObject' not in para:
        return para
    num = para_num(para)
    if num is None:
        return para
    tbl = build_table(para)
    # 段落里的书签标记（MTBlankEqn）挪到表格之后，保持纯 body 级范围
    marks = ''.join(re.findall(r'<w:bookmark(?:Start|End)\b[^>]*/>', para))
    hits.append((len(para), num))
    return tbl + marks


out = PARA.sub(repl, xml)

print('replaced paragraphs:', len(hits))
if len(hits) != 7:
    print('!! expected 7, got', len(hits))
    sys.exit(1)

print('tbl count  %d -> %d' % (xml.count('<w:tbl>'), out.count('<w:tbl>')))
print('OLEObject  %d -> %d' % (xml.count('<o:OLEObject'), out.count('<o:OLEObject')))
print('rId refs   %d -> %d' % (xml.count('r:id="rId'), out.count('r:id="rId')))

# 写新 docx：其余条目原样复制
zin = zipfile.ZipFile(SRC)
zout = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename == 'word/document.xml':
        data = out.encode('utf-8')
    zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
    zi.compress_type = item.compress_type
    zi.external_attr = item.external_attr
    zout.writestr(zi, data)
zout.close()
print('written ->', DST)
print('size =', os.path.getsize(DST))

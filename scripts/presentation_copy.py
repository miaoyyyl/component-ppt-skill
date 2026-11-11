"""展示文字的通用整理与实际PPT检查；不处理演讲备注或来源元数据。"""
from __future__ import annotations
import argparse,json,re,zipfile
from pathlib import Path
from lxml import etree

POLICY_VERSION='2026-11-02'
CITATION=re.compile(r'[\[［]\s*\d+(?:\s*[,，、–—\-]\s*\d+)*\s*[\]］]')
TERMINAL=re.compile(r'[。．.!！?？,，;；:：、…]+(?=[”’」』）)】]*$)')
QUESTION=re.compile(r'哪些|为什么|如何|是否|什么是|是什么|怎么|[?？]')
NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}

def clean_display_text(value):
    """保留句中标点、小数、单位及科学符号；来源编号只在展示副本中清理。"""
    return '\n'.join(TERMINAL.sub('',CITATION.sub('',line).rstrip()).rstrip() for line in value.split('\n'))

def inspect_text(value,*,title=False):
    issues=[]
    for line in value.split('\n'):
        if CITATION.search(line):issues.append('visible-citation')
        if TERMINAL.search(line.rstrip()):issues.append('terminal-punctuation')
    if title and QUESTION.search(value):issues.append('question-title')
    return sorted(set(issues))

def audit_pptx(path):
    issues=[];checked=0
    with zipfile.ZipFile(path) as z:
        parts=sorted((n for n in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)),key=lambda n:int(re.search(r'(\d+)\.xml',n)[1]))
        for page,part in enumerate(parts,1):
            root=etree.fromstring(z.read(part))
            for shape in root.findall('.//p:sp',NS):
                name=shape.find('p:nvSpPr/p:cNvPr',NS).get('name','')
                for paragraph in shape.findall('p:txBody/a:p',NS):
                    text=''.join(paragraph.xpath('./a:r/a:t/text() | ./a:fld/a:t/text()',namespaces=NS)).strip()
                    if not text:continue
                    checked+=1
                    for code in inspect_text(text,title=name=='prod-61000'):
                        issues.append({'page':page,'object':name,'text':text,'code':code})
    return {'policyVersion':POLICY_VERSION,'passed':not issues,'checkedParagraphs':checked,'issues':issues,'scope':'实际幻灯片原生文字；备注不清理，图片内文字另检查生成标签与实际渲染'}

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('pptx');parser.add_argument('--report',required=True);args=parser.parse_args()
    result=audit_pptx(args.pptx);Path(args.report).write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False));raise SystemExit(0 if result['passed'] else 1)

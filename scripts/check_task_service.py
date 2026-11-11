"""自动任务后端边界检查；测试夹具与真实模型检查分别记录。"""
import argparse
import copy
import io
import json
import subprocess
import tempfile
import threading
import time
import unittest
from unittest.mock import patch
from urllib.error import HTTPError
from urllib.request import Request,urlopen
import zipfile
from pathlib import Path
from xml.etree import ElementTree as E

from production_tasks import (ROOT,CONSTRAINTS,WORKFLOW,catalog_data,component_options,digest,extract_docx,
                              normalize_task,now,validate_result)
from task_server import TaskStore,TaskHTTPServer,git_snapshot


def docx_bytes(paragraphs):
    w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"; document=E.Element("{"+w+"}document");body=E.SubElement(document,"{"+w+"}body")
    for text in paragraphs: E.SubElement(E.SubElement(E.SubElement(body,"{"+w+"}p"),"{"+w+"}r"),"{"+w+"}t").text=text
    buffer=io.BytesIO()
    with zipfile.ZipFile(buffer,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels",'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml",E.tostring(document,encoding="utf-8",xml_declaration=True))
    return buffer.getvalue()


def smoke_document():
    return docx_bytes(["自动分析功能测试原稿", "以下方法A、B、C是为测试排版选择而编写的假设性占位，不是科研结论或实验数据。",
      "方法A：关注输入资料的来源与质量。方法B：关注分析步骤和条件的一致性。方法C：关注结果解释和资料追溯。",
      "共同需求：三类方法均需要统一记录信息来源，说明条件与限制，并在输出前进行质量检查。不得写出未给出的性能提升数值。",
      "资料补充测试：请自主查阅NASA或USGS官网，核对Landsat 8发射日期并提供支持这一日期的官方原文链接；该项作为信息来源核验示例，不宣称与上述虚构方法存在技术因果。",
      "页面分配：2至3页即可，不需要封面、结束页或参考文献专页。三类方法的共同需求适合三类并列与共同特征归纳。外部核验示例可以放在相关说明页。"])


def fixture_result(options):
    component=next(c for c in options if c["id"]=="BW-TRIPLE-COMMON-01");form=next(f for f in component["forms"] if f["id"]=="all-text")
    slot=next(s for s in form["slots"] if s["kind"] in {"text","rich-text"})
    return {"title":"后端测试夹具","summary":"未调用模型，仅验证任务执行与结果约束。","contentCards":[{"id":"C1","statement":"测试内容","relation":"共同归纳","sourceRefs":["D1:p1"],"limitations":"仅测试"}],"gaps":[],"sources":[],"researchLog":[],
      "pages":[{"pageNumber":1,"section":"测试","title":"测试页","purpose":"测试任务执行","relation":"共同归纳","componentId":component["id"],"variantId":"all-text","selectionTier":"preferred","selectionReason":"人工优先且关系适配","contentCardIds":["C1"],"slots":[{"key":slot["key"],"content":"测试内容","sourceRefs":["D1:p1"],"emphasisPhrases":[]}],"images":[],"notes":"测试夹具","unresolved":[]}],"warnings":["测试夹具，未调用模型"]}


def make_task(document):
    catalog=catalog_data();stamp=now()
    return {"schema":"ppt-production-task","schemaVersion":1,"id":"test-task-12345678","libraryVersion":catalog["version"],"createdAt":stamp,"updatedAt":stamp,"savedAt":None,
      "brief":{"title":"后端测试","audience":"","durationMinutes":None,"pageRange":"","mustKeep":"","sourceNotes":"","researchScope":""},
      "preferredComponents":[{"componentId":"BW-TRIPLE-COMMON-01","variantId":"all-text","priority":1}],"documents":[{k:document[k] for k in ["sourceId","name","sha256"]}],
      "workflow":dict(WORKFLOW),"constraints":dict(CONSTRAINTS),"fallback":{"mode":"unselected-library-components","componentIds":[],"variantPolicy":"model-auto"}}


class ServiceChecks(unittest.TestCase):
    def setUp(self):
        qa=ROOT/"_qa";qa.mkdir(exist_ok=True);self.temp=tempfile.TemporaryDirectory(dir=qa)
        self.directory=Path(self.temp.name).resolve();self.assertTrue(self.directory.is_relative_to(qa.resolve()))
        self.stores=[]

    def tearDown(self):
        for store in self.stores: store.close()
        self.temp.cleanup()

    def store(self,executor):
        store=TaskStore(self.directory/str(len(self.stores)),executor,False);self.stores.append(store);return store

    def wait_terminal(self,store,ident):
        deadline=time.monotonic()+8
        while time.monotonic()<deadline:
            job=store.get(ident)
            if job["status"] in {"completed","failed","cancelled","needs-review"}:return job
            time.sleep(.02)
        self.fail("测试任务未结束")

    def test_word_extraction_and_original_hash(self):
        payload=docx_bytes(["标题","正文与原稿定位"]);before=digest(payload)
        result=extract_docx(payload,"原稿.docx","D1",self.directory/"assets")
        self.assertEqual(result["blocks"][1]["id"],"D1:p2");self.assertEqual(result["blocks"][1]["text"],"正文与原稿定位");self.assertEqual(before,digest(payload))

    def test_invalid_file_and_selection(self):
        store=self.store(lambda *args:None)
        with self.assertRaises(ValueError):store.upload("bad.docx",b"not a docx")
        doc=store.upload("sample.docx",docx_bytes(["测试"]))
        for version in catalog_data().get('compatibleTaskVersions',[]):
            legacy=make_task(doc);legacy['libraryVersion']=version
            upgraded=normalize_task(legacy,catalog_data())
            self.assertEqual(upgraded['libraryVersion'],catalog_data()['version']);self.assertEqual(upgraded['preferredComponents'],legacy['preferredComponents'])
        task=make_task(doc);task["preferredComponents"]=[]
        with self.assertRaises(ValueError):normalize_task(task,catalog_data())
        task=make_task(doc);task["workflow"]["contentAnalysis"]="manual"
        with self.assertRaises(ValueError):normalize_task(task,catalog_data())
        task=make_task(doc);task["documents"][0]["sha256"]="0"*64
        with self.assertRaises(ValueError):store.submit(task)

    def test_model_completed_and_saved_snapshot(self):
        gate=threading.Event();result=fixture_result(component_options(catalog_data()))
        def executor(directory,cancel,progress,timeout):gate.wait(3);return copy.deepcopy(result)
        store=self.store(executor);doc=store.upload("sample.docx",docx_bytes(["测试内容"]))
        task=make_task(doc);job=store.submit(task);duplicate=store.submit(task)
        self.assertEqual(job["id"],duplicate["id"]);task["brief"]["title"]="提交后改名";gate.set()
        completed=self.wait_terminal(store,job["id"]);self.assertEqual(completed["status"],"completed");self.assertEqual(completed["executionMode"],"test-fixture")
        saved=json.loads((store.output/job["id"]/"selection.json").read_text(encoding="utf-8"));self.assertEqual(saved["brief"]["title"],"后端测试")
        self.assertTrue((store.runs/job['id']/'library/theme.json').is_file());self.assertTrue((store.runs/job['id']/'docs/组件拆解与复用规范.md').is_file())
        sources=json.loads((store.output/job['id']/'sources.json').read_text(encoding='utf-8'));self.assertEqual(sources['documents'][0]['usedWordBlocks'][0]['id'],'D1:p1')
        self.assertEqual(len(saved["fallback"]["componentIds"]),len(catalog_data()["components"])-1);self.assertTrue((store.output/job["id"]/"report.md").is_file())

    def test_failure_and_cancellation_are_not_completed(self):
        def broken(*args):raise RuntimeError("测试模型失败")
        store=self.store(broken);doc=store.upload("sample.docx",docx_bytes(["测试内容"]))
        job=store.submit(make_task(doc));self.assertEqual(self.wait_terminal(store,job["id"])["status"],"failed")
        def waiting(directory,cancel,*args):cancel.wait(5);raise InterruptedError()
        other=self.store(waiting);doc=other.upload("sample.docx",docx_bytes(["测试内容"]))
        job=other.submit(make_task(doc));other.cancel(job["id"]);self.assertEqual(self.wait_terminal(other,job["id"])["status"],"cancelled")

    def test_model_cannot_override_variant_or_invent_source(self):
        options=component_options(catalog_data());docs=[{"blocks":[{"id":"D1:p1"}],"assets":[]}]
        store=self.store(lambda *a:None);doc=store.upload("sample.docx",docx_bytes(["测试内容"]));task=make_task(doc)
        result=fixture_result(options);result["pages"][0]["variantId"]="base"
        with self.assertRaises(ValueError):validate_result(result,task,docs,options)
        result=fixture_result(options);result["contentCards"][0]["sourceRefs"]=["D99:p999"]
        with self.assertRaises(ValueError):validate_result(result,task,docs,options)

    def test_local_research_log_accepts_only_existing_word_refs(self):
        options=component_options(catalog_data());docs=[{'blocks':[{'id':'D1:p1'}],'assets':[{'id':'D1:image1'}]}]
        store=self.store(lambda *a:None);doc=store.upload('sample.docx',docx_bytes(['测试内容']));task=make_task(doc)
        result=fixture_result(options);result['researchLog']=[{'query':'本地材料核对','purpose':'确认已有证据','outcome':'无需外部检索','sourceIds':['D1:p1','D1:image1']}]
        validate_result(result,task,docs,options)
        result['researchLog'][0]['sourceIds'].append('D1:p999')
        with self.assertRaises(ValueError):validate_result(result,task,docs,options)

    def test_presentation_failure_is_review_state_and_preserves_analysis(self):
        def builder(*args):raise RuntimeError('测试：渲染尚未完成')
        store=TaskStore(self.directory/'deck-failure',lambda *a:fixture_result(component_options(catalog_data())),False,deck_builder=builder);self.stores.append(store)
        doc=store.upload('sample.docx',docx_bytes(['测试内容']));job=store.submit(make_task(doc));finished=self.wait_terminal(store,job['id'])
        self.assertEqual(finished['status'],'needs-review');self.assertFalse(finished['renderPassed'])
        self.assertTrue((store.output/job['id']/'analysis.json').is_file());self.assertTrue((store.output/job['id']/'presentation-check.json').is_file())
        inventory=json.loads((store.runs/job['id']/'library/inventory.json').read_text(encoding='utf-8'))
        self.assertIn('sourceTemplateSha256',inventory['components'][0])

    def test_retry_reuses_valid_analysis_with_original_selection(self):
        count=[]
        def executor(directory,*args):
            count.append(1);from production_tasks import write_json
            write_json(directory/'model-result.json',fixture_result(component_options(catalog_data())))
            raise RuntimeError('测试执行中断')
        store=self.store(executor);doc=store.upload('sample.docx',docx_bytes(['测试内容']));job=store.submit(make_task(doc))
        self.assertEqual(self.wait_terminal(store,job['id'])['status'],'failed');store.pending.join()
        before=(store.runs/job['id']/'selection.json').read_bytes();store.retry(job['id']);finished=self.wait_terminal(store,job['id'])
        self.assertEqual(finished['status'],'completed');self.assertTrue(finished['analysisReused']);self.assertEqual(len(count),1)
        self.assertEqual((store.runs/job['id']/'selection.json').read_bytes(),before)

    def test_http_origin_and_private_file_boundaries(self):
        store=self.store(lambda *a:fixture_result(component_options(catalog_data())))
        with patch('task_server.model_availability',return_value={"available":True,"loggedIn":True,"message":"测试夹具"}):server=TaskHTTPServer(('127.0.0.1',0),store)
        thread=threading.Thread(target=server.serve_forever,daemon=True);thread.start();origin=f'http://127.0.0.1:{server.server_port}'
        try:
            with urlopen(origin+'/api/bootstrap') as response:bootstrap=json.load(response)
            self.assertEqual(bootstrap['applicationId'],'ppt-component-task-server')
            with self.assertRaises(HTTPError) as caught:urlopen(origin+'/.git/config')
            self.assertEqual(caught.exception.code,404)
            request=Request(origin+'/api/jobs',data=b'{}',headers={'Origin':'http://unrelated.invalid','X-Task-Token':bootstrap['token']},method='POST')
            with self.assertRaises(HTTPError) as caught:urlopen(request)
            self.assertEqual(caught.exception.code,403)
            request=Request(origin+'/api/jobs',data=b'{}',headers={'Origin':origin,'X-Task-Token':bootstrap['token']},method='POST')
            with self.assertRaises(HTTPError) as caught:urlopen(request)
            self.assertEqual(caught.exception.code,400)
        finally:server.shutdown();server.server_close();thread.join(timeout=3)

    def test_cancel_during_result_save_stays_cancelled(self):
        store=self.store(lambda *a:fixture_result(component_options(catalog_data())))
        original_update=store.update
        def update(ident,**values):
            original_update(ident,**values)
            if values.get('status')=='saving':store.cancel(ident)
        store.update=update
        doc=store.upload('test.docx',docx_bytes(['测试内容']));job=store.submit(make_task(doc))
        self.assertEqual(self.wait_terminal(store,job['id'])['status'],'cancelled')

    def test_git_snapshot_preserves_unrelated_staged_changes(self):
        repository=self.directory/"repo";repository.mkdir()
        def git(*args):return subprocess.run(["git",*args],cwd=repository,capture_output=True,check=True).stdout.decode("utf-8").strip()
        git("init");git("config","user.name","Local task test");git("config","user.email","test@local.invalid")
        unrelated=repository/"unrelated.txt";unrelated.write_text("original",encoding="utf-8");git("add","--","unrelated.txt");git("commit","-m","baseline")
        unrelated.write_text("user change",encoding="utf-8");git("add","--","unrelated.txt")
        output=repository/"jobs/test/report.md";output.parent.mkdir(parents=True);output.write_text("model result",encoding="utf-8")
        sha=git_snapshot(repository,[output],"保存测试任务")
        self.assertTrue(sha);self.assertEqual(git("show","HEAD:unrelated.txt"),"original");self.assertEqual(git("diff","--cached","--name-only"),"unrelated.txt")


if __name__=="__main__":
    parser=argparse.ArgumentParser();parser.add_argument("--write-smoke",action="store_true");args=parser.parse_args()
    if args.write_smoke:
        path=ROOT/"_qa/auto-analysis-smoke.docx";path.parent.mkdir(exist_ok=True);path.write_bytes(smoke_document());print(path)
    else:
        suite=unittest.defaultTestLoader.loadTestsFromTestCase(ServiceChecks);run=unittest.TextTestRunner(verbosity=2).run(suite)
        from production_tasks import write_json
        write_json(ROOT/"docs/task-service-check.json",{"passed":run.wasSuccessful(),"testsRun":run.testsRun,"mode":"test-fixtures-no-model","failures":len(run.failures),"errors":len(run.errors)})
        raise SystemExit(0 if run.wasSuccessful() else 1)

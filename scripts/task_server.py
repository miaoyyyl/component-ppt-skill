"""仅在本机运行的制作任务入口：选型保存后自动调用已登录模型。"""
from __future__ import annotations
import argparse
import json
import mimetypes
import os
import platform
import queue
import re
import secrets
import shutil
import subprocess
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlsplit

from palette_assets import theme_path
from production_tasks import (ROOT, MAX_DOCUMENT_BYTES, catalog_data, component_options, digest, extract_docx,
                              model_prompt, normalize_task, now, open_docx, read_json, report_markdown,
                              response_schema, validate_result, write_json)

ACTIVE={"queued","extracting","model-running","validating","saving","preparing-ppt","binding-ppt","collecting-images","assembling-ppt","rendering-ppt","reviewing-ppt"}
HIDDEN_FLAGS=subprocess.CREATE_NO_WINDOW if os.name=="nt" else 0


def codex_executable():
    direct=shutil.which("codex.exe")
    if direct: return direct
    if os.name=="nt":
        arch="arm64" if platform.machine().lower() in {"arm64","aarch64"} else "x64"
        triple="aarch64-pc-windows-msvc" if arch=="arm64" else "x86_64-pc-windows-msvc"
        package=Path(os.environ.get("APPDATA",""))/"npm/node_modules/@openai/codex"
        candidates=[package/f"node_modules/@openai/codex-win32-{arch}/vendor/{triple}/bin/codex.exe",package/f"vendor/{triple}/bin/codex.exe"]
        return next((str(p) for p in candidates if p.is_file()),None)
    return shutil.which("codex")


def model_availability():
    executable=codex_executable()
    if not executable: return {"available":False,"loggedIn":False,"message":"未找到本机模型入口，请先安装并登录Codex命令行。"}
    try:
        result=subprocess.run([executable,"login","status"],capture_output=True,timeout=15,creationflags=HIDDEN_FLAGS)
        logged_in=result.returncode==0
        return {"available":True,"loggedIn":logged_in,"message":"本机模型已登录，可以开始自动制作。" if logged_in else "本机模型尚未登录，请先完成Codex命令行登录。"}
    except (OSError,subprocess.TimeoutExpired):
        return {"available":True,"loggedIn":False,"message":"暂时无法确认本机模型连接，请重启本地任务入口后重试。"}


def stop_owned_process(process):
    if process is None or process.poll() is not None: return
    if os.name=="nt":
        subprocess.run(["taskkill","/PID",str(process.pid),"/T","/F"],capture_output=True,timeout=15,creationflags=HIDDEN_FLAGS)
    else: process.terminate()
    try: process.wait(timeout=10)
    except subprocess.TimeoutExpired: process.kill()


def run_model_request(directory,cancel,on_progress,timeout_seconds,prompt,schema,name='model',images=None):
    executable=codex_executable()
    if not executable: raise RuntimeError("没有可用的本机模型入口")
    directory=Path(directory)
    if not re.fullmatch('[a-zA-Z0-9-]{1,80}',name):raise ValueError('模型阶段名称不正确')
    schema_path=directory/(name+'-schema.json');write_json(schema_path,schema)
    output_path=directory/(name+'-result.json')
    command=[executable,"-a","never","--search","exec","--ephemeral","--sandbox","read-only","--json","--color","never",
             "--skip-git-repo-check","--cd",str(directory)]
    if images:command.extend(['--image',*[str(Path(p).resolve()) for p in images]])
    command.extend(["--output-schema",str(schema_path),"--output-last-message",str(output_path),"-"])
    environment=os.environ.copy(); environment.pop("CODEX_THREAD_ID",None)
    started=time.monotonic(); process=None
    try:
        with (directory/(name+'-events.jsonl')).open("wb") as stdout,(directory/(name+'-stderr.log')).open("wb") as stderr:
            process=subprocess.Popen(command,stdin=subprocess.PIPE,stdout=stdout,stderr=stderr,cwd=directory,env=environment,creationflags=HIDDEN_FLAGS)
            (directory/(name+'-process.json')).write_text(json.dumps({"pid":process.pid,"startedAt":now(),'images':[str(p) for p in images or []]}),encoding="utf-8")
            process.stdin.write(prompt.encode("utf-8")); process.stdin.close()
            while process.poll() is None:
                if cancel.wait(.5): raise InterruptedError("任务已取消")
                if time.monotonic()-started>timeout_seconds: raise TimeoutError("本阶段模型执行超时，记录已保留，可按原快照继续")
                on_progress(int(time.monotonic()-started))
        if process.returncode!=0: raise RuntimeError(f"模型执行失败（退出码{process.returncode}），本机任务日志已保留")
        output=output_path
        if not output.is_file() or output.stat().st_size>6*1024*1024: raise RuntimeError("模型未返回有效的初版结果")
        return read_json(output)
    finally: stop_owned_process(process)


def run_codex(directory,cancel,on_progress,timeout_seconds):
    return run_model_request(directory,cancel,on_progress,timeout_seconds,model_prompt(),response_schema())


def git_snapshot(repository,paths,message):
    repository=Path(repository).resolve()
    relative=[]
    for path in paths:
        resolved=Path(path).resolve()
        if not resolved.is_relative_to(repository): raise ValueError("提交目标超出本仓库")
        relative.append(resolved.relative_to(repository).as_posix())
    def git(args):
        result=subprocess.run(["git",*args],cwd=repository,capture_output=True,timeout=60,creationflags=HIDDEN_FLAGS)
        if result.returncode: raise RuntimeError("本地Git版本保存失败；分析结果仍保留在任务目录")
        return result.stdout.decode("utf-8",errors="replace").strip()
    git(["add","--",*relative])
    # --only防止带入用户已暂存的其他修改。
    git(["commit","--only","-m",message,"--",*relative])
    return git(["rev-parse","--short","HEAD"])


class TaskStore:
    def __init__(self,data_dir=None,executor=run_codex,auto_commit=True,timeout_seconds=900,deck_builder=None):
        self.data=Path(data_dir or ROOT/".runtime").resolve()
        if not self.data.is_relative_to(ROOT): raise ValueError("任务存储必须位于本项目目录")
        self.data.mkdir(parents=True,exist_ok=True)
        self.output=ROOT/"jobs" if self.data==ROOT/".runtime" else self.data/"results"
        self.runs=self.data/"runs"; self.runs.mkdir(exist_ok=True)
        self.uploads=self.data/"uploads";self.uploads.mkdir(exist_ok=True)
        self.upload_index=self.uploads/"index.json"
        self.lock=threading.RLock();self.pending=queue.Queue();self.jobs={};self.cancellations={};self.running_job=None
        self.executor=executor;self.auto_commit=auto_commit;self.timeout=timeout_seconds;self.closed=False
        if deck_builder is None and executor is run_codex:
            from deck_pipeline import build_presentation
            deck_builder=build_presentation
        self.deck_builder=deck_builder
        for path in self.runs.glob("*/job.json"):
            try:
                job=read_json(path)
                if job["status"] in ACTIVE:
                    job.update(status="interrupted",message="本地服务已重启，可按原快照继续制作",updatedAt=now());write_json(path,job)
                self.jobs[job["id"]]=job
            except (ValueError,KeyError,OSError): continue
        self.worker=threading.Thread(target=self.work,daemon=True);self.worker.start()

    def document_entries(self):
        entries={}; source_root=(ROOT/"by_human").resolve()
        if source_root.is_dir():
            for path in sorted(source_root.rglob("*.docx")):
                if path.name.startswith("~$") or not path.resolve().is_relative_to(source_root): continue
                source_id="workspace:"+digest(path.relative_to(ROOT).as_posix().encode("utf-8"))
                entries[source_id]={"sourceId":source_id,"name":path.name,"sha256":digest(path.read_bytes()),"bytes":path.stat().st_size,"path":path}
        if self.upload_index.is_file():
            for item in read_json(self.upload_index):
                path=self.uploads/(item["sha256"]+".docx")
                if path.is_file(): entries[item["sourceId"]]=item|{"path":path}
        return entries

    def public_documents(self):
        with self.lock: return [{k:v for k,v in d.items() if k!="path"} for d in self.document_entries().values()]

    def upload(self,name,payload):
        name=name.replace("\\","/").rsplit("/",1)[-1]
        if not name.lower().endswith(".docx") or not 1<=len(name)<=240 or "\x00" in name: raise ValueError("请选择有效的DOCX文件")
        with open_docx(payload): pass
        sha=digest(payload); item={"sourceId":"upload:"+sha,"name":name,"sha256":sha,"bytes":len(payload)}
        with self.lock:
            destination=self.uploads/(sha+".docx")
            if not destination.exists(): destination.write_bytes(payload)
            index=read_json(self.upload_index) if self.upload_index.exists() else []
            if not any(d["sourceId"]==item["sourceId"] for d in index): index.append(item);write_json(self.upload_index,index)
            return item

    def update(self,job_id,**values):
        with self.lock:
            job=self.jobs[job_id];job.update(values,updatedAt=now());write_json(self.runs/job_id/"job.json",job)

    def submit(self,raw):
        catalog_payload=(ROOT/"library/components.json").read_bytes();catalog=json.loads(catalog_payload);task=normalize_task(raw,catalog)
        with self.lock:
            entries=self.document_entries(); snapshots=[]
            for doc in task["documents"]:
                item=entries.get(doc["sourceId"])
                if not item: raise ValueError("任务原稿不在本机，请重新选择或上传")
                payload=item["path"].read_bytes()
                if digest(payload)!=doc["sha256"]: raise ValueError("原稿已更新，请重新选择原稿后开始")
                with open_docx(payload): pass
                doc["name"]=item["name"];snapshots.append(payload)
            fingerprint=digest(json.dumps({k:v for k,v in task.items() if k not in {"createdAt","updatedAt","savedAt"}},sort_keys=True,ensure_ascii=False).encode("utf-8"))
            existing=next((j for j in self.jobs.values() if j["fingerprint"]==fingerprint and j["status"] in ACTIVE),None)
            if existing: return self.public(existing)
            job_id="job-"+time.strftime("%Y%m%dT%H%M%S",time.gmtime()).lower()+"-"+secrets.token_hex(4)
            directory=self.runs/job_id;directory.mkdir()
            for index,payload in enumerate(snapshots,1): (directory/f"source-{index}.docx").write_bytes(payload)
            write_json(directory/"selection.json",task)
            write_json(directory/"component-options.json",component_options(catalog))
            write_json(directory/"library/components.json",catalog)
            write_json(directory/"library/theme.json",read_json(theme_path(ROOT/"library",task["paletteId"])))
            rules=directory/"docs/组件拆解与复用规范.md";rules.parent.mkdir(exist_ok=True)
            rules.write_bytes((ROOT/"docs/组件拆解与复用规范.md").read_bytes())
            if self.deck_builder is not None:
                from deck_pipeline import snapshot_templates
                snapshot_templates(directory,catalog,task["paletteId"])
            job={"id":job_id,"taskId":task["id"],"title":task["brief"]["title"] or task["documents"][0]["name"],"status":"queued","message":"等待自动分析",
                 "createdAt":now(),"updatedAt":now(),"fingerprint":fingerprint,"elapsedSeconds":0,"pageCount":0,"sourceCount":0,"versionCommit":None,"versionError":None,
                 "librarySha256":digest(catalog_payload),"executionMode":"live-model" if self.executor is run_codex else "test-fixture",'presentationAvailable':False,'renderPassed':False}
            self.jobs[job_id]=job;self.cancellations[job_id]=threading.Event();write_json(directory/"job.json",job);self.pending.put(job_id)
            return self.public(job)

    def public(self,job):
        return {k:v for k,v in job.items() if k not in {"fingerprint"}}

    def get(self,job_id):
        with self.lock:
            if job_id not in self.jobs: raise KeyError("任务不存在")
            return self.public(self.jobs[job_id])

    def latest(self,task_id):
        with self.lock:
            jobs=[j for j in self.jobs.values() if j["taskId"]==task_id]
            return self.public(max(jobs,key=lambda j:j["createdAt"])) if jobs else None

    def cancel(self,job_id):
        with self.lock:
            if job_id not in self.jobs: raise KeyError("任务不存在")
            if self.jobs[job_id]["status"] in ACTIVE:
                self.cancellations[job_id].set();self.update(job_id,status="cancelled",message="已取消自动分析，已有记录保留")
            return self.get(job_id)

    def retry(self,job_id):
        with self.lock:
            if job_id not in self.jobs:raise KeyError('任务不存在')
            job=self.jobs[job_id]
            if job['status'] in ACTIVE or self.running_job==job_id:raise ValueError('后台尚在处理或结束本次任务，请稍后再继续')
            if job['status'] not in {'failed','interrupted','cancelled','needs-review'}:raise ValueError('本次任务无需恢复，可按当前选型新建一次制作')
            self.cancellations[job_id]=threading.Event();self.update(job_id,status='queued',message='按启动快照继续制作，复用有效的初版分析')
            self.pending.put(job_id);return self.get(job_id)

    def work(self):
        while True:
            job_id=self.pending.get()
            if job_id is None: break
            self.running_job=job_id
            try: self.execute(job_id)
            except InterruptedError: self.update(job_id,status="cancelled",message="已取消自动分析，已有记录保留")
            except Exception as error:
                if self.cancellations[job_id].is_set(): self.update(job_id,status="cancelled",message="已取消自动分析，已有记录保留")
                else: self.update(job_id,status="failed",message=str(error)[:600])
            finally:self.running_job=None;self.pending.task_done()

    def execute(self,job_id):
        directory=self.runs/job_id; cancel=self.cancellations[job_id]
        if cancel.is_set(): raise InterruptedError()
        self.update(job_id,status="extracting",message="正在读取Word原稿，提取正文、表格和图像定位")
        task=read_json(directory/"selection.json");documents=[]
        for index,source in enumerate(task["documents"],1):
            documents.append(extract_docx((directory/f"source-{index}.docx").read_bytes(),source["name"],f"D{index}",directory/"assets"))
        if sum(len(b["text"]) for d in documents for b in d["blocks"])>180000:
            raise ValueError("本次原稿正文超过18万字，请明确要制作的文档或章节范围；没有截断原文")
        options=read_json(directory/"component-options.json")
        write_json(directory/"word-content.json",documents);write_json(directory/"component-options.json",options);write_json(directory/"output.schema.json",response_schema())
        (directory/"model-instructions.txt").write_text(model_prompt(),encoding="utf-8")
        if cancel.is_set(): raise InterruptedError()
        self.update(job_id,status="model-running",message="模型正在连续完成内容分析、资料检索和逐页分配")
        def progress(elapsed):
            if elapsed%5==0: self.update(job_id,elapsedSeconds=elapsed)
        result=None;cached=directory/'model-result.json'
        if cached.is_file():
            try:
                candidate=read_json(cached);validate_result(candidate,task,documents,options);result=candidate
                self.update(job_id,analysisReused=True,message='已复用本快照的有效初版分析，继续后续制作')
            except (ValueError,KeyError,OSError):pass
        if result is None:result=self.executor(directory,cancel,progress,self.timeout)
        if cancel.is_set(): raise InterruptedError()
        self.update(job_id,status="validating",message="模型已返回，正在核对组件、人工变体与来源定位")
        validate_result(result,task,documents,options)
        destination=self.output/job_id;destination.mkdir(parents=True,exist_ok=True)
        write_json(destination/"selection.json",task)
        write_json(destination/"analysis.json",result)
        used_refs={ref for c in result["contentCards"] for ref in c["sourceRefs"]}|{ref for p in result["pages"] for s in p["slots"] for ref in s["sourceRefs"]}
        source_documents=[{k:v for k,v in d.items() if k!="blocks"}|{"usedWordBlocks":[b for b in d["blocks"] if b["id"] in used_refs]} for d in documents]
        write_json(destination/"sources.json",{"documents":source_documents,"contentCache":(directory/"word-content.json").relative_to(ROOT).as_posix(),"externalSources":result["sources"],"researchLog":result["researchLog"]})
        write_json(destination/"plan.json",{"title":result["title"],"pages":result["pages"],"workflow":task["workflow"],"librarySha256":self.jobs[job_id]["librarySha256"]})
        (destination/"report.md").write_text(report_markdown(result,task,documents),encoding="utf-8")
        deck_report=None
        if self.deck_builder is not None:
            def stage(status,message):
                if cancel.is_set():raise InterruptedError()
                self.update(job_id,status=status,message=message,elapsedSeconds=0)
            def model_call(prompt,schema,name,images):
                if cancel.is_set():raise InterruptedError()
                return run_model_request(directory,cancel,progress,self.timeout,prompt,schema,name,images)
            try:
                deck_report=self.deck_builder(directory,destination,task,result,documents,cancel,model_call,stage)
            except InterruptedError:raise
            except Exception as error:
                deck_report={'passed':False,'error':str(error)[:1200],'rendered':False}
                write_json(destination/'presentation-check.json',deck_report)
        self.update(job_id,status="saving",message="正在保存初版方案与本地版本")
        if cancel.is_set(): raise InterruptedError()
        commit=None;version_error=None
        if self.auto_commit:
            files=[destination/n for n in ['selection.json','analysis.json','sources.json','plan.json','report.md','presentation.pptx','presentation-check.json','render-plan.json','image-bindings.json'] if (destination/n).is_file()]
            files+=sorted((destination/'previews').glob('slide-*.png'))
            try: commit=git_snapshot(ROOT,files,"保存PPT制作任务结果："+job_id)
            except Exception as error: version_error=str(error)
        with self.lock:
            if cancel.is_set(): raise InterruptedError()
            ready=deck_report is None or deck_report['passed']
            message='可编辑PPTX已生成，PowerPoint渲染和模型逐页检查通过' if deck_report and ready else '初版内容分析、资料补充与逐页分配已完成' if deck_report is None else '初版已保存；PPT制作或渲染仍有待检查项'
            self.update(job_id,status='completed' if ready else 'needs-review',message=message,pageCount=len(result["pages"]),sourceCount=len(result["sources"]),
                        unresolvedCount=sum(g["remainsOpen"] for g in result["gaps"])+sum(len(p["unresolved"]) for p in result["pages"]),versionCommit=commit,versionError=version_error,
                        presentationAvailable=(destination/'presentation.pptx').is_file(),renderPassed=bool(deck_report and deck_report.get('passed')),presentationError=deck_report.get('error') if deck_report else None)

    def close(self):
        self.closed=True
        for job_id in list(self.jobs):
            if self.jobs[job_id]["status"] in ACTIVE: self.cancel(job_id)
        self.pending.put(None);self.worker.join(timeout=20)


class TaskHTTPServer(ThreadingHTTPServer):
    daemon_threads=True
    def __init__(self,address,store):
        self.store=store;self.token=secrets.token_urlsafe(32);self.availability=model_availability()
        super().__init__(address,Handler)
        self.origins={f"http://127.0.0.1:{self.server_port}",f"http://localhost:{self.server_port}"}


class Handler(BaseHTTPRequestHandler):
    def log_message(self,format,*args): pass

    def send(self,status,payload,content_type="application/json; charset=utf-8"):
        body=json.dumps(payload,ensure_ascii=False).encode("utf-8") if isinstance(payload,(dict,list)) else payload
        self.send_response(status);self.send_header("Content-Type",content_type);self.send_header("Content-Length",str(len(body)))
        self.send_header("Cache-Control","no-store");self.send_header("X-Content-Type-Options","nosniff")
        self.send_header("Content-Security-Policy","default-src 'self' blob: data:; script-src 'self'; style-src 'self' 'unsafe-inline'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'")
        self.end_headers();self.wfile.write(body)

    def check_request(self,mutation=False):
        if "http://"+self.headers.get("Host","") not in self.server.origins: raise PermissionError("只接受本机任务入口请求")
        if mutation:
            if self.headers.get("Origin") not in self.server.origins or self.headers.get("X-Task-Token")!=self.server.token: raise PermissionError("请求来源或本次连接无效，请刷新本地任务页面")

    def payload(self,limit=2*1024*1024):
        try: size=int(self.headers.get("Content-Length","-1"))
        except ValueError: raise ValueError("请求长度不正确")
        if not 0<=size<=limit: raise ValueError("提交内容超过允许大小")
        payload=self.rfile.read(size)
        if len(payload)!=size: raise ValueError("提交内容不完整")
        return payload

    def do_GET(self):
        try:
            self.check_request();url=urlsplit(self.path);route=unquote(url.path);store=self.server.store
            if route=="/api/bootstrap":
                self.send(200,{"applicationId":"ppt-component-task-server","workspace":str(ROOT),"token":self.server.token,"documents":store.public_documents(),"runner":self.server.availability,"autoCommit":store.auto_commit});return
            if route=="/api/jobs/latest":
                task_id=parse_qs(url.query).get("taskId",[""])[0];self.send(200,{"job":store.latest(task_id)});return
            match=re.fullmatch(r"/api/jobs/(job-[a-z0-9-]+)(?:/(report\.md|analysis\.json|sources\.json|plan\.json|selection\.json|presentation\.pptx|presentation-check\.json|render-plan\.json|image-bindings\.json|previews/slide-[0-9]{3}\.png))?",route)
            if match:
                job_id,name=match.groups();store.get(job_id)
                if name:
                    path=store.output/job_id/name
                    if not path.is_file(): raise FileNotFoundError("任务结果尚未生成")
                    content_type='application/vnd.openxmlformats-officedocument.presentationml.presentation' if name.endswith('.pptx') else 'image/png' if name.endswith('.png') else 'text/plain; charset=utf-8' if name.endswith('.md') else 'application/json; charset=utf-8'
                    self.send(200,path.read_bytes(),content_type)
                else: self.send(200,{"job":store.get(job_id)})
                return
            if route=="/": route="/library/index.html"
            path=(ROOT/route.lstrip("/")).resolve()
            allowed=any(path.is_relative_to(ROOT/folder) for folder in ["library","docs","templates"])
            if not allowed or not path.is_file() or path.suffix.lower() not in {".html",".css",".js",".json",".png",".jpg",".jpeg",".pptx",".md"}: raise FileNotFoundError("页面不存在")
            content_type=mimetypes.guess_type(path.name)[0] or "application/octet-stream"
            if path.suffix in {".html",".css",".js",".json",".md"}: content_type+="; charset=utf-8"
            self.send(200,path.read_bytes(),content_type)
        except PermissionError as error: self.send(403,{"error":str(error)})
        except (KeyError,FileNotFoundError) as error: self.send(404,{"error":str(error)})
        except (OSError,ValueError) as error: self.send(400,{"error":str(error)})

    def do_POST(self):
        try:
            self.check_request(True);route=urlsplit(self.path).path;store=self.server.store
            if route=="/api/uploads":
                result=store.upload(unquote(self.headers.get("X-File-Name","")),self.payload(MAX_DOCUMENT_BYTES));self.send(201,{"document":result});return
            if route=="/api/jobs":
                if not all(self.server.availability.get(k) for k in ["available","loggedIn"]): raise ValueError(self.server.availability["message"])
                self.send(202,{"job":store.submit(json.loads(self.payload().decode("utf-8-sig")))});return
            match=re.fullmatch(r"/api/jobs/(job-[a-z0-9-]+)/cancel",route)
            if match: self.payload();self.send(200,{"job":store.cancel(match.group(1))});return
            match=re.fullmatch(r'/api/jobs/(job-[a-z0-9-]+)/retry',route)
            if match:self.payload();self.send(202,{'job':store.retry(match.group(1))});return
            if route=="/api/shutdown":
                self.payload();self.send(200,{"message":"本地任务入口正在停止"})
                threading.Thread(target=self.server.shutdown,daemon=True).start();return
            raise FileNotFoundError("操作不存在")
        except PermissionError as error: self.send(403,{"error":str(error)})
        except (KeyError,FileNotFoundError) as error: self.send(404,{"error":str(error)})
        except (ValueError,OSError) as error: self.send(400,{"error":str(error)})


def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument("--port",type=int,default=8871)
    parser.add_argument("--data-dir",type=Path);parser.add_argument("--no-auto-commit",action="store_true");parser.add_argument("--timeout",type=int,default=900)
    args=parser.parse_args()
    store=TaskStore(args.data_dir,auto_commit=not args.no_auto_commit,timeout_seconds=args.timeout)
    server=TaskHTTPServer(("127.0.0.1",args.port),store)
    info={"pid":os.getpid(),"url":f"http://127.0.0.1:{server.server_port}/library/index.html","startedAt":now()}
    write_json(store.data/"server.json",info);print(json.dumps(info,ensure_ascii=False),flush=True)
    try: server.serve_forever(poll_interval=.2)
    except KeyboardInterrupt: pass
    finally: server.server_close();store.close()


if __name__=="__main__": main()

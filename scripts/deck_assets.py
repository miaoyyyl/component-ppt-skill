"""收集本次实际图像：Word原图、公开图片文件或来源网页的实际截图。"""
from __future__ import annotations
import base64
import ipaddress
import json
import os
import shutil
import socket
import subprocess
from pathlib import Path
from urllib.parse import urlsplit
from urllib.request import Request,build_opener,HTTPRedirectHandler

from production_tasks import ROOT,digest,write_json
from deck_native import image_details

MAX_IMAGE=20*1024*1024


def public_url(url):
    parsed=urlsplit(url)
    if parsed.scheme not in {"http","https"} or not parsed.hostname or parsed.username or parsed.password:raise ValueError("只支持无需凭据的公开HTTP图片或来源页面")
    if parsed.port not in {None,80,443}:raise ValueError("来源页面使用了不支持的端口")
    for info in socket.getaddrinfo(parsed.hostname,parsed.port or (443 if parsed.scheme=="https" else 80),type=socket.SOCK_STREAM):
        if not ipaddress.ip_address(info[4][0]).is_global:raise ValueError("图片来源不能访问本机或内部网络")
    return url


class PublicRedirects(HTTPRedirectHandler):
    def redirect_request(self,request,fp,code,message,headers,newurl):
        public_url(newurl)
        return super().redirect_request(request,fp,code,message,headers,newurl)


def download_image(url,destination):
    public_url(url);request=Request(url,headers={"User-Agent":"PPTComponentLibrary/1.0"})
    with build_opener(PublicRedirects()).open(request,timeout=30) as response:
        public_url(response.geturl());payload=response.read(MAX_IMAGE+1)
        if len(payload)>MAX_IMAGE:raise ValueError("图片文件超过20MB")
    destination.write_bytes(payload)
    width,height,_,_=image_details(destination)
    if width*height>50_000_000:raise ValueError("图片像素过大，需提供适合演示的原始图")
    return destination


def browser_runtime():
    module=Path(os.environ.get("PPT_PLAYWRIGHT_DIR",Path.home()/".cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright"))
    browser=next((p for p in [Path(os.environ.get("PROGRAMFILES(X86)","C:/Program Files (x86)"))/"Microsoft/Edge/Application/msedge.exe",Path(os.environ.get("PROGRAMFILES","C:/Program Files"))/"Microsoft/Edge/Application/msedge.exe"] if p.is_file()),None)
    node=shutil.which("node")
    if not node or not module.exists() or browser is None:raise ValueError("本机网页截图运行库不可用；可继续使用Word原图或公开图片文件")
    return node,module,browser


def screenshot_source(url,destination):
    public_url(url);node,module,browser=browser_runtime();request=destination.with_suffix(".request.json")
    write_json(request,{"url":url,"output":str(destination),"viewport":{"width":1280,"height":720}})
    result=subprocess.run([node,str(ROOT/"scripts/capture_source.cjs"),str(module),str(browser),str(request)],capture_output=True,timeout=65,
                          creationflags=subprocess.CREATE_NO_WINDOW if os.name=="nt" else 0)
    if result.returncode or not destination.is_file():raise ValueError("未取得可用的来源页面截图："+result.stderr.decode("utf-8",errors="replace")[-300:])
    return destination


def collect_assets(directory,documents,render_plan,cancel=None):
    directory=Path(directory);folder=directory/"collected-assets";folder.mkdir(exist_ok=True)
    available={a["id"]:directory/"assets"/a["file"] for d in documents for a in d["assets"]}
    expected={a["id"]:a["sha256"] for d in documents for a in d["assets"]};assets={};records=[];missing=[];cached={}
    for page in render_plan["pages"]:
        for image in page["images"]:
            if cancel is not None and cancel.is_set():raise InterruptedError()
            key=image["assetId"] or f"missing:{page['pageNumber']}:{image['targetId']}"
            image["assetId"]=key
            try:
                mode=image["mode"]
                if mode=="word":
                    path=available.get(key)
                    if path is None or digest(path.read_bytes())!=expected[key]:raise ValueError("本次Word没有该原图，或原图哈希变化")
                elif mode in {"image-url","web-screenshot"}:
                    cache_key=mode+":"+image["url"]
                    if cache_key in cached:path=cached[cache_key]
                    else:
                        path=folder/(digest(cache_key.encode("utf-8"))+(".png" if mode=="web-screenshot" else ".image"))
                        if not path.exists():
                            if mode=="web-screenshot":screenshot_source(image["url"],path)
                            else:download_image(image["url"],path)
                        cached[cache_key]=path
                else:raise ValueError("本次没有可用的真实图像")
                width,height,_,_=image_details(path)
                if key in assets and assets[key]!=path:raise ValueError("同一资产编号对应不同图像")
                assets[key]=path
                records.append({"pageNumber":page["pageNumber"],"targetId":image["targetId"],"assetId":key,"sourceRef":image["sourceRef"],"mode":mode,"url":image["url"],"path":str(path),"sha256":digest(path.read_bytes()),"pixels":[width,height]})
            except (ValueError,OSError,subprocess.TimeoutExpired) as error:
                # 透明像素只供诊断占位；最终报告明确缺图，不借用母版案例图片。
                placeholder=folder/"missing-transparent.png"
                if not placeholder.exists():placeholder.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII="))
                assets[key]=placeholder
                missing.append({"pageNumber":page["pageNumber"],"targetId":image["targetId"],"description":image["description"],"reason":str(error)})
    return assets,records,missing

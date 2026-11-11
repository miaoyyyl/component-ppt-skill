"""独立后台PowerPoint工作进程：读取实际样式、渲染、尺寸检查和保存重开。"""
from __future__ import annotations
import argparse
import contextlib
import json
import os
import time
from pathlib import Path

from production_tasks import ROOT,digest,read_json,write_json


@contextlib.contextmanager
def powerpoint_lock():
    path=ROOT/".runtime/powerpoint.lock";path.parent.mkdir(exist_ok=True)
    with path.open("a+b") as lock:
        lock.seek(0);lock.write(b"1");lock.flush();deadline=time.monotonic()+180
        if os.name=="nt":
            import msvcrt
            while True:
                try:lock.seek(0);msvcrt.locking(lock.fileno(),msvcrt.LK_NBLCK,1);break
                except OSError:
                    if time.monotonic()>deadline:raise TimeoutError("PowerPoint后台正在处理其他任务，请稍后重试")
                    time.sleep(.2)
        try:yield
        finally:
            if os.name=="nt":lock.seek(0);msvcrt.locking(lock.fileno(),msvcrt.LK_UNLCK,1)


def walk_shapes(shapes):
    for index in range(1,shapes.Count+1):
        shape=shapes.Item(index);yield shape
        if shape.Type==6:yield from walk_shapes(shape.GroupItems)


def color_hex(value):
    value=int(value)
    return f"{value&255:02X}{(value>>8)&255:02X}{(value>>16)&255:02X}" if value>=0 else "000000"


def inspect(application,request):
    profiles={}
    for form in request["forms"]:
        presentation=None
        try:
            presentation=application.Presentations.Open(str(Path(form["template"]).resolve()),True,False,False)
            shapes={int(s.Id):s for s in walk_shapes(presentation.Slides(1).Shapes)};profile={"templateSha256":digest(Path(form["template"]).read_bytes()),"shapes":{}}
            for target in form["targets"]:
                if target["kind"]!="text":continue
                shape=shapes[target["shapeId"]];text=shape.TextFrame.TextRange;paras=[]
                for index,prototype in enumerate(target["paragraphs"],1):
                    paragraph=text.Paragraphs(index,1)
                    char=paragraph.Characters(min(prototype["sampleOffset"]+1,max(int(paragraph.Length),1)),1)
                    size=float(char.Font.Size)
                    if size<=0:raise ValueError(f"无法确认对象{shape.Id}第{index}段的字号")
                    paras.append({"sizePt":size,"bold":int(char.Font.Bold)==-1,"italic":int(char.Font.Italic)==-1,"color":color_hex(char.Font.Color.RGB),"isBlank":prototype["isBlank"],"bullet":prototype["bullet"],"runIndex":prototype.get("runIndex",0)})
                profile["shapes"][str(target["shapeId"]) ]={"paragraphs":paras}
            profiles[form["key"]]=profile
        finally:
            if presentation is not None:presentation.Close()
    return {"application":"Microsoft PowerPoint","version":str(application.Version),"profiles":profiles}


def render(application,request):
    path=Path(request["input"]).resolve();folder=Path(request["output"]).resolve();folder.mkdir(parents=True,exist_ok=True)
    presentation=None;records=[];issues=[]
    try:
        presentation=application.Presentations.Open(str(path),True,False,False)
        if presentation.Slides.Count!=len(request["pages"]):raise ValueError("PowerPoint打开后页数变化")
        for index,page in enumerate(request["pages"],1):
            slide=presentation.Slides(index);shapes={int(s.Name[5:]) if str(s.Name).startswith('prod-') and str(s.Name)[5:].isdigit() else int(s.Id):s for s in walk_shapes(slide.Shapes)}
            image=folder/f"slide-{index:03d}.png";slide.Export(str(image),"PNG",1280,720);checks=[]
            for ident in page.get('hiddenShapeIds',[]):
                if ident not in shapes or int(shapes[ident].Visible)!=0:issues.append({'pageNumber':index,'targetId':'prod-'+str(ident),'kind':'visibility-change','description':'已授权隐藏的连接对象仍然显示'})
            for target in page["targets"]:
                if target["kind"]!="text":continue
                shape=shapes[target["shapeId"]];frame=shape.TextFrame2;text=frame.TextRange
                if not str(text.Text).strip():continue
                width=float(text.BoundWidth);height=float(text.BoundHeight)
                usable_width=float(shape.Width)-float(frame.MarginLeft)-float(frame.MarginRight)
                usable_height=float(shape.Height)-float(frame.MarginTop)-float(frame.MarginBottom)
                # 旋转文本的外接尺寸与本地文字框互换；竖排文本保留PowerPoint实际外接尺寸。
                rotation=float(shape.Rotation)%360
                if min(abs(rotation-90),abs(rotation-270))<.1:width,height=height,width
                outer_width=float(shape.Width);outer_height=float(shape.Height)
                overflow=width>outer_width+1.5 or height>outer_height+1.5
                item={"targetId":target["targetId"],"shapeId":target["shapeId"],"boundPt":[round(width,2),round(height,2)],"availablePt":[round(usable_width,2),round(usable_height,2)],"framePt":[round(outer_width,2),round(outer_height,2)],"touchesMargins":width>usable_width+1.5 or height>usable_height+1.5,"overflow":overflow}
                for pindex,expected in enumerate(target["paragraphs"],1):
                    actual=float(shape.TextFrame.TextRange.Paragraphs(pindex,1).Characters(1,1).Font.Size)
                    if abs(actual-expected["sizePt"])>.02:issues.append({"pageNumber":index,"targetId":target["targetId"],"kind":"font-change","description":f"字号应为{expected['sizePt']}pt，实际{actual}pt"})
                    actual_color=color_hex(shape.TextFrame.TextRange.Paragraphs(pindex,1).Characters(1,1).Font.Color.RGB)
                    if expected.get('color') and actual_color!=expected['color']:issues.append({'pageNumber':index,'targetId':target['targetId'],'kind':'color-change','description':'文字颜色与绑定样式不一致'})
                if overflow:issues.append({"pageNumber":index,"targetId":target["targetId"],"kind":"overflow","description":f"文字外接尺寸{item['boundPt']}超过框体{item['framePt']}，保持字号重新组织内容"})
                checks.append(item)
            records.append({"pageNumber":index,"preview":str(image),"sha256":digest(image.read_bytes()),"textChecks":checks})
        roundtrip=Path(request.get('roundtrip',folder/"roundtrip.pptx"))
        if roundtrip.exists():roundtrip.unlink()
        presentation.SaveCopyAs(str(roundtrip),24);presentation.Close();presentation=None
        presentation=application.Presentations.Open(str(roundtrip),True,False,False)
        if presentation.Slides.Count!=len(request["pages"]):raise ValueError("PowerPoint保存副本重开后页数变化")
        return {"application":"Microsoft PowerPoint","version":str(application.Version),"inputSha256":digest(path.read_bytes()),"opened":True,"rendered":True,"savedCopyAndReopened":True,"passed":not issues,"issues":issues,"pages":records}
    finally:
        if presentation is not None:presentation.Close()


def assemble(application,request):
    presentation=None
    try:
        presentation=application.Presentations.Add(False)
        while presentation.Slides.Count:presentation.Slides(1).Delete()
        presentation.PageSetup.SlideWidth=request['dimensions'][0]/12700
        presentation.PageSetup.SlideHeight=request['dimensions'][1]/12700
        for index,path in enumerate(request['inputs'],1):
            source=None
            try:
                source=application.Presentations.Open(str(Path(path).resolve()),True,False,False)
                design=presentation.Designs.Clone(source.Slides(1).Design)
                inserted=presentation.Slides.InsertFromFile(str(Path(path).resolve()),presentation.Slides.Count,1,1)
                if int(inserted)!=1:raise ValueError('PowerPoint原生导入页数异常')
                slide=presentation.Slides(index);slide.Design=design
                layout=source.Slides(1).CustomLayout
                for layout_index in range(1,design.SlideMaster.CustomLayouts.Count+1):
                    candidate=design.SlideMaster.CustomLayouts(layout_index)
                    if candidate.Name==layout.Name:slide.CustomLayout=candidate;break
                slide.FollowMasterBackground=source.Slides(1).FollowMasterBackground
                for shape_index in range(slide.Shapes.Count,0,-1):
                    shape=slide.Shapes(shape_index)
                    if not str(shape.Name).startswith('prod-'):shape.Delete()
            finally:
                if source is not None:source.Close()
            note=request['notes'][index-1]
            for shape in walk_shapes(presentation.Slides(index).NotesPage.Shapes):
                if shape.Type==14 and shape.PlaceholderFormat.Type==2 and shape.HasTextFrame:
                    shape.TextFrame.TextRange.Text=note;break
        presentation.SaveAs(str(Path(request['output']).resolve()),24)
        return {'application':'Microsoft PowerPoint','version':str(application.Version),'slides':presentation.Slides.Count,'nativeImport':True}
    finally:
        if presentation is not None:presentation.Close()


def placeholder_batch(application,request):
    records={}
    for item in request['items']:
        source=Path(item['input']);image=Path(item['image']);image.parent.mkdir(parents=True,exist_ok=True)
        report=render(application,{'input':str(source),'output':item['output'],'pages':[{'targets':[]}]})
        import shutil
        shutil.copyfile(report['pages'][0]['preview'],image)
        records[item['id']]={'inputSha256':report['inputSha256'],'previewSha256':digest(image.read_bytes()),'opened':True,'rendered':True,'savedCopyAndReopened':True}
    return {'application':'Microsoft PowerPoint','version':str(application.Version),'items':records}


def main():
    import pythoncom
    from win32com.client import DispatchEx
    parser=argparse.ArgumentParser();parser.add_argument("request",type=Path);args=parser.parse_args()
    request=read_json(args.request);application=None
    with powerpoint_lock():
        pythoncom.CoInitialize()
        try:
            application=DispatchEx("PowerPoint.Application")
            operation=request['operation']
            result=inspect(application,request) if operation=='inspect' else assemble(application,request) if operation=='assemble' else placeholder_batch(application,request) if operation=='placeholder-batch' else render(application,request)
            write_json(request["report"],result)
            print(json.dumps({"operation":request["operation"],"passed":result.get("passed",True)},ensure_ascii=False),flush=True)
        finally:
            if application is not None and application.Presentations.Count==0:application.Quit()
            pythoncom.CoUninitialize()


if __name__=="__main__":main()

# 组件PPT与科研工作流运行包

本包整理自2026-11-02的本地工作环境，提供当前组件库、Agent执行约定、制作工具和五个科研skill。主流程输入为“核心方向与重点＋论文PDF＋组件库选型JSON”，Word可选

这是一份可迁移的运行资料包。Python、Office、浏览器、Node及Codex宿主工具需要在目标电脑准备，不是免安装软件。包内没有本次主题的Word、论文、已导出任务、生成的主题PPT及outputs

## 1. 包内目录

```text
README.md                         本说明
MANIFEST.json                     文件清单、逐文件SHA-256、来源版本与排除范围
requirements-ppt.txt              PPT工具已核对的直接依赖版本
requirements-research.lock.txt    科研Python环境依赖快照
project/                          当前组件库项目
  AGENTS.md                       当前主流程与固定结构规则
  library/index.html              组件展示与任务选型入口
  library/components.json         33类组件、16个变体，共49种形式
  library/templates/              原生案例母版
  library/placeholders/           X占位母版、容量及预览
  library/previews/               案例预览
  library/references/             公开参考图与来源清单
  scripts/                        PDF提取、原生PPT制作、渲染及检查工具
  docs/                           工作流、组件规范和历史检查说明
  templates/                      任务说明与PDF交接提示词
  packaging/                      本运行包的说明源文件
skills/                           主流程按需使用的五个科研skill
  research-literature/            缺口文献检索、去重和元数据核对
  research-reading/               定向阅读、关键图表与证据评价
  research-writing/               汇报主线、论点和表达组织
  research-figures/               真实数据科研绘图
  research-design/                需要时生成研究设想与实验方案
optional-skills/
  academic-blue-ppt-design/       本机另一个独立PPT skill的原样快照
```

组件案例母版、X占位母版、示例图片及预览是组件库的运行资产，会随包保留。它们用于理解和复用布局，正式制作时须替换为本次真实图文，不能沿用示例数据或单位身份

## 2. 解压后先做什么

1. 完整解压，保持上述相对目录，不只取出HTML、JSON或某几个PPTX
2. 双击`project/library/index.html`，离线查看组件、选择优先级与变体、填写大致要求并导出JSON
3. 在能访问本目录的Codex/Agent中提供核心方向、相关PDF及选型JSON，要求先读本README、`project/AGENTS.md`和`project/docs/PDF文献驱动的Agent制作流程.md`
4. 主Agent按实际任务读取`skills/<名称>/SKILL.md`，连续完成阅读、图文提取、补充检索、汇报层次、逐页安排、原生PPTX制作与实际渲染检查

默认不要求安装skill到全局目录，Agent可直接按包内路径读取。需要在Codex中以安装skill的方式调用时，将五个目录复制到该电脑配置的skills目录，保留目录内references、assets、scripts、LICENSE和NOTICE等文件；已有同名目录先比较并另存，不覆盖旧版本

组件库只负责人工选型与JSON导出。离线HTML不上传PDF/Word、不连接模型，也不启动PPT制作。首次打开没有原电脑的浏览器任务记录；跨电脑迁移选型需要另行提供导出的JSON

## 3. Python与系统依赖

### PPT制作环境

原环境使用Windows、Python 3.11以及桌面版Microsoft PowerPoint。需有微软雅黑字体。实际渲染、保存重开和COM操作依赖PowerPoint，不能以只生成PPTX文件代替渲染验收

在本包根目录，可按已安装的Python版本调整以下命令

```powershell
py -3.11 -m venv .venv-ppt
& ./.venv-ppt/Scripts/python.exe -m pip install -r requirements-ppt.txt
```

包内没有Office安装程序、字体文件或Python解释器。PowerPoint后台检查串行执行，不关闭用户已打开的文稿、不向真实窗口发送鼠标键盘输入

### 科研环境

原科研环境使用Python 3.12，主要库为PyMuPDF、Matplotlib、Pillow、bibtexparser、Pyzotero和PyYAML。依赖锁定文件来自当前科研环境的快照

```powershell
py -3.12 -m venv .venv-research
& ./.venv-research/Scripts/python.exe -m pip install -r requirements-research.lock.txt
```

上述安装需要目标电脑可访问所用Python包源，本包没有离线wheel。版本快照不保证所有操作系统、Python版本及未来包源都能直接安装；安装失败应保留错误，由Agent核对兼容性，不静默替换全部依赖

科研skill内的`references/local-runtime.md`与`run-literature.ps1`、`run-figures.ps1`保留原电脑的用户目录约定。迁移时可以直接用新环境的Python运行其脚本，不必建立原作者的绝对路径

```powershell
& ./.venv-research/Scripts/python.exe skills/research-literature/scripts/literature_tools.py --help
& ./.venv-research/Scripts/python.exe skills/research-figures/scripts/plot_csv.py --help
```

联网检索、网页读取和AI生成图需要接手Agent实际具备对应工具。Zotero、本机认证或数据库API key按需在目标机另行配置；包内不含用户账号、密钥、模型认证及Zotero文献库。不要把密钥写进任务JSON或日志

## 4. 无Word任务的执行

用户可直接使用`project/templates/PDF汇报交接提示词.md`。最少给出核心方向、PDF及选型JSON，听众/时长/页数可沿用JSON，未填时由Agent记录合理假设

工具示例从本包根目录执行，替换成真实文件路径

```powershell
Push-Location project
& ../.venv-research/Scripts/python.exe -B scripts/research_bundle.py init --selection "选型.json" --pdf "论文A.pdf" "论文B.pdf" --direction "本次核心方向与重点" --output "outputs/本次任务"
& ../.venv-research/Scripts/python.exe -B scripts/research_bundle.py check "outputs/本次任务"
Pop-Location
```

`init`保存选型快照、原PDF哈希、分页文字位置、原页预览、嵌入图资源及空证据台账。它不调用模型，也不把提取完成当成已经精读。多面板、矢量图与带遮罩图应回到原页核对并按需要裁切

Agent补齐重点分析、来源/证据/图片记录和逐页方案后，运行`check --stage plan`检查来源定位、图片状态、指定变体与槽位覆盖。科学结论和图片可读性仍由Agent回看原文确认

制作阶段使用包内原生母版，保持结构、字号、位置、图框及渐变/标签等局部样式。蓝白主色与红黄强调保留，内容页无页码。通过低层原生PPT函数绑定图文并拼装，随后使用实际PowerPoint渲染、保存重开、逐页查看和修订。完整步骤见`project/docs/PDF文献驱动的Agent制作流程.md`

允许围绕已确认缺口联网补充。AI可补概念或场景插图，参照本次图组统一风格并记录生成性质；真实数据图使用科研绘图工具，不用AI图伪造实验照片、测量曲线或定量结果

## 5. 两个PPT入口的区别

默认入口是`project/`中的人工选型组件工作流，组件库版本0.7.0。指定JSON与变体、槽位位置和字号为本次制作约束

`optional-skills/academic-blue-ppt-design/`是本机另一个独立skill的快照，自己的发布版本为0.1，内部含54类基础组件与33个组合。它以Word和自己的规划接口、原生组件及Artifact Tool为主，不能用来替代本项目选型JSON或重新设计固定布局。只有用户明确选择该独立入口时再按其SKILL.md运行

该可选入口的Codex Artifact Tool、bundled Node等宿主依赖不在本包中，需在目标Codex环境通过`load_workspace_dependencies`定位。具体环境变量、构建和定稿命令见其`references/runtime-guide.md`。本包没有改写或安装这份快照，也不把两套入口视为同一版本

## 6. 包含与排除的边界

已包含当前主流程的代码、母版、容量、预览、参考图、规范、任务模板，以及五个科研skill的脚本、schema、参考说明和来源/许可文件

已排除

- 当前`outputs/`及其他主题制作结果、主题PPT、原稿Word/PDF、导出的个人选型任务
- `by_human/`、原始`materials/`整套参考PPT
- `.git/`、`.runtime/`、`_qa/`、历史测试成品和检查截图
- Python虚拟环境、node_modules、缓存、日志、Office临时锁文件及密钥配置

当前49种组件形式的母版、X占位与预览均已包含，可用于制作新稿。完整重建组件库需要 `materials/` 中的原始PPT和归档参考图，部分历史基线检查还需要Git历史。

原项目README与部分历史文档记录了过去的目录、版本、检查截图或本机路径，包中保留说明但不包含对应主题成品；当前任务以本README、AGENTS及PDF流程为准

## 7. 完整性检查与版本管理

在解压后的根目录运行

```powershell
python project/scripts/check_runtime_package.py .
```

该命令只需Python标准库，校验MANIFEST列出的文件、字节数、SHA-256、目录边界和组件资产。可在解压前直接检查ZIP

```powershell
python project/scripts/check_runtime_package.py "运行包.zip"
```

科研入口功能检查可运行

```powershell
& ./.venv-research/Scripts/python.exe -B project/scripts/check_research_bundle.py
```

浏览器回归还需Node、Playwright和Edge/Chromium，在目标环境提供实际模块与浏览器路径后运行`project/scripts/check_tasks.cjs`。这些均不能代替新稿的实际PowerPoint视觉检查

本包没有Git历史。可在解压后的新目录运行`git init`后提交基线，之后每项完成并检查的修改单独提交。不要在用户其他项目内重复初始化或把无关资料一起提交

MANIFEST中的源Git版本指本次打包前的项目快照，不代表目标机已经具备原环境。本次打包的校验结果与ZIP SHA-256附在压缩包旁边，便于复核

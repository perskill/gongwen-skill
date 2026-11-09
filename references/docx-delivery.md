# Word交付

## 选择现有工具或独立脚本

先完成并核对正文，再进行排版。用户提供已有Word、红头、印章、复杂表格或修订稿时，优先使用宿主实际具备的文档编辑能力保留结构。本包独立脚本用于**新建纯文本及普通矩形表格的DOCX**，不假装支持所有Word内容。

使用用户模板时只继承版式，不继承历史正文；已有文档的修改保留原件或另存新版本。禁止仅改后缀伪造DOCX，禁止把接口返回字典或路径字符串当成文件已生成的证据。

独立工具为`scripts/gongwen_docx.py`。生成需要Python 3.9+及python-docx；提取、内容比对和环境诊断用标准库；渲染可选依赖本机LibreOffice，PNG另需pdftoppm。不依赖WorkBuddy SDK、TaxAudit、特定MCP名称、网络或后台服务。

## 内容输入

完整示例见`assets/example-document.json`。输入UTF-8 JSON，根对象为：

```json
{"schema_version":1,"profile":"internal-report","blocks":[
  {"type":"title","text":"关于开展服务窗口试行的报告"},
  {"type":"paragraph","role":"recipient","text":"相关单位："},
  {"type":"heading","level":1,"text":"一、拟开展事项"},
  {"type":"paragraph","text":"以下为虚构示例。拟选择2个窗口试行，具体时间待确认。"},
  {"type":"table","header_rows":1,"rows":[["项目","状态"],["新增金额","0元"]]},
  {"type":"paragraph","role":"signature","text":"【待确认：署名】"},
  {"type":"paragraph","role":"date","text":"【待确认：成文日期】"}
]}
```

`blocks`顺序就是正文顺序。支持`title`、`heading`（level为1—3）、`paragraph`、`table`及`page_break`。段落role可为body、recipient、signature、date、attachment、note；未指定为body。表格rows为矩形二维字符串，header_rows为0或1。请把`A&B`等内容写为实际Unicode文本，不先转成HTML实体。脚本不把全半角数字、小数、否定词或空白随意视为等价。

JSON必须覆盖已经核对的全文，不能为让检查通过删去原稿的难处理内容。它用于排版和回读比对，不是政策或事实来源。

## 命令

以下`<本技能目录>`须替换为实际绝对路径，输入输出路径可包含中文和空格。示例中的`python`按环境选择`python3`或`py -3`。

```text
python "<本技能目录>/scripts/gongwen_docx.py" doctor
python "<本技能目录>/scripts/gongwen_docx.py" build --source "document.json" --output "正文.docx"
python "<本技能目录>/scripts/gongwen_docx.py" verify --source "document.json" --docx "正文.docx" --check-layout
python "<本技能目录>/scripts/gongwen_docx.py" extract --docx "正文.docx"
python "<本技能目录>/scripts/gongwen_docx.py" render --docx "正文.docx" --output-dir "新建预览目录" --png
```

`build`可用`--profile internal-report|research-report|government-body`选择预设，以显式参数为准；自定义样式用`--profiles 配置.json`。已存在的输出默认拒绝覆盖；确认这是本任务允许替换的文件后才传`--overwrite`。构建先验证临时文件，验证通过后再落到目标路径。

`verify`比对主文档正文中的受支持块，不检查页眉、脚注、批注等所有Word部件。`verify --check-layout`只检查本脚本基础版式已实现的参数，不认证任意单位模板或所有有效样式；使用自定义参数时用`--profile`、`--profiles`传入相同配置。页眉页脚的文字内容及其他部件使用原生工具另核。`render`要求输出目录新建或为空；必要时用`--soffice`、`--pdftoppm`指定已有程序位置，不能自动安装软件。

退出码：0表示该命令所述检查完成或通过，1表示内容/受检版式不一致，2表示输入、依赖或执行错误。读取完整JSON中的状态与范围，不能仅看退出码宣布事实正确或公文已合规。

## 内容与页面检查

1. 回读最终文件，检查完整有序段落、短句、数字、否定、重复段、表格每格、附件和落款。保留段首而删去后文必须判失败，删除短金额或署名也必须判失败。
2. 不忽略未支持内容。本工具遇到修订、文本框、图片、公式、正文域、嵌套或合并表格、内容控件等会拒绝或报告不能完整验证；改用能处理该结构的工具并人工核对，不删内容来绕过检查。
3. 检查字体、字号、边距、段落与页码；样式引用不等于字体已经在运行或接收环境安装。
4. 有渲染条件时生成PDF/PNG，逐页查看文字截断、空白、孤立标题、表格跨页、页码和字体替换。脚本输出PDF或PNG不代表页面已被看过。
5. 无渲染或图像阅读能力时照常交付可生成文件，并简短说明已完成文本/结构核对、页面尚未验证；不要冒称逐页检查完成。

外部宿主生成的文件仍需真实存在且可解析，再作适合其结构的全文和页面核对。若生成失败，保留可用正文并准确说明失败环节，不给不存在的下载链接。

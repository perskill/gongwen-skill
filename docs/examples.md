# 使用示例

以下请求可直接交给支持技能的助手，所有案例均为虚构材料。把技能名保留为 `draft-official-documents`，将材料换为自己的内容即可。

## 起草待补报告

```text
请使用 $draft-official-documents，起草约 600 字的工作进展报告，报送上级主管部门。
按目前进展、存在问题、下步安排组织，只交付正文。
已知事实：本月印发了填报说明；收到 2 个单位关于重复提交材料的反馈。
其他单位尚未报送，不代表没有开展工作。培训、走访情况暂缺，可保留 XXX 和 XX。
下步准备收集材料、核对问题并研究试点，地区和时间尚未确定。
不要把旧稿中的成绩、数字或候选做法写成本期已确认事实。
```

## 只修改指定段落

```text
请使用 $draft-official-documents，只修改下面这一段，使普通办事群众更容易理解。
保留业务条件、否定词、时间范围和工作状态，不扩写其他段落。
原文：维修期间实施服务空间功能迁移，窗口业务临时调整至二楼服务区；办理条件沿用现行规定。
```

## 压缩并核对

```text
请使用 $draft-official-documents，将所附草稿压缩到 1000 字以内。
保留主要事实、数字口径、适用条件和待决定事项；不补造来源。
输出压缩后的正文，必要时另列会影响结论的待核事项。
```

## Word 交付

```text
请使用 $draft-official-documents，把已确认正文排成内部汇报 Word。
采用 internal-report 预设，不改变正文内容；输出 DOCX 后回读核对全文和表格。
如环境能够渲染，再检查页面；如不能渲染，说明本次实际完成的核对范围。
正文：……
```

需要直接运行工具时，以 [JSON 示例](../assets/example-document.json) 为输入：

```sh
python scripts/gongwen_docx.py doctor
python scripts/gongwen_docx.py build --source assets/example-document.json --output "示例.docx"
python scripts/gongwen_docx.py verify --source assets/example-document.json --docx "示例.docx" --check-layout
```

`build` 默认保护已有同名文件。输出为基础 DOCX，不包括红头、印章、版记或字体文件；完整参数见 [Word 交付](../references/docx-delivery.md)。

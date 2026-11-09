# 跨工具使用与能力分级

## 目录与调用

每个包独立包含一个技能目录，目录根部直接是 `SKILL.md`，并保留 `references/`、`scripts/`、`assets/`、`agents/`、许可和测试文件。不要只复制入口，也不要把整份ZIP当成一段提示词。`agents/openai.yaml`是可选界面信息，通用指令只依赖标准入口与相对路径。

两个技能名分别是 `draft-official-documents`（通用版）与 `draft-tax-official-documents`（税务核验版）。可以只装一个，也可以并存；一次任务选一版，避免重复执行。税务版已含全部通用能力。

按2026年10月29日查阅的官方文档，Codex和DSH都支持用户级 `~/.agents/skills`。默认配置下，将解压得到的技能目录直接放入此目录，可供两者发现。Windows通常对应 `%USERPROFILE%\.agents\skills`；Linux/macOS对应 `~/.agents/skills`。如果工具在WSL、容器、远程主机中运行，应放在该运行环境可见的目录，不能只放在Windows侧后假定容器已经读到。

- **Codex**：用户级目录 `~/.agents/skills/<技能名>/SKILL.md`；项目级在项目的 `.agents/skills/<技能名>/SKILL.md`。可用 `$draft-official-documents` 或 `$draft-tax-official-documents`明确选择。若选择器尚未更新，刷新技能页面或开启新会话；也可直接要求读取具体路径验证内容。
- **DSH / DeepSeek Harness**：用户级默认 `~/.dsh/skills`或`~/.agents/skills`；`DSH_HOME`、`DSH_AGENTS_HOME`等自定义设置会改变实际位置。项目级为项目根目录的`.dsh/skills`或`.agents/skills`；项目根按最近的`.git`祖先判断，无Git时为当前工作目录。只扫描技能根下一层的`<技能名>/SKILL.md`，避免多套一层包装目录。已有ZIP导入界面时可直接导入本包；没有该界面时用上述文件夹方式。需当前配置启用了本地技能发现。调用时直接说“使用draft-official-documents技能……”即可，具体快捷手势遵从实际版本。
- **其他支持Agent Skills的工具**：使用其文档指定的技能根目录，保持本包目录完整。不要假设所有工具都有同样的斜杠命令、MCP工具名或技能UI。
- **仅能读文件的助手**：明确要求先读入口，再按链接读必要参考；有命令执行能力时才运行脚本。仅能聊天时，可把所需规则和材料提供给模型，文件生成能力取决于该平台。

DSH同名技能按来源优先级处理；Codex同名技能可能同时出现在选择器。安装新版本时只替换对应技能目录并保留旧副本，避免把旧文件混进新包；不要自动删除其他技能或修改已有AGENTS.md。

## 能力分级

| 实际能力 | 可以完成 | 不可声称已完成 |
| --- | --- | --- |
| 仅文本与材料阅读 | 起草、修改、摘要、人工审稿 | 已执行脚本、已生成Word或已查外网 |
| 可运行Python 3.9+ | 文本诊断、DOCX提取与内容比对 | 已生成DOCX，除非生成依赖齐备 |
| 有python-docx | 新建DOCX正文基础版式并回读验证 | 字体已安装、实际分页已检查 |
| 有LibreOffice | 转成PDF供检查 | PDF已经逐页看过 |
| 有LibreOffice与pdftoppm | PDF与逐页PNG | 已经人工验证所有页面 |
| 有可用搜索/浏览工具 | 在授权范围内核对现行来源 | 无工具时假装联网，或认为官网政策证明本地工作 |

通过宿主现有工具读写和搜索，不写死工具名、端口或操作系统路径。所有脚本使用技能自身相对位置定位资源，输出路径由用户任务确定；从其他工作目录启动时传入脚本绝对路径。不要修改系统环境变量、配置凭据、自动安装软件或启动后台服务来完成普通写作。

`python`在示例中代表实际可用的Python 3.9+：Windows可用`py -3`，Linux常用`python3`。使用已存在的环境；需要安装依赖时按用户既有授权处理。文本写作不需要任何依赖。Word生成的可选依赖见根目录`requirements.txt`，未包含离线安装包或字体。内网环境可预先准备与操作系统、CPU架构和Python版本匹配的依赖；无需联网也可使用已安装的工具。

## 快速确认

让助手读取本技能入口并说明技能名与版本，再处理一个真实小任务。可运行`python "<本技能目录>/scripts/gongwen_docx.py" doctor`查看Python、Word依赖与渲染工具；此命令不安装软件、不联网。正式使用前再检查本机字体与最终页面。

本包按上述发现规则和跨平台路径设计；验证记录只表示已执行的范围，不等于已在用户每个DSH/Codex版本、Windows、UOS或每套字体中完成端到端运行。

## 官方依据

- OpenAI：[Build skills](https://learn.chatgpt.com/docs/build-skills)，本地目录、标准入口、显式调用与可选agents元数据。
- DeepSeek官方仓库：[skill-filesystem](https://github.com/deepseek-ai/deepseek-harness/blob/master/packages/skill/skill-filesystem/README.md)，目录形态、扫描范围、默认根目录与优先级。

查阅日期：2026-10-29。工具更新后应以其当前实际配置和官方文档为准。

# 许可与来源

## 当前状态

本项目按用户明确选择采用 **PolyForm Noncommercial 1.0.0**。原包附有 [同一许可全文](../LICENSE-upstream.txt)，根目录 [LICENSE](../LICENSE) 提供完整条款；新增的项目整理内容也按同一许可提供，原包来源链和声明继续保留。

可以在 GitHub 公开托管、协作和发布安装包。GitHub 的 [仓库许可说明](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/licensing-a-repository) 区分公开托管与开源许可；非商业限制不符合 [OSI 开源定义](https://opensource.org/osd) 第 6 项。因此本项目使用“源码公开、非商业许可”表述。具体使用、修改和分发范围以 [官方完整条款](https://polyformproject.org/licenses/noncommercial/1.0.0) 为准，不把“非商业”简单等同于所有单位或所有免费用途。

## 来源清单

| 来源 | 包中记载的用途 | 包内可确认的许可信息 |
| --- | --- | --- |
| 用户提供的通用版 1.0.0 | 本项目的直接输入 | 附上游非商业许可；未提供所有贡献者的独立授权文件 |
| lieflat-gongwen，提交 `e0a5aba6b6ce66402aff799ea213af2d987d99dd` | 通过 gongwen V4 形成来源链 | 原包明确记载 PolyForm Noncommercial 1.0.0 |
| gongwen V4 / integrated-writing-v4 | 工作流、文种、事实与风格区分 | 见上述来源链；未提供完整逐文件归属 |
| tax-dsh V3 writing-first | 通用写作方法 | 原包称未附独立许可声明，不能代其授权 |
| WorkBuddy DOCX | 正文、排版和交付的分工思路 | 原包称脚本为新实现、未复制 SDK；本次没有原素材进行独立比对 |

该清单记录原包陈述，不等于对第三方权利的独立鉴定。用户于 2026-11-09 确认拥有包内全部内容的版权和重新许可权，随后明确选择保留非商业许可；整理工作据此继续，确认记录见 [导入记录](provenance/import-record.json)。上游许可证和来源说明中的保留义务随再分发继续保留。

## 再分发

按许可保留条款、来源和适用声明。贡献者只能提供自己有权提交的内容，新的第三方素材另行记录来源和适用许可。当前发布不附带 MIT 或 Apache-2.0 的替代授权；未来许可变更应由有相应权利的维护者明确作出并同步更新相关文档。

## 依赖

文本诊断、DOCX 提取与核对使用 Python 标准库；DOCX 生成的可选依赖由 `requirements.txt` 声明，开发验证依赖由 `requirements-dev.txt` 声明。本项目不打包这些第三方库、字体或 LibreOffice。安装与再分发依赖时按各自许可处理。

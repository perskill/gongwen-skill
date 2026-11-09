# 发布到 GitHub

项目仓库：[perskill/gongwen-skill](https://github.com/perskill/gongwen-skill)；技能名保持 `draft-official-documents`。

建议仓库简介：

> 中文公文起草、改稿、审稿与 Word 交付 Agent Skill。支持本地工具与可核对的文稿输出，采用非商业许可。

建议 Topics：`agent-skills`、`chinese-writing`、`gongwen`、`document-generation`、`python`、`docx`。

## 当前许可

用户已确认拥有包内全部版权及重新许可权，并选择保留 PolyForm Noncommercial 1.0.0。发布时随附现有 LICENSE、NOTICE 和来源记录，使用“源码公开、非商业许可”的介绍。完整背景见 [许可与来源](licensing.md)。

## 本地验证与打包

1. 使用 Python 3.9+ 的独立环境安装 `requirements-dev.txt`。
2. 运行以下命令并处理实际失败；记录实际测试平台，不把配置中的平台当成已验证结果。

```sh
python scripts/validate_project.py
python -m unittest discover -s tests -v
python scripts/build_release.py --source-archive
```

3. 核对 `dist/` 中两种 ZIP：安装包顶层为 `draft-official-documents/`；项目源码包顶层为 `gongwen-skill/`，含 `.github/`。两者都不包含本机环境、凭据、生成文稿和 `.git/`。
4. 修改版本时更新 `assets/package-info.json`，在 `CHANGELOG.md` 添加对应记录；打包名称自动读取版本。原包历史记录不随之改写。

## 创建仓库与上传

如需发布自己的副本，在 GitHub 新建空仓库 `gongwen-skill`，先不要添加 README、LICENSE 或 `.gitignore`。本目录已经提供这些文件。

已安装 Git 和 GitHub CLI 的维护者可在项目根目录执行：

```sh
git init -b main
git add .
git diff --cached --stat
git diff --cached
git commit -m "Prepare public skill project"
gh repo create gongwen-skill --public --source=. --remote=origin --push
```

提交前检查暂存内容；`gh repo create` 会创建公开仓库并上传。也可解压 `gongwen-skill-github-v<版本>.zip` 后，通过 GitHub 网页上传其内部文件；确认隐藏目录 `.github/` 也已上传，避免额外套上外层目录。

## 发布安装包

上传后查看 **Actions → CI**，等待实际测试结果。需要在 GitHub 构建时，手动运行 **Prepare release**，从该次运行的 Artifacts 下载 ZIP 与校验文件。该工作流只构建产物，不自动创建 Release 或推送代码。

在仓库 Releases 页面新建 Release，标签与 `assets/package-info.json` 的版本一致，例如 `v1.0.1`。添加对应更新记录，并上传：

- `draft-official-documents-v1.0.1.zip` 及同名 `.sha256`；
- `gongwen-skill-github-v1.0.1.zip` 及同名 `.sha256`（可选，便于迁移完整项目）。

发布说明保留非商业许可提示和已知限制。GitHub 自动生成的 Source code ZIP 是完整仓库快照，与提供单层技能目录的安装包用途不同。

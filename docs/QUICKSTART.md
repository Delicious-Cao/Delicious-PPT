# 安装与使用

## 先选版本

- 有图像生成工具、文件处理和 PPTX 制作能力的宿主，选择 ChatGPT 版。
- 有联网检索、图片获取、文件处理和 PPTX 制作能力的宿主，选择 DeepSeek 版。
- 两版都需要实际实现文字动画、翻页和指定字体。名称表示背景策略，不能仅凭模型名称认定工具已经具备。

先阅读 [运行环境与兼容性](RUNTIME.md)。仅加载 Skill 不需要安装本仓库的 Python 示例依赖；依赖只用于运行示例构建脚本。

## 方法一：安装到支持 Skill 的宿主

1. 从 [Releases](https://github.com/Delicious-Cao/Delicious-PPT/releases) 下载对应版本的独立 ZIP。
2. 解压后保留整个 `delicious-ppt-chatgpt` 或 `delicious-ppt-deepseek` 目录，不能只取走 `SKILL.md`。
3. 将该目录放入宿主文档指定的技能目录，保留 `references/` 和可用的 `agents/` 元数据。
4. 按宿主要求重新发现技能，再使用名称调用或让其自动匹配。不同宿主的导入入口可能不同。

例如，Codex 默认个人技能目录为 `~/.codex/skills/`。本次发布没有修改使用者的全局技能配置，也没有对所有宿主的自动发现行为逐一测试。

## 方法二：普通对话中使用指令文件

将所选版本的 `SKILL.md`、该文件夹内的参考文件及任务资料一起提供给助手。界面不支持这些附件时，可以按需粘贴正文。宿主仍须拥有相应的检索、生成、文件与导出工具；粘贴文字不会增加工具能力。

```text
请按 Delicious-PPT 的所选版本执行。
先完整阅读我的资料并充分询问需求，复述制作需求确认单，确认后再开工。
我需要一份面向……的 PPT，主要目标是……，目前的原始材料是……。
```

用途、受众、时长、详略、风格和交付位置没有说清时，助手应继续访谈。固定的字体、背景来源、融合淡化和逐段动画要求直接执行。

## 先看示例

在 [示例目录](../examples/README.md) 查看两组完整案例。每组都有原始材料、需求设定、4 页 PPTX、静态 PDF、逐页预览和原生动画视频。动画视频为缩时效果预览，不是完整授课录音。

## 重建公开示例

此路径使用本机已安装的 Microsoft PowerPoint，限 Windows，不需要在幻灯片中启用宏。示例构建脚本只读取已提供的图片，不会替你生成或联网搜索新的背景。

```powershell
git clone https://github.com/Delicious-Cao/Delicious-PPT.git
Set-Location Delicious-PPT
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r examples/requirements.txt
.\.venv\Scripts\python.exe examples/build_samples.py --output-root work/rebuild --video
```

运行前安装可合法使用的小标宋、仿宋与 Times New Roman，确保 PowerPoint 已激活且能够正常启动。输出目录已有同名文件时，脚本会停止，需另选目录。字体文件不随仓库提供。

此命令是维护者示例的参考构建方式，不代表 Skill 在所有环境中都必须调用这个脚本。实际任务仍需要宿主按对应版本取得新背景、访谈并制作本次内容。

## 常见问题

- **无法生成或搜索背景：** 检查所选版本和宿主工具；报告具体缺项，不能用纯色背景假装满足要求。
- **字体被替换：** 核实实际字体名称及放映电脑的安装情况，不能只看编辑器中保存的字体名称。
- **PDF 看不到动画：** PDF 是附加静态版，动态成品以 PPTX 和对应动画预览为准。
- **播放顺序不对：** 先在本机 PowerPoint 中核对，反馈软件版本、操作步骤和实际表现，不仅描述“动画坏了”。

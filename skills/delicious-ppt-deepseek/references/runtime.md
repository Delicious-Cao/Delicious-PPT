# 运行环境与兼容性

版本：1.0.0。核实日期：2026-10-21。

## 必需能力

| 能力 | ChatGPT 版 | DeepSeek 版 |
| --- | --- | --- |
| 读取用户资料并进行多轮对话 | 必需 | 必需 |
| 每页背景获取 | 实际调用图像生成工具 | 实际联网检索、核验并获取图片 |
| 创建可编辑 PPTX | 必需 | 必需 |
| 真实文字动画和翻页过渡 | 必需 | 必需 |
| 指定字体可用 | 小标宋、仿宋、Times New Roman | 小标宋、仿宋、Times New Roman |
| 文件保存、下载与预览 | 必需 | 必需 |
| 额外事实检索 | 按资料边界执行 | 按资料边界执行，与背景搜图区分 |

Skill 是工作指令与参考资源，不能自动增加底层工具。版本名称表示背景策略，不构成任何客户端或模型版本的能力保证。缺失能力时报告具体未满足项，不静默降级。

## 本次实际使用的参考环境

| 项目 | 实际情况 |
| --- | --- |
| 维护宿主 | Codex 桌面环境，能够读写文件、检索网络及调用内置图像生成工具 |
| 操作系统 | Windows，使用本机 Microsoft PowerPoint |
| PowerPoint | Version 16.0，Build 20430 |
| 示例制作方式 | 原生 PowerPoint COM 自动化，文字、表格、图片及动画为实际 PPTX 对象 |
| Python | 3.11.15 |
| 示例构建依赖 | pywin32 311、Pillow 12.3.0 |
| 实际中文标题字体 | 方正小标宋简体（字库内部名称 FZXiaoBiaoSong-B05S） |
| 正文与拉丁字体 | 仿宋（FangSong）、Times New Roman |
| 视频检查工具 | FFmpeg / FFprobe 8.1.1，用于检查原生导出视频及提取审阅帧 |
| 实时演示会话 | 本次连接失败，未使用；未因此宣称其已通过测试 |

字体名称、版本和输出结果对应本次维护机器，不意味着其他机器默认安装这些字体。仓库不包含字体文件。

## 已验证与未验证

| 项目 | 状态 | 证据与范围 |
| --- | --- | --- |
| 实际生成图片背景 | 已执行 | ChatGPT 策略示例的 4 张背景及生成提示词记录 |
| 实际网络图片背景 | 已执行 | DeepSeek 策略示例的 4 张图片、原始页面与 CC0 许可记录 |
| 本机 PowerPoint 打开与导出 | 已验证 | 两组原生 PPTX、PDF、逐页 PNG 和动画 MP4 |
| 指定字体、非加粗正文、图片嵌入与动画对象 | 已检查 | 参见示例构建记录及发布验证说明 |
| 缩时动画视频 | 已导出并审阅 | PowerPoint 原生渲染；自动触发预览，PPTX 保留点击推进 |
| 完整授课时长 | 未实测 | 每组短例约 5 分钟为讲述规划，非排练记录 |
| ChatGPT 网页或移动客户端中的导入及端到端执行 | 未独立验证 | 当前生成能力来自维护宿主，不能外推至其他客户端 |
| DeepSeek 官方客户端中的端到端执行 | 未独立验证 | 公开示例由维护宿主按其搜图策略制作，不是 DeepSeek 模型实测结果 |
| WPS、macOS PowerPoint、Google Slides、LibreOffice | 未验证 | 不保证字体映射、动画或导入表现一致 |

这两组示例展示资源策略与成品结构，不是两个模型之间的效果、耗时或成功率对比。

## 安装与故障处理

完整仓库的安装步骤见 [快速开始](https://github.com/Delicious-Cao/Delicious-PPT/blob/main/docs/QUICKSTART.md)。独立 Skill ZIP 应保留整个技能目录；普通对话方式需要同时提供 `SKILL.md` 与其参考资源。

示例构建脚本依赖 Windows 桌面版 PowerPoint。先确认应用可以正常启动、字体存在、输出位置可写，再使用 `examples/requirements.txt` 中的依赖。不要把运行该脚本当作对任意模型客户端的端到端验证。

反馈时提供所选 Skill 版本、宿主、放映软件版本、复现步骤与脱敏的小文件。不要上传访问凭据、私有文档或无权分享的字体。

原生接口参考：[PowerPoint CreateVideo](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.createvideo)、[文字效果 AddEffect](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.sequence.addeffect)。

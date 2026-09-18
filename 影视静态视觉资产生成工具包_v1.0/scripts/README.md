# 脚本说明

## New-VisualProject.ps1

从`project_template`创建全新的视觉资产项目，不覆盖已有目录，并自动替换模板中的项目名称。

## Validate-VisualProject.ps1

读取新项目的`09_正式资产索引.json`，检查：

- JSON是否有效。
- 正式角色、道具、场景和关键帧文件是否存在。
- PNG/JPG是否可读取。
- 是否符合索引指定尺寸；未指定尺寸时检查16:9比例。

脚本只做文件和技术规格检查，人物身份、空间关系和剧情逻辑仍需使用`07_连续性审核表.md`人工复审。

## Build-ContactSheets.ps1

按照索引中的`official_keyframes`顺序生成联系表。默认每张联系表为2列×3行，并显示关键帧ID和正式版本。

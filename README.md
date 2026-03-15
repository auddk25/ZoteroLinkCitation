# ZoteroLinkCitation

一个 VBA 宏，用于在 Microsoft Word 文档中为 Zotero 生成的正文引注自动创建超链接，点击引注即可跳转到文末对应的参考文献条目。

## 功能

- 自动识别文档中所有 Zotero 引注域（`ZOTERO_ITEM`）和参考文献列表（`ZOTERO_BIBL`）
- 为每条正文引注创建指向对应参考文献的超链接（书签跳转）
- 支持合并引注（如 `[1-3]`、`[Author, 2024; Author, 2025]`）
- 自动处理 Unicode 转义字符（`\uXXXX`）和 JSON 转义引号
- 兼容 en-dash / em-dash 等破折号变体
- 超链接样式保持黑色无下划线，不影响论文排版
- 支持重复运行（自动清除已有超链接后重新生成）

## 使用方法

1. 在 Word 中按 `Alt + F11` 打开 VBA 编辑器
2. 将 `ZoteroLinkCitation.bas` 导入为模块（文件 → 导入文件），或将代码复制到新模块中
3. 关闭 VBA 编辑器，回到 Word 文档
4. 按 `Alt + F8`，选择 `ZoteroLinkCitation`，点击"运行"

## 前置条件

- 文档中已通过 Zotero 插件插入引注和参考文献列表
- 引注使用 Zotero 域代码格式（非纯文本）

## 注意事项

- 单条引注中最多支持 200 个引用项
- 书签名长度限制为 40 字符，超长标题会自动截断并避免冲突
- 运行前请先保存文档，以便在需要时撤销更改

## License

MIT

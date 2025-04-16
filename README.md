# 手写模拟器

一个将电子文本转换为逼真手写效果的应用程序，基于Python和Handright库开发。

![手写模拟器预览](https://user-images.githubusercontent.com/your_username/your_repository/raw/main/screenshot.png)

## 功能特色

- **逼真的手写效果**：使用Handright库生成接近真人书写的效果
- **多种格式支持**：
  - 导入：支持TXT、DOCX和PDF格式文件
  - 导出：支持PNG、JPEG、PDF和DOCX格式
- **高度可定制**：
  - 可调整字体、字号、边距等基本参数
  - 支持多种扰动参数设置，使手写效果更加自然
  - 可添加自定义背景，模拟不同类型的纸张
- **批量处理**：
  - 批量导入：一次导入多个文件并合并
  - 批量导出：支持导出多种格式或将长文本分块导出
- **预设系统**：保存和加载不同风格的手写参数
- **实时预览**：立即查看参数变化效果，无需等待导出
- **深浅主题**：支持深色和浅色两种界面主题

## 安装说明

### 环境要求

- Python 3.6+
- PySide6
- Handright
- Pillow (PIL)
- PyMuPDF (fitz)
- python-docx

### 安装步骤

1. 克隆项目仓库
   ```bash
   git clone https://github.com/your_username/handwrite-simulator.git
   cd handwrite-simulator
   ```

2. 安装依赖
   ```bash
   pip install -r requirements.txt
   ```

3. 运行程序
   ```bash
   python handwrite_app.py
   ```

### 可选：创建虚拟环境

```bash
python -m venv venv
source venv/bin/activate  # 在Windows上使用 venv\Scripts\activate
pip install -r requirements.txt
```

## 使用指南

### 基本操作

1. **导入文本**：
   - 点击"导入文件"按钮，选择TXT、DOCX或PDF文件
   - 或直接在文本框中输入内容

2. **调整参数**：
   - 在左侧面板调整字体、边距和扰动设置
   - 如果有自定义字体，请将TTF字体文件放入"fonts"目录
   - 如需自定义背景，请将图片放入"backgrounds"目录

3. **预览效果**：
   - 点击"预览"按钮或选择"预览"选项卡即可查看效果
   - 使用导航按钮浏览多页预览

4. **导出**：
   - 点击"导出"按钮，选择所需的导出格式
   - 支持PNG、JPEG、PDF和DOCX格式

### 批量处理

1. **批量导入**：
   - 点击"批量导入文件"，可一次选择多个文件
   - 可选择替换或追加到现有文本

2. **批量导出**：
   - 点击"批量导出"按钮
   - 可同时选择多种导出格式
   - 可设置是否拆分长文本

### 参数说明

- **边距设置**：
  - 上下左右边距：调整文字在页面中的位置
  - 字间距：调整字符之间的间距
  - 行间距：调整行与行之间的距离（必须大于字体大小）

- **扰动设置**：
  - 字间距扰动：使字间距产生随机变化
  - 行间距扰动：使行间距产生随机变化
  - 字体大小随机扰动：使字体大小产生随机变化
  - 笔画横向/纵向偏移随机扰动：使笔画位置产生随机偏移
  - 笔画旋转偏移随机扰动：使笔画角度产生随机变化

## 常见问题

1. **为什么预览/导出时报错？**
   - 最常见的原因是行间距小于字体大小，程序会自动调整并提示
   - 检查字体文件是否可用

2. **如何获取更多手写字体？**
   - 将TTF格式的手写字体文件放入"fonts"目录
   - 重启程序或点击刷新按钮

3. **如何使用自定义背景？**
   - 将PNG或JPG格式的背景图像放入"backgrounds"目录
   - 在基本设置中选择背景

## 贡献指南

欢迎贡献代码、报告问题或提出改进建议！请遵循以下步骤：

1. Fork 本仓库
2. 创建您的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开一个 Pull Request

## 许可证

本项目采用 MIT 许可证 - 详情见 [LICENSE](LICENSE) 文件

## 致谢

- [Handright](https://github.com/Gsllchb/Handright) - 手写效果生成库
- [PySide6](https://doc.qt.io/qtforpython-6/) - Qt for Python
- 感谢所有贡献者和用户的支持！

---

**作者：** 赵伟
**联系方式：** 851906121@qq.com 
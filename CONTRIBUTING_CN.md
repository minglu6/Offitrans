# 贡献指南

感谢您对 Offitrans 项目的关注！我们欢迎任何形式的贡献，包括但不限于：

- 🐛 报告 Bug
- 💡 提出新功能建议
- 📝 改进文档
- 🔧 提交代码修复
- ✨ 添加新功能
- 🌍 翻译和本地化

[English Contributing Guide](CONTRIBUTING.md)

## 📋 贡献前准备

### 开发环境设置

1. **Fork 项目**
   ```bash
   # 在GitHub上Fork项目到您的账户
   # 然后克隆您的Fork
   git clone https://github.com/minglu6/Offitrans.git
   cd Offitrans
   ```

2. **设置开发环境**
   ```bash
   # 创建虚拟环境
   python -m venv venv
   
   # 激活虚拟环境
   # Windows
   venv\Scripts\activate
   # macOS/Linux
   source venv/bin/activate
   
   # 安装依赖
   pip install -r requirements.txt
   ```

3. **创建功能分支**
   ```bash
   git checkout -b feature/your-feature-name
   # 或者
   git checkout -b fix/your-fix-name
   ```

## 🐛 报告 Bug

如果您发现了 Bug，请通过 [GitHub Issues](https://github.com/your-username/Offitrans/issues) 报告。

**Bug 报告应包含：**

- 🔍 **清晰的标题和描述**
- 📱 **运行环境信息**（Python版本、操作系统等）
- 📝 **重现步骤**
- 🎯 **期望行为 vs 实际行为**
- 📋 **相关的错误日志或截图**
- 📄 **示例文件**（如果涉及特定的Office文件）

### Bug 报告模板

```markdown
## Bug 描述
简洁清晰地描述这个bug。

## 重现步骤
1. 执行 '...'
2. 点击 '....'
3. 滚动到 '....'
4. 看到错误

## 期望行为
清晰简洁地描述您期望发生的事情。

## 实际行为
清晰简洁地描述实际发生的事情。

## 环境信息
- OS: [例如 Windows 10, macOS 12.1, Ubuntu 20.04]
- Python版本: [例如 3.9.7]
- Offitrans版本: [例如 1.0.0]

## 附加信息
添加其他关于问题的上下文信息。
```

## 💡 功能建议

我们欢迎新功能建议！请通过 [GitHub Issues](https://github.com/your-username/Offitrans/issues) 提交。

**功能建议应包含：**

- 🎯 **问题描述**：您想解决什么问题？
- 💡 **解决方案**：您建议的功能如何解决这个问题？
- 🔄 **替代方案**：您考虑过其他解决方案吗？
- 📊 **使用场景**：谁会使用这个功能，在什么情况下使用？

## 🔧 代码贡献

### 编码规范

1. **代码风格**
   - 遵循 [PEP 8](https://www.python.org/dev/peps/pep-0008/) 编码规范
   - 使用 4 个空格进行缩进
   - 行长度限制为 88 字符（Black 默认设置）

2. **命名规范**
   - 类名使用 `PascalCase`
   - 函数和变量名使用 `snake_case`
   - 常量使用 `UPPER_CASE`
   - 私有方法和属性以 `_` 开头

3. **文档字符串**
   ```python
   def translate_text(self, text: str, target_language: str = 'en') -> str:
       """
       翻译文本内容
       
       Args:
           text: 要翻译的文本
           target_language: 目标语言代码
           
       Returns:
           翻译后的文本
           
       Raises:
           ValueError: 当输入参数无效时
           TranslationError: 当翻译失败时
       """
   ```

4. **类型提示**
   - 所有公共方法都应该有类型提示
   - 使用 `typing` 模块进行复杂类型定义

### 代码质量工具

在提交代码前，请使用以下工具检查代码质量：

```bash
# 代码格式化
black .

# 代码风格检查
flake8 .

# 运行测试
pytest tests/ -v --cov=.
```

### 提交规范

使用 [Conventional Commits](https://www.conventionalcommits.org/) 规范：

- `feat:` 新功能
- `fix:` Bug修复
- `docs:` 文档更新
- `style:` 代码格式调整
- `refactor:` 代码重构
- `test:` 测试相关
- `chore:` 构建过程或辅助工具的变动

**示例：**
```bash
git commit -m "feat: 添加PDF翻译支持图片保护功能"
git commit -m "fix: 修复Excel合并单元格翻译格式问题"
git commit -m "docs: 更新API使用文档"
```

### Pull Request 流程

1. **确保代码质量**
   - 所有测试通过
   - 代码风格检查通过
   - 新功能有对应的测试

2. **创建 Pull Request**
   - 提供清晰的标题和描述
   - 解释更改的原因和内容
   - 如果修复了 Issue，请在 PR 中引用

3. **PR 描述模板**
   ```markdown
   ## 更改类型
   - [ ] Bug 修复
   - [ ] 新功能
   - [ ] 文档更新
   - [ ] 代码重构
   - [ ] 性能改进
   
   ## 更改描述
   清晰描述本次PR的更改内容
   
   ## 相关Issue
   修复 #123
   
   ## 测试
   - [ ] 新增了单元测试
   - [ ] 现有测试全部通过
   - [ ] 手动测试通过
   
   ## 检查清单
   - [ ] 代码遵循项目编码规范
   - [ ] 已添加必要的文档和注释
   - [ ] 所有测试通过
   ```

## 🧪 测试

### 运行测试

```bash
# 运行所有测试
pytest

# 运行特定测试文件
pytest tests/test_excel_translate.py

# 运行测试并生成覆盖率报告
pytest --cov=. --cov-report=html
```

### 编写测试

- 每个新功能都应该有对应的测试
- 测试文件命名为 `test_*.py`
- 测试方法命名为 `test_*`

```python
def test_translate_excel_basic():
    """测试基本Excel翻译功能"""
    translator = ExcelTranslator()
    result = translator.translate_text("你好", "en")
    assert result == "Hello"
```

## 📝 文档贡献

### 文档类型

- **API 文档**：代码中的文档字符串
- **用户指南**：README.md 和使用示例
- **开发文档**：CONTRIBUTING.md 和技术说明

### 文档规范

- 使用简洁明了的语言
- 提供实际的代码示例
- 保持中英文文档同步更新

## 🌍 国际化贡献

我们欢迎多语言支持的贡献：

- 翻译文档到其他语言
- 添加新的翻译语言支持
- 改进现有语言的翻译质量

## 🎯 项目优先级

当前项目的重点关注领域：

1. **稳定性改进** - 修复现有功能的Bug
2. **性能优化** - 提高翻译速度和内存使用效率
3. **格式保持** - 改进各种Office格式的样式保持
4. **新格式支持** - 添加对更多文件格式的支持
5. **多翻译引擎** - 集成更多翻译服务

## 📞 获取帮助

如果您有任何问题或需要帮助：

- 📝 创建 [GitHub Issue](https://github.com/your-username/Offitrans/issues)
- 💬 参与 [Discussions](https://github.com/your-username/Offitrans/discussions)

## 🙏 贡献者名单

感谢所有为 Offitrans 做出贡献的开发者！

<!-- 这里可以添加贡献者列表，或者使用GitHub的contributors API -->

---

再次感谢您对 Offitrans 项目的贡献！🚀
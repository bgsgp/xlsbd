# 🎉 积分备份器 v2.2.1.3 发布（独立 exe 版）

现在，**积分备份器** 已打包为独立 Windows 可执行文件，无需安装 Python 或任何依赖，下载即可使用。

---

### 🆕 本次亮点

- **独立 exe 运行**  
  下载 `积分备份器.exe` 即可直接使用，不再需要配置 Python 环境，适合任何 Windows 电脑。

- **按日期自动备份**  
  自动将源文件复制到目标目录，按 `score_YYYYMMDD_序号.xlsx` 命名，支持每日多次备份，不会相互覆盖。

- **全屏查看日志**  
  备份完成后自动调用 Excel 并全屏打开，定位到 LOG 工作簿的最后一行，方便快速回顾历史记录。

- **文件属性修复**  
  自动移除备份文件的隐藏和系统属性，确保在任意文件管理器中都能看到。

---

### 📥 下载

- **主程序（exe）**：  
  [积分备份器.exe](https://github.com/bgsgp/xlsbd/releases/download/v2.2.1.3/积分备份器.exe)

- **源代码（Python）**：  
  [积分备份器.py](https://github.com/bgsgp/xlsbd/releases/download/v2.2.1.3/积分备份器.py) （供有 Python 环境的用户定制）

---

### ⚙️ 如何使用

1. **配置备份路径**  
   用文本编辑器（如记事本）打开 `积分备份器.py` 源代码文件，修改以下两行路径：
   ```python
   SRC = r"D:\Program Files\bggp\1\scorer\score.xlsx"   # 源文件绝对路径
   DSTDIR = os.path.join(os.environ['USERPROFILE'], 'OneDrive', '文档', 'xslbd')  # 备份存放目录
   ```
   *注：exe 版本已将上述路径硬编码打包，若有不同路径需求，请直接使用 `.py` 源码修改并运行。*

2. **开始备份**  
   双击 `积分备份器.exe`，它将：
   - 检查源文件存在
   - 自动创建备份目录（如有需要）
   - 生成带日期和序号的备份文件
   - 打开备份文件并全屏定位到 LOG 表最新行

3. **自动运行（可选）**  
   可将该 exe 加入 Windows 任务计划程序，实现每天定时自动备份。

---

### 🧩 与“乞分君”联动

该工具是 [ScoreLauncher（乞分君）](https://github.com/bgsgp/scorerlauncher) 的官方配套备份方案。  
日常工作流推荐：

　① 在“乞分君”中完成积分录入  
　② 双击运行 `积分备份器.exe`，立即生成今日备份  
　③ 弹出的 Excel 窗口会自动展示最新 LOG，便于核对

---

### 📝 完整更新日志

- 提供独立 `.exe` 文件，移除 Python 依赖
- 支持全屏打开及自动滚动到最新记录
- 自动清除文件隐藏/系统属性
- 包含 `version.txt` 用于记录版本信息

---

> 🔗 项目仓库：[https://github.com/bgsgp/xlsbd](https://github.com/bgsgp/xlsbd)  
> 如有问题，欢迎提交 Issue。

# StartFuckingWork

一个自动化脚本，用于快速启动常用办公软件，包括飞书、Teams和微信多开。

## 功能特点

- 🚀 **一键启动**：自动查找并启动飞书、Teams和微信
- 💬 **微信多开**：支持微信多账户同时登录
- 🔍 **智能查找**：自动搜索软件安装路径，无需手动配置
- 🎯 **支持多版本**：支持Windows桌面版和Windows App版Teams
- 🧹 **自动清理**：执行完毕后自动关闭临时窗口
- ⏱️ **快速退出**：执行完毕后2秒自动退出脚本

## 使用方法

### 方法一：直接运行

1. 双击 `StartFuckingWork.ps1` 文件
2. 如果出现安全提示，点击 "更多信息" → "运行 anyway"
3. 脚本会自动执行并启动相关软件

### 方法二：通过PowerShell运行

1. 打开PowerShell终端
2. 导航到脚本所在目录：
   ```powershell
   cd "C:\Users\boe\Desktop\StartFuckingWork"
   ```
3. 执行脚本：
   ```powershell
   .\StartFuckingWork.ps1
   ```

## 脚本工作原理

1. **微信多开**：
   - 查找微信安装路径
   - 打开微信所在目录
   - 模拟按键操作启动多个微信实例
   - 直接多次启动微信进程（最可靠的方法）

2. **启动Teams**：
   - 首先尝试启动桌面版Teams
   - 如果失败，尝试通过URI协议启动Windows App版
   - 如果仍失败，尝试通过AppxPackage启动

3. **启动飞书**：
   - 查找飞书安装路径并启动

4. **清理操作**：
   - 关闭微信资源管理器窗口
   - 显示启动状态汇总
   - 2秒后自动退出脚本

## 支持的软件路径

### 微信路径
- `C:\Program Files (x86)\Tencent\WeChat\WeChat.exe`
- `C:\Program Files\Tencent\WeChat\WeChat.exe`
- `C:\Program Files\Tencent\Weixin\Weixin.exe`
- `C:\Program Files (x86)\Tencent\Weixin\Weixin.exe`
- `%LOCALAPPDATA%\Programs\Tencent\WeChat\WeChat.exe`
- `%APPDATA%\Tencent\WeChat\WeChat.exe`
- `%LOCALAPPDATA%\Programs\Tencent\Weixin\Weixin.exe`
- `%APPDATA%\Tencent\Weixin\Weixin.exe`

### 飞书路径
- `%LOCALAPPDATA%\Feishu\Feishu.exe`
- `%APPDATA%\Feishu\Feishu.exe`
- `C:\Program Files (x86)\Feishu\Feishu.exe`
- `%USERPROFILE%\AppData\Local\Feishu\Feishu.exe`

### Teams路径
- `%LOCALAPPDATA%\Microsoft\Teams\current\Teams.exe`
- `%APPDATA%\Microsoft\Teams\current\Teams.exe`
- `C:\Program Files (x86)\Microsoft\Teams\current\Teams.exe`
- `C:\Program Files\Microsoft\Teams\current\Teams.exe`
- Windows App版（通过URI协议或AppxPackage）

## 注意事项

1. **安全提示**：首次运行时可能会出现PowerShell执行策略提示，需要允许脚本运行
2. **权限要求**：脚本需要基本的文件系统访问权限
3. **多开效果**：微信多开可能需要在第一次登录时手动扫码或输入密码
4. **启动顺序**：脚本会先执行微信多开，然后启动Teams，最后启动飞书
5. **路径查找**：如果软件安装在非标准路径，可能需要手动修改脚本中的路径列表

## 故障排除

### 微信多开失败
- 检查微信是否已安装
- 确认微信安装路径是否在支持列表中
- 尝试手动运行微信，确保微信可以正常启动

### Teams启动失败
- 检查Teams是否已安装
- 确认Teams版本（桌面版或Windows App版）
- 尝试手动启动Teams，确保Teams可以正常运行

### 飞书启动失败
- 检查飞书是否已安装
- 确认飞书安装路径是否在支持列表中
- 尝试手动启动飞书，确保飞书可以正常运行

## 自定义配置

如果需要修改脚本行为，可以编辑 `StartFuckingWork.ps1` 文件：

1. **添加新的软件路径**：在相应的路径数组中添加新路径
2. **调整启动顺序**：修改脚本中的执行顺序
3. **修改等待时间**：调整 `Start-Sleep` 命令的参数
4. **添加新软件**：按照现有模式添加新软件的查找和启动逻辑

## 许可证

本项目采用MIT许可证，详见 [LICENSE](LICENSE) 文件。

## 贡献

欢迎提交Issue和Pull Request来改进这个项目！

## 更新日志

### v1.0.0
- 初始版本
- 支持微信多开
- 支持启动Teams（桌面版和Windows App版）
- 支持启动飞书
- 自动查找软件路径
- 自动清理临时窗口

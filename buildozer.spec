[app]

# (str) 应用名称（显示在手机图标下方）
title = 记账软件

# (str) 包名（必须唯一，请将 org.example 改为你自己的域名）
package.name = accountbook
package.domain =com.zhh11281128

android.sdk_path = /usr/local/lib/android/sdk
android.ndk_path = /usr/local/lib/android/sdk/ndk/25.2.9519653
# (str) 应用版本号
version = 0.1

# (list) 应用依赖的 Python 模块（根据 import 添加）
requirements = python3,kivy,plyer,openpyxl

# (str) 源代码目录
source.dir = .

# (list) 需要包含的额外文件扩展名（确保字体 .ttf 被包含）
source.include_exts = py,png,jpg,kv,atlas,ttf

# (str) 入口文件
source.main = main.py

# ----------------- Android 配置 -----------------
# (int) 目标 API 级别（与工作流中一致）
android.api = 33

# (int) 最低支持 API 级别
android.minapi = 21

# (str) NDK 版本（工作流中会通过环境变量指定具体路径）
android.ndk = 25b

# (str) 构建工具版本（与工作流中一致）
android.build_tools = 34.0.0

# (bool) 启用 AndroidX
android.use_support_library = True

# (str) 可选：指定 SDK 路径（由环境变量覆盖，这里留空）
# android.sdk_path = 
# android.ndk_path = 
# 强制竖屏
android.orientation = portrait

# 添加存储权限（读写外部存储）
android.permissions = INTERNET, READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE
[buildozer]
log_level = 2
warn_on_root = 1

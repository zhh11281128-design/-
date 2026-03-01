[app]

# (str) 应用名称（显示在手机图标下方）
title = 记账软件

# (str) 包名（必须唯一，通常使用反向域名，如 com.yourname.accountbook）
package.name = accountbook
package.domain = com.zhh11281128

# (str) 应用版本号
version = 0.1

# (list) 应用依赖的 Python 模块（根据你的 main.py 中 import 的内容添加）
requirements = python3,kivy,plyer,openpyxl

# (str) 应用源代码目录（默认为当前目录）
source.dir = .

# (list) 需要包含的额外文件扩展名（确保字体文件 .ttf 被包含）
source.include_exts = py,png,jpg,kv,atlas,ttf

# (list) 需要排除的目录或文件（可选）
# source.exclude = docs, tests

# (str) 应用入口文件（默认为 main.py）
source.main = main.py

# (str) 应用图标（建议使用 512x512 的 PNG 图片，放置于项目目录）
# icon.filename = %(source.dir)s/icon.png

# (str) 应用启动图片（可选）
# presplash.filename = %(source.dir)s/splash.png

# ----------------- Android 配置 -----------------
# (int) 目标 Android API 级别（应与 GitHub Actions 中的 api-level 一致）
android.api = 33

# (int) 最低支持的 Android API 级别
android.minapi = 21

# (str) NDK 版本（应与 GitHub Actions 中的 ndk 版本对应，25b 对应 25.2.9519653）
android.ndk = 25b

# (str) 构建工具版本（与 GitHub Actions 中的 build-tools 一致）
android.build_tools = 34.0.0

# (list) 应用需要的 Android 权限
android.permissions = INTERNET

# (bool) 是否允许应用读取外部存储（如需要读写文件）
# android.permissions = WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE

# (bool) 是否启用 AndroidX（Kivy 默认使用支持库）
android.use_support_library = True

# (str) Android 应用分类（可选）
# android.category = GAME

# (list) 添加额外的 JAR 文件或 AAR 库（可选）
# android.add_src =

# (list) 添加额外的资源文件到 res 目录（可选）
# android.add_res =

# (list) 添加额外的资产文件到 assets 目录（可选）
# android.add_assets =

# (bool) 是否生成 AAB 文件（Android App Bundle，用于 Google Play 发布）
# android.aab = False

# ----------------- 其他平台配置（可忽略） -----------------
[buildozer]

# (int) 日志级别（0=仅错误,1=警告,2=信息,3=调试）
log_level = 2

# (bool) 当以 root 权限运行时是否警告（在 Docker 或 CI 中可设为 0 以忽略）
warn_on_root = 1

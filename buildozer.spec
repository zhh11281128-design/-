[app]

# 应用名称
title = 记账软件

# 包名（请将 org.example 改为你自己的域名，如 com.zhh11281128）
package.name = accountbook
package.domain = org.example

# 源代码目录
source.dir = .

# 需要包含的文件扩展名
source.include_exts = py,png,jpg,kv,atlas,ttf

# 应用依赖库
requirements = python3,kivy,plyer,openpyxl

# 版本号
version = 0.1

# -------------------- Android 配置 --------------------
# 目标 API 级别（使用稳定的 33，不要用 34 或更高，因为可能没有对应 platform）
android.api = 33

# 最低支持 API 级别
android.minapi = 21

# NDK 版本（与 python-for-android 推荐一致）
android.ndk = 25b

# 明确指定 build-tools 版本（避免使用预览版）
android.build_tools = 34.0.0

# 权限（根据你的应用需要添加）
android.permissions = INTERNET

# 其他配置保持默认
# ----------------------------------------------------

[buildozer]
log_level = 2
warn_on_root = 1

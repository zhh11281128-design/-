[app]

# 应用名称（中文也可以，但最好用英文）
title = 记账软件

# 包名（必须唯一，请将 "org.example" 替换为你自己的域名，如 "com.yourname"）
package.name = accountbook
package.domain = com.xs.accountbook

# 源代码目录
source.dir = .

# 需要包含的文件扩展名（确保 .ttf 被包含）
source.include_exts = py,png,jpg,kv,atlas,ttf

# 应用依赖库（根据你的 main.py 中 import 的内容添加）
requirements = python3,kivy,plyer,openpyxl

# 版本号
version = 0.1

# 是否为发布版（debug 版用于测试）
osx.python_version = 3
osx.kivy_version = 2.3.0

# Android 特定配置
android.permissions = INTERNET
android.api = 33
android.minapi = 21
android.ndk = 25b
android.sdk = 34
android.build_tools = 34.0.0
# 如果你的应用使用中文，可能需要设置默认编码
android.add_src =

# 其他保持默认即可

[buildozer]
log_level = 2
warn_on_root = 1

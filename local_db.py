[app]
title = KSMS Billing
package.name = ksmsbilling
package.domain = com.guhanenterprises
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,ttf
version = 1.0.0

# Everything the app imports at runtime. reportlab/openpyxl/Pillow all have
# working python-for-android recipes as of Kivy 2.3 / p4a master -- if a build
# fails on one of these, see the troubleshooting notes in WINDOWS_AND_ANDROID_BUILD.md.
requirements = python3,kivy==2.3.0,kivymd==1.2.0,reportlab,pillow,openpyxl,pyjnius,android

orientation = portrait
fullscreen = 0

android.permissions = INTERNET, WRITE_EXTERNAL_STORAGE, READ_EXTERNAL_STORAGE
android.api = 34
android.minapi = 24
android.ndk = 25b
android.archs = arm64-v8a, armeabi-v7a
android.allow_backup = True

# Uses scoped/app-private storage (see core/billing_core.get_app_base_dir),
# so no special "all files access" permission is needed.

[buildozer]
log_level = 2
warn_on_root = 1

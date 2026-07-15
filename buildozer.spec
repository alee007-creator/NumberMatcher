[app]

# App identity
title = Bond Matcher
package.name = bondmatcher
package.domain = org.aliraza

# Source
source.dir = .
source.include_exts = py,png,jpg,kv,atlas
version = 1.0

# Python/Kivy requirements (NO pandas/numpy on purpose - openpyxl is pure Python
# and builds cleanly for Android; pandas is huge and usually fails to compile).
requirements = python3,kivy==2.3.1,openpyxl,et_xmlfile,plyer,pyjnius,android

# Portrait phone app
orientation = portrait
fullscreen = 0

# Permissions: read files the user picks; MANAGE_EXTERNAL_STORAGE covers Android 11+
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,MANAGE_EXTERNAL_STORAGE

# Android API / build targets
android.api = 34
android.minapi = 24
android.ndk_api = 24
android.archs = arm64-v8a,armeabi-v7a

# Auto-accept the SDK license (needed for unattended CI builds)
android.accept_sdk_license = True

# Bootstrap
p4a.bootstrap = sdl2

[buildozer]
log_level = 2
warn_on_root = 1

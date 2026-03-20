#!/bin/bash
# 打包为 macOS .app + .dmg
# 运行方式：bash build_app.sh

set -e
cd "$(dirname "$0")"

APP_NAME="PPT转图片"
ARCH=$(uname -m)

echo "=========================================="
echo "  PPT 批量导出图片工具 — macOS 打包脚本"
echo "  当前架构: $ARCH"
echo "=========================================="
echo ""
echo "▶ 步骤 1/3  PyInstaller 打包 .app ..."

# 注入构建标识
BUILD=$(git rev-parse --short HEAD 2>/dev/null || echo "unknown")
echo "BUILD = \"${BUILD}\"" > _build_info.py
echo "  构建标识: $BUILD"

pyinstaller \
  --windowed \
  --name "$APP_NAME" \
  --noconfirm \
  --add-data "_build_info.py:." \
  main.py

echo ""
echo "▶ 步骤 2/3  移除隔离属性（本机测试用）..."
xattr -cr "dist/$APP_NAME.app" 2>/dev/null || true

echo ""
echo "▶ 步骤 3/3  打包为 .dmg ..."

DMG_NAME="${APP_NAME}_${ARCH}.dmg"
DMG_TMP="dist/dmg_tmp"
DMG_OUT="dist/$DMG_NAME"

rm -rf "$DMG_TMP" "$DMG_OUT"
mkdir -p "$DMG_TMP"
cp -r "dist/$APP_NAME.app" "$DMG_TMP/"

hdiutil create \
  -volname "$APP_NAME" \
  -srcfolder "$DMG_TMP" \
  -ov \
  -format UDZO \
  "$DMG_OUT"

rm -rf "$DMG_TMP"

echo ""
echo "=========================================="
echo "  打包完成！"
echo ""
echo "  .app 路径：dist/$APP_NAME.app"
echo "  .dmg 路径：dist/$DMG_NAME"
echo ""
echo "  发给其他 Mac 用户时："
echo "  1. 发送 $DMG_NAME 文件"
echo "  2. 双击挂载 DMG，将 .app 拖入「应用程序」"
echo "  3. 首次打开：右键点击 .app → 打开 → 打开"
echo ""
echo "  当前架构 $ARCH — 此包只能在同架构 Mac 上运行"
echo "=========================================="

#!/usr/bin/env bash
# 用 LibreOffice 把目录下所有 docx 转成 PDF（注意：手册的 Word 目录域需先在 Word 里更新并保存）
# 用法: bash docx_to_pdf.sh 正式资料/
set -euo pipefail
dir="${1:-.}"
for f in "$dir"/*.docx; do
  [ -e "$f" ] || continue
  soffice --headless --convert-to pdf --outdir "$dir" "$f" >/dev/null 2>&1 && echo "✅ ${f%.docx}.pdf"
done
for p in "$dir"/*.pdf; do
  [ -e "$p" ] || continue
  sz=$(du -m "$p" | cut -f1); echo "   $(basename "$p")  ${sz}MB $( [ "$sz" -gt 50 ] && echo '⚠️ 超过 50MB' )"
done

#!/bin/bash
# 雙擊執行：把 dict-tool 的最新修改 commit + push → GitHub Actions 在 Windows 上打包
# → 下載「國文辭典查詢工具.exe」放到桌面（舊的移到垃圾桶）。
# Mac 沒辦法直接做 Windows exe，所以一定要走 GitHub Actions（workflow 在 repo root 的
# .github/workflows/build-dict-exe.yml）。需要 gh 已登入（gh auth status）。

export PATH="/opt/homebrew/bin:/usr/local/bin:${PATH}"
export LANG="${LANG:-zh_TW.UTF-8}" LC_ALL="${LC_ALL:-zh_TW.UTF-8}"
EXE_NAME="國文辭典查詢工具.exe"
ARTIFACT="國文辭典查詢工具-exe"
WORKFLOW="build-dict-exe.yml"

TOOL_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO="$(git -C "${TOOL_DIR}" rev-parse --show-toplevel)"
REL="${TOOL_DIR#${REPO}/}"

fail() {
  echo ""
  echo "❌ $1"
  osascript -e "display dialog \"打包失敗：$1\" buttons {\"好\"} default button 1 with icon stop with title \"國文辭典查詢工具\"" >/dev/null 2>&1
  echo "（按 Enter 關閉視窗）"; read -r
  exit 1
}
step() { echo ""; echo "▶ $1"; }

cd "${REPO}" || fail "找不到 git 專案"
command -v gh >/dev/null || fail "沒有安裝 gh（GitHub CLI）"
gh auth status >/dev/null 2>&1 || fail "gh 沒有登入，請先在終端機執行 gh auth login"

step "1/4 檢查程式有沒有語法錯誤"
python3 -m py_compile "${TOOL_DIR}/CNdictionary.py" "${TOOL_DIR}/moe_video.py" || fail "程式有語法錯誤，先修好再打包"
echo "  沒問題"

step "2/4 上傳最新修改到 GitHub"
START_TIME="$(date -u +%Y-%m-%dT%H:%M:%SZ)"
git add -- "${REL}" ".github/workflows/${WORKFLOW}"
if git diff --cached --quiet; then
  echo "  沒有新修改，直接用目前版本重新打包"
  git push -q origin master 2>/dev/null   # 之前 commit 了但還沒 push 的也一起送上去
  gh workflow run "${WORKFLOW}" --ref master >/dev/null || fail "無法啟動 GitHub 打包"
  SHA="$(git rev-parse HEAD)"
else
  git commit -q -m "Update 國文辭典查詢工具 (一鍵打包 $(date '+%Y-%m-%d %H:%M'))" || fail "commit 失敗"
  if ! git push -q origin master 2>/dev/null; then
    echo "  GitHub 上有比較新的版本，先合併再上傳"
    git pull -q --rebase --autostash origin master || fail "合併 GitHub 上的新版本失敗，需要手動處理"
    git push -q origin master || fail "上傳到 GitHub 失敗"
  fi
  SHA="$(git rev-parse HEAD)"
  echo "  已上傳 ${SHA:0:7}"
fi

step "3/4 等 GitHub 在 Windows 上打包（約 2 分鐘）"
RUN_ID=""
for _ in $(seq 1 30); do
  RUN_ID="$(gh run list --workflow "${WORKFLOW}" --limit 5 --json databaseId,headSha,createdAt \
    -q "[.[] | select(.headSha==\"${SHA}\" and .createdAt>=\"${START_TIME}\")][0].databaseId")"
  [ -n "${RUN_ID}" ] && break
  sleep 4
done
[ -n "${RUN_ID}" ] || fail "找不到這次的打包工作"
echo "  打包工作 #${RUN_ID}"
gh run watch "${RUN_ID}" --exit-status --interval 10 >/dev/null 2>&1 \
  || fail "GitHub 打包失敗，詳細錯誤：gh run view ${RUN_ID} --log-failed"
echo "  打包完成"

step "4/4 下載到桌面"
TMP="$(mktemp -d)"
gh run download "${RUN_ID}" -n "${ARTIFACT}" -D "${TMP}" || fail "下載失敗"
[ -f "${TMP}/${EXE_NAME}" ] || fail "下載的檔案裡沒有 ${EXE_NAME}"
if [ -f "${HOME}/Desktop/${EXE_NAME}" ]; then
  mv "${HOME}/Desktop/${EXE_NAME}" "${HOME}/.Trash/${EXE_NAME%.exe}_舊版_$(date '+%m%d-%H%M').exe" 2>/dev/null \
    || rm -f "${HOME}/Desktop/${EXE_NAME}"
  echo "  舊版已移到垃圾桶"
fi
mv "${TMP}/${EXE_NAME}" "${HOME}/Desktop/${EXE_NAME}"
rmdir "${TMP}" 2>/dev/null

SIZE="$(du -h "${HOME}/Desktop/${EXE_NAME}" | cut -f1)"
echo ""
echo "✅ 完成！桌面上的 ${EXE_NAME}（${SIZE}）已經是最新版"
osascript -e "display notification \"桌面上的 ${EXE_NAME} 已更新（${SIZE}）\" with title \"國文辭典查詢工具\" sound name \"Glass\"" >/dev/null 2>&1
open -R "${HOME}/Desktop/${EXE_NAME}"
echo "（按 Enter 關閉視窗）"; read -r

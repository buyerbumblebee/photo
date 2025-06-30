@echo off
setlocal

REM === 設定路徑與 GitHub 倉庫 ===
set REPO_URL=https://github.com/buyerbumblebee/photo.git
set LOCAL_DIR=E:\github\target
set BRANCH=main

cd /d %LOCAL_DIR%

REM === 初始化 Git（如尚未） ===
if not exist ".git" (
    git init
    git remote add origin %REPO_URL%
)

REM === 抓取最新資料，強制切換分支 ===
git fetch origin
git checkout -B %BRANCH% origin/%BRANCH%

REM === 刪除舊檔案 ===
git rm -rf . >nul 2>&1

REM === 加入所有新檔案（圖片） ===
git add .

REM === 提交並強制推送 ===
git commit -m "?? Replace all files with new images"
git push origin %BRANCH% --force

echo.
echo ? 強制推送完成！GitHub 倉庫已經完全同步 %LOCAL_DIR% 的內容。
pause

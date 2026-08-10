# Phase 1 v0.53.0 UI/UX再確認手順

対象ペルソナは、IT操作に不慣れで老眼の忙しい管理職、弱視でキーボード／Narratorを使う人、手指が震えやすく作業記憶に負荷をかけにくい人です。

## 自動確認

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\phase1\Test-Static.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests\phase1\Test-WebRender.ps1
node tests\phase1\Test-WebAssets.mjs
powershell -NoProfile -ExecutionPolicy Bypass -File tests\phase1\Test-RecorderCompanionCompile.ps1
```

## 画面確認

1. 1366x768・Windows 125～150%表示、またはブラウザー幅1093px・高さ614pxで、ヘッダー、左ナビ、確認ツールバーが重ならないことを確認する。
2. 200%ズームでページ全体に横スクロールが出ず、新規作成、記録開始、記録終了、Excel作成の各操作へ到達できることを確認する。
3. マニュアル一覧で「新しいマニュアル名」が見え、名前あり／空欄の両方をEnterで作成できることを確認する。
4. 記録中に×またはEscを押し、「記録を続ける」「終了して確認」「記録を捨てて閉じる」が結果の分かる言葉で表示されること、Escで記録を続けることを確認する。
5. 保存通信を失敗させ、入力値が残り、「保存できません」「再試行」が成功まで表示されることを確認する。
6. 出力前確認でExcelが唯一の主操作、Wordが印刷向け副操作として表示されることを確認する。
7. キーボードで記録の各工程へ進み、フォーカスが非表示要素や本文へ失われないことを確認する。

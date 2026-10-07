# acm-pptx

ACM Lab (NYCU) PowerPoint 投影片生成工具 / Codex 與 Claude Skill。

`main` 分支永遠是最新版本；每個歷史版本都有對應的 git tag，方便在新版不好用時退回舊版。

## 版本列表

| 版本 | 說明 |
|---|---|
| `v1.0.0` | 初版 |
| `v1.1.0` | lab-rules / outline-schema 更新 |
| `v1.2.0` | SKILL.md 與規則微調 |
| `v2.0.0` | 新增 paper-study 支援、blueprint-paper 等 references |
| `v2.1.0` | 新增 `qa_check --review`、Codex skill 支援 |
| `v2.2.0` | 紅框改為稀用、標題承載論點；QA 只渲染需要看的頁面；輸出只留一份 .pptx |
| `v2.3.0` | 支援嵌入影片；QA 改成大部分不用渲染就查得出來；contact sheet 把看圖成本降到約三分之一 |
| `v2.4.0` | **原生可編輯流程圖**：nodes / edges / groups 自動排版成 PowerPoint shapes 與 connectors；表格依內容分配欄寬；QA 驗證 diagram 不是扁平圖片 |
| `v2.5.0` | 目前最新版。**新增圖形版型** `cards` / `flow` / `bignum` / `quadrant`，表格改成會隨列數放大字；**每頁可指定自己的繪圖函式** `draw`；紅框座標統一（`at` 是 xywh，`xyxy` 是角點，寫錯會被擋）；`"ours": true` 標記自己的內容；QA 會抓「整份都是表格」「寬圖被塞進半欄」「論文報告沒有自己的觀點」 |

## 使用方式

抓最新版:

```bash
git clone https://github.com/wyattsheu/acm-pptx.git
```

只抓某個特定舊版本（不需要完整 git 歷史）:

```bash
git clone --branch v1.2.0 --depth 1 https://github.com/wyattsheu/acm-pptx.git
```

已經 clone 過，想切回某個舊版本:

```bash
git checkout v1.2.0
```

想切回最新版:

```bash
git checkout main
```

## v2.5.0 有什麼

這版是針對第一批用它做出來的論文報告的回饋：16 頁裡 7 頁表格、2 頁純條列，
表格字小、下半頁空白；論文的架構圖被壓到 5 吋寬看不清；整份都是論文的內容，
沒有我們自己做了什麼；紅框座標照文件寫還是畫錯兩次。

**圖形版型。** 以前一頁只能放條列、表格或一張圖，所以流程和對照全部變成表格。
現在 `outline.json` 多了四種 exhibit，都畫在該段落的色調裡，並把要講的那一格標成紅色：

| 欄位 | 畫什麼 | 用在 |
|---|---|---|
| `cards` | 2–6 張卡片，各有標題與幾行字 | related work 的幾個陣營、contributions、優缺點、能搬 / 不能搬到我們這邊的 |
| `flow` | 方框加箭頭的流程，一格打亮 | 方法總覽，之後每頁講一格；`"style": "chevron"`、`"direction": "column"` 可選 |
| `bignum` | 1–4 個大數字加標籤與 baseline | 整場報告的那一個結果，放在完整表格前面 |
| `quadrant` | 兩軸定位圖，prior work 是點，我們的是紅點 | 這篇論文在領域裡的位置 |

```jsonc
"flow": {"steps": [{"label": "FLAME tracking", "sub": "UV maps"}, "MGPM",
                   {"label": "Enhancer", "sub": "diffusion"}], "highlight": 2}
"bignum": {"items": [{"value": "20 min", "label": "per-identity fitting",
                      "sub": "CAP4D: 400 min", "highlight": true}]}
```

`matrix` 也重做了：字級隨列數決定（四列 18pt），列高撐到區域的七成，
短表格會置中而不是縮在上面；多了 `highlight_col`、`col_widths`、`caption`。

**自訂出口。** 四種版型不夠用時，一頁可以寫 `"draw": "design.py:timeline"`，
`compose.py` 會把那一頁的 slide、留給 exhibit 的區域（英寸）、該頁的 outline
和一組畫圖 helper 交給你的函式，標題、副標、條列、callout 都已經排好，
在區域裡畫就不會疊到。`assets/examples/design.py` 是參考實作，複製到 outline 旁邊改。

**座標只剩一種講法。** `annotations` 的 `at` 是 `[x, y, w, h]`，`figure.py crop`
的 `--box` 是 `x0,y0,x1,y1`，以前文件沒把這兩件事放在一起講。現在 annotation
和 `stage_figure` 都接受 `"xyxy": [x0, y0, x1, y1]`，有角點就直接寫角點不要換算；
`x + w` 超出圖片邊緣（就是把角點寫進 `at` 的樣子）會被 `qa_check.py` 和
`compose.py` 直接擋下來，訊息裡附對照表。

**標記自己的東西。** `"ours": true` 會在左下角畫一個紅色 `OUR TAKE` 小標
（給字串就畫那個字串），對應實驗室「作者主張 / 證據 / 報告者詮釋要分開」的規則。
論文報告沒有任何一頁是 `ours` 會被 QA 警告。

**QA 多抓的。** 表格超過內容頁的三分之一、連續三頁表格或三頁純文字、
長寬比超過 1.8 的圖被排進半欄（會算出它實際落在幾吋寬）、`draw` 指到不存在的檔案或函式。

**預設版型變了一點。** 沒寫 `layout` 時：有條列的 `matrix` / `cards` / `flow` /
`bignum` 排在條列下面，沒條列就佔滿；`quadrant` 排在右邊；
`text-only` 同時有條列和圖形時會自動改成上下排，不再疊在一起。

## v2.4.0 有什麼

**流程圖不再是圖片。** 在 outline 使用 `diagram`，`compose.py` 會建立原生
PowerPoint 方塊、決策菱形、資料庫、文字、群組框與連接線；每個物件都能在
PowerPoint 裡單獨修改。LR/RL/TB/BT 會自動分層排版，也可用 0–1 比例座標微調
單一節點。預設風格刻意克制：中性填色、細灰線、無陰影與漸層，只有目前論點的
節點使用 ACM 紅色，避免常見的「每個概念都是彩色圓角卡片」生成感。

**表格仍是原生 PowerPoint table。** `matrix` 會依內容長度配置欄寬，標籤欄較寬、
短數值欄較窄，只突出正在討論的 row。`qa_check.py` 會檢查 diagram 宣告的每個
node、edge、group 是否真的以命名原生物件存在於 `.pptx`。

## v2.3.0 有什麼

**影片可以放進去了。** 以前 `outline.json` 沒有 `video` 欄位，只能自己用
python-pptx 硬塞；而 python-pptx 的 `add_movie()` 預設把媒體標成
`video/unknown`，PowerPoint 會當成不認識的檔案，投影片在會議上就是一塊黑的。
現在：

```bash
python scripts/video.py probe raw.mov                        # PowerPoint 放得出來嗎
python scripts/video.py prep raw.mov -o figs/demo.mp4 --clip 0:03-0:18
```

然後在 outline 裡寫 `"video": {"src": "figs/demo.mp4", "autoplay": true, "loop": true}`。
`compose.py` 會用正確的 content type 嵌入、自動抽一張 poster、標上 `▶ 0:15`，
並且**拒絕**任何 PowerPoint 解不開的檔案，直接把該跑的 `prep` 指令印給你。

**省 token。** 以前很多錯誤只有把投影片渲染出來用眼睛看才抓得到，一張圖約 1600
tokens，25 頁看完一次就是 4 萬。現在 `qa_check.py` 直接用幾何和 XML 查出：殘留的
`XXX` / `20XX` / `Ur Name`、標題換行撞到副標、文字爆框、PowerPoint 會自動縮小的
文字。剩下真正要用眼睛看的，`render_qa.py --contact 6` 把六頁拼成一張圖，25 頁
從約 40k 降到約 13k。

**其他修掉的東西**：`compose.py` 跑第二次會把所有圖疊上去（現在會擋下來）；
結論頁加 `subtitle` 會把範本的 Todolist 標籤清空並拉滿整頁（現在不會）；標題框
的底邊本來就壓在副標上面 0.07 吋；範本自己會爆框的那個標籤讓 QA 永遠無法歸零
（現在以範本為基準扣掉）。

## 安裝為 Codex Skill

參考 repo 內的 `INSTALL.md`。

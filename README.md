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
| `v2.3.0` | 目前最新版。**支援嵌入影片**；QA 改成大部分不用渲染就查得出來；contact sheet 把看圖成本降到約三分之一 |

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

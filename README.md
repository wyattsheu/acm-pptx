# acm-pptx

ACM Lab (NYCU) PowerPoint 投影片生成工具 / Claude Skill。

`main` 分支永遠是最新版本；每個歷史版本都有對應的 git tag，方便在新版不好用時退回舊版。

## 版本列表

| 版本 | 說明 |
|---|---|
| `v1.0.0` | 初版 |
| `v1.1.0` | lab-rules / outline-schema 更新 |
| `v1.2.0` | SKILL.md 與規則微調 |
| `v2.0.0` | 新增 paper-study 支援、blueprint-paper 等 references |
| `v2.1.0` | 目前最新版 |

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

## 安裝為 Claude Skill

參考 repo 內的 `INSTALL.md`。

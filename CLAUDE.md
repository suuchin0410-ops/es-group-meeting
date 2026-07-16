# esグループ月次会議 — 3社財務ダッシュボード & 議事録管理システム

## 概要
esグループ3社（esエンターテイメント・エスクリエイト・esライフワーク）の月次財務データを分析・可視化し、
月次会議の議事録とAIフィードバックで改善サイクルを回すシステム。

## 目的
- 3社の財務データ（BS/PL/CF）をStreamlitダッシュボードで可視化
- 月次会議の議事録を構造化して蓄積
- 財務データ × 議事録の突き合わせによるAIフィードバック
- 前回フィードバック → 翌月の議論 → 改善の月次PDCAサイクル

## 対象法人

| 法人 | 事業 | 決算月 | ファイル |
|------|------|--------|----------|
| ㈱esエンターテイメント | 飲食業 | 6月（期首7月） | 月次報告-YYYYMM.xlsx |
| ㈱エスクリエイト | コンサルティング | 6月（期首7月） | SC-月次検証-YYYYMM.xlsx |
| ㈱esライフワーク | 就労継続支援A型/B型 | 3月（期首4月） | LW-月次検証-YYYYMM.xlsx |

各社の詳細は `app/config/companies.yaml` で管理。

## 月次サイクル（データフロー）

```
Phase 1: 月次Excelを data/monthly/ に配置
  ↓ ダッシュボード更新（BS/PL/CF/セグメント/予実）
Phase 2: 税理士・経営者で月次会議を実施（人間）
  ↓ 文字起こし → Notionにアップ
Phase 3: Claude Code が構造化議事録に変換 → Notion書き込み
  ↓
Phase 4: 財務データ × 議事録 → AIフィードバック生成 → Notion書き込み
  ↓
翌月Phase 1: 前回フィードバックを参照して改善を確認
```

## 技術基盤

### ダッシュボード
- **アプリ:** `app/dashboard.py`（Streamlit + Plotly）
- **公開URL（Streamlit Cloud）:** https://suuchin0410-ops-es-group-meeting-appdashboard-6ichlx.streamlit.app/
  - mainブランチへのpushで自動更新。スリープ時は「Yes, get this app back up!」で起動（1〜2分）
- **ローカル起動:** `cd app && streamlit run dashboard.py`
- 3社のBS/PL/CF推移、セグメント別売上、予実対比、BS詳細ドリルダウンを可視化

### Notion連携（2つのワークスペース）

| 用途 | ワークスペース | トークン | 親ページID |
|------|----------------|----------|------------|
| 議事録・AIフィードバック | esグループ会議 | `app/config/.notion_token_es_group` | `3808a77d-6ccc-808e-b17c-e0bf93ac67e6` |
| 過去の議事録参照（読み取りのみ） | bokashi | `app/config/.notion_token` | `38005528-aa05-8136-bddb-e11e4acdf0a3` |

- **書き込みモジュール:** `app/notion_writer.py`
- **読み取りモジュール:** `app/notion_reader.py`

## フォルダ構成
```
es-group-meeting/
├── app/
│   ├── dashboard.py            # Streamlit 3社財務ダッシュボード
│   ├── data_loader.py          # Excel読み込み・データ抽出
│   ├── notion_writer.py        # Notion書き込み
│   ├── notion_reader.py        # Notion読み取り
│   └── config/
│       ├── .notion_token           # bokashiワークスペース用トークン
│       ├── .notion_token_es_group  # esグループ会議ワークスペース用トークン
│       └── companies.yaml          # 各社情報・シート名・データソース定義
├── data/
│   └── monthly/                # 月次Excel格納先
├── docs/
│   ├── meeting-notes/          # 構造化議事録（YYYY-MM-DD.md）
│   └── reports/                # 生成レポート
├── skills/
│   └── es-group-meeting/
│       └── SKILL.md            # 月次会議フロースキル
└── CLAUDE.md
```

## 合言葉ワークフロー

### 「esグループ」「月次会議」「esMTG」→ 月次会議フロー

`skills/es-group-meeting/SKILL.md` を参照。
4フェーズの対話型ワークフローで月次会議を管理する:

1. **データ更新** — 月次Excelからダッシュボードを最新化、前回FBの追跡
2. **会議実施** — 人間作業。ダッシュボードを見ながら会議 → 文字起こしをNotionへ
3. **議事録作成** — 文字起こしを法人別・議題別に構造化 → Notion書き込み
4. **AIフィードバック** — 財務データ×議事録で分析・提案 → Notion書き込み

途中参加: 「議事録をアップした」→ Phase 3、「フィードバック」→ Phase 4

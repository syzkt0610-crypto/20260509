from pathlib import Path

from pptx import Presentation
from pptx.util import Pt

TITLE = "令和8年度（2026年）コンプライアンス＆情報セキュリティ教育（初回）"
SUBTITLE = "〜プロとして知っておくべき共通言語と、明日からの行動〜"

slides = [
    {
        "title": "研修の目的とスコープ",
        "bullets": [
            "目的：コンプライアンスとガバナンスを共通言語として理解する",
            "対象：全社員（特にITエンジニア）",
            "今回の重点：個人情報保護と生成AI利用ルール",
            "注意：法令の最終確認は必ず最新の公的情報・社内規程で行う",
        ],
    },
    {
        "title": "コンプライアンスの3層構造",
        "bullets": [
            "① 法令（法律・省令・条例等）：違反すると法的責任が生じ得る",
            "② 社内規程：就業規則・情報セキュリティ規程等の社内ルール",
            "③ 社会的規範：倫理・モラル・説明責任",
            "実務では『合法でも不適切』な行為も信用失墜につながるため注意",
        ],
    },
    {
        "title": "ガバナンスとコンプライアンスの関係",
        "bullets": [
            "ガバナンス：組織が健全に運営されるための仕組み（主語：組織）",
            "コンプライアンス：仕組みに沿って適正に行動すること（主語：個人）",
            "関係性：『仕組み』と『行動』の両輪でリスクを低減する",
        ],
    },
    {
        "title": "IT現場での主要テーマ",
        "bullets": [
            "個人情報保護・不正アクセス防止・著作権・労務・インサイダー対策",
            "クラウド設定ミス、権限管理不備、生成AIへの入力ミスは事故要因",
            "本研修では『個人情報』『生成AI』を優先テーマとして扱う",
        ],
    },
    {
        "title": "個人情報保護：基本用語の整理",
        "bullets": [
            "個人情報：特定の個人を識別できる情報",
            "個人データ：検索可能な形で体系化された個人情報",
            "保有個人データ：開示・訂正等の権限を事業者が持つ個人データ",
            "漏えい等が疑われる場合は、法令と社内手順に沿って速やかに報告",
        ],
    },
    {
        "title": "2026年時点での実務上の注意点",
        "bullets": [
            "AI活用・プロファイリングは利用目的と同意範囲を厳密に確認",
            "顔・指紋・音声等の生体情報は要配慮データとして厳格管理",
            "クラウド公開設定の誤りは『漏えいのおそれ』として初動報告対象",
            "※具体的要件は個人情報保護委員会等の最新公表資料を確認",
        ],
    },
    {
        "title": "明日からの3アクション（個人情報）",
        "bullets": [
            "1) 目的外利用をしない（勝手な二次利用を止める）",
            "2) テストに実データを使わない（マスキング・ダミー化）",
            "3) 設定ミスや誤送信の疑いは即時報告（目安：1分以内に一次報告）",
        ],
    },
    {
        "title": "生成AI利用のOK/NG境界",
        "bullets": [
            "NG：個人情報、未公開情報、機密ソースコード、秘密情報",
            "条件付きOK：匿名化・要約化済みデータ（社内ルールに従う）",
            "OK：一般公開情報、公開仕様、一般的な技術質問",
            "原則：入力前に『外部公開されても問題ないか』を確認する",
        ],
    },
    {
        "title": "ケーススタディからの学び",
        "bullets": [
            "生成AIへの機密入力は漏えい事故・信用毀損・懲戒の原因になり得る",
            "無断の個人情報入力は法令違反・契約違反・対外紛争のリスク",
            "教訓：技術利用前に『データ分類』『利用可否』『承認経路』を確認",
        ],
    },
    {
        "title": "会社の体制整備（ガバナンス）",
        "bullets": [
            "規程類（コンプライアンス・情報セキュリティ）を継続更新",
            "台帳管理（アカウント・資産）は事故時の追跡性と説明責任を担保",
            "ルールは抑止ではなく『社員と会社を守る防具』として運用する",
        ],
    },
    {
        "title": "36協定の正しい理解",
        "bullets": [
            "36協定は、時間外・休日労働に関する労使協定（許可制ではない）",
            "法定上限の遵守と健康確保のため、勤怠の正確な申告が必須",
            "社内公開場所：［社内フォルダの実パスを記入］",
        ],
    },
    {
        "title": "報告・連絡の原則",
        "bullets": [
            "重要事実の報告は雇用契約上の基本義務",
            "報告対象は『成果・変化』と『ミス・リスク』の両方",
            "早期報告で被害最小化・再発防止・会社による組織対応が可能になる",
        ],
    },
    {
        "title": "報告が自分を守る理由",
        "bullets": [
            "ミスの隠蔽は『過失』を『重大な規律違反』へ悪化させる",
            "適時報告は事実関係の保全と適切な法務・顧客対応につながる",
            "成果報告は評価・単価交渉・待遇改善の根拠データになる",
        ],
    },
    {
        "title": "未報告・隠蔽のリスク",
        "bullets": [
            "懲戒処分や信用失墜の可能性",
            "対応遅延により損害が拡大した場合は法的紛争リスクが増大",
            "結論：迷ったら即相談。自己判断で抱え込まない",
        ],
    },
    {
        "title": "まとめ：明日からの行動",
        "bullets": [
            "(1) データを扱う前に分類・目的・同意範囲を確認",
            "(2) 生成AI入力前に機密性と匿名化状態を確認",
            "(3) 事故の兆候は一次報告→エスカレーション→記録化",
            "(4) 成果もリスクも『見える化』して、強い現場を作る",
        ],
    },
]

html_parts = [
    "<!doctype html>",
    '<html lang="ja">',
    "<head>",
    '<meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    f"<title>{TITLE}</title>",
    "<style>",
    "body{font-family:'Yu Gothic','Hiragino Kaku Gothic ProN',sans-serif;line-height:1.7;margin:0;background:#f5f7fb;color:#1f2937}",
    ".wrap{max-width:1100px;margin:0 auto;padding:24px}",
    "h1{margin:0 0 8px;font-size:30px}",
    "h2{font-size:24px;border-left:6px solid #1d4ed8;padding-left:10px;margin-top:34px}",
    ".lead{color:#374151;margin-bottom:24px}",
    "section{background:#fff;border:1px solid #dbe3ef;border-radius:10px;padding:16px 20px;margin:14px 0}",
    "li{margin:6px 0}",
    ".note{background:#fffbe6;border:1px solid #f8dea0;padding:12px;border-radius:8px}",
    "</style>",
    "</head>",
    "<body>",
    '<div class="wrap">',
    f"<h1>{TITLE}</h1>",
    f'<p class="lead">{SUBTITLE}</p>',
    '<p class="note">本資料は社内教育向け整理版です。法令・ガイドラインの運用は、必ず最新の公的情報と社内規程を確認してください。</p>',
]

for slide in slides:
    html_parts.append(f"<section><h2>{slide['title']}</h2><ul>")
    for b in slide["bullets"]:
        html_parts.append(f"<li>{b}</li>")
    html_parts.append("</ul></section>")

html_parts.extend(["</div>", "</body>", "</html>"])

root = Path(__file__).resolve().parent
html_path = root / "compliance_security_training_2026.html"
html_path.write_text("\n".join(html_parts), encoding="utf-8")

prs = Presentation()
for idx, s in enumerate(slides):
    layout = prs.slide_layouts[0] if idx == 0 else prs.slide_layouts[1]
    slide = prs.slides.add_slide(layout)
    slide.shapes.title.text = TITLE if idx == 0 else s["title"]

    if idx == 0:
        subtitle = slide.placeholders[1]
        subtitle.text = SUBTITLE + "\n\n重点：個人情報保護 / 生成AI / 報告連絡"
        for paragraph in subtitle.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(20 if paragraph == subtitle.text_frame.paragraphs[0] else 16)
    else:
        body = slide.shapes.placeholders[1].text_frame
        body.clear()
        for i, b in enumerate(s["bullets"]):
            p = body.paragraphs[0] if i == 0 else body.add_paragraph()
            p.text = b
            p.level = 0
            p.font.size = Pt(20)

pptx_path = root / "compliance_security_training_2026.pptx"
prs.save(pptx_path)

print(f"Created: {html_path}")
print(f"Created: {pptx_path}")

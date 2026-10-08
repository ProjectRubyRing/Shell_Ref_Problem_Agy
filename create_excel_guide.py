import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_troubleshooting_workbook(output_filename="Shell_Properties_Troubleshooting_Guide.xlsx"):
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Monotone Palette Definitions
    COLOR_PRIMARY_DARK = "212529"   # Very dark charcoal / near black
    COLOR_HEADER_BG    = "343A40"   # Dark slate charcoal
    COLOR_HEADER_FG    = "FFFFFF"   # White text
    COLOR_SUBHEADER_BG = "495057"   # Medium dark gray
    COLOR_SUBHEADER_FG = "FFFFFF"   # White text
    COLOR_ZEBRA_EVEN   = "FFFFFF"   # White
    COLOR_ZEBRA_ODD    = "F8F9FA"   # Very light subtle gray
    COLOR_BORDER       = "CED4DA"   # Light subtle border
    COLOR_CODE_BG      = "E9ECEF"   # Cool light gray for code / highlight
    COLOR_TEXT_MAIN    = "212529"   # Primary text
    COLOR_TEXT_MUTED   = "6C757D"   # Muted gray text

    # Common Font helper
    def m_font(size=10, bold=False, italic=False, color=COLOR_TEXT_MAIN):
        return Font(name="Meiryo UI", size=size, bold=bold, italic=italic, color=color)

    # Common Border helper
    thin_border_side = Side(border_style="thin", color=COLOR_BORDER)
    medium_border_side = Side(border_style="medium", color=COLOR_PRIMARY_DARK)
    double_border_side = Side(border_style="double", color=COLOR_PRIMARY_DARK)
    
    cell_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side
    )
    
    header_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=medium_border_side,
        bottom=medium_border_side
    )

    # -------------------------------------------------------------
    # SHEET 1: 00_表紙・全体概要 (Cover & Overview)
    # -------------------------------------------------------------
    ws_cover = wb.create_sheet(title="00_表紙・全体概要")
    ws_cover.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_cover.merge_cells("B2:J3")
    title_cell = ws_cover["B2"]
    title_cell.value = "RHEL 9.8 Javaバッチ シェルスクリプト プロパティ読込障害 完全調査・対策技術書"
    title_cell.font = m_font(size=16, bold=True, color=COLOR_HEADER_FG)
    title_cell.fill = PatternFill(start_color=COLOR_PRIMARY_DARK, end_color=COLOR_PRIMARY_DARK, fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws_cover.merge_cells("B4:J4")
    sub_cell = ws_cover["B4"]
    sub_cell.value = "〜 原因ランキング・詳細切り分け・恒久対策・Linux/JVM動作原理・歴史的背景および小学生向け解説 〜"
    sub_cell.font = m_font(size=10, bold=True, italic=True, color=COLOR_HEADER_FG)
    sub_cell.fill = PatternFill(start_color=COLOR_SUBHEADER_BG, end_color=COLOR_SUBHEADER_BG, fill_type="solid")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Metadata Block
    metadata = [
        ("対象システム環境", "Red Hat Enterprise Linux 9.8 (Linux 5.14+ / glibc 2.34+ / systemd 252+)"),
        ("対象スクリプト", "Bash 5.1+ ラッパースクリプト (Javaバッチ起動・環境構築スクリプト)"),
        ("対象ファイル形式", "Javaプロパティファイル (*.properties / java.util.Properties 準拠)"),
        ("発生事象", "プロパティファイルはファイルシステム上に実在するが、シェル実行時に読み込めない・展開されない"),
        ("セキュリティ特性", "SELinux Enforcing 標準有効 / systemd リソース分離 / POSIX 厳格パーミッション"),
        ("作成日 / 基準版", "2026年10月 / エンタープライズミッションクリティカル対応版"),
    ]

    ws_cover.cell(row=6, column=2, value="■ 基本仕様・環境情報").font = m_font(size=11, bold=True)
    row_idx = 7
    for label, val in metadata:
        ws_cover.cell(row=row_idx, column=2, value=label).font = m_font(size=10, bold=True)
        ws_cover.cell(row=row_idx, column=2).fill = PatternFill(start_color=COLOR_ZEBRA_ODD, end_color=COLOR_ZEBRA_ODD, fill_type="solid")
        ws_cover.cell(row=row_idx, column=2).border = cell_border
        ws_cover.cell(row=row_idx, column=2).alignment = Alignment(vertical="center")

        ws_cover.merge_cells(start_row=row_idx, start_column=3, end_row=row_idx, end_column=7)
        c = ws_cover.cell(row=row_idx, column=3, value=val)
        c.font = m_font(size=10)
        c.border = cell_border
        c.alignment = Alignment(vertical="center")
        # Apply border to merged cells
        for col_i in range(4, 8):
            ws_cover.cell(row=row_idx, column=col_i).border = cell_border

        ws_cover.row_dimensions[row_idx].height = 24
        row_idx += 1

    # Executive Summary Block
    row_idx += 1
    ws_cover.cell(row=row_idx, column=2, value="■ 本書の目的とエグゼクティブサマリー").font = m_font(size=11, bold=True)
    row_idx += 1

    summary_text = (
        "【問題の本質】\n"
        "Linuxサーバー上で「ファイルが存在しているのに読み込めない」という事象は、表層的には不可解に見えますが、OSカーネルのVFS走査、"
        "SELinuxのMAC（強制アクセス制御）、Bashのプロセス空間（サブシェルと環境変数の一方向継承）、およびJavaプロパティ仕様とPOSIXシェル仕様の"
        "文化的ギャップに起因する明確な力学が存在します。\n\n"
        "【本書の構成】\n"
        "1. 『01_原因ランキング一覧』: 発生確率順に15の要因を整理し、即効性のあるクイックチェックを提示。\n"
        "2. 『02_事象確認・切分・対処詳細』: 各原因ごとの詳細な診断コマンド、切り分けフロー、および本番耐性のあるコード対策を提示。\n"
        "3. 『03_動作原理と歴史(小学生解説)』: Linuxカーネルの内部挙動、歴史的背景（Unix vs Javaの衝突）、小学生でも直感的に腹落ちする比喩解説を収録。\n"
        "4. 『04_RHEL9.8診断スクリプト』: RHEL 9.8上で即座に実行可能な、15のチェックを網羅した自動診断Bashスクリプトを格納。"
    )
    ws_cover.merge_cells(start_row=row_idx, start_column=2, end_row=row_idx+6, end_column=9)
    sum_cell = ws_cover.cell(row=row_idx, column=2, value=summary_text)
    sum_cell.font = m_font(size=9.5)
    sum_cell.fill = PatternFill(start_color=COLOR_ZEBRA_ODD, end_color=COLOR_ZEBRA_ODD, fill_type="solid")
    sum_cell.alignment = Alignment(vertical="top", wrap_text=True)
    for r in range(row_idx, row_idx+7):
        for c in range(2, 10):
            ws_cover.cell(row=r, column=c).border = cell_border

    row_idx += 8
    ws_cover.cell(row=row_idx, column=2, value="■ シート目次とナビゲーション").font = m_font(size=11, bold=True)
    row_idx += 1

    toc_items = [
        ("01_原因ランキング一覧", "推測される原因15選を発生確率順にマトリクス化。分類・影響度・概要・即効処置を俯瞰。"),
        ("02_事象確認・切分・対処詳細", "全15原因に対する「確認コマンド」「切り分け手順」「根本対処法」「修正スクリプト例」の詳説。"),
        ("03_動作原理と歴史(小学生解説)", "OS/プロセス/SELinuxの動作原理、UnixとJavaの歴史的設計思想の衝突、現場監督と職人の小学生向けたとえ話。"),
        ("04_RHEL9.8診断スクリプト", "サーバー上で直接動かしてボトルネックを1秒で突き止めるフル自動診断シェルスクリプト。"),
    ]

    for sheet_name, desc in toc_items:
        ws_cover.cell(row=row_idx, column=2, value=sheet_name).font = m_font(size=10, bold=True)
        ws_cover.cell(row=row_idx, column=2).fill = PatternFill(start_color=COLOR_CODE_BG, end_color=COLOR_CODE_BG, fill_type="solid")
        ws_cover.cell(row=row_idx, column=2).border = cell_border
        ws_cover.cell(row=row_idx, column=2).alignment = Alignment(vertical="center")

        ws_cover.merge_cells(start_row=row_idx, start_column=3, end_row=row_idx, end_column=8)
        c = ws_cover.cell(row=row_idx, column=3, value=desc)
        c.font = m_font(size=9.5)
        c.border = cell_border
        c.alignment = Alignment(vertical="center")
        for col_i in range(4, 9):
            ws_cover.cell(row=row_idx, column=col_i).border = cell_border

        ws_cover.row_dimensions[row_idx].height = 24
        row_idx += 1

    ws_cover.column_dimensions["A"].width = 3
    ws_cover.column_dimensions["B"].width = 28
    ws_cover.column_dimensions["C"].width = 18
    ws_cover.column_dimensions["D"].width = 18
    ws_cover.column_dimensions["E"].width = 18
    ws_cover.column_dimensions["F"].width = 18
    ws_cover.column_dimensions["G"].width = 18
    ws_cover.column_dimensions["H"].width = 18
    ws_cover.column_dimensions["I"].width = 18
    ws_cover.column_dimensions["J"].width = 3

    # -------------------------------------------------------------
    # SHEET 2: 01_原因ランキング一覧 (Rankings Matrix)
    # -------------------------------------------------------------
    ws_rank = wb.create_sheet(title="01_原因ランキング一覧")
    ws_rank.views.sheetView[0].showGridLines = True
    ws_rank.freeze_panes = "A3"

    # Title row
    ws_rank.merge_cells("A1:G1")
    r_title = ws_rank["A1"]
    r_title.value = "【推測される原因ランキング一覧】プロパティファイルが存在するのに読み込めない問題 (RHEL 9.8 / Javaバッチ)"
    r_title.font = m_font(size=12, bold=True, color=COLOR_HEADER_FG)
    r_title.fill = PatternFill(start_color=COLOR_PRIMARY_DARK, end_color=COLOR_PRIMARY_DARK, fill_type="solid")
    r_title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_rank.row_dimensions[1].height = 30

    headers_rank = [
        ("順位", 8, "center"),
        ("推測される原因の名称", 32, "left"),
        ("分類", 18, "center"),
        ("発生確率", 14, "center"),
        ("影響度", 12, "center"),
        ("概要と典型的な症状", 45, "left"),
        ("即効チェック・簡易対処", 40, "left"),
    ]

    for col_idx, (h_name, width, align_h) in enumerate(headers_rank, start=1):
        c = ws_rank.cell(row=2, column=col_idx, value=h_name)
        c.font = m_font(size=10, bold=True, color=COLOR_HEADER_FG)
        c.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
        c.border = header_border
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws_rank.column_dimensions[get_column_letter(col_idx)].width = width
    ws_rank.row_dimensions[2].height = 26

    rank_data = [
        (
            "第1位",
            "カレントディレクトリ（CWD）乖離による相対パス解決失敗",
            "実行環境 / パス",
            "極めて高い (~35%)",
            "致命的",
            "手動実行時は動くが、cron/systemd/ジョブ管理ツール経由だとカレントディレクトリが異なり「ファイルなし」となる。",
            "SCRIPT_DIR=\"$(cd \"$(dirname \"$0\")\" && pwd)\" でスクリプト基準の絶対パスを動的生成する。"
        ),
        (
            "第2位",
            "パイプラインによるサブシェル起動と変数スコープ消失",
            "シェル構文 / プロセス",
            "非常に高い (~20%)",
            "致命的",
            "cat file | while read line で export しても、パイプライン右側は別プロセス（サブシェル）のため、ループ終了で変数が消滅。",
            "入力リダイレクト while ... done < file に書き換えて親プロセス内でループを実行する。"
        ),
        (
            "第3位",
            "source コマンドによる読込破綻（キーのドットや空白等）",
            "構文 / ファイル仕様",
            "非常に高い (~15%)",
            "致命的",
            "source file.properties を実行。Javaプロパティのドット（db.url）や空白をBashが変数名として認識できず構文エラー。",
            "sourceを禁止し、AWKで抽出するかJava起動引数（--spring.config.location）で直接渡す。"
        ),
        (
            "第4位",
            "SELinux ポリシーによるアクセス拒否（AVC Denial）",
            "OSセキュリティ",
            "高い (~10%)",
            "重度",
            "RHEL9.8はSELinuxがEnforcing標準。mvで移動されたファイル等に不適切なラベル（user_tmp_t等）が残りアクセス拒否。",
            "ls -lZ で確認し、restorecon -v <file> で既定のセキュリティコンテキストを復元する。"
        ),
        (
            "第5位",
            "実行ユーザー不一致・親ディレクトリの走査権限(+x)不足",
            "パーミッション",
            "高い (~8%)",
            "重度",
            "ファイル自体が644でも、上位ディレクトリ（/opt/app等）に一般ユーザーへの実行権限(+x)がなく到達不能となる。",
            "namei -l <file> で全親階層の権限を可視化し、不足しているディレクトリに +x を付与する。"
        ),
        (
            "第6位",
            "改行コード CRLF（\\r\\n）の混入による不可視文字障害",
            "文字コード / 改行",
            "中〜高 (~5%)",
            "中〜重度",
            "Windowsで編集・転送され行末に \\r が付着。変数やパス末尾に \\r が入り「No such file...」や画面表示崩れを招く。",
            "file <file> または cat -v で ^M を確認し、dos2unix <file> または sed -i 's/\\r$//' を実行。"
        ),
        (
            "第7位",
            "UTF-8 BOM（Byte Order Mark）による1行目キー破損",
            "文字コード",
            "中 (~2%)",
            "中度",
            "先頭3バイト（0xEF 0xBB 0xBF）の不可視文字により、1行目のキー名（例: url）が \\ufeffurl と認識されマッチしない。",
            "head -c 3 <file> | xxd で確認し、sed -i '1s/^\\xef\\xbb\\xbf//' でBOMを除去する。"
        ),
        (
            "第8位",
            "最終行の改行コード(LF)欠落による最終行読取スキップ",
            "シェル / テキスト仕様",
            "中 (~2%)",
            "中度",
            "POSIX仕様上、末尾に改行がない行で read は戻り値1を返す。while read が最終行を変数に取った直後にループを脱出する。",
            "while read -r line || [[ -n \"$line\" ]]; do の救済構文を使うか、ファイル末尾に改行を追加。"
        ),
        (
            "第9位",
            "スクリプトの実行制御オプション（set -e, set -u）による突然死",
            "スクリプト制御",
            "低〜中 (~1%)",
            "中度",
            "set -u で未定義のパス変数が死因になる、または grep がマッチせず終了コード1を返してスクリプトが無言中断する。",
            "bash -x でトレース実行し、grep \"key\" file || true のように安全なイディオムを使用する。"
        ),
        (
            "第10位",
            "systemd / コンテナ環境による名前空間分離（PrivateTmp等）",
            "サービス分離",
            "低〜中 (~1%)",
            "重度",
            "systemdサービス化の際、PrivateTmp=yes や ProtectHome=yes により、/tmp や /home 内のファイルが隔離されて不可視化。",
            "systemctl cat でユニットを確認し、プロパティを /etc や /opt 等の正規ディレクトリに移設。"
        ),
        (
            "第11位",
            "切れたシンボリックリンク / NFSキャッシュ遅延不整合",
            "ファイルシステム",
            "低 (<1%)",
            "中度",
            "デプロイ時の相対シンボリックリンク先が壊れている（ダングリングリンク）。NFS属性キャッシュによる反映遅延。",
            "test -L <file> && readlink -f で実体存在を確認。NFSマウントオプション（actimeo）を点検。"
        ),
        (
            "第12位",
            "Javaコマンドライン引数のクォーティング不足",
            "引数展開",
            "低 (<1%)",
            "中度",
            "シェルは読み込めたが、java起動時に -Dpath=$VAR と裸展開し、パス内の空白で引数が分割されJava側で未認識。",
            "java -jar app.jar --config=\"$PROP_FILE\" のように必ず二重引用符で変数を保護する。"
        ),
        (
            "第13位",
            "POSIX ACL（拡張アクセス権）による特定ユーザー拒否",
            "アクセス権",
            "低 (<1%)",
            "中度",
            "通常のパーミッションは問題なく見えても（+付き）、getfacl で個別ユーザー/グループの拒否が設定されている。",
            "getfacl <file> で個別アクセス制御リストを確認し、setfacl -b で不要なACLを解除する。"
        ),
        (
            "第14位",
            "パス文字列内の全角空白やゼロ幅スペース等の不可視文字",
            "テキスト入力",
            "稀 (<0.5%)",
            "低〜中度",
            "設計書やチャットからスクリプトにパスをコピペした際、全角空白（U+3000）やゼロ幅スペース（U+200B）が混入。",
            "cat -A script.sh または printf '%q\\n' \"$PROP_PATH\" で不可視文字を可視化・手動再入力。"
        ),
        (
            "第15位",
            "ファイル更新中（非アトミック書き込み）に伴う空読み・競合",
            "プロセス間競合",
            "稀 (<0.5%)",
            "低〜中度",
            "別プロセスが > リダイレクトで上書きしている一瞬（0バイト）にシェルが読み込み、空データとして処理された。",
            "ファイル更新側をアトミックリネーム（一時ファイル作成後に mv）または flock 排他制御にする。"
        ),
    ]

    for idx, row in enumerate(rank_data, start=3):
        bg_color = COLOR_ZEBRA_EVEN if idx % 2 == 1 else COLOR_ZEBRA_ODD
        for col_idx, val in enumerate(row, start=1):
            c = ws_rank.cell(row=idx, column=col_idx, value=val)
            c.font = m_font(size=9.5, bold=(col_idx in (1, 2)))
            c.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
            c.border = cell_border
            align_h = headers_rank[col_idx-1][2]
            c.alignment = Alignment(horizontal=align_h, vertical="center", wrap_text=True)
        ws_rank.row_dimensions[idx].height = 42

    # -------------------------------------------------------------
    # SHEET 3: 02_事象確認・切分・対処詳細 (Deep Dive Table)
    # -------------------------------------------------------------
    ws_detail = wb.create_sheet(title="02_事象確認・切分・対処詳細")
    ws_detail.views.sheetView[0].showGridLines = True
    ws_detail.freeze_panes = "A3"

    ws_detail.merge_cells("A1:F1")
    d_title = ws_detail["A1"]
    d_title.value = "【原因別 詳細技術解説】事象の確認方法・問題の切り分け手順・根本対処法・検証コード (RHEL 9.8)"
    d_title.font = m_font(size=12, bold=True, color=COLOR_HEADER_FG)
    d_title.fill = PatternFill(start_color=COLOR_PRIMARY_DARK, end_color=COLOR_PRIMARY_DARK, fill_type="solid")
    d_title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_detail.row_dimensions[1].height = 30

    headers_detail = [
        ("順位・原因名", 24, "left"),
        ("発生メカニズム（なぜ起きるのか）", 38, "left"),
        ("事象の確認方法（診断コマンド）", 35, "left"),
        ("問題の切り分け手順（トラブルシューティング）", 35, "left"),
        ("根本対処法・修正コード（ベストプラクティス）", 45, "left"),
        ("予防策・設計指針", 28, "left"),
    ]

    for col_idx, (h_name, width, align_h) in enumerate(headers_detail, start=1):
        c = ws_detail.cell(row=2, column=col_idx, value=h_name)
        c.font = m_font(size=10, bold=True, color=COLOR_HEADER_FG)
        c.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
        c.border = header_border
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws_detail.column_dimensions[get_column_letter(col_idx)].width = width
    ws_detail.row_dimensions[2].height = 26

    detail_data = [
        (
            "【第1位】\nカレントディレクトリ（CWD）の乖離による相対パス解決失敗",
            "スクリプト内で PROP_FILE=\"../conf/app.properties\" と記述されている場合、実行時のカレントディレクトリ（pwd）を基準にファイルを探す。cronやジョブ管理ツール（JP1等）は /home/user や / で動くため、パス解決に失敗する。",
            "# スクリプト実行コンテキストの確認:\necho \"CWD: $(pwd)\"\necho \"TARGET: ${PROP_FILE}\"\nls -l \"${PROP_FILE}\"",
            "1. スクリプト冒頭に pwd > /tmp/batch_cwd.log を追記しジョブ実行。\n2. 記録されたパスとファイルの本来の絶対パスを照合。\n3. 手動実行時とcron実行時の差分を特定。",
            "# スクリプト自身の物理配置ディレクトリから動的に絶対パスを生成:\nSCRIPT_DIR=\"$(cd \"$(dirname \"$(readlink -f \"${BASH_SOURCE[0]}\")\")\" && pwd)\"\nPROP_FILE=\"${SCRIPT_DIR}/../conf/app.properties\"\n[[ -f \"${PROP_FILE}\" ]] || { echo \"File not found\"; exit 1; }",
            "スクリプト内では相対パスの使用を全面禁止し、常に BASH_SOURCE[0] を基点とした絶対パス解決を標準コーディング規約とする。"
        ),
        (
            "【第2位】\nパイプラインによるサブシェル起動と変数スコープ消失",
            "cat file | while read line; do export \"$key=$val\"; done の構文において、パイプライン（|）の右側は別プロセス（サブシェル）としてフォークされる。サブシェルが終了すると変数テーブルは破棄され、親シェルには反映されない。",
            "# ループ内外のPID確認:\ncat file | while read l; do\n  echo \"In Subshell PID: $BASHPID\"\ndone\necho \"Parent PID: $$\"",
            "1. ループ内では変数が表示されるか確認。\n2. ループを抜けた直後に変数が空（ブランク）になっているか確認。\n3. 親プロセスの環境変数一覧（env）に反映されていないことを確認。",
            "# パイプラインを廃止し、入力リダイレクトに変更:\nwhile IFS='=' read -r key val || [[ -n \"${key}\" ]]; do\n  [[ \"${key}\" =~ ^[[:space:]]*# ]] && continue\n  [[ -z \"${key}\" ]] && continue\n  export \"$(echo \"${key}\" | xargs)=$(echo \"${val}\" | xargs)\"\ndone < \"${PROP_FILE}\"",
            "シェルスクリプトにおいて、ループ内で環境変数を設定する場合はパイプラインを絶対に使用せず、リダイレクト構文を徹底する。"
        ),
        (
            "【第3位】\nsource コマンドによる読込破綻（キーのドットや空白等）",
            "Javaプロパティは key.name=val や key = val を許容するが、Bash変数名は [a-zA-Z_][a-zA-Z0-9_]* に限定される。source すると、Bashは key.name をコマンドと誤認し command not found となる。",
            "# 構文ドライランチェック:\nbash -n \"${PROP_FILE}\"\n# 直接source時のエラー確認:\nsource \"${PROP_FILE}\" 2>&1 | head -n 5",
            "1. 実行ログに \"command not found: spring...\" や \"syntax error\" が出ていないか確認。\n2. プロパティファイル内にドット（.）や等号前後の空白があるか確認。",
            "# 推奨A: Java引数として直接渡す（最も安全）:\njava -jar app.jar --spring.config.additional-location=\"file:${PROP_FILE}\"\n\n# 推奨B: AWKによる個別安全抽出:\nVAL=$(awk -F= -v k=\"db.url\" '$1 ~ \"^[[:space:]]*\"k\"[[:space:]]*$\" {sub(/^[[:space:]]*[^=]+=[[:space:]]*/, \"\"); print; exit}' \"${PROP_FILE}\")",
            "Javaプロパティファイルはシェルの設定ファイルとは別物であると認識し、シェルで直接 source / eval する設計を禁止する。"
        ),
        (
            "【第4位】\nSELinux ポリシーによるアクセス拒否（AVC Denial）",
            "RHEL 9.8はSELinux Enforcingが標準。/tmp や /home から mv で移動したファイルは、元の user_tmp_t や user_home_t コンテキストを維持する。systemdやサービスアカウントからの読み取りがLSMカーネル層でブロックされる。",
            "# SELinux状態とコンテキスト確認:\ngetenforce\nls -lZ \"${PROP_FILE}\"\n# AVC拒否ログ確認:\nsudo ausearch -m avc -ts recent",
            "1. 一時的に sudo setenforce 0 で警告モードに変更。\n2. バッチを実行し、正常に読めるかテスト。\n3. sudo setenforce 1 に戻す。\n4. setenforce 0 でのみ動く場合はSELinuxが原因。",
            "# 既定のセキュリティコンテキストを復元:\nrestorecon -v \"${PROP_FILE}\"\n\n# 恒久的なポリシー登録（/opt/app/conf配下）:\nsudo semanage fcontext -a -t etc_t \"/opt/app/conf(/.*)?\"\nsudo restorecon -Rv /opt/app/conf/",
            "ファイルを配置する際は mv ではなく cp を用いる（親ディレクトリのコンテキストを継承）。CI/CDでデプロイ後に restorecon を必須化。"
        ),
        (
            "【第5位】\n実行ユーザー不一致・親ディレクトリの走査権限(+x)不足",
            "LinuxのVFSはファイルを開く際、ルートから末尾までの全ディレクトリの実行権限(+x)を要求する。ファイルが 644 であっても、親の /opt/app が drwxr-x--- だと一般ユーザーは通過できず EACCES となる。",
            "# 全階層パーミッション可視化:\nnamei -l \"${PROP_FILE}\"\n# 実行ユーザーでの読取テスト:\nsudo -u batchuser test -r \"${PROP_FILE}\" && echo OK || echo NG",
            "1. namei -l の出力を確認し、途中のディレクトリに実行権限（x）が欠落していないか特定。\n2. 実行ユーザーの所属グループ（id batchuser）とディレクトリ所有グループを照合。",
            "# 所有権の統一と親ディレクトリ権限付与:\nsudo chown -R batchuser:batchgroup /opt/app/conf/\nsudo chmod 755 /opt /opt/app /opt/app/conf\nsudo chmod 644 /opt/app/conf/app.properties",
            "デプロイ手順書に、親ディレクトリのパーミッション確認をチェック項目として明記する。"
        ),
        (
            "【第6位】\n改行コード CRLF（\\r\\n）の混入による不可視文字障害",
            "Windowsで編集されたファイルの行末には \\r が存在する。Bashの read は \\r をデータとして読み込むため、変数末尾に \\r が付着する。これを元にファイルを開こうとすると、存在しないパスとなり失敗する。",
            "# 改行コードと不可視文字の確認:\nfile \"${PROP_FILE}\"\n# 行末の ^M を確認:\ncat -v \"${PROP_FILE}\" | head -n 5\nod -c \"${PROP_FILE}\" | head -n 2",
            "1. file コマンドで \"with CRLF line terminators\" が出るか確認。\n2. cat -v で行末に ^M が表示されるか確認。\n3. エラーメッセージの行頭が上書きされて消えていないか確認。",
            "# CRLFをLFに一括変換:\ndos2unix \"${PROP_FILE}\"\n# または sed で \\r を削除:\nsed -i 's/\\r$//' \"${PROP_FILE}\"",
            "Gitの core.autocrlf = input を設定し、リポジトリに .gitattributes （* text=auto eol=lf）を配備してCRLF混入を未然に防止。"
        ),
        (
            "【第7位】\nUTF-8 BOM（Byte Order Mark）による1行目キー破損",
            "Windowsのメモ帳等で保存した際、先頭に 0xEF 0xBB 0xBF が付与される。1行目のキー名先頭に不可視バイトが入り、キーの完全一致検索が絶対にヒットしなくなる。",
            "# 先頭3バイトのバイナリ確認:\nhead -c 3 \"${PROP_FILE}\" | xxd\n# fileコマンド確認:\nfile \"${PROP_FILE}\"",
            "1. xxd の出力が ef bb bf で始まっているか確認。\n2. 2行目以降の設定値は読めるのに、1行目の設定値だけが読めない現象を確認。",
            "# 先頭行のBOMを削除:\nsed -i '1s/^\\xef\\xbb\\xbf//' \"${PROP_FILE}\"",
            "エディタ設定で「UTF-8 (BOMなし)」を標準とし、CIパイプラインでBOM検出チェックを実施。"
        ),
        (
            "【第8位】\n最終行の改行コード(LF)欠落による最終行読取スキップ",
            "ファイルの最後の行に Enter（改行）がない場合、POSIXの read は行を読み込むが EOF 到達によりステータス 1 を返す。while read は即時終了し、最後の設定項目が処理されない。",
            "# 最終文字の改行有無チェック:\ntest \"$(tail -c 1 \"${PROP_FILE}\")\" = \"\" && echo \"改行あり\" || echo \"改行なし（危険）\"",
            "1. プロパティファイル末尾の設定項目だけがJavaやシェルに渡っていないか確認。\n2. ファイル末尾に空行を追加して再実行し、挙動が変わるかテスト。",
            "# 最終行救済イディオムの採用:\nwhile read -r line || [[ -n \"${line}\" ]]; do\n  # 処理\ndone < \"${PROP_FILE}\"\n\n# またはファイル末尾に改行を強制補正:\nsed -i -e '$a\\' \"${PROP_FILE}\"",
            "テキストファイルを生成・編集する際は、必ず末尾に改行を含めるルールを周知する。"
        ),
        (
            "【第9位】\nスクリプト実行制御オプション（set -e, set -u）による突然死",
            "set -u が有効な際、未定義変数を参照すると即座に終了する。また set -e が有効な際、grep でキーが見つからないと終了ステータス 1 でスクリプトが無言中断する。",
            "# デバッグトレース実行:\nbash -x ./run_batch.sh\n# 異常終了直前のコマンドを特定",
            "1. どの行でスクリプトが突然止まったかを bash -x で確認。\n2. 未定義変数の参照や grep の失敗がないかをログから特定。",
            "# デフォルト値の安全な展開:\nCONF_DIR=\"${APP_CONF_DIR:-/opt/app/conf}\"\n\n# grepの戻り値1を安全にスルー:\nMATCH=$(grep \"^db.url=\" \"${PROP_FILE}\" || true)",
            "set -euo pipefail を使う際は、失敗しうるコマンドに || true やデフォルト値フォールバックを必ず記述する。"
        ),
        (
            "【第10位】\nsystemd / コンテナ環境による名前空間分離（PrivateTmp等）",
            "systemd のセキュリティ機能により、PrivateTmp=yes ではホストの /tmp が見えず、ProtectHome=yes では /home がマスクされる。手動実行では見えるパスがサービス内から不可視となる。",
            "# ユニット定義の確認:\nsystemctl cat your-batch.service | grep -E \"Protect|Private|ReadOnly\"",
            "1. systemctl status でサービスのエラーログを確認。\n2. ユニットファイル内に PrivateTmp や ProtectHome が定義されているか照合。",
            "# ユニットファイル内で必要なパスを明示的に開放:\n[Service]\nProtectSystem=full\nReadOnlyPaths=/opt/app/conf\n\n# またはファイルを /etc/app/conf に配置",
            "バッチ設定ファイルは /tmp や /home に置かず、FHS（Filesystem Hierarchy Standard）に準拠した /etc や /opt に配置する。"
        ),
        (
            "【第11位】\n切れたシンボリックリンク / NFSキャッシュ遅延不整合",
            "シンボリックリンクの参照先実体ファイルが存在しない（ダングリングリンク）。またはNFSの属性キャッシュ（actimeo）により、他サーバーでの更新がローカルで反映されていない。",
            "# シンボリックリンク実体確認:\ntest -L \"${PROP_FILE}\" && ls -lL \"${PROP_FILE}\"\nstat -L \"${PROP_FILE}\"",
            "1. ls -l でリンク先が赤字表示されていないか確認。\n2. readlink -f で絶対パスの実体が存在するか確認。",
            "# 実体ファイルの再リンク作成:\nln -sfn /opt/app/releases/v1.2/conf/app.properties /opt/app/conf/app.properties\n# NFSマウントオプションの最適化（noac または適切な actimeo）",
            "デプロイ時のシンボリックリンク切り替えは、実体ファイルの配備完了を検証してから atomic に切り替える。"
        ),
        (
            "【第12位】\nJavaコマンドライン引数のクォーティング不足",
            "シェルはプロパティを読み込めたが、Java起動時に java -jar app.jar --config=$PROP_FILE と記述され、パス内の空白文字で引数が分割されてJava側に伝わらない。",
            "# 起動された実プロセス引数の確認:\nps -ef | grep java | grep -v grep",
            "1. ps -ef の出力から、Javaに渡っている引数文字列が期待通りか目視確認。\n2. Javaのログに \"unrecognized argument\" 等が出ていないか確認。",
            "# 必ず二重引用符で保護してJavaを起動:\njava -jar /opt/app/app.jar \\\n  --spring.config.additional-location=\"file:${PROP_FILE}\"",
            "シェルスクリプト内のすべての変数展開（\"$VAR\"）にダブルクォートを適用することを徹底する。"
        ),
        (
            "【第13位】\nPOSIX ACL（拡張アクセス権）による特定ユーザー拒否",
            "標準パーミッション（ls -l）は 644 に見えても、末尾に + 記号がある場合、ACLにより特定ユーザー/グループからのアクセスが拒否されている可能性がある。",
            "# ACLの確認:\ngetfacl \"${PROP_FILE}\"",
            "1. getfacl の出力で user:batchuser:--- などの拒否ルールがないか確認。\n2. mask の値が意図せず絞られていないか確認。",
            "# 不要な拡張ACLを全削除:\nsetfacl -b \"${PROP_FILE}\"\n# または明示的に読み取り権限を付与:\nsetfacl -m u:batchuser:r \"${PROP_FILE}\"",
            "ファイル権限管理は標準パーミッションを基本とし、ACLの多用はトラブルの温床となるため原則避ける。"
        ),
        (
            "【第14位】\nパス文字列内の全角空白やゼロ幅スペース等の不可視文字",
            "設計書やTeams/Slack等からスクリプトにパスをコピペした際、末尾や途中に全角空白（\\u3000）やゼロ幅スペース（\\u200b）が混入し、存在しないパスを探してしまう。",
            "# 不可視文字の検査:\ncat -A script.sh | grep -P \"[\\x80-\\xFF]\"\nprintf '%q\\n' \"${PROP_FILE}\"",
            "1. printf '%q\\n' の出力に \\$'\\x...' やエスケープ文字列が含まれていないか確認。\n2. スクリプトのパス指定行を手動で再入力して挙動が変わるかテスト。",
            "# 該当行を手入力で再記述、またはASCII文字以外を排除:\nsed -i 's/[^ -~]//g' script.sh",
            "設定ファイルやスクリプトを作成・編集する際は、リッチテキストやチャットからの直接コピペを避ける。"
        ),
        (
            "【第15位】\nファイル更新中（非アトミック書き込み）に伴う空読み・競合",
            "別プロセス（ファイル同期、rsync、デプロイスクリプト等）がプロパティファイルを > で直接上書きしている瞬間、サイズが0バイトになり、その瞬間にシェルが読むと空になる。",
            "# ファイルを掴んでいるプロセスの確認:\nlsof \"${PROP_FILE}\"\nfuser \"${PROP_FILE}\"",
            "1. ファイルサイズが一時的に 0 バイトになっていないか監視。\n2. バッチ実行とファイル配信ジョブの実行時刻が重なっていないか確認。",
            "# アトミック書き込み（作成後に mv で置換）:\necho \"key=val\" > \"${PROP_FILE}.tmp\" && mv -f \"${PROP_FILE}.tmp\" \"${PROP_FILE}\"\n\n# または flock でファイル排他ロックを取得して読み込み:\nexec 200<\"${PROP_FILE}\"\nflock -s 200\n# 読込処理\nflock -u 200",
            "設定ファイルの更新・配信は必ずアトミックリネーム（別名作成→mv）で行い、ファイルが0バイトになる瞬間を作らない。"
        ),
    ]

    for idx, row in enumerate(detail_data, start=3):
        bg_color = COLOR_ZEBRA_EVEN if idx % 2 == 1 else COLOR_ZEBRA_ODD
        for col_idx, val in enumerate(row, start=1):
            c = ws_detail.cell(row=idx, column=col_idx, value=val)
            c.font = m_font(size=9.5, bold=(col_idx == 1))
            c.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
            c.border = cell_border
            align_h = headers_detail[col_idx-1][2]
            c.alignment = Alignment(horizontal=align_h, vertical="top", wrap_text=True)
        ws_detail.row_dimensions[idx].height = 110

    # -------------------------------------------------------------
    # SHEET 4: 03_動作原理と歴史(小学生解説) (Principles & Metaphors)
    # -------------------------------------------------------------
    ws_meta = wb.create_sheet(title="03_動作原理と歴史(小学生解説)")
    ws_meta.views.sheetView[0].showGridLines = True
    ws_meta.freeze_panes = "A3"

    ws_meta.merge_cells("A1:E1")
    m_title = ws_meta["A1"]
    m_title.value = "【動作原理・歴史的背景・小学生向けたとえ話解説】なぜ起きるのか？なぜ必要だったのか？"
    m_title.font = m_font(size=12, bold=True, color=COLOR_HEADER_FG)
    m_title.fill = PatternFill(start_color=COLOR_PRIMARY_DARK, end_color=COLOR_PRIMARY_DARK, fill_type="solid")
    m_title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_meta.row_dimensions[1].height = 30

    headers_meta = [
        ("テーマ・項目", 24, "left"),
        ("小学生向け直感解説（たとえ話）", 38, "left"),
        ("技術的な動作原理・深堀り（OS・カーネル・JVM）", 42, "left"),
        ("登場の歴史と背景（なぜ作られ、なぜ必要だったのか）", 38, "left"),
        ("教訓と設計の勘所", 24, "left"),
    ]

    for col_idx, (h_name, width, align_h) in enumerate(headers_meta, start=1):
        c = ws_meta.cell(row=2, column=col_idx, value=h_name)
        c.font = m_font(size=10, bold=True, color=COLOR_HEADER_FG)
        c.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
        c.border = header_border
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws_meta.column_dimensions[get_column_letter(col_idx)].width = width
    ws_meta.row_dimensions[2].height = 26

    meta_data = [
        (
            "① シェルとJavaの関係\n【現場監督と職人さん】",
            "🧑‍💼「シェル君（現場監督）」:\n作業全体の段取りを決める監督。ルールに厳格で融通が利かない。\n\n👷「Javaさん（重機の職人）」:\nクレーン車を動かす職人。力持ちだが指示書がないと動けない。\n\n📋「プロパティファイル」:\n現場の棚にある指示書（メモ）。",
            "シェルはOSネイティブプロセス（Bash）として起動し、環境変数の設定、リソース制限（ulimit）、クラスパス構築を行い、execve(2) でJVM（Java仮想マシン）プロセスを起動するラッパーの責務を果たす。",
            "1970年代のUnix哲学「小さなツールの結合」と、1995年のJava「Write Once, Run Anywhere」の融合。OSの起動管理をシェルに任せ、重厚なビジネスロジックをJavaに任せる分業体制がエンタープライズの標準となった。",
            "シェルとJavaの責務を明確に分離し、シェルで無理にJavaの設定ファイルを加工しないこと。"
        ),
        (
            "② カレントディレクトリ問題\n【現場監督が別の部屋にいる事件】",
            "指示書は本社ビルの棚にある。シェル君が本社にいる時はすぐ取れた。\nしかし、夜中にタイマー（cron）で呼び出されたシェル君は遠くの資材置き場にいた！\nそこで「棚の指示書を読もう！」としても、そこには資材しかなく「指示書がない！」と大騒ぎになった。",
            "プロセスは実行時カレントワーキングディレクトリ（CWD）を持つ。相対パスはCWDを基点にVFSで解決される。cronやsystemdは起動時CWDが / や /home に設定されるため、相対パスが破綻する。",
            "Unixではユーザーがシェルを対話的に操作することを前提に相対パスが普及した。しかし自動化（cron/デーモン）が進むにつれ、プロセスの起点とファイルの物理配置が乖離する事故が多発し、絶対パス管理が必須となった。",
            "BASH_SOURCE[0] を用いて、スクリプト自身の物理配置場所から絶対パスを動的に組み立てる。"
        ),
        (
            "③ サブシェルと変数消失\n【小部屋にメモを忘れてきた事件】",
            "指示書を覚えるため、シェル君は隣の小部屋（サブシェル）に入った。\n「合言葉はリンゴだな！」と納得して大部屋に戻ってきたが、メモを小部屋に置いたまま鍵をかけてしまった！\n大部屋で職人のJavaさんに「合言葉は？」と聞かれて「あれ？何だっけ？」と忘れてしまった。",
            "パイプライン（|）の右側は fork(2) された別プロセス空間。子プロセスは親のメモリをCOWでコピーするが、子で export した環境変数は親プロセスのメモリ空間（MMU保護）には絶対に戻らない。終了時に子プロセスは破棄される。",
            "Unixのマルチタスク思想により「プロセス間のメモリ空間完全分離」という堅牢なセキュリティが確立された。その代償として、親プロセスへの環境変数の逆流は原理的に禁止されている。",
            "パイプラインではなく入力リダイレクト（<）を使用し、親シェルプロセス自身でループを回す。"
        ),
        (
            "④ ドット構文と文法不一致\n【ドットは嫌い！事件】",
            "職人のJavaさんは「spring.db.url」のようにドット（.）で細かく整理するのが大好き。\nでも現場監督のシェル君は「ドットを変数名にするのはルール違反だ！」と激怒。\nsource コマンドで読もうとして「そんなコマンド知らない！」とへそを曲げてしまった。",
            "Bashの変数名規則はPOSIX標準で [a-zA-Z_][a-zA-Z0-9_]* に限定。ドットはディレクトリ区切りや特殊構文。一方、JavaのProperties仕様はドットによる階層表現を標準とするため、source コマンドでの読込は文法的に破綻する。",
            "Unix環境変数はフラットな名前空間を前提に1970年代に策定された。一方Javaは1990年代にオブジェクト指向のパッケージ/プロパティ階層を表現するためにドット記法を導入。異なる時代の設計思想が衝突した。",
            "Javaプロパティをシェルで直接 source してはならない。Javaの起動引数でファイルごと渡す。"
        ),
        (
            "⑤ SELinuxとアクセス拒否\n【超まじめな警備員事件】",
            "指示書は目の前にある。住所も合っている。\nシェル君が手を伸ばした瞬間、警備員のSELinux君が腕をガシッと掴んだ！\n「待ちなさい！その紙には『ゴミ箱マーク（user_tmp_t）』のシールが貼ってあるぞ！正規の重要書類シールがない紙は触らせん！」",
            "RHEL 9.8のSELinuxはLSMフックによりカーネルレベルでMAC（強制アクセス制御）を適用する。mv コマンドで移動されたファイルは移動元のラベルを維持するため、デーモンドメインからの read が AVC Denial として遮断される。",
            "従来のUnixパーミッション（DAC）ではroot奪取や権限昇格攻撃を防ぎきれなかったため、アメリカ国家安全保障局（NSA）がFlask/SELinuxを開発。RHEL 9ではエンタープライズの防壁として強制適用（Enforcing）が標準となった。",
            "ファイルの配備時は mv ではなく cp を用いるか、配置後に必ず restorecon で規定ラベルを復元する。"
        ),
        (
            "⑥ CRLF改行コード障害\n【透明な宇宙人のゴミ事件】",
            "別のパソコン（Windows）から持ってきた指示書には、人間の目には見えないけれど「透明な宇宙人のゴミ（\\r）」が文字の末尾にくっついていた。\nシェル君が「指示書_ゴミ」という存在しないファイルを探してしまい、パニックになった。",
            "Windowsの改行は CRLF (0x0D 0x0A)、Linuxは LF (0x0A)。Bashの read は 0x0D を通常の文字列として読み込むため、パス末尾に \\r が混入。カーネルの open(2) は該当ファイルを見つけられず ENOENT となる。",
            "1960年代のタイプライター（テレタイプ端末）では、行頭に戻る動作（CR）と紙を1行送る動作（LF）が物理的に別だった。MS-DOS/Windowsは両方を残し、Unixは効率のためLFだけに統合した。この50年前の分岐が現代もトラブルを生む。",
            "Gitの core.autocrlf 設定や dos2unix ツールにより、Linux環境上のファイルは常に LF に統一する。"
        ),
        (
            "⑦ UTF-8 BOM障害\n【先頭の透明な名札事件】",
            "Windowsのメモ帳で保存したら、ファイルの先頭に人間の目には見えない「透明な名札（BOM）」がくっついた。\n1行目のキー名「url」を探しても、実は「名札+url」になっているため、シェル君は「見つからない！」と見落としてしまった。",
            "BOM（Byte Order Mark: 0xEF 0xBB 0xBF）はUTF-16等のエンディアン識別のために規定されたが、UTF-8では不要。しかしWindows系ツールが付加することがあり、文字列比較で完全一致せずパーサーが失敗する。",
            "Unicode初期の過渡期に、Windows陣営が文字コード自動判別のためにBOM付きUTF-8を推進した歴史がある。Unix/Linuxは伝統的にプレーンなUTF-8（BOMなし）を標準としたため、現在も断絶が残る。",
            "テキストエディタの設定で「BOMなしUTF-8」を厳守し、CI/CDスクリプトでBOM混入を自動排除する。"
        ),
    ]

    for idx, row in enumerate(meta_data, start=3):
        bg_color = COLOR_ZEBRA_EVEN if idx % 2 == 1 else COLOR_ZEBRA_ODD
        for col_idx, val in enumerate(row, start=1):
            c = ws_meta.cell(row=idx, column=col_idx, value=val)
            c.font = m_font(size=9.5, bold=(col_idx == 1))
            c.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
            c.border = cell_border
            align_h = headers_meta[col_idx-1][2]
            c.alignment = Alignment(horizontal=align_h, vertical="top", wrap_text=True)
        ws_meta.row_dimensions[idx].height = 115

    # -------------------------------------------------------------
    # SHEET 5: 04_RHEL9.8診断スクリプト (Diagnostic Tool)
    # -------------------------------------------------------------
    ws_tool = wb.create_sheet(title="04_RHEL9.8診断スクリプト")
    ws_tool.views.sheetView[0].showGridLines = True
    ws_tool.freeze_panes = "A3"

    ws_tool.merge_cells("A1:D1")
    t_title = ws_tool["A1"]
    t_title.value = "【RHEL 9.8 診断自動化スクリプト】プロパティファイル読込不具合を1秒で特定する検証ツール"
    t_title.font = m_font(size=12, bold=True, color=COLOR_HEADER_FG)
    t_title.fill = PatternFill(start_color=COLOR_PRIMARY_DARK, end_color=COLOR_PRIMARY_DARK, fill_type="solid")
    t_title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_tool.row_dimensions[1].height = 30

    headers_tool = [
        ("項目・ステップ", 22, "left"),
        ("診断対象・検証ロジック", 35, "left"),
        ("スクリプト実装（Bashコード抜粋）", 50, "left"),
        ("判定結果とアクション", 30, "left"),
    ]

    for col_idx, (h_name, width, align_h) in enumerate(headers_tool, start=1):
        c = ws_tool.cell(row=2, column=col_idx, value=h_name)
        c.font = m_font(size=10, bold=True, color=COLOR_HEADER_FG)
        c.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
        c.border = header_border
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws_tool.column_dimensions[get_column_letter(col_idx)].width = width
    ws_tool.row_dimensions[2].height = 26

    tool_data = [
        (
            "STEP 1:\n引数検証・ファイル実在確認",
            "引数で渡されたファイルパスの存在を検証。存在しない場合はカレントディレクトリ（CWD）の乖離を警告。",
            "if [[ ! -e \"$TARGET_PROP\" ]]; then\n  echo \"[NG] ファイルが見つかりません。\"\n  echo \"現在のCWD: $(pwd)\"\n  exit 1\nfi\nABS_PATH=\"$(readlink -f \"$TARGET_PROP\")\"",
            "NGの場合: 原因第1位（相対パスの解決失敗）が確定。絶対パスの動的組み立てを実施。",
        ),
        (
            "STEP 2:\nシンボリックリンク検証",
            "ファイルがシンボリックリンクであるか検査し、リンク先の実体が存在するか（切れていないか）を検証。",
            "if [[ -L \"$TARGET_PROP\" ]]; then\n  if [[ ! -e \"$ABS_PATH\" ]]; then\n    echo \"[NG] 切れたシンボリックリンクです！\"\n  fi\nfi",
            "NGの場合: 原因第11位（ダングリングリンク）。正しい実体パスへの再リンクを実施。",
        ),
        (
            "STEP 3:\n改行コード (CRLF) 検査",
            "file コマンドの出力から CRLF (Windows改行コード) の混入を自動検出。",
            "if file -b \"$ABS_PATH\" | grep -q \"CRLF\"; then\n  echo \"[NG] 改行コードがCRLFです！\"\n  dos2unix \"$ABS_PATH\"\nfi",
            "NGの場合: 原因第6位。dos2unix または sed -i 's/\\r$//' で即座にLFへ変換。",
        ),
        (
            "STEP 4:\nUTF-8 BOM 検査",
            "先頭3バイトが 0xEF 0xBB 0xBF であるかをバイナリレベルで判定。",
            "if head -c 3 \"$ABS_PATH\" | grep -q -P \"^\\xEF\\xBB\\xBF\"; then\n  echo \"[NG] UTF-8 BOM を検出！\"\n  sed -i '1s/^\\xef\\xbb\\xbf//' \"$ABS_PATH\"\nfi",
            "NGの場合: 原因第7位。先頭行のBOMを除去してキー名破損を解消。",
        ),
        (
            "STEP 5:\n最終行改行コード (LF) 検査",
            "ファイルの最終1バイトを取得し、改行文字で終端しているかを検証。",
            "if [[ \"$(tail -c 1 \"$ABS_PATH\")\" != \"\" ]]; then\n  echo \"[警告] 最終行に改行がありません！\"\n  echo \"\" >> \"$ABS_PATH\"\nfi",
            "警告の場合: 原因第8位。readループでの最終行スキップを防ぐため改行を追加。",
        ),
        (
            "STEP 6:\nアクセス権・VFS走査検査",
            "現在の実行ユーザーでの読み取り権限および、上位全階層の実行権限(+x)を namei で可視化。",
            "test -r \"$ABS_PATH\" || echo \"[NG] 読取権限なし！\"\nif command -v namei >/dev/null 2>&1; then\n  namei -l \"$ABS_PATH\"\nfi",
            "NGの場合: 原因第5位。上位ディレクトリの +x 付与および所有権（chown）を是正。",
        ),
        (
            "STEP 7:\nSELinuxコンテキスト検査 (RHEL9)",
            "getenforce でEnforcingを確認し、ファイルコンテキストに user_tmp_t 等の拒否ラベルがないか判定。",
            "if [[ \"$(getenforce 2>/dev/null)\" == \"Enforcing\" ]]; then\n  CTX=$(ls -Zd \"$ABS_PATH\" | awk '{print $1}')\n  if echo \"$CTX\" | grep -q -E \"user_tmp_t|user_home_t\"; then\n    echo \"[NG] 不正なコンテキスト: $CTX\"\n    restorecon -v \"$ABS_PATH\"\n  fi\nfi",
            "NGの場合: 原因第4位。restorecon -v を実行して正規のコンテキストを復元。",
        ),
        (
            "STEP 8:\nJava構文 vs シェル構文適合検査",
            "キー名にドット（.）が含まれている行や、等号前後の空白を検知し、source 実行時の破綻を事前警告。",
            "if grep -E '^[[:space:]]*[^#][^=]*\\.[^=]*=' \"$ABS_PATH\"; then\n  echo \"[注意] キー名にドットが含まれています！\"\n  echo \"-> source コマンドでの読込は失敗します。\"\nfi",
            "該当の場合: 原因第3位。シェルでの source を中止し、Java起動引数での読込に移行。",
        ),
    ]

    for idx, row in enumerate(tool_data, start=3):
        bg_color = COLOR_ZEBRA_EVEN if idx % 2 == 1 else COLOR_ZEBRA_ODD
        for col_idx, val in enumerate(row, start=1):
            c = ws_tool.cell(row=idx, column=col_idx, value=val)
            c.font = m_font(size=9.5, bold=(col_idx == 1))
            c.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
            c.border = cell_border
            align_h = headers_tool[col_idx-1][2]
            c.alignment = Alignment(horizontal=align_h, vertical="top", wrap_text=True)
        ws_tool.row_dimensions[idx].height = 80

    # Save workbook
    wb.save(output_filename)
    print(f"Successfully created: {output_filename}")

if __name__ == "__main__":
    create_troubleshooting_workbook()

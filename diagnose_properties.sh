#!/usr/bin/env bash
# ==============================================================================
# RHEL 9.8 プロパティファイル読込障害 完全自動診断スクリプト
# ==============================================================================
# 目的: RHEL 9.8環境において、「プロパティファイルが存在するのにシェルから
#       正常に読み込めない」事象の原因を全自動で網羅的にテストし特定します。
# 使用法: ./diagnose_properties.sh /path/to/your/app.properties
# ==============================================================================
set -u

TARGET_PROP="${1:-}"

echo "================================================================================"
echo " [診断開始] RHEL 9.8 プロパティファイル読込トラブルシューティング診断ツール"
echo "================================================================================"

if [[ -z "${TARGET_PROP}" ]]; then
    echo "[!] エラー: 診断対象のプロパティファイルパスが指定されていません。"
    echo "    使用法: $0 <プロパティファイルのパス>"
    echo "    例:     $0 /opt/app/conf/application.properties"
    exit 1
fi

echo "■ 診断対象入力パス: ${TARGET_PROP}"
echo "■ 実行ユーザー: $(id -un) (UID: $(id -u), GID: $(id -g))"
echo "■ 実行時カレントディレクトリ (PWD): $(pwd)"
echo "--------------------------------------------------------------------------------"

# ------------------------------------------------------------------------------
# STEP 1: ファイルの物理的存在および相対パス検証 (原因第1位)
# ------------------------------------------------------------------------------
echo "[STEP 1] ファイルの存在および相対パス解決チェック"
if [[ -e "${TARGET_PROP}" ]]; then
    echo "  [OK] 指定されたパスにファイルが存在します。"
    ABS_PATH="$(readlink -f "${TARGET_PROP}" 2>/dev/null || realpath "${TARGET_PROP}" 2>/dev/null || echo "")"
    echo "  [INFO] 解決された絶対パス: ${ABS_PATH}"
else
    echo "  [NG] 【原因第1位の可能性大】ファイルが存在しません！"
    echo "       指定されたパス: '${TARGET_PROP}'"
    echo "       現在の作業ディレクトリ: '$(pwd)'"
    echo "       -> cronやジョブ管理ツール経由の実行でCWDがずれている可能性があります。"
    echo "       -> 対処法: SCRIPT_DIR=\"\$(cd \"\$(dirname \"\$0\")\" && pwd)\" で絶対パスを構築してください。"
    exit 1
fi

# ------------------------------------------------------------------------------
# STEP 2: シンボリックリンクの整合性検証 (原因第11位)
# ------------------------------------------------------------------------------
echo ""
echo "[STEP 2] シンボリックリンク整合性チェック"
if [[ -L "${TARGET_PROP}" ]]; then
    LINK_TARGET="$(readlink "${TARGET_PROP}")"
    echo "  [WARN] 対象はシンボリックリンクです -> 参照先: ${LINK_TARGET}"
    if [[ ! -e "${ABS_PATH}" ]]; then
        echo "  [NG] 【原因第11位】切れたシンボリックリンク（ダングリングリンク）です！"
        echo "       参照先の実体ファイルが存在しません。"
        exit 1
    else
        echo "  [OK] シンボリックリンクの実体が存在します。"
    fi
else
    echo "  [OK] 通常のファイル（シンボリックリンクではありません）。"
fi

FILE_TYPE="$(file -b "${ABS_PATH}")"
echo "  [INFO] ファイル種別判定: ${FILE_TYPE}"

# ------------------------------------------------------------------------------
# STEP 3: 改行コード CRLF (\r\n) 検査 (原因第6位)
# ------------------------------------------------------------------------------
echo ""
echo "[STEP 3] 改行コード検査 (CRLF / LF)"
if echo "${FILE_TYPE}" | grep -q "CRLF"; then
    echo "  [NG] 【原因第6位】改行コードが Windows形式 (CRLF: \\r\\n) です！"
    echo "       -> シェルが読み込んだ変数末尾に \\r が付着し、パス未検出や構文エラーになります。"
    echo "       -> 対処法: dos2unix \"${ABS_PATH}\" または sed -i 's/\\r$//' \"${ABS_PATH}\""
else
    echo "  [OK] 改行コードは正常な Linux形式 (LF) です。"
fi

# ------------------------------------------------------------------------------
# STEP 4: UTF-8 BOM 検査 (原因第7位)
# ------------------------------------------------------------------------------
echo ""
echo "[STEP 4] 文字コード・BOM (Byte Order Mark) 検査"
if head -c 3 "${ABS_PATH}" 2>/dev/null | grep -q -P "^\xEF\xBB\xBF" 2>/dev/null; then
    echo "  [NG] 【原因第7位】UTF-8 BOM (0xEF 0xBB 0xBF) が検出されました！"
    echo "       -> 1行目のキー名の先頭に不可視バイトが入り、キー検索に失敗します。"
    echo "       -> 対処法: sed -i '1s/^\\xef\\xbb\\xbf//' \"${ABS_PATH}\""
else
    echo "  [OK] UTF-8 BOM は検出されませんでした。"
fi

# ------------------------------------------------------------------------------
# STEP 5: ファイル末尾の改行コード有無検査 (原因第8位)
# ------------------------------------------------------------------------------
echo ""
echo "[STEP 5] 最終行の終端改行コード検査"
LAST_BYTE="$(tail -c 1 "${ABS_PATH}" 2>/dev/null)"
if [[ "${LAST_BYTE}" != "" ]]; then
    echo "  [WARN] 【原因第8位】ファイルの最終行に改行コードがありません！"
    echo "         -> 'while read line' ループで最終行が処理されずにスキップされます。"
    echo "         -> 対処法: sed -i -e '\$a\\' \"${ABS_PATH}\""
else
    echo "  [OK] 最終行は正常に改行で終端されています。"
fi

# ------------------------------------------------------------------------------
# STEP 6: パーミッションおよび上位ディレクトリ走査権限検査 (原因第5位)
# ------------------------------------------------------------------------------
echo ""
echo "[STEP 6] パーミッションおよびVFSディレクトリ走査権限チェック"
if [[ -r "${ABS_PATH}" ]]; then
    echo "  [OK] 現在の実行ユーザーでファイルへの読み取り権限（Read）があります。"
else
    echo "  [NG] 【原因第5位】ファイル自体の読み取り権限がありません！"
    echo "       詳細: $(ls -l "${ABS_PATH}")"
    echo "       -> 対処法: chmod 644 \"${ABS_PATH}\" または所有者変更 (chown)"
fi

echo "  [INFO] 上位全ディレクトリ階層のトラバース権限 (+x) 解析:"
if command -v namei >/dev/null 2>&1; then
    namei -l "${ABS_PATH}"
else
    # namei がない場合のフォールバック
    CHECK_DIR="$(dirname "${ABS_PATH}")"
    while [[ "${CHECK_DIR}" != "/" && "${CHECK_DIR}" != "." ]]; do
        if [[ ! -x "${CHECK_DIR}" ]]; then
            echo "  [NG] 【原因第5位】上位ディレクトリ '${CHECK_DIR}' に実行権限(+x)がありません！"
        fi
        CHECK_DIR="$(dirname "${CHECK_DIR}")"
    done
fi

# ------------------------------------------------------------------------------
# STEP 7: SELinux セキュリティコンテキスト検査 (RHEL 9.8特有: 原因第4位)
# ------------------------------------------------------------------------------
echo ""
echo "[STEP 7] RHEL 9.8 SELinux セキュリティコンテキスト検査"
if command -v getenforce >/dev/null 2>&1; then
    SE_MODE="$(getenforce)"
    echo "  [INFO] SELinux 動作モード: ${SE_MODE}"
    CTX="$(ls -Zd "${ABS_PATH}" 2>/dev/null | awk '{print $1}')"
    echo "  [INFO] ファイルのセキュリティコンテキスト: ${CTX}"
    
    if [[ "${SE_MODE}" == "Enforcing" ]]; then
        if echo "${CTX}" | grep -q -E "user_tmp_t|user_home_t|admin_home_t"; then
            echo "  [NG] 【原因第4位】不適切なSELinuxコンテキスト (${CTX}) です！"
            echo "       -> mv コマンド等で移動されたため、一時ファイル用ラベルが付着しています。"
            echo "       -> systemdやサービスアカウントからの読み取りがカーネルで遮断されます。"
            echo "       -> 対処法: restorecon -v \"${ABS_PATH}\""
        else
            echo "  [OK] 不正な一時ファイルラベルではありません。"
        fi
        
        # 監査ログの確認 (root権限がある場合)
        if [[ $(id -u) -eq 0 ]] && command -v ausearch >/dev/null 2>&1; then
            DENIALS="$(ausearch -m avc -ts recent 2>/dev/null | grep -c "${ABS_PATH}" || true)"
            if [[ ${DENIALS} -gt 0 ]]; then
                echo "  [CRITICAL] 直近でこのファイルに対する AVC Denial (アクセス拒否) が ${DENIALS} 件発生しています！"
            fi
        fi
    fi
else
    echo "  [INFO] SELinux コマンドが未検出または無効です。"
fi

# ------------------------------------------------------------------------------
# STEP 8: Javaプロパティ構文 vs シェル構文適合検査 (原因第3位)
# ------------------------------------------------------------------------------
echo ""
echo "[STEP 8] Javaプロパティ構文適合性検査 (source 読込破綻リスク)"
DOT_LINES="$(grep -E '^[[:space:]]*[^#][^=]*\.[^=]*=' "${ABS_PATH}" 2>/dev/null | head -n 3 || true)"
if [[ -n "${DOT_LINES}" ]]; then
    echo "  [WARN] 【原因第3位】キー名にドット（.）を含むプロパティが存在します！"
    echo "         該当例:"
    echo "${DOT_LINES}" | sed 's/^/           /'
    echo "         -> Bash変数名にドットは使用できないため、'source \"${ABS_PATH}\"' すると"
    echo "            'command not found' となり読み込みに失敗します。"
    echo "         -> 対処法: source を廃止し、AWKで抽出するか、Java起動引数（--spring.config.location）で渡してください。"
fi

SPACE_LINES="$(grep -E '^[[:space:]]*[^#]+[[:space:]]+=' "${ABS_PATH}" 2>/dev/null | head -n 3 || true)"
if [[ -n "${SPACE_LINES}" ]]; then
    echo "  [WARN] 【原因第3位】等号（=）の前に空白が存在します！"
    echo "         該当例:"
    echo "${SPACE_LINES}" | sed 's/^/           /'
    echo "         -> シェルで source すると構文エラーになります。"
fi

echo ""
echo "================================================================================"
echo " [診断終了] 上記の [NG] または [WARN] 項目に対処を実施してください。"
echo "================================================================================"

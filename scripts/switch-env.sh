#!/bin/bash

# 環境切り替えスクリプト
# 使用方法: ./scripts/switch-env.sh [dev|prod]

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
GAS_DIR="$SCRIPT_DIR/../Google Script"

if [ -z "$1" ]; then
    echo "使用方法: $0 [dev|prod]"
    echo "  dev  - 開発環境に切り替え"
    echo "  prod - 本番環境に切り替え"
    exit 1
fi

case "$1" in
    dev)
        if [ -f "$GAS_DIR/.clasp.json.dev" ]; then
            cp "$GAS_DIR/.clasp.json.dev" "$GAS_DIR/.clasp.json"
            echo "✓ 開発環境に切り替えました"
        else
            echo "エラー: .clasp.json.dev が見つかりません"
            exit 1
        fi
        ;;
    prod)
        if [ -f "$GAS_DIR/.clasp.json.prod" ]; then
            cp "$GAS_DIR/.clasp.json.prod" "$GAS_DIR/.clasp.json"
            echo "✓ 本番環境に切り替えました"
        else
            echo "エラー: .clasp.json.prod が見つかりません"
            exit 1
        fi
        ;;
    *)
        echo "エラー: 無効な環境です。'dev' または 'prod' を指定してください"
        exit 1
        ;;
esac

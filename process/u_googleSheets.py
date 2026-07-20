"""
Googleスプレッドシート読み取り用の共通ヘルパー。

サービスアカウント認証(st.connection + GSheetsConnection)経由で、
非公開スプレッドシート(閲覧者を特定アカウントのみに制限したもの)を読み取る。

これまで各所で個別に組み立てていた `gviz` の公開CSVエクスポートURL
(pd.read_csv(csv_url)) は、匿名アクセス前提のため「特定アカウントのみ閲覧可」
という制限と両立しない。本モジュールは secrets.toml の [connections.gsheets]
に登録したサービスアカウントを介して同じ処理を行う。

接続先スプレッドシートは secrets.toml 側 ([connections.gsheets].spreadsheet)
で固定しているため、呼び出し側はワークシート名だけを指定すればよい。
"""
import streamlit as st
from streamlit_gsheets import GSheetsConnection


def read_sheet(worksheet_name: str, header="infer", ttl: int = 0):
    """
    secrets.toml の [connections.gsheets] で指定したスプレッドシート内の
    任意のワークシートを DataFrame として取得する。

    Parameters:
        worksheet_name: シート名 (例: "アクセス管理")
        header: pandas.read_csv 互換。ヘッダー行なしで全行取得したい場合は None を指定
        ttl: キャッシュ保持秒数。0 (デフォルト) はキャッシュ無効＝毎回最新を取得
    """
    conn = st.connection("gsheets", type=GSheetsConnection)
    return conn.read(worksheet=worksheet_name, header=header, ttl=ttl, use_spinner=False)

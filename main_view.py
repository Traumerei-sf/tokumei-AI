import streamlit as st
import pandas as pd
import os
import importlib
from datetime import datetime
from process.p2_2_Template_DiagnosticPDF import ESSENCE_MAP

def show_main():
    # PC前提のワイドレイアウト設定
    st.set_page_config(page_title="特命AI", layout="wide", initial_sidebar_state="collapsed")
    
    # CSS: サイドバー削除およびUI調整
    st.markdown(
        """
        <style>
            /* サイドバーとトグルの完全非表示 */
            [data-testid="stSidebar"] { display: none; }
            [data-testid="collapsedControl"] { display: none; }
            
            /* 特命AI タイトルスタイル（左上） */
            .app-title {
                font-size: 3.5rem;
                font-weight: 900;
                color: #1E1E1E;
                margin-bottom: 0.5rem;
                text-align: left;
            }
            
            /* 説明文スタイル（中央揃え） */
            .description-container {
                text-align: center;
                margin-top: 2rem;
                margin-bottom: 3rem;
            }
            .description-text {
                font-size: 1.4rem;
                color: #444444;
                font-weight: 500;
            }

            /* ボタンのスタイリング */
            .stButton > button {
                height: 3.5rem;
                font-size: 1.2rem !important;
                background-color: #000000;
                color: #FFFFFF;
                border: none;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # 1. アプリ名とアイコン（構成: アイコン[1], タイトル[10]）
    col_icon, col_title = st.columns([1, 12], vertical_alignment="center")

    with col_icon:
        st.image("assets/特命AI_ロゴ_osakacapital_900×600.png", width=80)

    with col_title:
        st.markdown("<h1 style='margin:0;'>特命AI</h1>", unsafe_allow_html=True)

    # 2. 説明文（中央揃え）
    st.markdown("あなたの12～24か月の会計データから、診断レポートと営業先リストを作成いたします")

    # 3. メインアクションエリア (PC向けに中央に寄せる)
    _, col_center, _ = st.columns([0.5, 4, 0.5])

    with col_center:
        st.markdown("### 【任意】会社情報の入力")
        st.markdown("営業先候補リスト・仕入先候補リストの精度を高めたい場合は、以下会社情報を入力してください")
        st.text_input("会社名", key="company_name")
        st.text_input("会社の業種", key="company_industry")
        st.markdown("<br>", unsafe_allow_html=True)
        
        st.markdown("### 会計データのアップロード")
        st.markdown("※csvまたはxlsxファイルのみ受け付けます。xlsxファイルは1つ目のシートのみ読み込みます")
        st.markdown("※貸借対照表をアップロードする場合は、「仕訳帳の最終月」と「貸借対照表の期末月」を合わせてください")
        st.markdown("※アップロードされた会計データは、診断後すぐに破棄されますのでご安心ください")
        
        # --- 非同期処理用のヘルパー関数 ---
        import threading
        import importlib
        from streamlit.runtime.scriptrunner import add_script_run_ctx
        
        # 数字から始まるモジュールは直接importできないためimportlibを使用
        SAD_module = importlib.import_module("process.1_standardizeAccountingData")
        process_journal_single = SAD_module.process_journal_single
        process_bs_single = SAD_module.process_bs_single

        def start_async_process(file, key_prefix):
            def task():
                try:
                    file_num = 1 if key_prefix == "j1" else 2
                    df, error = process_journal_single(file, file_num=file_num)
                    if error:
                        st.session_state[f"{key_prefix}_error"] = error
                        st.session_state[f"{key_prefix}_status"] = "error"
                        return

                    # --- 期間バリデーション ---
                    min_d = df["date"].min()
                    max_d = df["date"].max()
                    months = (max_d.year - min_d.year) * 12 + (max_d.month - min_d.month) + 1

                    if key_prefix == "j2":
                        # 2枚目の場合：1枚目(J1)との統合期間をチェック
                        if st.session_state.get("j1_status") == "success":
                            df1 = st.session_state["j1_data"]
                            j1_start = df1["date"].min()
                            j2_end = max_d
                            total_months = (j2_end.year - j1_start.year) * 12 + (j2_end.month - j1_start.month) + 1
                            
                            if not (12 <= total_months <= 24):
                                st.session_state[f"{key_prefix}_error"] = f"2ファイルの合計期間が不適切です（{total_months}ヶ月分。12〜24ヶ月にする必要があります）"
                                st.session_state[f"{key_prefix}_status"] = "error"
                                return
                            st.session_state["total_months_msg"] = f"✅ 2ファイル合計で {total_months}ヶ月分を確認しました"
                        else:
                            # J1がない状態でJ2が上がった場合
                            st.session_state[f"{key_prefix}_error"] = "先に1枚目の仕訳帳をアップロードしてください。"
                            st.session_state[f"{key_prefix}_status"] = "error"
                            return
                    else:
                        # 1枚目(J1)単体の場合
                        if months > 24:
                                st.session_state[f"{key_prefix}_error"] = f"仕訳帳1枚の期間が長すぎます（{months}ヶ月分）。"
                                st.session_state[f"{key_prefix}_status"] = "error"
                                return

                    # 全てOKなら保存
                    st.session_state[f"{key_prefix}_data"] = df
                    st.session_state[f"{key_prefix}_status"] = "success"
                    st.session_state[f"{key_prefix}_error"] = None

                except Exception as e:
                    st.session_state[f"{key_prefix}_error"] = f"システムエラー: {str(e)}"
                    st.session_state[f"{key_prefix}_status"] = "error"

            st.session_state[f"{key_prefix}_status"] = "processing"
            st.session_state[f"{key_prefix}_error"] = None
            st.session_state[f"{key_prefix}_file_id"] = f"{file.name}_{file.size}"
            thread = threading.Thread(target=task)
            add_script_run_ctx(thread)
            thread.start()

        def start_async_process_bs(file, key_prefix):
            def task():
                try:
                    data, error = process_bs_single(file)
                    if error:
                        st.session_state[f"{key_prefix}_error"] = error
                        st.session_state[f"{key_prefix}_status"] = "error"
                        return
                    
                    st.session_state[f"{key_prefix}_data"] = data
                    st.session_state[f"{key_prefix}_status"] = "success"
                    st.session_state[f"{key_prefix}_error"] = None

                except Exception as e:
                    st.session_state[f"{key_prefix}_error"] = f"システムエラー: {str(e)}"
                    st.session_state[f"{key_prefix}_status"] = "error"

            st.session_state[f"{key_prefix}_status"] = "processing"
            st.session_state[f"{key_prefix}_error"] = None
            st.session_state[f"{key_prefix}_file_id"] = f"{file.name}_{file.size}"
            thread = threading.Thread(target=task)
            add_script_run_ctx(thread)
            thread.start()

        @st.fragment(run_every="5s")
        def show_file_status(key_prefix, name):
            status = st.session_state.get(f"{key_prefix}_status", "idle")
            if status == "processing":
                st.info(f"⏳ {name} を解析中...（最大30秒程度かかります）")
            elif status == "error":
                st.error(f"❌ {st.session_state.get(f'{key_prefix}_error')}")
            elif status == "success":
                # J2の場合は統合期間メッセージを表示
                msg = "解析が完了しました"
                if key_prefix == "j2" and st.session_state.get("total_months_msg"):
                    st.success(st.session_state["total_months_msg"])
                else:
                    st.success(f"✅ {name} の{msg}")

        # 3つのアップロード枠
        # --- 仕訳帳 1 ---
        file_journal1 = st.file_uploader("① 仕訳帳（1期目または2期分） 【必須】", type=["csv", "xlsx"], help="必須項目です。")
        if file_journal1:
            fid = f"{file_journal1.name}_{file_journal1.size}"
            if st.session_state.get("j1_file_id") != fid:
                start_async_process(file_journal1, "j1")
            show_file_status("j1", "仕訳帳（1つ目）")
        else:
            # ファイルが削除された場合は状態をリセット
            st.session_state["j1_file_id"] = None
            st.session_state["j1_status"] = "idle"
            st.session_state["j1_data"] = None
            st.session_state["j1_error"] = None

        # --- 仕訳帳 2 ---
        file_journal2 = st.file_uploader("② 仕訳帳（2期目） 【任意】", type=["csv", "xlsx"], help="ファイルが分かれている場合はこちらもアップロードしてください。")
        if file_journal2:
            fid2 = f"{file_journal2.name}_{file_journal2.size}"
            if st.session_state.get("j2_file_id") != fid2:
                start_async_process(file_journal2, "j2")
            show_file_status("j2", "仕訳帳（2つ目）")
        else:
            st.session_state["j2_file_id"] = None
            st.session_state["j2_status"] = "idle"
            st.session_state["j2_data"] = None
            st.session_state["j2_error"] = None
            st.session_state["total_months_msg"] = None

        # --- 貸借対照表 ---
        file_bs = st.file_uploader("③ 貸借対照表 【任意】", type=["csv", "xlsx"])
        if file_bs:
            fid_bs = f"{file_bs.name}_{file_bs.size}"
            if st.session_state.get("bs_file_id") != fid_bs:
                start_async_process_bs(file_bs, "bs")
            show_file_status("bs", "貸借対照表")
        else:
            st.session_state["bs_file_id"] = None
            st.session_state["bs_status"] = "idle"
            st.session_state["bs_data"] = None
            st.session_state["bs_error"] = None
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 4. 作成ボタン
        if st.button("診断レポートと営業先リストを作成", use_container_width=True):
            if not file_journal1:
                st.error("❌ 最低でも「仕訳帳（1つ目）」のアップロードが必要です。")
            elif st.session_state.get("j1_status") == "processing":
                st.warning("⚠️ 仕訳帳（1つ目）の解析がまだ完了していません。あと数秒お待ちください。")
            elif st.session_state.get("j1_status") == "error":
                st.error("❌ 仕訳帳（1つ目）にエラーがあるため、進めません。")
            elif st.session_state.get("bs_status") == "processing":
                st.warning("⚠️ 貸借対照表の解析がまだ完了していません。あと数秒お待ちください。")
            elif st.session_state.get("bs_status") == "error":
                st.error(f"❌ 貸借対照表の解析エラー: {st.session_state.get('bs_error')}")
            else:
                # BSのバリデーション
                bs_invalid = False
                if file_bs and st.session_state.get("bs_status") == "success":
                    bs_data = st.session_state.get("bs_data")
                    journal_max_date = None
                    if st.session_state.get("j2_status") == "success":
                        journal_max_date = st.session_state["j2_data"]["date"].max()
                    elif st.session_state.get("j1_status") == "success":
                        journal_max_date = st.session_state["j1_data"]["date"].max()
                    
                    if journal_max_date is not None and bs_data and bs_data.get("year_month"):
                        try:
                            bs_ym = datetime.strptime(bs_data["year_month"], "%Y/%m")
                            j_year = journal_max_date.year
                            j_month = journal_max_date.month
                            b_year = bs_ym.year
                            b_month = bs_ym.month
                            
                            month_diff = abs((j_year - b_year) * 12 + (j_month - b_month))
                            print(f"--- DEBUG: Month Diff: J: {j_year}/{j_month}, B: {b_year}/{b_month}, Diff: {month_diff} ---")
                            
                            if month_diff > 3:
                                st.error(f"❌ 「仕訳帳の最後の取引年月({j_year}/{j_month})」と「貸借対照表の期末年月({b_year}/{b_month})」が3ヶ月を超えて離れています。（差: {month_diff}ヶ月）")
                                bs_invalid = True
                        except Exception as e:
                            print(f"Date Parse debug Error: {e}")
                            st.error("❌ 貸借対照表の期末年月フォーマットが不正です。")
                            bs_invalid = True

                if not bs_invalid:
                    st.success("解析を開始します。しばらくお待ちください...")
                    
                    try:
                        # 1. データの準備
                        std_data = {
                            "journal": st.session_state.get("j1_data"),
                            "bs": pd.DataFrame() # BSは一旦空フレーム
                        }
                        # もしBSデータがあれば組み込む(後続処理用)
                        if file_bs and st.session_state.get("bs_status") == "success":
                            bs_d = st.session_state.get("bs_data")
                            if bs_d and bs_d.get("cash_amount") is not None:
                                std_data["bs"] = pd.DataFrame({"期末現預金合計": [bs_d.get("cash_amount")]})

                        if st.session_state.get("j2_status") == "success":
                            st.info("2つの仕訳帳を結合しています...")
                            std_data["journal"] = pd.concat([std_data["journal"], st.session_state.get("j2_data")]).sort_values("date")
                        
                        # 2. 診断レポートの作成
                        report_mod = importlib.import_module("process.2_createDiagnosticReportPdf")
                        report_data = report_mod.create_diagnostic_report(
                            df_journal=std_data["journal"],
                            df_bs=std_data["bs"]
                        )
                        
                        # セッション状態に保存
                        st.session_state["standardized_journal"] = std_data["journal"]
                        st.session_state["standardized_bs"] = std_data["bs"]
                        st.session_state["report_pdf_bytes"] = report_data["pdf_bytes"]
                        st.session_state["report_preview_md"] = report_data["preview_md"]
                        st.session_state["report_analysis_df"] = report_data["analysis_df"]
                        
                        # フラグ管理
                        st.session_state["report_ready"] = True
                        st.session_state["biz_list_ready"] = False
                        st.session_state["supplier_list_ready"] = False
                        st.session_state["is_processed"] = True
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"❌ 処理中にエラーが発生しました: {e}")

    # 5. 処理後のUI（ボタン押下後のみ表示）
    if st.session_state.get("is_processed", False):
        st.divider()
        st.subheader("診断結果プレビュー")
        st.info("解析が完了しました。以下に各プレビューが表示されます。")
        
        # --- レイアウト変更: 1列構成 ---
        
        # 5.1 診断レポート プレビュー
        st.markdown("### 診断レポート (プレビュー)")
        st.markdown("---")
        
        analysis_df = st.session_state.get("report_analysis_df")
        
        if analysis_df is not None:
            # ② 分析結果の概要
            red_count = len(analysis_df[analysis_df["color"] == "red"])
            if red_count == 0:
                summary_msg = "経営の存続に関わる【赤信号】はありませんでした。"
            elif red_count == 1:
                summary_msg = "経営の存続に関わる【赤信号】が一つありました。"
            else:
                summary_msg = "経営の存続に関わる複数の【赤信号】がありました。"
            
            st.markdown(f"<div style='text-align: center;'><h4>貴社の会計データを「特命AI財務分析ロジック」で分析した結果</h4>", unsafe_allow_html=True)
            color = "red" if red_count > 0 else "black"
            st.markdown(f"<div style='text-align: center; color: {color};'><h3>{summary_msg}</h3></div>", unsafe_allow_html=True)
            
            # ③ 画像の挿入 (中央揃え)
            img_path = os.path.join("assets", "特命AI_レポート画像_1.jpeg")
            if os.path.exists(img_path):
                _, img_col, _ = st.columns([1, 3, 1])
                with img_col:
                    st.image(img_path, width=500)
            
            st.markdown("<br>", unsafe_allow_html=True)

            # ④ 指定された3項目のみを抽出
            # 赤項目を優先し、合計3項目を抽出
            red_rows = analysis_df[analysis_df["color"] == "red"]
            other_rows = analysis_df[analysis_df["color"] != "red"]
            preview_df = pd.concat([red_rows, other_rows]).head(3).copy()
            
            # --- 「問題の本質」の追加 ---
            preview_df["essence"] = preview_df["item"].map(ESSENCE_MAP)
            
            # 行全体の背景色を設定するスタイル関数
            def style_rows(row):
                color_key = row["color"]
                if color_key == "red":
                    bg_color = "background-color: #ffeef0; color: #000000"
                elif color_key == "blue":
                    bg_color = "background-color: #eef6ff; color: #000000"
                elif color_key == "grey":
                    bg_color = "background-color: #f8f9fa; color: #000000"
                else:
                    bg_color = "color: #000000"
                return [bg_color] * len(row)

            st.dataframe(
                preview_df.style.apply(style_rows, axis=1),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "category": "カテゴリ",
                    "item": "評価項目",
                    "result": "結果",
                    "comment": "コメント",
                    "essence": "問題の本質",
                    "color": None 
                }
            )
        else:
            st.markdown(st.session_state.get("report_preview_md", "レポートを作成できませんでした。"))
        
        # 5.2 診断レポート ダウンロードボタン
        now_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        pdf_filename = f"特命AI_診断レポート_{now_str}.pdf"

        st.download_button(
            label="診断レポート(PDF)をダウンロード",
            data=st.session_state.get("report_pdf_bytes", b""),
            file_name=pdf_filename,
            mime="application/pdf",
            use_container_width=True
        )
        
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.divider()

        # --- 営業先リスト作成のバックグラウンド実行（レポート表示後） ---
        if not st.session_state.get("biz_list_ready", False):
            with st.status("AIがおすすめ営業先候補を探しています...", expanded=True) as status:
                biz_mod = importlib.import_module("process.3_createBusinessList")
                df_full, df_preview = biz_mod.create_business_list(st.session_state["standardized_journal"])
                st.session_state["business_list_full"] = df_full
                st.session_state["business_list_preview"] = df_preview
                st.session_state["biz_list_ready"] = True
                status.update(label="営業先候補リストの作成が完了しました！", state="complete", expanded=False)
            st.rerun()

        # 5.3 営業先リスト プレビュー
        st.markdown("### 営業先候補リスト (プレビュー)")
        st.markdown("---")
        
        df_biz_preview = st.session_state.get("business_list_preview")
        df_biz_full = st.session_state.get("business_list_full")

        if df_biz_preview is not None:
            st.dataframe(
                df_biz_preview,
                use_container_width=True,
                hide_index=True
            )
            
            # 5.4 営業先リスト ダウンロードボタン
            csv_now_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            csv_filename = f"特命AI_営業先候補リスト_{csv_now_str}.csv"
            
            # CSV変換 (Shift-JIS for Excel compatibility, fallback to UTF-8-sig)
            try:
                csv_data = df_biz_full.to_csv(index=False).encode('utf-8-sig')
            except:
                csv_data = df_biz_full.to_csv(index=False).encode('utf-8')

            st.download_button(
                label="全ての新規営業先候補リストをダウンロード",
                data=csv_data,
                file_name=csv_filename,
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.info("営業先リストは生成されませんでした。")

        st.markdown("<br><br>", unsafe_allow_html=True)
        st.divider()

        # --- 仕入先リスト作成のバックグラウンド実行 ---
        if st.session_state.get("biz_list_ready", False) and not st.session_state.get("supplier_list_ready", False):
            with st.status("AIがおすすめ仕入先候補を探しています...", expanded=True) as status:
                biz_mod = importlib.import_module("process.3_createBusinessList")
                df_full_s, df_prev_s = biz_mod.create_supplier_list(st.session_state["standardized_journal"])
                st.session_state["supplier_list_full"] = df_full_s
                st.session_state["supplier_list_preview"] = df_prev_s
                st.session_state["supplier_list_ready"] = True
                status.update(label="仕入先候補リストの作成が完了しました！", state="complete", expanded=False)
            st.rerun()

        # 5.5 仕入先リスト プレビュー
        st.markdown("### 仕入先候補リスト (プレビュー)")
        st.markdown("---")
        
        df_sup_preview = st.session_state.get("supplier_list_preview")
        df_sup_full = st.session_state.get("supplier_list_full")

        if df_sup_preview is not None:
            st.dataframe(
                df_sup_preview,
                use_container_width=True,
                hide_index=True
            )
            
            # 5.6 仕入先リスト ダウンロードボタン
            csv_now_str_s = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            csv_filename_s = f"特命AI_仕入先候補リスト_{csv_now_str_s}.csv"
            
            try:
                csv_data_s = df_sup_full.to_csv(index=False).encode('utf-8-sig')
            except:
                csv_data_s = df_sup_full.to_csv(index=False).encode('utf-8')

            st.download_button(
                label="全ての新規仕入先候補リストをダウンロード",
                data=csv_data_s,
                file_name=csv_filename_s,
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.info("仕入先リストは生成されませんでした。")

    # 6. フッター
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    st.divider()
    st.markdown(
        """
        <div style="text-align: center; color: #888888; font-size: 0.9rem; padding-bottom: 2rem;">
            大阪キャピタル株式会社<br>
            〒103-0026 東京都中央区日本橋兜町５－１兜町第一平和ビル３階
        </div>
        """,
        unsafe_allow_html=True
    )

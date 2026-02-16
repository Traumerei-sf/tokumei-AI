import streamlit as st
import pandas as pd

def show_main():
    import os
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

    # 1. アプリ名（左上）
    st.markdown('<h1 class="app-title">特命AI</h1>', unsafe_allow_html=True)

    # 2. 説明文（中央揃え）
    st.markdown(
        '<div class="description-container">'
        '<p class="description-text">あなたの12～24か月の会計データから、診断レポートと営業先リストを作成いたします</p>'
        '</div>', 
        unsafe_allow_html=True
    )

    # 3. メインアクションエリア (PC向けに中央に寄せる)
    _, col_center, _ = st.columns([1, 2, 1])

    with col_center:
        st.markdown("### 会計データのアップロード")
        
        # 2つのアップロード枠
        file_journal = st.file_uploader("① 仕訳帳（CSV） 【必須】", type=["csv"], help="必須項目です。")
        if file_journal:
            st.success("✅ 仕訳帳が読み込まれました。")
            
        file_bs = st.file_uploader("② 貸借対照表（CSV） 【任意】", type=["csv"])
        if file_bs:
            st.success("✅ 貸借対照表が読み込まれました。")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 4. 作成ボタン
        if st.button("診断レポートと営業先リストを作成", use_container_width=True):
            if not file_journal:
                st.error("❌ 最低でも「仕訳帳」のアップロードが必要です。")
            else:
                st.session_state["is_processed"] = True
                st.success("解析を開始します。しばらくお待ちください...")
                
                # 処理エンジンの呼び出し
                import importlib
                process_logic = importlib.import_module("process.1_standardizeAccountingData")
                standardize_logic = process_logic.standardize_logic
                check_accounting_files = process_logic.check_accounting_files
                
                try:
                    # 1. バリデーションチェック
                    validation_results = check_accounting_files(file_journal, file_bs)
                    
                    is_all_ok = True
                    for res in validation_results:
                        if res["color"] == "green":
                            st.success(f"✅ {res['message']}")
                        elif res["color"] == "red":
                            st.error(f"❌ {res['message']}")
                            is_all_ok = False
                        else:
                            st.info(res['message'])
                    
                    if not is_all_ok:
                        st.stop() # 処理をストップ

                    # 2. 標準化処理（仕訳帳は必須）
                    std_data = standardize_logic(
                        file_journal=file_journal,
                        file_bs=file_bs
                    )
                    
                    # 3. 診断レポートの作成
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
                    st.session_state["biz_list_ready"] = False # まだ
                    st.session_state["is_processed"] = True
                    st.rerun() # 画面を更新してプレビューを表示
                    
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
            
            st.markdown(f"<div style='text-align: center;'><h4>貴社の会計データを「宮田ロジック」で分析した結果</h4>", unsafe_allow_html=True)
            color = "red" if red_count > 0 else "black"
            st.markdown(f"<div style='text-align: center; color: {color};'><h3>{summary_msg}</h3></div>", unsafe_allow_html=True)
            
            # ③ 画像の挿入 (中央揃え)
            img_path = r"C:\Users\User\.antigravity\特命AI\assets\特命AI_レポート画像_1.jpeg"
            if os.path.exists(img_path):
                # st.image で中央寄せにするには、カラムを使うかCSS
                _, img_col, _ = st.columns([1, 2, 1])
                with img_col:
                    st.image(img_path, use_container_width=True) # 这里还是用这个，或者宽度
            
            st.markdown("<br>", unsafe_allow_html=True)

            # ④ 指定された3項目のみを抽出
            target_items = ["現金薄さ", "買掛・未払残高", "仕訳入力遅延"]
            preview_df = analysis_df[analysis_df["item"].isin(target_items)].copy()
            
            # --- 「問題の本質」の追加 ---
            from process.p2_2_Template_DiagnosticPDF import ESSENCE_MAP
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
        from datetime import datetime
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
            import importlib
            with st.status("AIがおすすめ営業先を探しています...", expanded=True) as status:
                biz_mod = importlib.import_module("process.3_createBusinessList")
                df_full, df_preview = biz_mod.create_business_list(st.session_state["standardized_journal"])
                st.session_state["business_list_full"] = df_full
                st.session_state["business_list_preview"] = df_preview
                st.session_state["biz_list_ready"] = True
                status.update(label="営業先リストの作成が完了しました！", state="complete", expanded=False)
            st.rerun()

        # 5.3 営業先リスト プレビュー
        st.markdown("### 営業先リスト (プレビュー)")
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
                label="10件全ての新規営業先候補リストをダウンロード",
                data=csv_data,
                file_name=csv_filename,
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.info("営業先リストは生成されませんでした。")

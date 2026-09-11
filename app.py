import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="카페24 재고-주문 연동 체크기",
    page_icon="📦",
    layout="wide"
)

st.title("📦 카페24 주문-재고 연동 및 출고 가능 여부 체크기")
st.caption("주문 엑셀과 재고 엑셀을 업로드하면, 총 주문 수량을 합산하여 재고 부족 및 발주 필요 수량을 정밀하게 분석합니다.")

with st.expander("ℹ️ 사용 가이드 및 기준 엑셀 컬럼 안내"):
    st.markdown("""
    * **카페24 주문 엑셀 기본 열 이름**: `상품명`, `옵션정보`, `수량` (또는 `주문수량`)
    * **재고 엑셀 기본 열 이름**: `상품명`, `옵션정보`, `재고수량` (또는 `현재재고`)
    * **자동 처리 특징**:
        1. 동일한 상품/옵션으로 들어온 여러 주문 건의 수량을 **하나로 자동 합산**합니다.
        2. [현재 재고 - 총 주문 수량]을 계산하여 **출고 가능(Pass)** 및 **재고 부족(Fail)**을 판별합니다.
        3. 부족한 품목은 **필요 발주 수량**을 자동으로 계산해 줍니다.
    """)

col1, col2 = st.columns(2)

with col1:
    st.subheader("1) 카페24 주문 엑셀 파일")
    orders_file = st.file_uploader("카페24 주문 목록 엑셀", type=["xlsx", "xls"], key="orders")

with col2:
    st.subheader("2) 창고/ERP 재고 엑셀 파일")
    inventory_file = st.file_uploader("현재 재고 목록 엑셀", type=["xlsx", "xls"], key="inventory")

if orders_file and inventory_file:
    try:
        df_orders = pd.read_excel(orders_file)
        df_inventory = pd.read_excel(inventory_file)

        order_qty_col = next((c for c in ['수량', '주문수량', '구매수량'] if c in df_orders.columns), None)
        inv_qty_col = next((c for c in ['재고수량', '현재재고', '재고'] if c in df_inventory.columns), None)
        
        if '상품명' not in df_orders.columns or order_qty_col is None:
            st.error("❌ 주문 엑셀에서 '상품명' 또는 '수량' 열을 찾을 수 없습니다.")
            st.stop()
            
        if '상품명' not in df_inventory.columns or inv_qty_col is None:
            st.error("❌ 재고 엑셀에서 '상품명' 또는 '재고수량' 열을 찾을 수 없습니다.")
            st.stop()

        if '옵션정보' not in df_orders.columns:
            df_orders['옵션정보'] = '-'
        if '옵션정보' not in df_inventory.columns:
            df_inventory['옵션정보'] = '-'

        df_orders['옵션정보'] = df_orders['옵션정보'].fillna('-')
        df_inventory['옵션정보'] = df_inventory['옵션정보'].fillna('-')

        df_orders['매칭키'] = df_orders['상품명'].astype(str).str.strip() + " || " + df_orders['옵션정보'].astype(str).str.strip()
        df_inventory['매칭키'] = df_inventory['상품명'].astype(str).str.strip() + " || " + df_inventory['옵션정보'].astype(str).str.strip()

        orders_summary = df_orders.groupby(['매칭키', '상품명', '옵션정보'])[order_qty_col].sum().reset_index()
        orders_summary.rename(columns={order_qty_col: '총주문수량'}, inplace=True)

        inventory_summary = df_inventory.groupby(['매칭키'])[inv_qty_col].sum().reset_index()
        inventory_summary.rename(columns={inv_qty_col: '현재재고수량'}, inplace=True)

        merged = pd.merge(orders_summary, inventory_summary, on='매칭키', how='left')
        merged['현재재고수량'] = merged['현재재고수량'].fillna(0).astype(int)
        
        merged['차감후재고'] = merged['현재재고수량'] - merged['총주문수량']
        merged['출고가능여부'] = merged['차감후재고'].apply(lambda x: '✅ 출고가능' if x >= 0 else '🚨 재고부족')
        merged['필요발주수량'] = merged['차감후재고'].apply(lambda x: abs(x) if x < 0 else 0)

        result_df = merged[['상품명', '옵션정보', '총주문수량', '현재재고수량', '차감후재고', '출고가능여부', '필요발주수량']]

        total_items = len(result_df)
        shortage_items = len(result_df[result_df['차감후재고'] < 0])
        total_shortage_qty = result_df['필요발주수량'].sum()

        st.divider()
        st.markdown("### 📊 정밀 분석 결과 요약")
        
        m1, m2, m3 = st.columns(3)
        m1.metric("총 주문 상품 종류", f"{total_items} 개")
        m2.metric("재고 부족(품절 위험) 품목", f"{shortage_items} 개", delta_color="inverse")
        m3.metric("총 부족 수량 (발주 필요)", f"{total_shortage_qty} 개")

        tab1, tab2, tab3 = st.tabs(["🔴 재고 부족 품목 (발주 필요)", "🟢 출고 가능 품목", "📋 전체 분석 목록"])

        with tab1:
            shortage_df = result_df[result_df['차감후재고'] < 0]
            if not shortage_df.empty:
                st.warning(f"총 {len(shortage_df)}개 품목의 재고가 부족합니다.")
                st.dataframe(shortage_df, use_container_width=True)
            else:
                st.success("🎉 모든 주문 건에 대해 재고가 충분합니다!")

        with tab2:
            pass_df = result_df[result_df['차감후재고'] >= 0]
            st.dataframe(pass_df, use_container_width=True)

        with tab3:
            st.dataframe(result_df, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, sheet_name='전체분석결과', index=False)
            if not shortage_df.empty:
                shortage_df.to_excel(writer, sheet_name='발주필요목록', index=False)
        output.seek(0)

        st.download_button(
            label="📥 분석 결과 엑셀 다운로드",
            data=output,
            file_name="재고_주문_연동_분석결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"파일을 처리하는 중 오류가 발생했습니다: {e}")
else:
    st.info("👆 상단의 두 입력 칸에 카페24 주문 엑셀과 재고 엑셀 파일을 모두 업로드해주세요.")
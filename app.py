# app.py
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="오답률 통계 생성기", layout="centered")
st.title("📊 오답률 통계 생성기")

# -----------------------
# 예시 엑셀/CSV 제공
# -----------------------
def example_df():
    # 예시: 이름, Module1, Module2 (X=응시/오답0, 빈칸=미응시)
    return pd.DataFrame({
        "이름": ["홍길동", "김철수", "이영희", "박민수"],
        "Module1": ["1,3,5", "X", "2,4,7", ""],   # "" 또는 NaN = 미응시
        "Module2": ["2,6", "1,3", "X", "5"]
    })

with st.expander("🧾 예시 입력 파일 보기 / 복사 / 다운로드"):
    ex = example_df()
    st.caption("열 이름은 반드시 **이름, Module1, Module2** 입니다. 값은 `1,3,5` 처럼 콤마로 구분하고, 오답이 없으면 `X`, 미응시는 빈칸으로 두세요.")
    st.dataframe(ex, use_container_width=True)
    # 복사용 CSV
    csv_text = ex.to_csv(index=False)
    st.text_area("복사용 CSV", csv_text, height=180)
    # 엑셀 다운로드
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        ex.to_excel(w, index=False, sheet_name="예시")
    buf.seek(0)
    st.download_button("📥 예시 엑셀 다운로드", buf, file_name="예시_오답현황_양식.xlsx")

# -----------------------
# 통계 함수
# -----------------------
def robust_parse_wrong_list(cell):
    """엑셀 셀 내용 → 틀린 문제 번호 리스트 변환
    None/빈칸 -> None(미응시), 'X' -> [] (응시/오답 0), '1,2,5' -> [1,2,5]
    전각 콤마/세미콜론도 허용
    """
    if pd.isna(cell) or str(cell).strip() == "":
        return None
    s = str(cell).strip()
    if s.lower() == "x":
        return []
    s = s.replace("，", ",").replace(";", ",")
    return [int(x.strip()) for x in s.split(",") if x.strip().isdigit()]

def compute_module_rates(series, total_questions):
    """모듈별 오답률 계산: 오답률(%) = (틀린 학생 수 / 응시자 수)*100"""
    attempted = sum(v is not None for v in series)  # 분모: 응시자 수
    rows = []
    for q in range(1, total_questions + 1):
        wrong = sum((v is not None and q in v) for v in series)
        rate = round((wrong / attempted) * 100, 1) if attempted > 0 else 0.0
        rows.append({"문제 번호": q, "오답률(%)": rate, "틀린 학생 수": wrong})
    return pd.DataFrame(rows)

# -----------------------
# 사용자 입력
# -----------------------
exam_title = st.text_input("통계 제목 입력 (예: 8월 Final mock 1)", value="8월 Final mock 1")
col1, col2 = st.columns(2)
with col1:
    m1_total = st.number_input("Module1 문제 수", min_value=1, value=22)
with col2:
    m2_total = st.number_input("Module2 문제 수", min_value=1, value=22)

uploaded_file = st.file_uploader("📂 학생 오답 현황 엑셀 업로드 (.xlsx)", type="xlsx")

# -----------------------
# 계산 & 다운로드
# -----------------------
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]
        required_cols = {"이름", "Module1", "Module2"}
        if not required_cols.issubset(df.columns):
            st.error(f"엑셀에 {required_cols} 컬럼이 모두 있어야 합니다.")
            st.stop()

        # 파싱
        df["M1_parsed"] = df["Module1"].apply(robust_parse_wrong_list)
        df["M2_parsed"] = df["Module2"].apply(robust_parse_wrong_list)

        # 통계
        m1_stats = compute_module_rates(df["M1_parsed"], m1_total)
        m1_stats["문제 번호"] = m1_stats["문제 번호"].apply(lambda x: f"m1-{x}")
        m2_stats = compute_module_rates(df["M2_parsed"], m2_total)
        m2_stats["문제 번호"] = m2_stats["문제 번호"].apply(lambda x: f"m2-{x}")
        combined = pd.concat([m1_stats, m2_stats], ignore_index=True)[["문제 번호", "오답률(%)", "틀린 학생 수"]]

        st.subheader("미리보기")
        st.dataframe(combined, use_container_width=True)

        # 엑셀 저장 (제목행 + 가운데정렬 + 오답률≥30% 강조)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            sheet_name = "오답률 통계"
            combined.to_excel(writer, index=False, sheet_name=sheet_name, startrow=2)
            wb = writer.book
            ws = writer.sheets[sheet_name]

            # 제목 행
            title_fmt = wb.add_format({"bold": True, "align": "center", "valign": "vcenter"})
            ws.merge_range(0, 0, 0, 2, f"<{exam_title}>", title_fmt)

            # 헤더
            header_fmt = wb.add_format({"bold": True, "align": "center", "valign": "vcenter"})
            ws.write(2, 0, "문제 번호", header_fmt)
            ws.write(2, 1, "오답률(%)", header_fmt)
            ws.write(2, 2, "틀린 학생 수", header_fmt)

            # 가운데 정렬
            center_fmt = wb.add_format({"align": "center", "valign": "vcenter"})
            ws.set_column(0, 2, 14, center_fmt)

            # 오답률 30% 이상 강조 (Bold + 폰트 15)
            cond_fmt = wb.add_format({"bold": True, "font_size": 15, "align": "center", "valign": "vcenter"})
            if len(combined) > 0:
                ws.conditional_format(3, 1, 3 + len(combined) - 1, 1, {
                    "type": "cell", "criteria": ">=", "value": 30, "format": cond_fmt
                })

        output.seek(0)
        st.download_button(
            "📥 오답률 통계 다운로드",
            data=output,
            file_name=f"오답률_통계_{exam_title}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
else:
    st.info("예시를 참고해 엑셀을 준비한 뒤 업로드하면 통계가 생성됩니다.")

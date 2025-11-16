import pandas as pd
from pathlib import Path


def _strip_spaces(s: str) -> str:
    """일반 공백 + NBSP 제거."""
    return s.replace("\xa0", "").strip()


def _is_blank(v) -> bool:
    """완전히 빈 셀인지 판정."""
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    if isinstance(v, str) and _strip_spaces(v) == "":
        return True
    return False


def _is_number(v) -> bool:
    try:
        float(v)
        return True
    except Exception:
        return False


def parse_diet_file(path: Path):
    """
    한 개의 엑셀 파일을 파싱해서 dict 리스트로 반환.
    (날짜, 이름, 양, 분류, 지방, 탄수화물, 단백질, 칼로리)
    """
    df = pd.read_excel(path, header=None)

    # 0행 0열: '화요일 2025년 11월 11일' 같은 날짜 문자열 그대로 사용
    raw_date = df.iloc[0, 0]
    date_str = _strip_spaces(str(raw_date)) if raw_date is not None else ""

    current_meal = None  # '아침 식사', '점심 식사', '저녁 식사', '간식/기타'
    n_rows = len(df)
    records = []

    for i in range(n_rows):
        first = df.iloc[i, 0]

        # ---------- 1) 식사 분류 업데이트 ----------
        if isinstance(first, str):
            cell = _strip_spaces(first)
            if "아침" in cell:
                current_meal = "아침 식사"
            elif "점심" in cell:
                current_meal = "점심 식사"
            elif "저녁" in cell:
                current_meal = "저녁 식사"
            elif "간식/기타" in cell or "간식기타" in cell:
                current_meal = "간식/기타"

        # ---------- 2) 매크로 행인지 판정 ----------
        # 조건: 첫 번째 셀은 비어 있고, 1~4열에 숫자가 하나 이상 있는 경우
        if not _is_blank(first):
            continue

        macro_vals = [df.iloc[i, j] for j in range(1, 5)]
        if not any(_is_number(v) for v in macro_vals):
            # 완전히 빈 줄이면 스킵
            continue

        if current_meal is None:
            # 파일 상단의 전체 합계 같은 건 스킵
            continue

        # ---------- 3) 이 매크로가 가리키는 이름/양 찾기 ----------
        name_row = i + 1
        while name_row < n_rows and _is_blank(df.iloc[name_row, 0]):
            name_row += 1
        if name_row >= n_rows:
            continue

        name_val = df.iloc[name_row, 0]
        name = _strip_spaces(str(name_val))

        qty_row = name_row + 1
        while qty_row < n_rows and _is_blank(df.iloc[qty_row, 0]):
            qty_row += 1

        qty = ""
        if qty_row < n_rows and not _is_blank(df.iloc[qty_row, 0]):
            qty = _strip_spaces(str(df.iloc[qty_row, 0]))

        fat, carb, protein, cal = [
            (float(v) if _is_number(v) else None) for v in macro_vals
        ]

        records.append(
            {
                "날짜": date_str,
                "이름": name,
                "양": qty,
                "분류": current_meal,
                "지방": fat,
                "탄수화물": carb,
                "단백질": protein,
                "칼로리": cal,
            }
        )

    return records


def main():
    base_dir = Path(".")  # 이 스크립트가 있는 디렉터리 기준
    excel_files = sorted(base_dir.glob("*.xlsx"))

    if not excel_files:
        print("*.xlsx 파일이 없습니다.")
        return

    for xlsx in excel_files:
        records = parse_diet_file(xlsx)
        if not records:
            print(f"{xlsx.name}: 변환할 레코드가 없습니다 (건너뜀).")
            continue

        df_out = pd.DataFrame(
            records,
            columns=["날짜", "이름", "양", "분류", "지방", "탄수화물", "단백질", "칼로리"],
        )

        # ---------- 4) 완전히 빈 매크로 줄 제거 ----------
        # (지방/탄/단/칼 네 칸이 모두 비어 있으면 의미 없는 줄이라고 보고 삭제)
        mask_all_empty = df_out[["지방", "탄수화물", "단백질", "칼로리"]].isna().all(axis=1)
        df_out = df_out[~mask_all_empty].reset_index(drop=True)

        out_path = xlsx.with_name(xlsx.stem + "_cleaned.csv")
        df_out.to_csv(out_path, index=False, encoding="utf-8-sig")

        print(f"{xlsx.name} → {out_path.name} ({len(df_out)}행) 저장 완료")


if __name__ == "__main__":
    main()

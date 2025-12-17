"""
재고표 생성 스크립트
본두리편의점 재고실사.xlsx의 정리표를 재고실사양식표 형식으로 변환
"""

import pandas as pd
from datetime import datetime


def format_date(date_value):
    """
    소비기한을 월/일 형식으로 변환
    예: 2026-11-25 -> 11/25
    """
    if pd.isna(date_value):
        return ""

    if isinstance(date_value, str):
        date_obj = pd.to_datetime(date_value)
    else:
        date_obj = date_value

    return date_obj.strftime('%m/%d')


def format_quantity_date(quantity, date_value):
    """
    수량(월/일) 형식으로 변환
    예: 16, 2026-11-25 -> 16(11/25)
    """
    if pd.isna(quantity) or quantity == '' or quantity == 0:
        return ""

    date_str = format_date(date_value)
    if date_str == "":
        return ""

    return f"{int(quantity)}({date_str})"


def create_inventory_report(input_file, output_file):
    """
    재고표 생성

    Args:
        input_file: 본두리편의점 재고실사.xlsx
        output_file: 출력 파일명 (재고표_출력용.xlsx)
    """
    print("📋 재고표 생성 시작...")
    print(f"   입력: {input_file}")

    # 정리표 시트 읽기
    df = pd.read_excel(input_file, sheet_name='정리표', header=1)

    print(f"   ✓ 정리표 데이터 읽기 완료: {len(df)}개 행")

    # 필요한 컬럼만 추출하고 형식 변환
    df['아주/날짜'] = df.apply(
        lambda row: format_quantity_date(row['아주'], row['소비기한']),
        axis=1
    )
    df['잔량/날짜'] = df.apply(
        lambda row: format_quantity_date(row['잔량'], row['소비기한']),
        axis=1
    )

    # 같은 제품코드 + 제품명 + 소비기한을 가진 행들을 그룹화
    # 아주/날짜와 잔량/날짜를 합침
    grouped = df.groupby(['제품코드', 'Brand Name', '소비기한']).agg({
        '아주/날짜': lambda x: ' '.join([v for v in x if v != '']),
        '잔량/날짜': lambda x: ' '.join([v for v in x if v != ''])
    }).reset_index()

    # 결과 DataFrame 생성
    result = pd.DataFrame()
    result['제품코드'] = grouped['제품코드']
    result['제품명'] = grouped['Brand Name']
    result['아주/날짜'] = grouped['아주/날짜']
    result['일반/날짜'] = ""  # 비워둠
    result['잔량/날짜'] = grouped['잔량/날짜']

    # 빈 행 제거 (아주와 잔량 둘 다 없는 행)
    result = result[
        (result['아주/날짜'] != "") | (result['잔량/날짜'] != "")
    ].copy()

    print(f"   ✓ 데이터 변환 완료: {len(result)}개 제품")

    # 엑셀로 저장
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        result.to_excel(writer, sheet_name='재고표', index=False)

    print(f"   ✓ 파일 저장 완료: {output_file}")

    # 결과 미리보기
    print("\n" + "="*70)
    print("📊 결과 미리보기 (상위 20개)")
    print("="*70)
    print(result.head(20).to_string(index=False))

    print("\n" + "="*70)
    print("✅ 재고표 생성 완료!")
    print("="*70)
    print(f"출력 파일: {output_file}")
    print(f"총 {len(result)}개 제품")
    print("\n이제 엑셀 파일을 열어서 확인하세요.")

    return result


if __name__ == "__main__":
    # 파일 경로
    input_file = "본두리편의점 재고실사.xlsx"
    output_file = "재고표_출력용.xlsx"

    try:
        result = create_inventory_report(input_file, output_file)
    except FileNotFoundError:
        print(f"❌ 오류: {input_file} 파일을 찾을 수 없습니다.")
        print(f"   현재 디렉토리에 파일이 있는지 확인하세요.")
    except Exception as e:
        print(f"❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()

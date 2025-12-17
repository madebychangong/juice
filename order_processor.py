"""
주문 데이터 처리 모듈
- SAP 주문 파일 읽기
- 업체명 매핑
- 마이너스 수량 제거
- 업체별/제품별 집계
"""

import pandas as pd


class OrderProcessor:
    def __init__(self, order_file, company_mapping_file):
        """
        주문 처리기 초기화

        Args:
            order_file: SAP 주문 파일 경로
            company_mapping_file: 업체명 정보 파일 경로
        """
        self.order_file = order_file
        self.company_mapping_file = company_mapping_file
        self.df_order = None
        self.df_company_mapping = None
        self.df_processed = None

    def load_data(self):
        """데이터 파일 로드"""
        print("📂 데이터 로드 중...")

        # SAP 주문 파일 읽기
        self.df_order = pd.read_excel(self.order_file)
        print(f"   ✓ 주문 데이터: {len(self.df_order)}개 행")

        # 업체명 매핑 파일 읽기
        self.df_company_mapping = pd.read_excel(
            self.company_mapping_file,
            sheet_name='업체명'
        )
        print(f"   ✓ 업체명 매핑: {len(self.df_company_mapping)}개 업체")

        return self

    def process_orders(self):
        """주문 데이터 처리"""
        print("\n🔄 주문 데이터 처리 중...")

        # 필요한 컬럼만 선택
        df = self.df_order[['납품처명', '자재내역', '주문수량']].copy()

        # 양수 수량만 필터링 (마이너스 제거)
        original_count = len(df)
        df = df[df['주문수량'] > 0].copy()
        removed_count = original_count - len(df)
        print(f"   ✓ 마이너스 수량 제거: {removed_count}개 항목 제외")

        # 업체명 매핑
        mapping_dict = dict(zip(
            self.df_company_mapping['센터명'],
            self.df_company_mapping['코드']
        ))
        df['업체코드'] = df['납품처명'].map(mapping_dict)

        # 매핑되지 않은 업체 확인
        unmapped = df[df['업체코드'].isna()]['납품처명'].unique()
        if len(unmapped) > 0:
            print(f"   ⚠️  매핑되지 않은 업체 ({len(unmapped)}개):")
            for company in unmapped:
                print(f"      - {company}")
            # 매핑되지 않은 업체는 원래 이름 사용
            df.loc[df['업체코드'].isna(), '업체코드'] = df['납품처명']

        # 컬럼 순서 정리
        df = df[['업체코드', '납품처명', '자재내역', '주문수량']]
        df.columns = ['업체코드', '업체명_원본', '제품명', '수량']

        self.df_processed = df
        print(f"   ✓ 처리 완료: {len(df)}개 주문 항목")
        print(f"   ✓ 고유 업체 수: {df['업체코드'].nunique()}개")
        print(f"   ✓ 고유 제품 수: {df['제품명'].nunique()}개")

        return self

    def get_summary_by_company(self):
        """업체별 주문 집계"""
        if self.df_processed is None:
            raise ValueError("먼저 process_orders()를 실행하세요")

        summary = self.df_processed.groupby('업체코드').agg({
            '업체명_원본': 'first',
            '제품명': 'count',
            '수량': 'sum'
        }).reset_index()
        summary.columns = ['업체코드', '업체명_원본', '제품종류수', '총수량']

        return summary.sort_values('업체코드')

    def get_orders_by_company(self, company_code):
        """특정 업체의 주문 내역 조회"""
        if self.df_processed is None:
            raise ValueError("먼저 process_orders()를 실행하세요")

        return self.df_processed[
            self.df_processed['업체코드'] == company_code
        ].copy()

    def get_all_companies(self):
        """모든 업체 코드 목록 조회"""
        if self.df_processed is None:
            raise ValueError("먼저 process_orders()를 실행하세요")

        return self.df_processed['업체코드'].unique()

    def save_processed_data(self, output_file):
        """처리된 데이터 저장"""
        if self.df_processed is None:
            raise ValueError("먼저 process_orders()를 실행하세요")

        self.df_processed.to_excel(output_file, index=False)
        print(f"\n💾 처리된 데이터 저장: {output_file}")

        return self


if __name__ == "__main__":
    # 테스트
    processor = OrderProcessor(
        order_file='TalkFile_SAP 주문파일.xlsx.xlsx',
        company_mapping_file='TalkFile_업체명 정보파일.xlsx.xlsx'
    )

    processor.load_data()
    processor.process_orders()

    print("\n" + "="*60)
    print("📊 업체별 주문 집계")
    print("="*60)
    summary = processor.get_summary_by_company()
    print(summary.to_string(index=False))

    # 샘플: 특정 업체 주문 내역
    companies = processor.get_all_companies()
    if len(companies) > 0:
        sample_company = companies[0]
        print(f"\n{'='*60}")
        print(f"📋 샘플 업체 주문 내역: {sample_company}")
        print("="*60)
        sample_orders = processor.get_orders_by_company(sample_company)
        print(sample_orders.to_string(index=False))

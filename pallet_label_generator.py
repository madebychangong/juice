"""
팔레트 라벨 PDF 생성 모듈
- 업체당 제품당 한 장의 A4 라벨 생성
- 큰 글씨로 제품명, 수량, 업체코드 표시
"""

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm
import os


class PalletLabelGenerator:
    def __init__(self, output_dir='labels'):
        """
        팔레트 라벨 생성기 초기화

        Args:
            output_dir: PDF 파일 저장 디렉토리
        """
        self.output_dir = output_dir
        self.width, self.height = A4

        # 출력 디렉토리 생성
        os.makedirs(output_dir, exist_ok=True)

        # 한글 폰트 등록 (나눔고딕)
        self._register_fonts()

    def _register_fonts(self):
        """한글 폰트 등록"""
        try:
            # 시스템 폰트 경로 시도
            font_paths = [
                '/usr/share/fonts/truetype/nanum/NanumGothic.ttf',
                '/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf',
                '/System/Library/Fonts/AppleSDGothicNeo.ttc',
                'C:\\Windows\\Fonts\\malgun.ttf',
            ]

            font_registered = False
            for font_path in font_paths:
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(TTFont('NanumGothic', font_path))
                        font_registered = True
                        print(f"   ✓ 폰트 등록: {font_path}")
                        break
                    except:
                        continue

            if not font_registered:
                print("   ⚠️  시스템 폰트를 찾을 수 없습니다. 기본 폰트를 사용합니다.")
                self.font_name = 'Helvetica'
            else:
                self.font_name = 'NanumGothic'

        except Exception as e:
            print(f"   ⚠️  폰트 등록 실패: {e}")
            self.font_name = 'Helvetica'

    def create_label(self, product_name, quantity, company_code, page_canvas):
        """
        단일 라벨 생성 (한 페이지)

        Args:
            product_name: 제품명
            quantity: 수량
            company_code: 업체코드
            page_canvas: reportlab canvas 객체
        """
        # 페이지 중앙 좌표
        center_x = self.width / 2
        center_y = self.height / 2

        # 제품명 (상단 중앙, 큰 글씨)
        page_canvas.setFont(self.font_name, 48)
        # 긴 제품명 처리 - 줄바꿈
        if len(product_name) > 20:
            # 제품명을 두 줄로 나누기
            mid = len(product_name) // 2
            # 공백 기준으로 나누기
            if ' ' in product_name:
                words = product_name.split(' ')
                mid_point = len(words) // 2
                line1 = ' '.join(words[:mid_point])
                line2 = ' '.join(words[mid_point:])
            else:
                line1 = product_name[:mid]
                line2 = product_name[mid:]

            page_canvas.drawCentredString(center_x, center_y + 5*cm, line1)
            page_canvas.drawCentredString(center_x, center_y + 3*cm, line2)
        else:
            page_canvas.drawCentredString(center_x, center_y + 4*cm, product_name)

        # 수량 (중앙, 매우 큰 글씨)
        page_canvas.setFont(self.font_name, 72)
        quantity_text = f"{quantity}"
        page_canvas.drawCentredString(center_x, center_y, quantity_text)

        # 업체코드 (하단 중앙, 중간 크기 글씨)
        page_canvas.setFont(self.font_name, 36)
        page_canvas.drawCentredString(center_x, center_y - 4*cm, str(company_code))

        # 테두리 (선택사항)
        page_canvas.rect(1*cm, 1*cm, self.width - 2*cm, self.height - 2*cm, stroke=1, fill=0)

    def generate_labels_for_company(self, company_code, orders_df):
        """
        특정 업체의 모든 제품 라벨 생성

        Args:
            company_code: 업체코드
            orders_df: 해당 업체의 주문 DataFrame (컬럼: 제품명, 수량)

        Returns:
            생성된 PDF 파일 경로
        """
        # 파일명 생성 (특수문자 제거)
        safe_company_code = str(company_code).replace('/', '_').replace('\\', '_')
        pdf_filename = f"{safe_company_code}_라벨.pdf"
        pdf_path = os.path.join(self.output_dir, pdf_filename)

        # PDF 생성
        c = canvas.Canvas(pdf_path, pagesize=A4)

        # 각 제품별로 라벨 생성
        for idx, row in orders_df.iterrows():
            product_name = row['제품명']
            quantity = row['수량']

            self.create_label(product_name, quantity, company_code, c)
            c.showPage()  # 다음 페이지로

        # PDF 저장
        c.save()

        return pdf_path

    def generate_all_labels(self, processed_orders_df):
        """
        모든 업체의 라벨 생성

        Args:
            processed_orders_df: 처리된 주문 DataFrame
                               (컬럼: 업체코드, 업체명_원본, 제품명, 수량)

        Returns:
            생성된 PDF 파일 경로 리스트
        """
        print("\n📄 팔레트 라벨 PDF 생성 중...")

        pdf_files = []
        companies = processed_orders_df['업체코드'].unique()

        for idx, company_code in enumerate(companies, 1):
            # 해당 업체의 주문만 필터링
            company_orders = processed_orders_df[
                processed_orders_df['업체코드'] == company_code
            ][['제품명', '수량']].copy()

            # 라벨 생성
            pdf_path = self.generate_labels_for_company(
                company_code,
                company_orders
            )

            pdf_files.append(pdf_path)
            pdf_filename = os.path.basename(pdf_path)
            print(f"   [{idx}/{len(companies)}] {company_code}: {len(company_orders)}개 라벨 → {pdf_filename}")

        print(f"\n   ✓ 총 {len(pdf_files)}개 PDF 파일 생성 완료")
        print(f"   ✓ 저장 위치: {self.output_dir}/")

        return pdf_files


if __name__ == "__main__":
    # 테스트
    from order_processor import OrderProcessor

    # 주문 데이터 처리
    processor = OrderProcessor(
        order_file='TalkFile_SAP 주문파일.xlsx.xlsx',
        company_mapping_file='TalkFile_업체명 정보파일.xlsx.xlsx'
    )
    processor.load_data().process_orders()

    # 라벨 생성
    label_gen = PalletLabelGenerator(output_dir='labels')
    pdf_files = label_gen.generate_all_labels(processor.df_processed)

    print("\n✅ 라벨 생성 완료!")
    print(f"   총 {len(pdf_files)}개 PDF 파일 생성")

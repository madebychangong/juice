"""
업체별 주문서 PDF 생성 모듈
- 업체별로 주문 내역을 정리한 주문서 생성
- 제품명과 수량을 표 형식으로 표시
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import os


class OrderSheetGenerator:
    def __init__(self, output_dir='order_sheets'):
        """
        주문서 생성기 초기화

        Args:
            output_dir: PDF 파일 저장 디렉토리
        """
        self.output_dir = output_dir

        # 출력 디렉토리 생성
        os.makedirs(output_dir, exist_ok=True)

        # 한글 폰트 등록
        self._register_fonts()

    def _register_fonts(self):
        """한글 폰트 등록"""
        try:
            font_paths = [
                '/usr/share/fonts/truetype/nanum/NanumGothic.ttf',
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
                print("   ⚠️  시스템 폰트를 찾을 수 없습니다.")
                self.font_available = False
            else:
                self.font_available = True

        except Exception as e:
            print(f"   ⚠️  폰트 등록 실패: {e}")
            self.font_available = False

    def create_order_sheet(self, company_code, company_name_original, orders_df, output_path):
        """
        업체별 주문서 PDF 생성

        Args:
            company_code: 업체코드
            company_name_original: 업체명 원본
            orders_df: 주문 DataFrame (컬럼: 제품명, 수량)
            output_path: 저장할 PDF 파일 경로
        """
        # PDF 문서 생성
        doc = SimpleDocTemplate(
            output_path,
            pagesize=A4,
            topMargin=1.5*cm,
            bottomMargin=1.5*cm,
            leftMargin=2*cm,
            rightMargin=2*cm
        )

        # 스타일 정의
        styles = getSampleStyleSheet()

        if self.font_available:
            # 한글 폰트 사용
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Title'],
                fontName='NanumGothic',
                fontSize=24,
                alignment=TA_CENTER,
                spaceAfter=12
            )
            heading_style = ParagraphStyle(
                'CustomHeading',
                parent=styles['Heading2'],
                fontName='NanumGothic',
                fontSize=14,
                alignment=TA_LEFT,
                spaceAfter=12
            )
            normal_style = ParagraphStyle(
                'CustomNormal',
                parent=styles['Normal'],
                fontName='NanumGothic',
                fontSize=11
            )
        else:
            title_style = styles['Title']
            heading_style = styles['Heading2']
            normal_style = styles['Normal']

        # 문서 요소 리스트
        elements = []

        # 제목
        title = Paragraph(f"주문서", title_style)
        elements.append(title)
        elements.append(Spacer(1, 0.5*cm))

        # 업체 정보
        company_info = Paragraph(
            f"<b>업체코드:</b> {company_code}<br/><b>업체명:</b> {company_name_original}",
            heading_style
        )
        elements.append(company_info)
        elements.append(Spacer(1, 0.3*cm))

        # 총 품목 수와 총 수량
        total_items = len(orders_df)
        total_quantity = orders_df['수량'].sum()
        summary = Paragraph(
            f"<b>총 품목 수:</b> {total_items}개 | <b>총 수량:</b> {total_quantity:,}",
            normal_style
        )
        elements.append(summary)
        elements.append(Spacer(1, 0.5*cm))

        # 주문 테이블 생성
        table_data = [['No.', '제품명', '수량']]

        for idx, row in orders_df.iterrows():
            table_data.append([
                str(idx + 1),
                row['제품명'],
                f"{row['수량']:,}"
            ])

        # 테이블 스타일 정의
        table = Table(table_data, colWidths=[2*cm, 12*cm, 3*cm])

        table_style = TableStyle([
            # 헤더 스타일
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'NanumGothic' if self.font_available else 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),

            # 데이터 스타일
            ('FONTNAME', (0, 1), (-1, -1), 'NanumGothic' if self.font_available else 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('ALIGN', (0, 1), (0, -1), 'CENTER'),  # No. 열
            ('ALIGN', (1, 1), (1, -1), 'LEFT'),    # 제품명 열
            ('ALIGN', (2, 1), (2, -1), 'RIGHT'),   # 수량 열

            # 테두리
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),

            # 번갈아 배경색
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
        ])

        table.setStyle(table_style)
        elements.append(table)

        # PDF 생성
        doc.build(elements)

    def generate_all_order_sheets(self, processed_orders_df):
        """
        모든 업체의 주문서 생성

        Args:
            processed_orders_df: 처리된 주문 DataFrame
                               (컬럼: 업체코드, 업체명_원본, 제품명, 수량)

        Returns:
            생성된 PDF 파일 경로 리스트
        """
        print("\n📋 업체별 주문서 PDF 생성 중...")

        pdf_files = []
        companies = processed_orders_df['업체코드'].unique()

        for idx, company_code in enumerate(companies, 1):
            # 해당 업체의 주문만 필터링
            company_orders = processed_orders_df[
                processed_orders_df['업체코드'] == company_code
            ].copy()

            company_name_original = company_orders['업체명_원본'].iloc[0]

            # 파일명 생성
            safe_company_code = str(company_code).replace('/', '_').replace('\\', '_')
            pdf_filename = f"{safe_company_code}_주문서.pdf"
            pdf_path = os.path.join(self.output_dir, pdf_filename)

            # 주문서 생성
            self.create_order_sheet(
                company_code,
                company_name_original,
                company_orders[['제품명', '수량']].reset_index(drop=True),
                pdf_path
            )

            pdf_files.append(pdf_path)
            print(f"   [{idx}/{len(companies)}] {company_code}: {len(company_orders)}개 품목 → {pdf_filename}")

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

    # 주문서 생성
    sheet_gen = OrderSheetGenerator(output_dir='order_sheets')
    pdf_files = sheet_gen.generate_all_order_sheets(processor.df_processed)

    print("\n✅ 주문서 생성 완료!")
    print(f"   총 {len(pdf_files)}개 PDF 파일 생성")

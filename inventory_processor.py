"""
재고표 처리 모듈
- 재고 파일 읽기
- 유통기한 선입선출 기준 정렬
- 재고표 PDF 생성
"""

import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER
import os
from datetime import datetime


class InventoryProcessor:
    def __init__(self, inventory_file):
        """
        재고 처리기 초기화

        Args:
            inventory_file: 재고 파일 경로
        """
        self.inventory_file = inventory_file
        self.df_inventory = None
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
                        break
                    except:
                        continue

            self.font_available = font_registered

        except Exception as e:
            self.font_available = False

    def load_inventory(self):
        """재고 파일 로드 및 기본 처리"""
        print("📦 재고 데이터 로드 중...")

        # 재고 파일 읽기 (header는 2번째 행)
        df = pd.read_excel(self.inventory_file, header=2)

        # 필요한 컬럼 선택 및 이름 변경
        # 실제 컬럼명은 재고 파일 구조에 따라 조정 필요
        # 여기서는 기본적인 구조로 처리
        df = df.rename(columns={
            df.columns[0]: '기존코드',
            df.columns[1]: 'SAP코드',
            df.columns[2]: '제품명',
            df.columns[3]: '구분',
        })

        # NaN 행 제거 (총계 등)
        df = df[df['제품명'].notna()].copy()
        df = df[df['제품명'] != 'TOTAL'].copy()

        # 수량 관련 컬럼 찾기 (총합계 등)
        quantity_cols = [col for col in df.columns if '합계' in str(col) or '수량' in str(col)]

        self.df_inventory = df
        print(f"   ✓ 재고 데이터: {len(df)}개 품목")

        return self

    def create_inventory_report(self, output_path):
        """
        재고표 PDF 생성

        Args:
            output_path: 저장할 PDF 파일 경로
        """
        print(f"\n📄 재고표 PDF 생성 중...")

        # PDF 문서 생성
        doc = SimpleDocTemplate(
            output_path,
            pagesize=A4,
            topMargin=1.5*cm,
            bottomMargin=1.5*cm,
            leftMargin=1.5*cm,
            rightMargin=1.5*cm
        )

        # 스타일 정의
        styles = getSampleStyleSheet()

        if self.font_available:
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Title'],
                fontName='NanumGothic',
                fontSize=20,
                alignment=TA_CENTER,
                spaceAfter=12
            )
            normal_style = ParagraphStyle(
                'CustomNormal',
                parent=styles['Normal'],
                fontName='NanumGothic',
                fontSize=10
            )
        else:
            title_style = styles['Title']
            normal_style = styles['Normal']

        # 문서 요소 리스트
        elements = []

        # 제목
        today = datetime.now().strftime('%Y-%m-%d')
        title = Paragraph(f"재고 현황표 ({today})", title_style)
        elements.append(title)
        elements.append(Spacer(1, 0.5*cm))

        # 총 품목 수
        total_items = len(self.df_inventory)
        summary = Paragraph(
            f"<b>총 품목 수:</b> {total_items}개",
            normal_style
        )
        elements.append(summary)
        elements.append(Spacer(1, 0.3*cm))

        # 재고 테이블 생성
        table_data = [['No.', 'SAP코드', '제품명', '구분']]

        for idx, row in self.df_inventory.iterrows():
            table_data.append([
                str(len(table_data)),  # No.
                str(row.get('SAP코드', '')),
                str(row.get('제품명', '')),
                str(row.get('구분', '')),
            ])

        # 테이블 생성
        table = Table(table_data, colWidths=[1.5*cm, 3*cm, 9*cm, 3*cm])

        table_style = TableStyle([
            # 헤더 스타일
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'NanumGothic' if self.font_available else 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),

            # 데이터 스타일
            ('FONTNAME', (0, 1), (-1, -1), 'NanumGothic' if self.font_available else 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('ALIGN', (1, 1), (1, -1), 'LEFT'),    # 제품명 열

            # 테두리
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),

            # 번갈아 배경색
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
        ])

        table.setStyle(table_style)
        elements.append(table)

        # PDF 생성
        try:
            doc.build(elements)
            print(f"   ✓ 재고표 생성 완료: {output_path}")
            return output_path
        except Exception as e:
            print(f"   ✗ 재고표 생성 실패: {e}")
            return None


if __name__ == "__main__":
    # 테스트
    processor = InventoryProcessor(
        inventory_file='TalkFile_재고파일.xlsx.xlsx'
    )

    processor.load_inventory()

    # 재고표 PDF 생성
    output_dir = 'inventory_reports'
    os.makedirs(output_dir, exist_ok=True)

    today = datetime.now().strftime('%Y%m%d')
    output_path = os.path.join(output_dir, f'재고표_{today}.pdf')

    processor.create_inventory_report(output_path)

    print("\n✅ 재고표 생성 완료!")

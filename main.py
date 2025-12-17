#!/usr/bin/env python3
"""
음료회사 주문/라벨/재고 자동화 시스템
메인 실행 스크립트

사용법:
    python main.py [옵션]

옵션:
    --all              모든 기능 실행 (기본값)
    --orders-only      업체별 주문서만 생성
    --labels-only      팔레트 라벨만 생성
    --inventory-only   재고표만 생성
    --order-file       SAP 주문 파일 경로 (기본: TalkFile_SAP 주문파일.xlsx.xlsx)
    --company-file     업체명 정보 파일 경로 (기본: TalkFile_업체명 정보파일.xlsx.xlsx)
    --inventory-file   재고 파일 경로 (기본: TalkFile_재고파일.xlsx.xlsx)
"""

import argparse
import os
from datetime import datetime
from order_processor import OrderProcessor
from order_sheet_generator import OrderSheetGenerator
from pallet_label_generator import PalletLabelGenerator
from inventory_processor import InventoryProcessor


def print_banner():
    """프로그램 시작 배너 출력"""
    print("=" * 70)
    print("  음료회사 주문/라벨/재고 자동화 시스템")
    print("  Order & Label & Inventory Automation System")
    print("=" * 70)
    print()


def main():
    # 명령행 인자 파싱
    parser = argparse.ArgumentParser(
        description='음료회사 주문/라벨/재고 자동화 시스템'
    )
    parser.add_argument('--all', action='store_true', help='모든 기능 실행 (기본값)')
    parser.add_argument('--orders-only', action='store_true', help='업체별 주문서만 생성')
    parser.add_argument('--labels-only', action='store_true', help='팔레트 라벨만 생성')
    parser.add_argument('--inventory-only', action='store_true', help='재고표만 생성')
    parser.add_argument('--order-file', default='TalkFile_SAP 주문파일.xlsx.xlsx',
                        help='SAP 주문 파일 경로')
    parser.add_argument('--company-file', default='TalkFile_업체명 정보파일.xlsx.xlsx',
                        help='업체명 정보 파일 경로')
    parser.add_argument('--inventory-file', default='TalkFile_재고파일.xlsx.xlsx',
                        help='재고 파일 경로')

    args = parser.parse_args()

    # 옵션이 하나도 선택되지 않았으면 --all로 설정
    if not any([args.orders_only, args.labels_only, args.inventory_only]):
        args.all = True

    # 배너 출력
    print_banner()

    # 타임스탬프
    start_time = datetime.now()
    timestamp = start_time.strftime('%Y%m%d_%H%M%S')
    print(f"실행 시작: {start_time.strftime('%Y-%m-%d %H:%M:%S')}\n")

    # 파일 존재 확인
    if args.all or args.orders_only or args.labels_only:
        if not os.path.exists(args.order_file):
            print(f"❌ 오류: 주문 파일을 찾을 수 없습니다: {args.order_file}")
            return
        if not os.path.exists(args.company_file):
            print(f"❌ 오류: 업체명 정보 파일을 찾을 수 없습니다: {args.company_file}")
            return

    if args.all or args.inventory_only:
        if not os.path.exists(args.inventory_file):
            print(f"❌ 오류: 재고 파일을 찾을 수 없습니다: {args.inventory_file}")
            return

    # 출력 디렉토리 생성
    output_base_dir = f'output_{timestamp}'
    os.makedirs(output_base_dir, exist_ok=True)

    try:
        # 1. 주문 관련 처리 (주문서 또는 라벨 생성)
        if args.all or args.orders_only or args.labels_only:
            print("🔄 STEP 1: 주문 데이터 처리")
            print("-" * 70)

            processor = OrderProcessor(
                order_file=args.order_file,
                company_mapping_file=args.company_file
            )
            processor.load_data().process_orders()

            # 처리된 데이터 저장
            processed_file = os.path.join(output_base_dir, '처리된_주문데이터.xlsx')
            processor.save_processed_data(processed_file)

            # 업체별 주문서 생성
            if args.all or args.orders_only:
                print("\n" + "=" * 70)
                print("🔄 STEP 2: 업체별 주문서 생성")
                print("-" * 70)

                order_sheet_dir = os.path.join(output_base_dir, 'order_sheets')
                sheet_gen = OrderSheetGenerator(output_dir=order_sheet_dir)
                order_pdfs = sheet_gen.generate_all_order_sheets(processor.df_processed)

            # 팔레트 라벨 생성
            if args.all or args.labels_only:
                print("\n" + "=" * 70)
                print("🔄 STEP 3: 팔레트 라벨 생성")
                print("-" * 70)

                label_dir = os.path.join(output_base_dir, 'labels')
                label_gen = PalletLabelGenerator(output_dir=label_dir)
                label_pdfs = label_gen.generate_all_labels(processor.df_processed)

        # 2. 재고표 생성
        if args.all or args.inventory_only:
            print("\n" + "=" * 70)
            print("🔄 STEP 4: 재고표 생성")
            print("-" * 70)

            inventory_processor = InventoryProcessor(
                inventory_file=args.inventory_file
            )
            inventory_processor.load_inventory()

            inventory_dir = os.path.join(output_base_dir, 'inventory_reports')
            os.makedirs(inventory_dir, exist_ok=True)

            today = datetime.now().strftime('%Y%m%d')
            inventory_pdf = os.path.join(inventory_dir, f'재고표_{today}.pdf')
            inventory_processor.create_inventory_report(inventory_pdf)

        # 완료 메시지
        end_time = datetime.now()
        elapsed_time = (end_time - start_time).total_seconds()

        print("\n" + "=" * 70)
        print("✅ 모든 작업 완료!")
        print("=" * 70)
        print(f"실행 시간: {elapsed_time:.2f}초")
        print(f"\n📁 출력 디렉토리: {output_base_dir}/")

        if args.all or args.orders_only:
            print(f"   ├── order_sheets/    (업체별 주문서)")
        if args.all or args.labels_only:
            print(f"   ├── labels/          (팔레트 라벨)")
        if args.all or args.inventory_only:
            print(f"   └── inventory_reports/ (재고표)")

        print("\n💡 사용 팁:")
        print("   1. 생성된 PDF 파일을 열어서 확인하세요")
        print("   2. 업체별로 필요한 PDF만 선택해서 인쇄할 수 있습니다")
        print("   3. 라벨은 A4 용지에 큰 글씨로 출력되어 현장에서 사용할 수 있습니다")

    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return 1

    return 0


if __name__ == "__main__":
    exit(main())

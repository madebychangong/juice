# 음료회사 주문/라벨/재고 자동화 시스템

음료회사의 입출고 업무를 자동화하는 시스템입니다.
SAP에서 받은 주문 데이터를 처리하여 업체별 주문서, 팔레트 라벨, 재고표를 자동으로 생성합니다.

## 주요 기능

### 1. 업체별 주문서 생성
- SAP 주문 데이터를 업체별로 정리
- 업체명을 코드로 자동 변환 (예: BGF화성 → 중부, GS발안 → 209)
- 마이너스 수량(반품) 자동 제거
- 업체별로 PDF 주문서 생성

### 2. 팔레트 라벨 생성
- 업체당 제품당 한 장의 A4 라벨 생성
- 큰 글씨로 제품명, 수량, 업체코드 표시
- 현장에서 지게차 작업자가 한눈에 확인 가능
- 업체별로 PDF 파일 분리 저장

### 3. 재고표 생성
- 유통기한 선입선출용 재고 현황표
- 제품별 재고 수량 정리
- PDF 형식으로 저장

## 시스템 구조

```
juice/
├── main.py                      # 메인 실행 스크립트
├── order_processor.py           # 주문 데이터 처리 모듈
├── order_sheet_generator.py    # 업체별 주문서 PDF 생성
├── pallet_label_generator.py   # 팔레트 라벨 PDF 생성
├── inventory_processor.py      # 재고표 처리 및 생성
├── TalkFile_SAP 주문파일.xlsx.xlsx     # SAP 주문 데이터 (입력)
├── TalkFile_업체명 정보파일.xlsx.xlsx  # 업체명 매핑 테이블 (입력)
├── TalkFile_재고파일.xlsx.xlsx         # 재고 데이터 (입력)
└── output_YYYYMMDD_HHMMSS/     # 출력 디렉토리 (자동 생성)
    ├── order_sheets/            # 업체별 주문서 PDF
    ├── labels/                  # 팔레트 라벨 PDF
    ├── inventory_reports/       # 재고표 PDF
    └── 처리된_주문데이터.xlsx     # 처리된 주문 데이터 (엑셀)
```

## 설치 방법

### 필수 요구사항
- Python 3.7 이상
- pip (Python 패키지 관리자)

### 1. 필요한 패키지 설치

```bash
pip install pandas openpyxl reportlab
```

### 2. 한글 폰트 설치 (Linux/Ubuntu)

```bash
sudo apt-get update
sudo apt-get install -y fonts-nanum fonts-nanum-coding
```

Windows나 macOS의 경우 시스템에 이미 한글 폰트가 설치되어 있으면 자동으로 사용됩니다.

## 사용 방법

### 기본 사용법

모든 기능을 한 번에 실행:

```bash
python3 main.py
```

또는

```bash
python3 main.py --all
```

### 선택적 실행

특정 기능만 실행하고 싶을 때:

```bash
# 업체별 주문서만 생성
python3 main.py --orders-only

# 팔레트 라벨만 생성
python3 main.py --labels-only

# 재고표만 생성
python3 main.py --inventory-only
```

### 파일 경로 지정

기본 파일명이 아닌 다른 파일을 사용하고 싶을 때:

```bash
python3 main.py \
  --order-file "주문데이터_20251217.xlsx" \
  --company-file "업체정보.xlsx" \
  --inventory-file "재고현황.xlsx"
```

## 입력 파일 형식

### 1. SAP 주문 파일 (TalkFile_SAP 주문파일.xlsx.xlsx)

**필수 컬럼:**
- `납품처명`: 업체명 (예: 신)BGF로지스 화성)
- `자재내역`: 제품명 (예: S)코크제로 250CAN 5X6)
- `주문수량`: 주문 수량 (양수: 주문, 음수: 반품)

**샘플:**
| 납품처명 | 자재내역 | 주문수량 |
|---------|---------|---------|
| 신)BGF로지스 화성 | S)코크제로 250CAN 5X6 | 12 |
| GS25 발안센터_C | S)조지아 오리지널 240CAN 5X6 | 24 |

### 2. 업체명 정보 파일 (TalkFile_업체명 정보파일.xlsx.xlsx)

**시트명:** `업체명`

**필수 컬럼:**
- `센터명`: 원본 업체명
- `코드`: 변환할 업체 코드

**샘플:**
| 센터명 | 코드 |
|-------|-----|
| 신)BGF로지스 화성 | 중부 |
| GS25 발안센터_C | 209 |
| 코리아세븐 양주물류센터 | 6 |

### 3. 재고 파일 (TalkFile_재고파일.xlsx.xlsx)

재고 현황 데이터가 포함된 엑셀 파일입니다.
시스템이 자동으로 헤더를 인식하고 처리합니다.

## 출력 파일

시스템 실행 시 `output_YYYYMMDD_HHMMSS/` 디렉토리가 자동으로 생성되며, 다음 파일들이 생성됩니다:

### 1. 업체별 주문서 (order_sheets/)
- 파일명 형식: `{업체코드}_주문서.pdf`
- 예시: `중부_주문서.pdf`, `209_주문서.pdf`
- 내용: 해당 업체의 모든 주문 품목과 수량을 표 형식으로 정리

### 2. 팔레트 라벨 (labels/)
- 파일명 형식: `{업체코드}_라벨.pdf`
- 예시: `중부_라벨.pdf`, `209_라벨.pdf`
- 내용: 업체의 각 제품별로 한 페이지씩 큰 글씨 라벨
  ```
  제품명
  수량
  업체코드
  ```

### 3. 재고표 (inventory_reports/)
- 파일명 형식: `재고표_YYYYMMDD.pdf`
- 예시: `재고표_20251217.pdf`
- 내용: 전체 재고 현황표

### 4. 처리된 주문 데이터
- 파일명: `처리된_주문데이터.xlsx`
- 내용: 마이너스 제거, 업체명 변환이 완료된 주문 데이터

## 업무 프로세스

### 기존 수동 프로세스
1. SAP에서 주문 데이터를 엑셀로 다운로드
2. 수동으로 업체명을 코드로 변경
3. 마이너스 수량 찾아서 수동으로 제거
4. 업체별로 데이터 정리
5. 워드/엑셀로 라벨 수동 작성
6. 라벨 하나씩 인쇄

### 자동화된 프로세스
1. SAP에서 주문 데이터를 엑셀로 다운로드
2. `python3 main.py` 실행
3. 생성된 PDF 파일 확인 및 인쇄

**시간 절약:** 수작업 2~3시간 → 자동화 2분

## 사용 예시

### 매일 아침 주문 처리

```bash
# 1. SAP에서 주문 데이터 다운로드 (기존 파일명으로 저장)
# 2. 프로그램 실행
python3 main.py

# 3. 생성된 PDF 확인
ls output_*/

# 4. 필요한 업체의 PDF만 선택해서 인쇄
# 예: 중부_라벨.pdf, 209_주문서.pdf 등
```

### 특정 업체만 재출력

```bash
# 1. 처리된 데이터 확인
python3 -c "
import pandas as pd
df = pd.read_excel('output_20251217_042314/처리된_주문데이터.xlsx')
print(df[df['업체코드'] == '중부'])
"

# 2. 해당 업체 PDF 재인쇄
# output_*/labels/중부_라벨.pdf
# output_*/order_sheets/중부_주문서.pdf
```

## 문제 해결

### Q: "매핑되지 않은 업체" 경고가 나옵니다

A: `TalkFile_업체명 정보파일.xlsx.xlsx`의 '업체명' 시트에 해당 업체를 추가하세요.

**예시:**
```
센터명: 신규업체명
코드: 새코드
```

### Q: 한글이 PDF에서 깨져 보입니다

A: 한글 폰트를 설치하세요.

```bash
# Linux/Ubuntu
sudo apt-get install fonts-nanum

# 폰트 확인
ls /usr/share/fonts/truetype/nanum/
```

### Q: 라벨의 글씨 크기를 조정하고 싶습니다

A: `pallet_label_generator.py` 파일에서 폰트 크기를 수정할 수 있습니다.

```python
# 61번째 줄 근처
page_canvas.setFont(self.font_name, 48)  # 제품명 크기
page_canvas.setFont(self.font_name, 72)  # 수량 크기
page_canvas.setFont(self.font_name, 36)  # 업체코드 크기
```

### Q: 마이너스 수량이 제대로 제거되지 않습니다

A: 시스템은 `주문수량 > 0`인 항목만 처리합니다.
만약 같은 제품의 +1과 -1을 합산해야 한다면, `order_processor.py`를 수정해야 합니다.

## 고급 사용

### 개별 모듈 사용

각 모듈을 개별적으로 테스트하거나 사용할 수 있습니다:

```python
# 주문 데이터만 처리
from order_processor import OrderProcessor

processor = OrderProcessor(
    order_file='TalkFile_SAP 주문파일.xlsx.xlsx',
    company_mapping_file='TalkFile_업체명 정보파일.xlsx.xlsx'
)
processor.load_data().process_orders()

# 업체별 집계 확인
summary = processor.get_summary_by_company()
print(summary)

# 특정 업체 주문만 확인
orders = processor.get_orders_by_company('중부')
print(orders)
```

### 데이터 분석

```python
import pandas as pd

# 처리된 데이터 읽기
df = pd.read_excel('output_20251217_042314/처리된_주문데이터.xlsx')

# 업체별 총 수량
print(df.groupby('업체코드')['수량'].sum().sort_values(ascending=False))

# 인기 제품 TOP 10
print(df.groupby('제품명')['수량'].sum().sort_values(ascending=False).head(10))
```

## 업데이트 이력

### v1.0 (2025-12-17)
- 초기 버전 릴리스
- 업체별 주문서 PDF 생성
- 팔레트 라벨 PDF 생성 (큰 글씨 A4)
- 재고표 PDF 생성
- 마이너스 수량 자동 제거
- 업체명 자동 변환

## 지원

문제가 발생하거나 개선 사항이 있으면 담당자에게 문의하세요.

## 라이선스

내부 사용 전용

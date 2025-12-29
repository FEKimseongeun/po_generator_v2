# MOM to PO Generator v2.0

MOM(Minutes of Meeting) 문서에서 **섹션별 데이터를 자동 추출**하고,  
**플레이스홀더 템플릿**을 사용하여 PO(Purchase Order) 문서를 자동 생성합니다.

---

## 🆕 v2.0 변경사항

- **플레이스홀더 방식**: `{{FIELD_NAME}}` 형태의 플레이스홀더 사용
- **섹션 기반 파싱**: MOM의 번호 체계(1, 2, 2.1, 2.2...)를 인식하여 자동 추출
- **하이라이트 불필요**: MOM 파일에 색상 하이라이트 없어도 작동
- **36개 필드 자동 추출**: Payment, Warranty, LD, Bond 등 모든 주요 항목

---

## 📦 설치 및 실행

### 요구사항
- Python 3.10 이상
- Windows 10/11

### 실행

```bash
# GUI 모드 (권장)
python main.py

# CLI 모드
python main.py --cli mom.docx template.docx output.docx

# MOM 분석만
python main.py --analyze mom.docx
```

---

## 📖 사용 방법

### 1. PO 템플릿 준비

템플릿 파일(`.docx`)에 값이 들어갈 위치에 플레이스홀더를 입력합니다:

```
PURCHASE ORDER

PO No.: {{PO_NO}}
Date: {{MOM_DATE}}

1. Basis of Agreement
   This Purchase Order (PO No. {{PO_NO}}) 
   Material Requisition ({{MR_NO}}) {{MR_DATE}}
   Vendor's Quotation ({{PI_NO}}) dated {{PI_DATE}}

2. Delivery Terms
   {{DELIVERY_TERMS}}

3. Payment
   {{PAYMENT_FULL}}

4. Warranty
   {{WARRANTY}}
...
```

### 2. MOM 파일 업로드

일반 MOM 문서를 그대로 사용합니다. (하이라이트 불필요!)

### 3. PO 생성

GUI에서 "PO 생성" 버튼 클릭 또는 CLI 실행

---

## 📋 플레이스홀더 목록 (36개)

### 헤더 정보
| 플레이스홀더 | 설명 | 예시 |
|-------------|------|------|
| `{{MOM_NO}}` | MOM 문서 번호 | 210025-28-126-001 |
| `{{PO_NO}}` | PO 번호 (자동생성) | 210025-28-126-001-A01 |
| `{{MOM_DATE}}` | MOM 미팅 날짜 | December 5th, 2024 |
| `{{SUBJECT}}` | MOM 제목 | Commercial Clarification... |

### 섹션 1: Inquiry/MR 정보
| 플레이스홀더 | 설명 |
|-------------|------|
| `{{MR_NO}}` | Material Requisition 번호 |
| `{{MR_DATE}}` | MR 날짜 |
| `{{PI_NO}}` | Vendor PI 번호 |
| `{{PI_DATE}}` | Vendor PI 날짜 |
| `{{ITEM_DESC}}` | 품목 설명 |

### 섹션 2: Payment
| 플레이스홀더 | 설명 |
|-------------|------|
| `{{PAYMENT}}` | Payment 서두 |
| `{{PAYMENT_FULL}}` | Payment 전체 (서브섹션 포함) |
| `{{ADVANCE_PAYMENT}}` | 선급금 조건 |
| `{{PROGRESS_PAYMENT_1ST}}` | 1차 기성금 |
| `{{PROGRESS_PAYMENT_2ND}}` | 2차 기성금 |
| `{{DELIVERY_PAYMENT}}` | 납품 결제 |
| `{{FINAL_PAYMENT}}` | 최종 결제 |

### 섹션 3: Warranty
| 플레이스홀더 | 설명 |
|-------------|------|
| `{{WARRANTY}}` | 보증 조건 전체 |

### 섹션 4: Liquidated Damages
| 플레이스홀더 | 설명 |
|-------------|------|
| `{{LIQUIDATED_DAMAGES}}` | LD 조건 전체 |
| `{{LD_DELIVERY}}` | 납기 지연 LD |
| `{{LD_ENGINEERING}}` | 설계문서 지연 LD |
| `{{LD_MAX}}` | 최대 LD 한도 |

### 섹션 5: Bond Requirements
| 플레이스홀더 | 설명 |
|-------------|------|
| `{{BOND_REQUIREMENTS}}` | Bond 조건 전체 |
| `{{BOND_APPLICATION}}` | Bond 적용 기준 |
| `{{ADVANCE_PROGRESS_BOND}}` | 선급/기성 보증 |
| `{{PERFORMANCE_BOND}}` | 이행 보증 |
| `{{WARRANTY_BOND}}` | 하자 보증 |
| `{{BOND_ISSUE}}` | Bond 발행 조건 |

### 섹션 6-10: 기타
| 플레이스홀더 | 설명 |
|-------------|------|
| `{{OPTIONAL}}` | 선택 조건 |
| `{{TRAINING_SUPERVISION}}` | 교육/감독 조건 |
| `{{DELIVERY_TERMS}}` | 배송 조건 전체 |
| `{{INCOTERMS}}` | INCOTERMS만 |
| `{{PRICE_SCOPE}}` | 가격 및 범위 |
| `{{SPECIAL_NOTE}}` | 특별 주의사항 |

### 섹션 13: Attachments
| 플레이스홀더 | 설명 |
|-------------|------|
| `{{ATTACHMENTS}}` | 첨부 문서 전체 |
| `{{ATTACHMENTS_GENERAL}}` | 일반 첨부 |
| `{{ATTACHMENTS_TECHNICAL}}` | 기술 문서 |

---

## 📁 프로젝트 구조

```
po_generator_v2/
├── main.py                    # 메인 실행 파일
├── README.md                  # 이 문서
├── config/
│   └── field_mapping.json     # 필드 매핑 설정
├── core/
│   ├── mom_parser.py          # MOM 파서 (섹션 기반)
│   └── po_generator.py        # PO 생성기 (플레이스홀더)
├── gui/
│   └── main_window.py         # GUI 인터페이스
├── templates/
│   └── PO_Template_Sample.docx # 샘플 템플릿
├── utils/
│   └── template_converter.py  # 템플릿 변환 유틸리티
└── output/                    # 생성된 PO 저장
```

---

## 🔧 개발자 가이드

### 새 필드 추가하기

1. `core/mom_parser.py`의 `_extract_fields()` 메서드에 추출 로직 추가
2. 템플릿에 새 플레이스홀더 사용

```python
# mom_parser.py 예시
if '14' in mom_data.sections:
    fields['NEW_SECTION'] = mom_data.get_section_content('14')
```

### MOM 구조가 변경된 경우

MOM의 섹션 번호 체계가 변경되면 `mom_parser.py`의 섹션 매핑을 수정합니다.

---

## ⚠️ 주의사항

1. **MOM 양식**: 기존 MOM 양식과 동일한 테이블 구조 필요
2. **섹션 번호**: NO. 열의 번호(1, 2, 2.1 등)를 기준으로 파싱
3. **파일 형식**: `.docx`만 지원
4. **인코딩**: UTF-8 권장

---

## 📝 버전 히스토리

### v2.0.0 (2024-12)
- 플레이스홀더 기반 시스템으로 전환
- 섹션 번호 기반 자동 파싱
- 36개 필드 자동 추출
- MOM 하이라이트 불필요

### v1.0.0 (2024-12)
- 초기 버전 (색상 하이라이트 기반)

---

© 2024 Procurement Team

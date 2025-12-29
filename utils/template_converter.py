"""
템플릿 변환 유틸리티
기존 하이라이트 기반 format 파일을 플레이스홀더 템플릿으로 변환합니다.
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import tempfile
import shutil
import os
import re
from typing import Dict, List, Tuple


# 색상 → 플레이스홀더 매핑
COLOR_TO_PLACEHOLDER = {
    'yellow': 'MOM_NO',  # 첫 번째 발견만
    'cyan': 'MOM_DATE',
    'green': 'MR_DATE', 
    'magenta': 'PI_INFO',
    'darkCyan': 'DELIVERY_TERMS',
    'red': 'CONTRACT_TERMS',  # 섹션별로 세분화 필요
    'darkYellow': 'SPECIAL_NOTE',
}

# red 색상 섹션별 세분화 (위치 기반)
RED_SECTION_MAPPING = {
    'Payment': 'PAYMENT_FULL',
    'Advance Payment': 'ADVANCE_PAYMENT',
    '1st Progress': 'PROGRESS_PAYMENT_1ST',
    '2nd Progress': 'PROGRESS_PAYMENT_2ND',
    'Delivery Payment': 'DELIVERY_PAYMENT',
    'Final Payment': 'FINAL_PAYMENT',
    'Warranty': 'WARRANTY',
    'Liquidated Damages': 'LIQUIDATED_DAMAGES',
    'Bond': 'BOND_REQUIREMENTS',
    'Optional': 'OPTIONAL',
    'Training': 'TRAINING_SUPERVISION',
    'Special Note': 'SPECIAL_NOTE',
    'Attachment': 'ATTACHMENTS',
}


def convert_to_template(input_path: str, output_path: str) -> Dict[str, List[str]]:
    """
    하이라이트된 format 파일을 플레이스홀더 템플릿으로 변환
    
    Returns:
        변환된 플레이스홀더와 원본 텍스트 매핑
    """
    NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    # 압축 해제
    temp_dir = Path(tempfile.mkdtemp())
    with zipfile.ZipFile(input_path, 'r') as zf:
        zf.extractall(temp_dir)
    
    document_xml = temp_dir / 'word' / 'document.xml'
    
    # XML 읽기
    with open(document_xml, 'r', encoding='utf-8') as f:
        content = f.read()
    
    tree = ET.parse(document_xml)
    root = tree.getroot()
    
    conversions = {}
    placeholder_counts = {}  # 같은 플레이스홀더 사용 횟수
    
    # 색상별 하이라이트 위치 추적
    current_section = None
    
    for para in root.iter(f"{{{NS['w']}}}p"):
        para_text_full = ''.join(
            t.text for t in para.iter(f"{{{NS['w']}}}t") if t.text
        )
        
        # 섹션 감지 (빨간색 처리용)
        for section_key in RED_SECTION_MAPPING.keys():
            if section_key.lower() in para_text_full.lower()[:50]:
                current_section = section_key
                break
        
        for run in para.iter(f"{{{NS['w']}}}r"):
            rPr = run.find(f"{{{NS['w']}}}rPr")
            if rPr is None:
                continue
            
            highlight = rPr.find(f"{{{NS['w']}}}highlight")
            if highlight is None:
                continue
            
            color = highlight.get(f"{{{NS['w']}}}val")
            text_elem = run.find(f"{{{NS['w']}}}t")
            
            if text_elem is None or not text_elem.text:
                continue
            
            original_text = text_elem.text
            
            # 플레이스홀더 결정
            if color == 'red' and current_section:
                placeholder = RED_SECTION_MAPPING.get(current_section, 'CONTRACT_TERMS')
            else:
                placeholder = COLOR_TO_PLACEHOLDER.get(color, f'FIELD_{color.upper()}')
            
            # 플레이스홀더로 변환
            placeholder_full = f"{{{{{placeholder}}}}}"
            
            # 기록
            if placeholder not in conversions:
                conversions[placeholder] = []
            conversions[placeholder].append(original_text)
            
            # 텍스트 교체 (첫 run만 플레이스홀더, 나머지는 빈 문자열)
            if placeholder not in placeholder_counts:
                placeholder_counts[placeholder] = 0
                text_elem.text = placeholder_full
            else:
                text_elem.text = ""  # 연속된 같은 색상은 비움
            
            placeholder_counts[placeholder] += 1
            
            # 하이라이트 제거
            rPr.remove(highlight)
    
    # 저장
    tree.write(document_xml, encoding='utf-8', xml_declaration=True)
    
    # 새 docx 생성
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root_dir, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = Path(root_dir) / file
                arcname = file_path.relative_to(temp_dir)
                zf.write(file_path, arcname)
    
    # 정리
    shutil.rmtree(temp_dir, ignore_errors=True)
    
    return conversions


def print_field_guide():
    """플레이스홀더 필드 가이드 출력"""
    print("""
╔══════════════════════════════════════════════════════════════════════════╗
║                    PO 템플릿 플레이스홀더 가이드                          ║
╠══════════════════════════════════════════════════════════════════════════╣
║                                                                          ║
║  【헤더 정보】                                                           ║
║    {{MOM_NO}}         - MOM/PO 번호 (예: 210025-28-126-001)             ║
║    {{PO_NO}}          - PO 번호 (MOM_NO + -A01)                         ║
║    {{MOM_DATE}}       - MOM 미팅 날짜                                    ║
║    {{SUBJECT}}        - MOM 제목/주제                                    ║
║                                                                          ║
║  【섹션 1: Inquiry/MR 정보】                                             ║
║    {{MR_NO}}          - Material Requisition 번호                        ║
║    {{MR_DATE}}        - MR 날짜                                          ║
║    {{PI_NO}}          - Vendor PI 번호                                   ║
║    {{PI_DATE}}        - Vendor PI 날짜                                   ║
║    {{ITEM_DESC}}      - 품목 설명                                        ║
║                                                                          ║
║  【섹션 2: Payment】                                                     ║
║    {{PAYMENT}}              - Payment 조건 전체                          ║
║    {{PAYMENT_FULL}}         - Payment 섹션 전체 (서브섹션 포함)          ║
║    {{ADVANCE_PAYMENT}}      - 선급금 조건                                ║
║    {{PROGRESS_PAYMENT_1ST}} - 1차 기성금 조건                            ║
║    {{PROGRESS_PAYMENT_2ND}} - 2차 기성금 조건                            ║
║    {{DELIVERY_PAYMENT}}     - 납품 결제 조건                             ║
║    {{FINAL_PAYMENT}}        - 최종 결제 조건                             ║
║                                                                          ║
║  【섹션 3: Warranty】                                                    ║
║    {{WARRANTY}}       - 보증 조건 전체                                   ║
║                                                                          ║
║  【섹션 4: Liquidated Damages】                                          ║
║    {{LIQUIDATED_DAMAGES}} - LD 조건 전체                                 ║
║    {{LD_DELIVERY}}        - 납기 지연 LD                                 ║
║    {{LD_ENGINEERING}}     - 설계문서 지연 LD                             ║
║    {{LD_MAX}}             - 최대 LD 한도                                 ║
║                                                                          ║
║  【섹션 5: Bond Requirements】                                           ║
║    {{BOND_REQUIREMENTS}}     - Bond 조건 전체                            ║
║    {{BOND_APPLICATION}}      - Bond 적용 기준                            ║
║    {{ADVANCE_PROGRESS_BOND}} - 선급/기성 보증                            ║
║    {{PERFORMANCE_BOND}}      - 이행 보증                                 ║
║    {{WARRANTY_BOND}}         - 하자 보증                                 ║
║                                                                          ║
║  【섹션 6-7: Optional/Training】                                         ║
║    {{OPTIONAL}}           - 선택 조건                                    ║
║    {{TRAINING_SUPERVISION}} - 교육/감독 조건                             ║
║                                                                          ║
║  【섹션 8: Delivery】                                                    ║
║    {{DELIVERY_TERMS}} - 배송 조건 (INCOTERMS 포함)                       ║
║    {{INCOTERMS}}      - INCOTERMS만                                      ║
║                                                                          ║
║  【섹션 9-10: Price/Special Note】                                       ║
║    {{PRICE_SCOPE}}    - 가격 및 범위                                     ║
║    {{SPECIAL_NOTE}}   - 특별 주의사항                                    ║
║                                                                          ║
║  【섹션 13: Attachments】                                                ║
║    {{ATTACHMENTS}}          - 첨부 문서 전체                             ║
║    {{ATTACHMENTS_GENERAL}}  - 일반 첨부                                  ║
║    {{ATTACHMENTS_TECHNICAL}} - 기술 문서 첨부                            ║
║                                                                          ║
╚══════════════════════════════════════════════════════════════════════════╝
""")


if __name__ == "__main__":
    import sys
    
    print_field_guide()
    
    if len(sys.argv) >= 3:
        input_file = sys.argv[1]
        output_file = sys.argv[2]
        
        print(f"\n변환 중: {input_file} -> {output_file}")
        conversions = convert_to_template(input_file, output_file)
        
        print(f"\n변환 완료! 플레이스홀더 {len(conversions)}개 생성됨:")
        for ph, texts in conversions.items():
            preview = texts[0][:40] + "..." if len(texts[0]) > 40 else texts[0]
            print(f"  {{{{{ph}}}}} ({len(texts)}회): {preview}")
    else:
        print("\n사용법: python template_converter.py <input.docx> <output_template.docx>")

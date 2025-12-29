"""
MOM 파서 모듈 v2.0
섹션 번호 기반으로 MOM 문서에서 데이터를 추출합니다.
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, field
import tempfile
import shutil
import re
import json


@dataclass
class MOMSection:
    """MOM 섹션 데이터"""
    number: str
    title: str
    content: str
    subsections: Dict[str, 'MOMSection'] = field(default_factory=dict)


@dataclass  
class MOMData:
    """MOM 문서에서 추출된 전체 데이터"""
    # 헤더 정보
    mom_no: str = ""
    mom_date: str = ""
    subject: str = ""
    
    # 섹션 데이터
    sections: Dict[str, MOMSection] = field(default_factory=dict)
    
    # 파싱된 필드 (플레이스홀더용)
    fields: Dict[str, str] = field(default_factory=dict)
    
    # 원본 파일 경로
    file_path: str = ""
    
    def get_field(self, field_name: str) -> str:
        """필드 값 조회"""
        return self.fields.get(field_name, "")
    
    def get_section_content(self, section_num: str) -> str:
        """섹션 전체 내용 조회 (서브섹션 포함)"""
        if section_num not in self.sections:
            return ""
        
        section = self.sections[section_num]
        content_parts = [section.content]
        
        for sub_num in sorted(section.subsections.keys()):
            sub = section.subsections[sub_num]
            content_parts.append(f"\n{sub.title}\n{sub.content}")
        
        return '\n'.join(content_parts)


class MOMParser:
    """MOM 문서 파서 (섹션 기반)"""
    
    NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    def __init__(self, file_path: str, config_path: Optional[str] = None):
        self.file_path = Path(file_path)
        self.temp_dir: Optional[Path] = None
        self.config = self._load_config(config_path)
        self._validate_file()
    
    def _load_config(self, config_path: Optional[str]) -> Dict:
        """설정 파일 로드"""
        if config_path:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {}
    
    def _validate_file(self):
        """파일 유효성 검사"""
        if not self.file_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {self.file_path}")
        if self.file_path.suffix.lower() not in ['.docx']:
            raise ValueError(f"지원하지 않는 파일 형식입니다. .docx 파일만 지원됩니다.")
    
    def _extract_docx(self) -> Path:
        """docx 압축 해제"""
        self.temp_dir = Path(tempfile.mkdtemp())
        with zipfile.ZipFile(self.file_path, 'r') as zf:
            zf.extractall(self.temp_dir)
        return self.temp_dir
    
    def _get_text(self, elem: ET.Element) -> str:
        """요소에서 텍스트 추출"""
        texts = []
        for t in elem.iter(f"{{{self.NS['w']}}}t"):
            if t.text:
                texts.append(t.text)
        return ''.join(texts)
    
    def _get_cell_texts(self, row: ET.Element) -> List[str]:
        """행에서 각 셀의 텍스트 추출"""
        cells = list(row.iter(f"{{{self.NS['w']}}}tc"))
        return [self._get_text(cell).strip() for cell in cells]
    
    def parse(self) -> MOMData:
        """MOM 파일 파싱"""
        try:
            extract_path = self._extract_docx()
            document_xml = extract_path / 'word' / 'document.xml'
            
            if not document_xml.exists():
                raise ValueError("유효한 Word 문서가 아닙니다.")
            
            tree = ET.parse(document_xml)
            root = tree.getroot()
            
            mom_data = MOMData(file_path=str(self.file_path))
            
            # 테이블 행 파싱
            rows = list(root.iter(f"{{{self.NS['w']}}}tr"))
            self._parse_rows(rows, mom_data)
            
            # 필드 추출
            self._extract_fields(mom_data)
            
            return mom_data
            
        finally:
            self._cleanup()
    
    def _parse_rows(self, rows: List[ET.Element], mom_data: MOMData):
        """테이블 행들 파싱"""
        current_section: Optional[str] = None
        current_subsection: Optional[str] = None
        
        for row in rows:
            cell_texts = self._get_cell_texts(row)
            if not cell_texts:
                continue
            
            first_cell = cell_texts[0]
            second_cell = cell_texts[1] if len(cell_texts) > 1 else ""
            
            # 헤더 파싱 (MOM NO, DATE, SUBJECT)
            if first_cell == 'MOM NO':
                mom_data.mom_no = second_cell.replace('MOM-', '').strip()
                # DATE 찾기
                for i, ct in enumerate(cell_texts):
                    if ct == 'DATE' and i + 1 < len(cell_texts):
                        mom_data.mom_date = cell_texts[i + 1].strip()
                        break
            
            elif first_cell == 'SUBJECT':
                mom_data.subject = second_cell.strip()
            
            # 섹션 번호 파싱
            elif re.match(r'^\d+$', first_cell):  # 메인 섹션 (1, 2, 3...)
                current_section = first_cell
                current_subsection = None
                
                mom_data.sections[current_section] = MOMSection(
                    number=current_section,
                    title=self._extract_title(second_cell),
                    content=second_cell
                )
            
            elif re.match(r'^\d+\.\d+$', first_cell):  # 서브섹션 (1.1, 2.1...)
                current_subsection = first_cell
                main_section = first_cell.split('.')[0]
                
                if main_section in mom_data.sections:
                    mom_data.sections[main_section].subsections[current_subsection] = MOMSection(
                        number=current_subsection,
                        title=self._extract_title(second_cell),
                        content=second_cell
                    )
            
            # 내용 행 (번호 없음)
            elif first_cell == '' and second_cell:
                if current_subsection and current_section in mom_data.sections:
                    sub = mom_data.sections[current_section].subsections.get(current_subsection)
                    if sub:
                        sub.content += ' ' + second_cell
                elif current_section and current_section in mom_data.sections:
                    mom_data.sections[current_section].content += ' ' + second_cell
    
    def _extract_title(self, text: str) -> str:
        """섹션 제목 추출"""
        # 첫 번째 문장이나 콜론까지
        match = re.match(r'^([^.:\n]+)', text)
        return match.group(1).strip() if match else text[:50]
    
    def _extract_fields(self, mom_data: MOMData):
        """플레이스홀더용 필드 추출"""
        fields = mom_data.fields
        
        # 헤더 필드
        fields['MOM_NO'] = mom_data.mom_no
        fields['MOM_DATE'] = mom_data.mom_date
        fields['SUBJECT'] = mom_data.subject
        fields['PO_NO'] = mom_data.mom_no + '-A01'  # PO 번호 생성
        
        # 섹션 1 - Inquiry/MR 정보
        if '1' in mom_data.sections:
            content = mom_data.sections['1'].content
            
            # MR 번호
            mr_match = re.search(r'MR-[\d-]+', content)
            if mr_match:
                fields['MR_NO'] = mr_match.group()
            
            # MR 날짜
            mr_date_match = re.search(r'MR-[\d-]+[^/]*dated\s+([A-Za-z]+\s+\d+[a-z]*,?\s+\d{4})', content)
            if mr_date_match:
                fields['MR_DATE'] = mr_date_match.group(1)
            
            # PI 정보 (있는 경우)
            pi_match = re.search(r'([A-Z]+-[A-Z]-[A-Z0-9]+)\)?.*?dated\s+([A-Za-z]+\s+\d+[a-z]*,?\s+\d{4})', content)
            if pi_match:
                fields['PI_NO'] = pi_match.group(1)
                fields['PI_DATE'] = pi_match.group(2)
            
            # 품목 설명
            item_match = re.search(r'/\s*([^/]+)$', content)
            if item_match:
                fields['ITEM_DESC'] = item_match.group(1).strip()
        
        # 섹션 2 - Payment
        if '2' in mom_data.sections:
            sec = mom_data.sections['2']
            fields['PAYMENT'] = sec.content
            fields['PAYMENT_FULL'] = mom_data.get_section_content('2')
            
            for sub_num, sub in sec.subsections.items():
                if sub_num == '2.1':
                    fields['ADVANCE_PAYMENT'] = sub.content
                elif sub_num == '2.2':
                    fields['PROGRESS_PAYMENT_1ST'] = sub.content
                elif sub_num == '2.3':
                    fields['PROGRESS_PAYMENT_2ND'] = sub.content
                elif sub_num == '2.4':
                    fields['DELIVERY_PAYMENT'] = sub.content
                elif sub_num == '2.5':
                    fields['FINAL_PAYMENT'] = sub.content
        
        # 섹션 3 - Warranty
        if '3' in mom_data.sections:
            fields['WARRANTY'] = mom_data.get_section_content('3')
        
        # 섹션 4 - Liquidated Damages
        if '4' in mom_data.sections:
            sec = mom_data.sections['4']
            fields['LIQUIDATED_DAMAGES'] = mom_data.get_section_content('4')
            
            for sub_num, sub in sec.subsections.items():
                if sub_num == '4.1':
                    fields['LD_DELIVERY'] = sub.content
                elif sub_num == '4.2':
                    fields['LD_ENGINEERING'] = sub.content
                elif sub_num == '4.3':
                    fields['LD_MAX'] = sub.content
        
        # 섹션 5 - Bond
        if '5' in mom_data.sections:
            sec = mom_data.sections['5']
            fields['BOND_REQUIREMENTS'] = mom_data.get_section_content('5')
            
            for sub_num, sub in sec.subsections.items():
                if sub_num == '5.1':
                    fields['BOND_APPLICATION'] = sub.content
                elif sub_num == '5.2':
                    fields['ADVANCE_PROGRESS_BOND'] = sub.content
                elif sub_num == '5.3':
                    fields['PERFORMANCE_BOND'] = sub.content
                elif sub_num == '5.4':
                    fields['WARRANTY_BOND'] = sub.content
                elif sub_num == '5.5':
                    fields['BOND_ISSUE'] = sub.content
        
        # 섹션 6 - Optional
        if '6' in mom_data.sections:
            fields['OPTIONAL'] = mom_data.get_section_content('6')
        
        # 섹션 7 - Training/Supervision
        if '7' in mom_data.sections:
            fields['TRAINING_SUPERVISION'] = mom_data.get_section_content('7')
        
        # 섹션 8 - Delivery Terms
        if '8' in mom_data.sections:
            content = mom_data.sections['8'].content
            fields['DELIVERY_TERMS'] = content
            
            # INCOTERMS 추출
            inco_match = re.search(r'(FCA|FOB|CIF|CFR|EXW|DAP|DDP)[^,]*,[^,]+', content)
            if inco_match:
                fields['INCOTERMS'] = inco_match.group()
        
        # 섹션 9 - Price/Scope
        if '9' in mom_data.sections:
            fields['PRICE_SCOPE'] = mom_data.get_section_content('9')
        
        # 섹션 10 - Special Note
        if '10' in mom_data.sections:
            fields['SPECIAL_NOTE'] = mom_data.get_section_content('10')
        
        # 섹션 13 - Attachments
        if '13' in mom_data.sections:
            sec = mom_data.sections['13']
            fields['ATTACHMENTS'] = mom_data.get_section_content('13')
            
            for sub_num, sub in sec.subsections.items():
                if sub_num == '13.1':
                    fields['ATTACHMENTS_GENERAL'] = sub.content
                elif sub_num == '13.2':
                    fields['ATTACHMENTS_TECHNICAL'] = sub.content
    
    def _cleanup(self):
        """임시 디렉토리 정리"""
        if self.temp_dir and self.temp_dir.exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)
            self.temp_dir = None


def parse_mom(file_path: str, config_path: Optional[str] = None) -> MOMData:
    """MOM 파일 파싱 편의 함수"""
    parser = MOMParser(file_path, config_path)
    return parser.parse()


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        result = parse_mom(sys.argv[1])
        print(f"MOM NO: {result.mom_no}")
        print(f"DATE: {result.mom_date}")
        print(f"SUBJECT: {result.subject}")
        print(f"\n추출된 필드 ({len(result.fields)}개):")
        for k, v in sorted(result.fields.items()):
            preview = v[:80] + "..." if len(v) > 80 else v
            print(f"  {k}: {preview}")

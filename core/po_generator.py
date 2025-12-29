"""
PO 생성기 모듈 v2.0
플레이스홀더({{FIELD_NAME}})를 MOM 데이터로 대체하여 PO 문서를 생성합니다.
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
import tempfile
import shutil
import os
import re

from .mom_parser import MOMData


@dataclass
class ReplacementResult:
    """교체 결과"""
    placeholder: str
    original: str
    replaced_with: str
    count: int


class POGenerator:
    """PO 문서 생성기 (플레이스홀더 기반)"""
    
    NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    # 네임스페이스 등록
    for prefix, uri in NS.items():
        ET.register_namespace(prefix, uri)
    
    ADDITIONAL_NS = {
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    }
    for prefix, uri in ADDITIONAL_NS.items():
        ET.register_namespace(prefix, uri)
    
    def __init__(self, template_path: str, placeholder_prefix: str = "{{", placeholder_suffix: str = "}}"):
        self.template_path = Path(template_path)
        self.prefix = placeholder_prefix
        self.suffix = placeholder_suffix
        self.temp_dir: Optional[Path] = None
        self.replacements: List[ReplacementResult] = []
        self._validate_file()
    
    def _validate_file(self):
        if not self.template_path.exists():
            raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {self.template_path}")
    
    def _extract_docx(self) -> Path:
        self.temp_dir = Path(tempfile.mkdtemp())
        with zipfile.ZipFile(self.template_path, 'r') as zf:
            zf.extractall(self.temp_dir)
        return self.temp_dir
    
    def generate(self, mom_data: MOMData, output_path: str) -> Tuple[str, List[ReplacementResult]]:
        """
        PO 문서 생성
        
        Args:
            mom_data: MOM에서 추출된 데이터
            output_path: 출력 파일 경로
            
        Returns:
            (출력 경로, 교체 결과 목록)
        """
        try:
            extract_path = self._extract_docx()
            document_xml = extract_path / 'word' / 'document.xml'
            
            if not document_xml.exists():
                raise ValueError("유효한 Word 템플릿이 아닙니다.")
            
            # 플레이스홀더 교체
            self._replace_placeholders(document_xml, mom_data.fields)
            
            # 새 docx 생성
            output = self._create_docx(extract_path, output_path)
            
            return output, self.replacements
            
        finally:
            self._cleanup()
    
    def _replace_placeholders(self, document_xml: Path, fields: Dict[str, str]):
        """플레이스홀더를 실제 값으로 교체"""
        
        # XML 파일 읽기
        with open(document_xml, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 플레이스홀더 패턴
        pattern = re.escape(self.prefix) + r'([A-Z_0-9]+)' + re.escape(self.suffix)
        
        # 발견된 플레이스홀더 목록
        found_placeholders = set(re.findall(pattern, content))
        
        # 각 플레이스홀더 교체
        for placeholder_name in found_placeholders:
            full_placeholder = f"{self.prefix}{placeholder_name}{self.suffix}"
            replacement_value = fields.get(placeholder_name, "")
            
            if replacement_value:
                # XML 특수문자 이스케이프
                replacement_value = self._escape_xml(replacement_value)
                
                count = content.count(full_placeholder)
                content = content.replace(full_placeholder, replacement_value)
                
                self.replacements.append(ReplacementResult(
                    placeholder=placeholder_name,
                    original=full_placeholder,
                    replaced_with=replacement_value[:50] + "..." if len(replacement_value) > 50 else replacement_value,
                    count=count
                ))
        
        # 파일 저장
        with open(document_xml, 'w', encoding='utf-8') as f:
            f.write(content)
    
    def _escape_xml(self, text: str) -> str:
        """XML 특수문자 이스케이프"""
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')
        text = text.replace('"', '&quot;')
        text = text.replace("'", '&apos;')
        return text
    
    def _create_docx(self, extract_path: Path, output_path: str) -> str:
        """docx 파일 생성"""
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(extract_path):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(extract_path)
                    zf.write(file_path, arcname)
        
        return str(output_path)
    
    def _cleanup(self):
        if self.temp_dir and self.temp_dir.exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)
            self.temp_dir = None
    
    def get_template_placeholders(self) -> List[str]:
        """템플릿에서 사용된 플레이스홀더 목록 조회"""
        try:
            extract_path = self._extract_docx()
            document_xml = extract_path / 'word' / 'document.xml'
            
            with open(document_xml, 'r', encoding='utf-8') as f:
                content = f.read()
            
            pattern = re.escape(self.prefix) + r'([A-Z_0-9]+)' + re.escape(self.suffix)
            placeholders = set(re.findall(pattern, content))
            
            return sorted(placeholders)
            
        finally:
            self._cleanup()


def generate_po(template_path: str, mom_data: MOMData, output_path: str) -> Tuple[str, List[ReplacementResult]]:
    """PO 생성 편의 함수"""
    generator = POGenerator(template_path)
    return generator.generate(mom_data, output_path)


if __name__ == "__main__":
    import sys
    from mom_parser import parse_mom
    
    if len(sys.argv) >= 3:
        mom_path = sys.argv[1]
        template_path = sys.argv[2]
        output_path = sys.argv[3] if len(sys.argv) > 3 else "output_PO.docx"
        
        print(f"MOM 파싱: {mom_path}")
        mom_data = parse_mom(mom_path)
        
        print(f"PO 생성: {output_path}")
        result_path, replacements = generate_po(template_path, mom_data, output_path)
        
        print(f"\n완료: {result_path}")
        print(f"교체된 플레이스홀더:")
        for r in replacements:
            print(f"  {{{{{r.placeholder}}}}} -> {r.replaced_with} ({r.count}회)")

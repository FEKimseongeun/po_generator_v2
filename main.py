#!/usr/bin/env python3
"""
MOM to PO Generator v2.0
========================

MOM 문서에서 섹션별 데이터를 추출하고,
플레이스홀더 템플릿을 사용하여 PO 문서를 자동 생성합니다.

사용법:
    GUI 모드:  python main.py
    CLI 모드:  python main.py --cli <mom_file> <template_file> [output_file]
    분석 모드: python main.py --analyze <mom_file>
"""

import sys
import argparse
from pathlib import Path
from datetime import datetime


def run_gui():
    """GUI 모드"""
    try:
        from gui.main_window import main
        main()
    except ImportError as e:
        print(f"GUI 모듈 로드 실패: {e}")
        print("tkinter가 설치되어 있는지 확인하세요.")
        sys.exit(1)


def run_cli(args):
    """CLI 모드"""
    from core.mom_parser import parse_mom
    from core.po_generator import generate_po
    
    mom_path = args.mom_file
    template_path = args.template_file
    
    # 출력 경로
    if args.output:
        output_path = args.output
    else:
        mom_file = Path(mom_path)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = str(mom_file.parent / f"PO_{mom_file.stem}_{timestamp}.docx")
    
    print("=" * 60)
    print("MOM to PO Generator v2.0")
    print("=" * 60)
    print(f"\nMOM 파일:    {mom_path}")
    print(f"템플릿:      {template_path}")
    print(f"출력 파일:   {output_path}")
    
    # MOM 파싱
    print("\n[1/2] MOM 파일 분석 중...")
    try:
        mom_data = parse_mom(mom_path)
        print(f"  ✓ MOM NO: {mom_data.mom_no}")
        print(f"  ✓ DATE: {mom_data.mom_date}")
        print(f"  ✓ 추출된 필드: {len(mom_data.fields)}개")
    except Exception as e:
        print(f"  ✗ 오류: {e}")
        sys.exit(1)
    
    # PO 생성
    print("\n[2/2] PO 문서 생성 중...")
    try:
        result_path, replacements = generate_po(template_path, mom_data, output_path)
        print(f"  ✓ 교체된 플레이스홀더: {len(replacements)}개")
        for r in replacements:
            print(f"    - {{{{{r.placeholder}}}}}")
    except Exception as e:
        print(f"  ✗ 오류: {e}")
        sys.exit(1)
    
    print("\n" + "=" * 60)
    print(f"✓ PO 생성 완료: {result_path}")
    print("=" * 60)


def run_analyze(args):
    """MOM 분석 모드"""
    from core.mom_parser import parse_mom
    
    print("=" * 60)
    print("MOM 문서 분석")
    print("=" * 60)
    
    try:
        mom_data = parse_mom(args.mom_file)
        
        print(f"\n📋 헤더 정보:")
        print(f"  MOM NO:  {mom_data.mom_no}")
        print(f"  DATE:    {mom_data.mom_date}")
        print(f"  SUBJECT: {mom_data.subject[:50]}...")
        
        print(f"\n📁 섹션 구조:")
        for num in sorted(mom_data.sections.keys(), key=lambda x: float(x) if '.' not in x else float(x.replace('.', ''))/10):
            sec = mom_data.sections[num]
            print(f"  [{num}] {sec.title}")
            for sub_num in sorted(sec.subsections.keys()):
                sub = sec.subsections[sub_num]
                print(f"    [{sub_num}] {sub.title}")
        
        print(f"\n📝 추출된 필드 ({len(mom_data.fields)}개):")
        print("-" * 60)
        for field, value in sorted(mom_data.fields.items()):
            preview = value[:60].replace('\n', ' ')
            if len(value) > 60:
                preview += "..."
            print(f"  {{{{{field:25s}}}}} = {preview}")
        
    except Exception as e:
        print(f"오류: {e}")
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description='MOM to PO Generator v2.0',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예제:
  GUI 실행:     python main.py
  CLI 실행:     python main.py --cli mom.docx template.docx output.docx
  MOM 분석:     python main.py --analyze mom.docx

플레이스홀더 예시:
  {{MOM_NO}}, {{MOM_DATE}}, {{PAYMENT_FULL}}, {{WARRANTY}}, {{DELIVERY_TERMS}} 등
        """
    )
    
    parser.add_argument('--cli', action='store_true', help='CLI 모드 실행')
    parser.add_argument('--analyze', action='store_true', help='MOM 분석 모드')
    parser.add_argument('mom_file', nargs='?', help='MOM 파일')
    parser.add_argument('template_file', nargs='?', help='PO 템플릿 파일')
    parser.add_argument('output', nargs='?', help='출력 파일')
    parser.add_argument('--version', action='version', version='MOM to PO Generator v2.0')
    
    args = parser.parse_args()
    
    if args.analyze:
        if not args.mom_file:
            parser.error("--analyze 모드에서는 MOM 파일이 필요합니다.")
        run_analyze(args)
    elif args.cli:
        if not args.mom_file or not args.template_file:
            parser.error("--cli 모드에서는 MOM 파일과 템플릿 파일이 필요합니다.")
        run_cli(args)
    else:
        run_gui()


if __name__ == "__main__":
    main()

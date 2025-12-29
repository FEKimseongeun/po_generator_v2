"""
JSON Validator and Fixer
Validates and fixes common JSON issues from LLM output
"""

import json
import re
import os


def clean_llm_json(json_string):
    """
    Clean JSON string from LLM output
    Removes markdown code blocks and extra text
    """
    
    # Remove markdown code blocks
    json_string = re.sub(r'^```json\s*', '', json_string, flags=re.MULTILINE)
    json_string = re.sub(r'^```\s*', '', json_string, flags=re.MULTILINE)
    
    # Find the actual JSON object (starts with { and ends with })
    start_idx = json_string.find('{')
    end_idx = json_string.rfind('}')
    
    if start_idx == -1 or end_idx == -1:
        raise ValueError("No valid JSON object found in input")
    
    json_string = json_string[start_idx:end_idx+1]
    
    # Remove common trailing commas (invalid in JSON)
    json_string = re.sub(r',(\s*[}\]])', r'\1', json_string)
    
    return json_string


def validate_and_fix_json_file(input_path, output_path=None):
    """
    Validate and fix JSON file
    
    Args:
        input_path: Path to input JSON file
        output_path: Path to save fixed JSON (optional, overwrites if None)
    """
    
    print(f"📂 Reading file: {input_path}")
    
    # Read file with different encodings
    content = None
    for encoding in ['utf-8', 'utf-8-sig', 'latin-1']:
        try:
            with open(input_path, 'r', encoding=encoding) as f:
                content = f.read()
            print(f"✅ File read successfully with {encoding} encoding")
            break
        except UnicodeDecodeError:
            continue
    
    if content is None:
        raise ValueError("Could not read file with any encoding")
    
    # Clean LLM output
    print("🧹 Cleaning JSON...")
    cleaned = clean_llm_json(content)
    
    # Parse JSON
    try:
        data = json.loads(cleaned)
        print("✅ JSON is valid!")
    except json.JSONDecodeError as e:
        print(f"❌ JSON Error at line {e.lineno}, column {e.colno}")
        print(f"   Message: {e.msg}")
        
        # Try to find the error location
        lines = cleaned.split('\n')
        if e.lineno <= len(lines):
            error_line = lines[e.lineno - 1]
            print(f"   Problem line: {error_line[:100]}")
        
        # Attempt to fix common issues
        print("\n🔧 Attempting auto-fix...")
        
        # Fix 1: Remove trailing commas
        cleaned = re.sub(r',(\s*[}\]])', r'\1', cleaned)
        
        # Fix 2: Escape unescaped quotes in strings
        # This is tricky and may not work perfectly
        
        # Try parsing again
        try:
            data = json.loads(cleaned)
            print("✅ Auto-fix successful!")
        except json.JSONDecodeError as e2:
            print(f"❌ Auto-fix failed: {e2}")
            
            # Save the cleaned version for manual inspection
            debug_path = input_path.replace('.json', '_debug.txt')
            with open(debug_path, 'w', encoding='utf-8') as f:
                f.write(cleaned)
            print(f"💾 Saved cleaned JSON to: {debug_path}")
            print("   Please inspect manually and fix errors")
            raise
    
    # Save fixed JSON
    if output_path is None:
        output_path = input_path
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    
    print(f"💾 Fixed JSON saved to: {output_path}")
    
    return data


def validate_html_json_structure(data):
    """
    Validate Step 3 HTML JSON structure
    """
    
    print("\n🔍 Validating JSON structure...")
    
    required_keys = ['conversion_metadata', 'html_data', 'html_validation']
    
    for key in required_keys:
        if key not in data:
            print(f"⚠️  Missing key: {key}")
        else:
            print(f"✅ Found key: {key}")
    
    # Check html_data fields
    if 'html_data' in data:
        html_data = data['html_data']
        print(f"\n📊 Found {len(html_data)} fields in html_data:")
        for field_name in html_data.keys():
            value_length = len(str(html_data[field_name]))
            print(f"   - {field_name}: {value_length} characters")
    
    print("\n✅ Validation complete!")


# CLI Usage
if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python json_validator.py <input_json_file> [output_json_file]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    try:
        data = validate_and_fix_json_file(input_file, output_file)
        validate_html_json_structure(data)
        print("\n🎉 All checks passed!")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        sys.exit(1)
"""
FastAPI App: Word to Excel Converter - Vercel Optimized
Only Groq API - Minimal dependencies
NO python-docx, NO pandas, NO openpyxl
"""

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
import json
import io
import re
from groq import Groq
import traceback
from zipfile import ZipFile
from xml.etree import ElementTree as ET

app = FastAPI(title="Word to Excel Converter - Groq")

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

MAX_FILE_SIZE = 4.5 * 1024 * 1024  # 4.5MB


def extract_text_from_docx(file_content: bytes) -> str:
    """Extract text from .docx using zipfile (no python-docx needed)"""
    try:
        text_parts = []
        
        with ZipFile(io.BytesIO(file_content)) as docx:
            # Extract main document
            xml_content = docx.read('word/document.xml')
            tree = ET.fromstring(xml_content)
            
            # Namespace
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            # Extract paragraphs
            for para in tree.findall('.//w:p', ns):
                texts = []
                for text in para.findall('.//w:t', ns):
                    if text.text:
                        texts.append(text.text)
                if texts:
                    text_parts.append(''.join(texts))
            
            # Extract tables
            for table in tree.findall('.//w:tbl', ns):
                text_parts.append("\n=== TABLE ===")
                for row in table.findall('.//w:tr', ns):
                    cells = []
                    for cell in row.findall('.//w:tc', ns):
                        cell_texts = []
                        for text in cell.findall('.//w:t', ns):
                            if text.text:
                                cell_texts.append(text.text)
                        cells.append(''.join(cell_texts).strip())
                    text_parts.append(" | ".join(cells))
                text_parts.append("=== END TABLE ===\n")
        
        return "\n".join(text_parts)
    
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Lỗi đọc file Word: {str(e)}")


def detect_structure(text: str) -> dict:
    """Phát hiện số câu hỏi và tên test"""
    # Đếm câu hỏi
    question_patterns = [
        r'^\s*\*?\*?\s*(\d+)\.\s+.+\?',
        r'###\s*\*?\*?\s*Câu\s+(\d+):',
        r'^\s*Câu\s+(\d+)[:：]',
        r'^\s*Question\s+(\d+)[:：]',
    ]
    
    all_question_numbers = []
    for pattern in question_patterns:
        matches = re.findall(pattern, text, re.MULTILINE | re.IGNORECASE)
        if matches:
            all_question_numbers.extend([int(m) for m in matches])
    
    unique_questions = sorted(set(all_question_numbers))
    num_questions = len(unique_questions) if unique_questions else 20
    
    print(f"✓ Detected {num_questions} questions")
    
    # Đếm topics
    topic_patterns = [
        r'Phần\s+(\d+)\.',
        r'Part\s+(\d+)\.',
        r'Section\s+(\d+)\.',
    ]
    
    all_topic_numbers = []
    for pattern in topic_patterns:
        matches = re.findall(pattern, text, re.MULTILINE | re.IGNORECASE)
        if matches:
            all_topic_numbers.extend([int(m) for m in matches])
    
    unique_topics = sorted(set(all_topic_numbers))
    num_topics = len(unique_topics) if unique_topics else 6
    
    # Trích xuất tên test
    test_name_patterns = [
        r'(?:\*\*)?Bài\s+Test\s+Đánh\s+giá\s+(.+?)(?:\*\*|\(|$)',
        r'(?:\*\*)?Bài\s+Kiểm\s+Tra\s+(.+?)(?:\*\*|\(|\n)',
        r'(?:\*\*)?(.+?Test.*?Đánh\s+giá.*?)(?:\*\*|\(|\n)',
        r'(?:\*\*)?(.+?)\s+(?:Check|Assessment|Health\s+Check)(?:\*\*|\s+\(|\n|$)',
    ]
    
    detected_test_name = None
    lines = text.split('\n')[:15]
    
    for line in lines:
        line = line.strip()
        if not line or len(line) < 5:
            continue
        
        for pattern in test_name_patterns:
            match = re.search(pattern, line, re.IGNORECASE)
            if match:
                detected_test_name = match.group(1).strip()
                detected_test_name = re.sub(r'\*+', '', detected_test_name)
                detected_test_name = re.sub(r'\[.*?\]', '', detected_test_name)
                detected_test_name = re.sub(r'\(.*?\)', '', detected_test_name)
                detected_test_name = detected_test_name.strip()
                
                if 5 <= len(detected_test_name) <= 100:
                    print(f"✓ Detected test name: '{detected_test_name}'")
                    break
        
        if detected_test_name:
            break
    
    if not detected_test_name:
        detected_test_name = 'Business Assessment'
    
    return {
        'num_questions': num_questions,
        'num_topics': num_topics,
        'detected_test_name': detected_test_name
    }


def analyze_with_groq(text: str, api_key: str, language: str, 
                      structure: dict, metadata: dict) -> list:
    """Sử dụng Groq để phân tích"""
    try:
        num_questions = structure['num_questions']
        
        # Chia batch nếu quá nhiều câu
        if num_questions > 15:
            print(f"Splitting {num_questions} questions into batches...")
            all_data = []
            batch_size = 10
            
            for i in range(0, num_questions, batch_size):
                batch_questions = min(batch_size, num_questions - i)
                print(f"Processing batch: questions {i+1} to {i+batch_questions}")
                
                batch_structure = {
                    'num_questions': batch_questions,
                    'num_topics': structure['num_topics']
                }
                
                batch_data = _call_groq_api(
                    text, api_key, language, 
                    batch_structure, metadata, i+1
                )
                all_data.extend(batch_data)
            
            return all_data[:num_questions]
        else:
            return _call_groq_api(
                text, api_key, language, 
                structure, metadata, 1
            )
    
    except Exception as e:
        error_msg = f"Lỗi gọi Groq API: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        raise HTTPException(status_code=500, detail=error_msg)


def _call_groq_api(text: str, api_key: str, language: str,
                   structure: dict, metadata: dict, start_question: int = 1) -> list:
    """Call Groq API"""
    num_questions = structure['num_questions']
    lang_full = 'Vietnamese' if language == 'vi' else 'English'
    detected_test_name = structure.get('detected_test_name', 'Business Assessment')
    
    # Lấy metadata từ form
    test_name = metadata.get('test_name') or detected_test_name
    test_category = metadata.get('test_category') or test_name
    
    # Tạo hashtag
    def make_hashtag(name: str) -> str:
        hashtag = re.sub(r'[^\w\s]', '', name.lower())
        hashtag = re.sub(r'\s+', '_', hashtag)
        return f"#{hashtag}"
    
    test_hashtag = metadata.get('test_hashtag')
    if test_hashtag and not test_hashtag.startswith('#'):
        test_hashtag = f"#{test_hashtag}"
    if not test_hashtag:
        test_hashtag = make_hashtag(test_name)
    
    test_cost = metadata.get('test_cost', '0')
    topic_name = metadata.get('topic_name') or test_name
    topic_category = metadata.get('topic_category') or test_category
    topic_hashtag = metadata.get('topic_hashtag')
    if topic_hashtag and not topic_hashtag.startswith('#'):
        topic_hashtag = f"#{topic_hashtag}"
    if not topic_hashtag:
        topic_hashtag = test_hashtag
    
    question_category = metadata.get('question_category') or test_category
    question_hashtag = metadata.get('question_hashtag')
    if question_hashtag and not question_hashtag.startswith('#'):
        question_hashtag = f"#{question_hashtag}"
    if not question_hashtag:
        question_hashtag = test_hashtag
    
    reference_link = metadata.get('reference_link_url', '')
    
    default_good = "Đã làm tốt!" if language == 'vi' else "Well done!"
    
    # Lấy phần text liên quan
    question_patterns = [
        rf'^\s*\*?\*?\s*{start_question}\.\s+',
        rf'###\s*\*?\*?\s*Câu\s+{start_question}:',
        rf'^\s*Câu\s+{start_question}[:：]',
    ]
    
    relevant_text = None
    for pattern in question_patterns:
        match = re.search(pattern, text, re.MULTILINE)
        if match:
            start_pos = max(0, match.start() - 800)
            relevant_text = text[start_pos:][:18000]
            break
    
    if not relevant_text:
        relevant_text = text[:18000]
    
    english_instruction = ""
    if language == 'en':
        english_instruction = """
  * SPECIAL FOR ENGLISH: Extract ONLY the English terms in parentheses
  * Example: "Nhận diện thương hiệu (Brand Identity)" → Extract: "Brand Identity"
  * Remove all Vietnamese text, keep only English terms
  * PRESERVE line breaks using \\n"""
    else:
        english_instruction = """
  * FOR VIETNAMESE: Keep the full text
  * PRESERVE line breaks using \\n"""
    
    prompt = f"""
Analyze the Word document and extract questions {start_question} to {start_question + num_questions - 1}.

CRITICAL: Return ONLY a pure JSON array [ ... ] with NO markdown, NO text.

QUESTION DETECTION:
- Questions may appear as: "**1. Question?**", "### **Câu 1:**", "Câu 1:", "Question 1:"
- Extract COMPLETE question text
- Question numbers: {start_question}, {start_question+1}, {start_question+2}, etc.
- Remove ALL emojis and icons from question text

ANSWER EXTRACTION:
- Each question may have 3, 4, or 5 answers (A, B, C, D, E)
- ALWAYS check for ALL 5 possible answers
- Look for patterns: "1.", "2.", "3.", "4.", "5." OR "A.", "B.", "C.", "D.", "E."
- If only 3 answers: leave answer4_text and answer5_text empty
- If 4 answers: leave answer5_text empty
- If 5 answers: fill ALL answer fields
- Remove ALL emojis from answer text

Each object must have ALL 31 keys:
{{
  "language": "{language}",
  "test_name": "{test_name}",
  "test_category": "{test_category}",
  "test_hashtag": "{test_hashtag}",
  "test_created_at": "",
  "test_updated_at": "",
  "test_cost": "{test_cost}",
  "topic_name": "{topic_name}",
  "topic_category": "{topic_category}",
  "topic_hashtag": "{topic_hashtag}",
  "topic_created_at": "",
  "topic_updated_at": "",
  "question_text": "...",
  "question_category": "{question_category}",
  "question_hashtag": "{question_hashtag}",
  "question_created_at": "",
  "question_updated_at": "",
  "answer1_text": "...",
  "answer1_info": "improve/review/good",
  "answer2_text": "...",
  "answer2_info": "improve/review/good",
  "answer3_text": "...",
  "answer3_info": "improve/review/good",
  "answer4_text": "...",
  "answer4_info": "improve/review/good",
  "answer5_text": "...",
  "answer5_info": "improve/review/good",
  "good_comment": "{default_good}",
  "improve_comment": "...",
  "read_more_comment": "...",
  "reference_link_url": "{reference_link}"
}}

RULES:
- Extract COMPLETE question_text in {lang_full}
- For 5 answers: "improve" (1-worst), "review" (2), "review" (3), "good" (4), "good" (5-best)
- For 3 answers: "improve" (A-weak), "review" (B-moderate), "good" (C-strong)
- good_comment: "{default_good}"
- reference_link_url: "{reference_link}"
- improve_comment and read_more_comment: Extract from reading suggestions
  * Look for: "📌 Gợi ý đọc thêm:", "📖 **Gợi ý đọc thêm:**", "**Gợi ý tìm hiểu thêm:**"{english_instruction}
  * Remove emojis and icons
  * If not found, use ""
- Empty strings for unused answers and *_created_at/*_updated_at

Text:
{relevant_text}
"""
    
    estimated_tokens = num_questions * 800 + 1500
    max_tokens = min(estimated_tokens, 16000)
    
    print(f"Requesting {max_tokens} tokens for {num_questions} questions")
    
    try:
        client = Groq(api_key=api_key)
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": "Return ONLY valid, complete JSON array."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=max_tokens
        )
        result_text = response.choices[0].message.content.strip()
    
    except Exception as api_error:
        error_msg = str(api_error)
        
        if 'rate_limit' in error_msg.lower() or '429' in error_msg:
            raise ValueError(
                "❌ Groq API đã hết quota (100k tokens/ngày). "
                "Giải pháp:\n"
                "1. Đợi 34 phút để reset quota\n"
                "2. HOẶC nâng cấp Groq lên Dev Tier tại: https://console.groq.com/settings/billing\n"
                "3. HOẶC tạo API key mới tại: https://console.groq.com/keys"
            )
        
        raise ValueError(f"❌ Lỗi gọi Groq API: {error_msg}")
    
    # Xử lý JSON
    result_text = result_text.strip()
    
    if result_text.startswith("```json"):
        result_text = result_text[7:].strip()
    if result_text.startswith("```"):
        result_text = result_text[3:].strip()
    if result_text.endswith("```"):
        result_text = result_text[:-3].strip()
    
    result_text = result_text.strip()
    
    # Fix truncated JSON
    if not result_text.endswith(']'):
        print("WARNING: JSON truncated, attempting fix...")
        last_complete = result_text.rfind('},')
        if last_complete > 0:
            result_text = result_text[:last_complete+1] + ']'
        else:
            raise ValueError("JSON bị cắt và không thể sửa. Thử giảm số câu hỏi.")
    
    try:
        data = json.loads(result_text)
    except json.JSONDecodeError as je:
        raise ValueError(f"JSON không hợp lệ: {str(je)}")
    
    if isinstance(data, dict):
        data = [data]
    
    if not isinstance(data, list) or not data:
        raise ValueError("AI trả về dữ liệu không hợp lệ")
    
    # Post-process: Remove emojis
    def remove_emojis(text):
        if not text:
            return text
        emoji_pattern = re.compile(
            "["
            u"\U0001F600-\U0001F64F"
            u"\U0001F300-\U0001F5FF"
            u"\U0001F680-\U0001F6FF"
            u"\U0001F1E0-\U0001F1FF"
            u"\U00002702-\U000027B0"
            u"\U000024C2-\U0001F251"
            u"\U0001F900-\U0001F9FF"
            u"\U0001FA00-\U0001FA6F"
            u"\U00002600-\U000026FF"
            u"\U00002700-\U000027BF"
            "]+", flags=re.UNICODE
        )
        text = emoji_pattern.sub('', text)
        text = re.sub(r'[📌📖✅📹💡⚠️✔❌🎯🚀📊📈📉💰🎪🌐⭐🔥💪🔝]', '', text)
        text = re.sub(r'\s+', ' ', text).strip()
        return text
    
    for item in data:
        for key in item.keys():
            if isinstance(item[key], str) and item[key]:
                item[key] = remove_emojis(item[key])
                if key in ['improve_comment', 'read_more_comment']:
                    item[key] = item[key].replace('\\n', '\n')
    
    print(f"✓ Successfully parsed {len(data)} questions")
    return data


def convert_json_to_csv(data: list) -> bytes:
    """Convert JSON to CSV (lighter than Excel)"""
    try:
        if not data or not isinstance(data, list):
            raise ValueError("Dữ liệu không hợp lệ")
        
        # 31 columns
        columns = [
            "language", "test_name", "test_category", "test_hashtag",
            "test_created_at", "test_updated_at", "test_cost",
            "topic_name", "topic_category", "topic_hashtag",
            "topic_created_at", "topic_updated_at",
            "question_text", "question_category", "question_hashtag",
            "question_created_at", "question_updated_at",
            "answer1_text", "answer1_info", "answer2_text", "answer2_info",
            "answer3_text", "answer3_info", "answer4_text", "answer4_info",
            "answer5_text", "answer5_info",
            "good_comment", "improve_comment", "read_more_comment",
            "reference_link_url"
        ]
        
        # Create CSV
        output = io.StringIO()
        
        # Write header
        output.write(','.join(f'"{col}"' for col in columns) + '\n')
        
        # Write rows
        for row in data:
            values = []
            for col in columns:
                value = str(row.get(col, '')).replace('"', '""')  # Escape quotes
                values.append(f'"{value}"')
            output.write(','.join(values) + '\n')
        
        return output.getvalue().encode('utf-8-sig')  # UTF-8 with BOM for Excel
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi tạo CSV: {str(e)}")


@app.get("/", response_class=HTMLResponse)
async def home():
    """Serve HTML interface"""
    try:
        with open("templates/index.html", "r", encoding="utf-8") as f:
            html_content = f.read()
        return HTMLResponse(content=html_content)
    except FileNotFoundError:
        return HTMLResponse(content="""
        <html>
            <body>
                <h1>Word to CSV Converter - Groq</h1>
                <p>Error: Template file not found. Please create templates/index.html</p>
            </body>
        </html>
        """, status_code=404)


@app.post("/convert")
async def convert_word_to_csv(
    file: UploadFile = File(...),
    api_key: str = Form(...),
    language: str = Form('vi'),
    test_name: str = Form(''),
    test_category: str = Form(''),
    test_hashtag: str = Form(''),
    test_cost: str = Form('0'),
    topic_name: str = Form(''),
    topic_category: str = Form(''),
    topic_hashtag: str = Form(''),
    question_category: str = Form(''),
    question_hashtag: str = Form(''),
    reference_link_url: str = Form('')
):
    """Convert Word to CSV with Groq"""
    
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Chỉ chấp nhận file .docx")
    
    file_content = await file.read()
    
    if len(file_content) > MAX_FILE_SIZE:
        raise HTTPException(status_code=400, detail=f"File quá lớn (max 4.5MB)")
    
    if not api_key or len(api_key) < 20:
        raise HTTPException(status_code=400, detail="API key không hợp lệ")
    
    if not api_key.startswith('gsk_'):
        raise HTTPException(status_code=400, detail='Groq API key phải bắt đầu bằng "gsk_"')
    
    metadata = {
        'test_name': test_name.strip(),
        'test_category': test_category.strip(),
        'test_hashtag': test_hashtag.strip(),
        'test_cost': test_cost.strip(),
        'topic_name': topic_name.strip(),
        'topic_category': topic_category.strip(),
        'topic_hashtag': topic_hashtag.strip(),
        'question_category': question_category.strip(),
        'question_hashtag': question_hashtag.strip(),
        'reference_link_url': reference_link_url.strip()
    }
    
    try:
        print("=" * 80)
        print(f"Processing file: {file.filename}")
        
        text = extract_text_from_docx(file_content)
        
        if not text.strip():
            raise HTTPException(status_code=400, detail="File Word không có nội dung")
        
        print(f"\nFirst 800 chars:\n{text[:800]}\n")
        
        structure = detect_structure(text)
        print(f"✓ Structure: {structure['num_questions']} questions, {structure['num_topics']} topics")
        
        print("\nAnalyzing with Groq...")
        structured_data = analyze_with_groq(
            text, api_key, language, structure, metadata
        )
        
        print(f"✓ Got {len(structured_data)} questions")
        
        print("\nCreating CSV...")
        csv_content = convert_json_to_csv(structured_data)
        
        output_filename = file.filename.replace('.docx', f'_converted_{language}.csv')
        
        print(f"✓ Success! Output: {output_filename}")
        print("=" * 80)
        
        return StreamingResponse(
            io.BytesIO(csv_content),
            media_type="text/csv",
            headers={
                "Content-Disposition": f"attachment; filename={output_filename}"
            }
        )
    
    except HTTPException:
        raise
    except Exception as e:
        error_detail = f"Lỗi: {str(e)}\n{traceback.format_exc()}"
        print(error_detail)
        raise HTTPException(status_code=500, detail=error_detail)


@app.get("/health")
async def health_check():
    """Health check"""
    return {
        "status": "healthy",
        "version": "3.2-vercel-optimized",
        "features": ["groq_only", "minimal_deps", "csv_output"]
    }


# Vercel handler
handler = app
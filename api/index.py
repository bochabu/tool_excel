"""
FastAPI App: Word to Excel Converter v3.1
Fixed: Hỗ trợ nhiều format câu hỏi (Brand Health + Customer Experience)
"""

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
import docx
import pandas as pd
import json
import io
import re
from openai import OpenAI
from groq import Groq
import google.generativeai as genai
import traceback

app = FastAPI(title="Word to Excel Converter v3.1")

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
    """Trích xuất text và bảng từ file .docx"""
    try:
        doc = docx.Document(io.BytesIO(file_content))
        extracted_text = []
        
        for para in doc.paragraphs:
            if para.text.strip():
                extracted_text.append(para.text)
        
        for table in doc.tables:
            extracted_text.append("\n=== TABLE ===")
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                extracted_text.append(" | ".join(row_data))
            extracted_text.append("=== END TABLE ===\n")
        
        return "\n".join(extracted_text)
    
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Lỗi đọc file Word: {str(e)}")


def detect_structure(text: str) -> dict:
    """
    Phát hiện số câu hỏi, topics và tên test từ Word
    FIXED: Hỗ trợ nhiều format câu hỏi
    """
    # ĐẾM CÂU HỎI - Hỗ trợ nhiều pattern
    question_patterns = [
        r'^\s*\*?\*?\s*(\d+)\.\s+.+\?',  # Format: **1. Câu hỏi?**
        r'###\s*\*?\*?\s*Câu\s+(\d+):',  # Format: ### **Câu 1:**
        r'^\s*Câu\s+(\d+)[:：]',          # Format: Câu 1: hoặc Câu 1：
        r'^\s*Question\s+(\d+)[:：]',     # Format: Question 1:
    ]
    
    all_question_numbers = []
    for pattern in question_patterns:
        matches = re.findall(pattern, text, re.MULTILINE | re.IGNORECASE)
        if matches:
            all_question_numbers.extend([int(m) for m in matches])
    
    # Loại bỏ trùng lặp và đếm số câu hỏi duy nhất
    unique_questions = sorted(set(all_question_numbers))
    num_questions = len(unique_questions) if unique_questions else 20
    
    print(f"✓ Detected question numbers: {unique_questions}")
    print(f"✓ Total unique questions: {num_questions}")
    
    # ĐẾM TOPICS
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
    
    print(f"✓ Detected topics: {unique_topics}")
    print(f"✓ Total topics: {num_topics}")
    
    # TRÍCH XUẤT TÊN TEST
    test_name_patterns = [
        r'(?:\*\*)?Bài\s+Test\s+Đánh\s+giá\s+(.+?)(?:\*\*|\(|$)',
        r'(?:\*\*)?Bài\s+Kiểm\s+Tra\s+(.+?)(?:\*\*|\(|\n)',
        r'(?:\*\*)?(.+?Test.*?Đánh\s+giá.*?)(?:\*\*|\(|\n)',
        r'(?:\*\*)?Bài\s+Test\s+(.+?)(?:\*\*|\(|$)',
        r'(?:\*\*)?(.+?)\s+(?:Check|Assessment|Health\s+Check)(?:\*\*|\s+\(|\n|$)',
        r'^(?:\*\*)?(.+?Test.+?)(?:\*\*|\n)',
        r'^\*\*(.+?)\*\*',
    ]
    
    detected_test_name = None
    lines = text.split('\n')[:15]  # Xem 15 dòng đầu
    
    for line in lines:
        line = line.strip()
        if not line or len(line) < 5:
            continue
            
        for pattern in test_name_patterns:
            match = re.search(pattern, line, re.IGNORECASE)
            if match:
                detected_test_name = match.group(1).strip()
                # Clean up
                detected_test_name = re.sub(r'\*+', '', detected_test_name)
                detected_test_name = re.sub(r'\[.*?\]', '', detected_test_name)
                detected_test_name = re.sub(r'\(.*?\)', '', detected_test_name)
                detected_test_name = detected_test_name.strip()
                
                # Kiểm tra độ dài hợp lý
                if 5 <= len(detected_test_name) <= 100:
                    print(f"✓ Detected test name: '{detected_test_name}' from line: '{line[:80]}'")
                    break
        
        if detected_test_name:
            break
    
    # Fallback
    if not detected_test_name:
        detected_test_name = 'Business Assessment'
        print(f"⚠ Could not detect test name, using fallback: '{detected_test_name}'")
    
    return {
        'num_questions': num_questions,
        'num_topics': num_topics,
        'detected_test_name': detected_test_name
    }


def analyze_with_llm(text: str, api_key: str, language: str, 
                     api_provider: str, structure: dict, metadata: dict) -> list:
    """
    Sử dụng LLM để phân tích và chuyển đổi
    FIXED: Prompt được cải thiện để xử lý nhiều format
    """
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
                
                batch_data = _call_llm_api(
                    text, api_key, language, api_provider, 
                    batch_structure, metadata, i+1
                )
                all_data.extend(batch_data)
            
            return all_data[:num_questions]
        else:
            return _call_llm_api(
                text, api_key, language, api_provider, 
                structure, metadata, 1
            )
    
    except Exception as e:
        error_msg = f"Lỗi gọi {api_provider} API: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        raise HTTPException(status_code=500, detail=error_msg)


def _call_llm_api(text: str, api_key: str, language: str, api_provider: str,
                  structure: dict, metadata: dict, start_question: int = 1) -> list:
    """Helper function gọi LLM API - FIXED prompt"""
    num_questions = structure['num_questions']
    lang_full = 'Vietnamese' if language == 'vi' else 'English'
    detected_test_name = structure.get('detected_test_name', 'Business Assessment')
    
    # Lấy metadata từ form
    test_name = metadata.get('test_name') or detected_test_name
    test_category = metadata.get('test_category') or test_name
    
    # Tạo hashtag từ tên nếu không có
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
    
    # Topic metadata
    topic_name = metadata.get('topic_name') or test_name
    topic_category = metadata.get('topic_category') or test_category
    topic_hashtag = metadata.get('topic_hashtag')
    if topic_hashtag and not topic_hashtag.startswith('#'):
        topic_hashtag = f"#{topic_hashtag}"
    if not topic_hashtag:
        topic_hashtag = test_hashtag
    
    # Question metadata
    question_category = metadata.get('question_category') or test_category
    question_hashtag = metadata.get('question_hashtag')
    if question_hashtag and not question_hashtag.startswith('#'):
        question_hashtag = f"#{question_hashtag}"
    if not question_hashtag:
        question_hashtag = test_hashtag
    
    # Reference link
    reference_link = metadata.get('reference_link_url', '')
    
    # Default comments
    if language == 'vi':
        default_good = "Đã làm tốt!"
    else:
        default_good = "Well done!"
    
    # Lấy phần text liên quan - MỞ RỘNG để bao gồm cả text trước câu hỏi
    question_patterns = [
        rf'^\s*\*?\*?\s*{start_question}\.\s+',
        rf'###\s*\*?\*?\s*Câu\s+{start_question}:',
        rf'^\s*Câu\s+{start_question}[:：]',
    ]
    
    relevant_text = None
    for pattern in question_patterns:
        match = re.search(pattern, text, re.MULTILINE)
        if match:
            # Lấy từ vị trí này đến hết, MỞ RỘNG HƠN để đủ 5 đáp án
            start_pos = max(0, match.start() - 800)
            relevant_text = text[start_pos:][:18000]  # Tăng lên 18000 ký tự cho đủ 5 answers
            break
    
    if not relevant_text:
        relevant_text = text[:18000]
    
    # IMPROVED PROMPT - Hỗ trợ nhiều format
    english_instruction = ""
    if language == 'en':
        english_instruction = """
  * SPECIAL FOR ENGLISH: Extract ONLY the English terms in parentheses from the reading suggestions
  * Example: "Nhận diện thương hiệu (Brand Identity), Sự nhất quán thương hiệu (Brand Consistency)" → Extract: "Brand Identity, Brand Consistency"
  * Remove all Vietnamese text, keep only English terms separated by commas
  * If a term has no parentheses (like "Brand Recognition"), keep it as is
  * PRESERVE line breaks between bullet points using \\n"""
    else:
        english_instruction = """
  * FOR VIETNAMESE: Keep the full text including both Vietnamese and English parts
  * PRESERVE line breaks between bullet points using \\n"""
    
    prompt = f"""
Analyze the Word document and extract questions {start_question} to {start_question + num_questions - 1}.

CRITICAL: Return ONLY a pure JSON array [ ... ] with NO markdown, NO text.

IMPORTANT - QUESTION DETECTION:
- Questions may appear in various formats:
  * "**1. Question text?**"
  * "### **Câu 1: Question text?**"
  * "Câu 1: Question text?"
  * "Question 1: Question text?"
- Extract the COMPLETE question text regardless of format
- Question numbers must be sequential: {start_question}, {start_question+1}, {start_question+2}, etc.
- IMPORTANT: Remove ALL emojis and icons (🔹, 📌, ✅, etc.) from question text

CRITICAL - ANSWER EXTRACTION:
- Each question may have 3, 4, or 5 answers (A, B, C, D, E)
- ALWAYS check for ALL 5 possible answers before moving to next question
- Look for patterns: "1.", "2.", "3.", "4.", "5." OR "A.", "B.", "C.", "D.", "E."
- If only 3 answers exist, leave answer4_text and answer5_text as empty strings
- If 4 answers exist, leave answer5_text as empty string
- If 5 answers exist, fill ALL answer fields (answer1-5)
- DO NOT stop at answer3 if answer4 and answer5 exist in the document
- IMPORTANT: Remove ALL emojis and icons (✅, 🔹, 📌, etc.) from answer text

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
- ALL metadata values are provided above
- Extract COMPLETE question_text in {lang_full}, handle all formats
- Look for answers marked with numbers or letters
- For 5-answer questions: Classify as "improve" (1-worst), "review" (2), "review" (3), "good" (4), "good" (5-best)
- For 3-answer questions: Classify as "improve" (A-weak), "review" (B-moderate), "good" (C-strong)
- good_comment is ALWAYS: "{default_good}"
- reference_link_url is ALWAYS: "{reference_link}"
- improve_comment and read_more_comment MUST BE IDENTICAL, extracted from reading suggestions:
  * Look for: "📌 Gợi ý đọc thêm:", "📖 **Gợi ý đọc thêm:**", "**Gợi ý tìm hiểu thêm:**"
  * Extract everything after this marker until the next question starts{english_instruction}
  * CRITICAL: Preserve line breaks between bullet points - use \\n for line breaks
  * Example input: "- Item 1\\n- Item 2\\n- Item 3"
  * If bullet points start with "-" or "•", keep them and add \\n after each item
  * IMPORTANT: DO NOT include any emojis, icons, or special symbols (📌, 📖, ✅, 🔹, etc.) in the extracted text
  * Remove all emojis and icons from the extracted content
  * SEARCH CAREFULLY between last answer and the next question
  * If not found, use "" for both fields
- Empty strings for unused answer fields and all *_created_at/*_updated_at fields

Text to analyze:
{relevant_text}
"""
    
    # Tính max_tokens - tăng lên cho 5 answers
    estimated_tokens = num_questions * 800 + 1500  # Tăng từ 600 lên 800 tokens/câu
    max_tokens = min(estimated_tokens, 16000)
    
    print(f"Requesting {max_tokens} tokens for {num_questions} questions")
    
    # Gọi API
    try:
        if api_provider == 'gemini':
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash')
            generation_config = {
                'temperature': 0.1,
                'max_output_tokens': max_tokens,
            }
            response = model.generate_content(prompt, generation_config=generation_config)
            result_text = response.text.strip()
            
        elif api_provider == 'groq':
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
            
        else:  # openai
            client = OpenAI(api_key=api_key)
            response = client.chat.completions.create(
                model="gpt-4o-mini",
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
        
        # Xử lý rate limit errors
        if 'rate_limit' in error_msg.lower() or '429' in error_msg:
            if api_provider == 'groq':
                raise ValueError(
                    "❌ Groq API đã hết quota (100k tokens/ngày). "
                    "Giải pháp:\n"
                    "1. Đợi 34 phút để reset quota\n"
                    "2. HOẶC chuyển sang Gemini (miễn phí, không giới hạn) - Khuyên dùng!\n"
                    "3. HOẶC nâng cấp Groq lên Dev Tier tại: https://console.groq.com/settings/billing\n"
                    "4. HOẶC dùng OpenAI (trả phí nhưng ổn định)"
                )
            elif api_provider == 'gemini':
                raise ValueError(
                    "❌ Gemini API đã hết quota. "
                    "Giải pháp:\n"
                    "1. Tạo API key mới tại: https://aistudio.google.com/app/apikey\n"
                    "2. HOẶC chuyển sang Groq (miễn phí)\n"
                    "3. HOẶC dùng OpenAI (trả phí)"
                )
            else:
                raise ValueError(
                    "❌ OpenAI API đã hết quota hoặc credit. "
                    "Giải pháp:\n"
                    "1. Nạp thêm credit tại: https://platform.openai.com/account/billing\n"
                    "2. HOẶC chuyển sang Gemini (miễn phí, không giới hạn) - Khuyên dùng!\n"
                    "3. HOẶC chuyển sang Groq (miễn phí)"
                )
        
        raise ValueError(f"❌ Lỗi gọi {api_provider} API: {error_msg}")
    
    # Xử lý JSON
    result_text = result_text.strip()
    
    # Remove markdown
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
    
    # Parse JSON
    try:
        data = json.loads(result_text)
    except json.JSONDecodeError as je:
        raise ValueError(f"JSON không hợp lệ: {str(je)}")
    
    if isinstance(data, dict):
        data = [data]
    
    if not isinstance(data, list) or not data:
        raise ValueError("AI trả về dữ liệu không hợp lệ")
    
    # Post-process: Đảm bảo line breaks và loại bỏ icons/emojis
    for item in data:
        # Hàm loại bỏ emojis và icons
        def remove_emojis(text):
            if not text:
                return text
            # Remove emojis, icons, và special symbols
            import re
            # Pattern loại bỏ emojis và icons Unicode
            emoji_pattern = re.compile(
                "["
                u"\U0001F600-\U0001F64F"  # emoticons
                u"\U0001F300-\U0001F5FF"  # symbols & pictographs
                u"\U0001F680-\U0001F6FF"  # transport & map symbols
                u"\U0001F1E0-\U0001F1FF"  # flags (iOS)
                u"\U00002702-\U000027B0"  # dingbats
                u"\U000024C2-\U0001F251"
                u"\U0001F900-\U0001F9FF"  # Supplemental Symbols and Pictographs
                u"\U0001FA00-\U0001FA6F"  # Chess Symbols
                u"\U00002600-\U000026FF"  # Miscellaneous Symbols
                u"\U00002700-\U000027BF"  # Dingbats
                "]+", flags=re.UNICODE
            )
            text = emoji_pattern.sub('', text)
            
            # Loại bỏ các ký tự đặc biệt như 📌, 📖, ✅, 🔹, etc.
            text = re.sub(r'[📌📖✅🔹💡⚠️✓❌🎯🚀📊📈📉💰🎪🌐⭐🔥💪👍]', '', text)
            
            # Loại bỏ khoảng trắng thừa
            text = re.sub(r'\s+', ' ', text).strip()
            
            return text
        
        # Áp dụng cho tất cả các trường text
        for key in item.keys():
            if isinstance(item[key], str) and item[key]:
                # Loại bỏ emojis
                item[key] = remove_emojis(item[key])
                
                # Xử lý line breaks cho improve_comment và read_more_comment
                if key in ['improve_comment', 'read_more_comment']:
                    item[key] = item[key].replace('\\n', '\n')
    
    print(f"✓ Successfully parsed {len(data)} questions")
    return data


def convert_json_to_excel(data: list) -> bytes:
    """Chuyển đổi JSON sang Excel"""
    try:
        if not data or not isinstance(data, list):
            raise ValueError("Dữ liệu không hợp lệ")
        
        df = pd.DataFrame(data)
        
        # 31 cột chuẩn
        required_columns = [
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
        
        # Thêm cột thiếu
        for col in required_columns:
            if col not in df.columns:
                df[col] = ""
        
        # Sắp xếp lại cột
        df = df[required_columns]
        
        # Chuyển test_cost sang dạng số
        def convert_to_number(value):
            """Chuyển đổi giá trị sang số"""
            if pd.isna(value) or value == '' or value is None:
                return 0
            try:
                # Loại bỏ dấu phẩy, khoảng trắng và chuyển sang float
                cleaned = str(value).replace(',', '').replace(' ', '').strip()
                return float(cleaned) if cleaned else 0
            except:
                return 0
        
        df['test_cost'] = df['test_cost'].apply(convert_to_number)
        
        # Tạo Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Questions')
            
            # Auto-adjust column width và enable text wrap
            worksheet = writer.sheets['Questions']
            from openpyxl.styles import Alignment
            
            for idx, col in enumerate(required_columns):
                if len(df) > 0:
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(col)
                    )
                else:
                    max_length = len(col)
                
                col_letter = ''
                num = idx + 1
                while num > 0:
                    num -= 1
                    col_letter = chr(65 + (num % 26)) + col_letter
                    num //= 26
                worksheet.column_dimensions[col_letter].width = min(max_length + 2, 50)
                
                # Enable text wrap cho improve_comment và read_more_comment
                if col in ['improve_comment', 'read_more_comment']:
                    for row in range(2, len(df) + 2):  # Start from row 2 (skip header)
                        cell = worksheet[f'{col_letter}{row}']
                        cell.alignment = Alignment(wrap_text=True, vertical='top')
        
        output.seek(0)
        return output.getvalue()
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi tạo Excel: {str(e)}")


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
                <h1>Error: Template not found</h1>
                <p>Please create templates/index.html file</p>
            </body>
        </html>
        """, status_code=404)


@app.post("/convert")
async def convert_word_to_excel(
    file: UploadFile = File(...),
    api_key: str = Form(...),
    language: str = Form('vi'),
    api_provider: str = Form('groq'),
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
    """Convert Word to Excel với metadata tùy chỉnh"""
    
    # Validate
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Chỉ chấp nhận file .docx")
    
    file_content = await file.read()
    
    if len(file_content) > MAX_FILE_SIZE:
        raise HTTPException(status_code=400, detail=f"File quá lớn (max 4.5MB)")
    
    if not api_key or len(api_key) < 20:
        raise HTTPException(status_code=400, detail="API key không hợp lệ")
    
    # Chuẩn bị metadata
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
        # Bước 1: Trích xuất text
        print("=" * 80)
        print(f"Processing file: {file.filename}")
        print("=" * 80)
        text = extract_text_from_docx(file_content)
        
        if not text.strip():
            raise HTTPException(status_code=400, detail="File Word không có nội dung")
        
        # Debug
        print("\nFIRST 800 CHARS FROM WORD:")
        print(text[:800])
        print("=" * 80)
        
        # Bước 2: Phát hiện cấu trúc
        print("\nDetecting structure...")
        structure = detect_structure(text)
        print(f"✓ Structure: {structure['num_questions']} questions, {structure['num_topics']} topics")
        print(f"✓ Test name: '{structure.get('detected_test_name')}'")
        
        # Bước 3: Phân tích với LLM
        print(f"\nAnalyzing with {api_provider}...")
        structured_data = analyze_with_llm(
            text, api_key, language, api_provider, structure, metadata
        )
        
        print(f"✓ Got {len(structured_data)} questions from LLM")
        
        # Bước 4: Tạo Excel
        print("\nCreating Excel...")
        excel_content = convert_json_to_excel(structured_data)
        
        # Output filename
        output_filename = file.filename.replace('.docx', f'_converted_{language}.xlsx')
        
        print(f"✓ Success! Output: {output_filename}")
        print("=" * 80)
        
        return StreamingResponse(
            io.BytesIO(excel_content),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
        "version": "3.1",
        "features": ["multi_format_support", "custom_metadata", "auto_detect", "3_providers"]
    }


# Handler cho Vercel deployment
handler = app
#!/usr/bin/env python3
"""
User Feedback Analyzer & Automated PowerPoint Generator
تطبيق Flask لتحليل ملاحظات الزائر السري وإنشاء تقارير PowerPoint
"""

from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import os
import tempfile
import socket

app = Flask(__name__, static_folder='.')
CORS(app)

TEMPLATE_PATH = 'template.pptx'

def find_free_port():
    """البحث عن بورت فاضي تلقائياً"""
    ports_to_try = [8080, 5000, 5001, 3000, 8000, 9000, 4000, 7000, 6000]
    for port in ports_to_try:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(('', port))
                return port
        except OSError:
            continue
    # إذا كل البورتات مشغولة، استخدم بورت عشوائي
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]

@app.route('/')
def index():
    """الصفحة الرئيسية"""
    return send_from_directory('.', 'index.html')

@app.route('/generate-pptx', methods=['POST'])
def generate_pptx():
    """إنشاء بوربوينت من البيانات"""
    try:
        # استلام البيانات
        data = request.get_json()
        
        if not data or 'categories' not in data or not data['categories']:
            return jsonify({'error': 'لا توجد ملاحظات للتحليل'}), 400
        
        # تحويل categories إلى list إذا كان dict
        categories_data = data.get('categories', [])
        
        if isinstance(categories_data, dict):
            # إذا كان dict، نحوله لـ list من الـ values
            # لكن كل value هو category كامل (dict)
            categories = []
            for key, value in categories_data.items():
                if isinstance(value, dict):
                    # إذا الـ value هو category كامل
                    categories.append(value)
                elif isinstance(value, list):
                    # إذا الـ value هو list of notes
                    categories.append({
                        'name': key,
                        'notes': value
                    })
        elif isinstance(categories_data, list):
            categories = categories_data
        else:
            categories = []
        
        # فتح القالب
        if not os.path.exists(TEMPLATE_PATH):
            return jsonify({'error': f'القالب غير موجود: {TEMPLATE_PATH}'}), 500
        
        prs = Presentation(TEMPLATE_PATH)
        
        # تنظيف الشرائح (حذف المحتوى، الخلفية تبقى)
        for slide in prs.slides:
            shapes_to_delete = [shape for shape in slide.shapes]
            for shape in shapes_to_delete:
                try:
                    sp = shape.element
                    sp.getparent().remove(sp)
                except:
                    pass
        
        # تجميع الملاحظات حسب الفئة الفرعية
        by_subcategory = {}
        
        for category in categories:
            # التأكد من أن category هو dict
            if isinstance(category, str):
                continue
            
            category_name = category.get('name', 'غير محدد')
            notes = category.get('notes', [])
            
            for note in notes:
                # التأكد من أن note هو dict
                if isinstance(note, str):
                    continue
                    
                subcategory = note.get('subCategory', 'غير محدد')
                key = f"{category_name} ( {subcategory} )"
                
                if key not in by_subcategory:
                    by_subcategory[key] = []
                    
                by_subcategory[key].append({
                    'category': category_name,
                    'subCategory': subcategory,
                    'observation': note.get('observation', ''),
                    'repeatCount': note.get('repeatCount', 1)
                })
        
        if not by_subcategory:
            return jsonify({'error': 'لم يتم العثور على ملاحظات صالحة'}), 400
        
        slide_index = 0
        
        # === شريحة العنوان ===
        slide = prs.slides[slide_index]
        slide_index += 1
        
        # العنوان الرئيسي
        title_box = slide.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(10.0), Inches(1.5))
        title_frame = title_box.text_frame
        title_frame.text = "الخطة التصحيحية لملاحظات الزائر السري"
        
        for paragraph in title_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.size = Pt(40)
                run.font.name = 'Calibri'
                run.font.bold = True
        
        # العنوان الفرعي (المنشأة ورقم التذكرة)
        subtitle_box = slide.shapes.add_textbox(Inches(1.5), Inches(4.2), Inches(10.0), Inches(1.0))
        subtitle_frame = subtitle_box.text_frame
        subtitle_text = f"{data.get('hospital', 'غير محدد')}\n{data.get('ticketId', 'غير محدد')}"
        subtitle_frame.text = subtitle_text
        
        for paragraph in subtitle_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.size = Pt(24)
                run.font.name = 'Calibri'
        
        # === شرائح البيانات ===
        for idx, (subcategory_full, notes) in enumerate(sorted(by_subcategory.items()), 1):
            
            # استخدام شريحة من القالب
            if slide_index < len(prs.slides):
                slide = prs.slides[slide_index]
                slide_index += 1
            else:
                # تكرار الشريحة 2 إذا انتهت الشرائح
                template_slide = prs.slides[1]
                slide_layout = template_slide.slide_layout
                slide = prs.slides.add_slide(slide_layout)
                
                # تنظيف الشريحة الجديدة
                shapes_to_delete = [shape for shape in slide.shapes]
                for shape in shapes_to_delete:
                    try:
                        sp = shape.element
                        sp.getparent().remove(sp)
                    except:
                        pass
            
            # حساب ارتفاع الجدول حسب عدد الملاحظات
            if len(notes) <= 3:
                height = Inches(1.5)
            elif len(notes) <= 6:
                height = Inches(2.5)
            elif len(notes) <= 9:
                height = Inches(3.5)
            else:
                height = Inches(4.0)
            
            # إنشاء الجدول
            rows = len(notes) + 1
            cols = 3
            table_shape = slide.shapes.add_table(
                rows, cols, 
                Inches(1.5), Inches(2.0), 
                Inches(9.0), height
            )
            table = table_shape.table
            
            # عرض الأعمدة
            table.columns[0].width = Inches(0.8)  # الحالة
            table.columns[1].width = Inches(2.0)  # الخطة
            table.columns[2].width = Inches(8.5)  # الملاحظة
            
            # رأس الجدول (أزرق غامق + نص أبيض)
            headers = ['الحالة', 'الخطة التصحيحية', 'الملاحظة']
            for col_idx, header_text in enumerate(headers):
                cell = table.rows[0].cells[col_idx]
                cell.text = header_text
                
                cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                cell.vertical_anchor = 1
                
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Calibri'
                        run.font.size = Pt(18)
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)
                
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(68, 114, 196)
                
                cell.margin_left = Inches(0.05)
                cell.margin_right = Inches(0.05)
                cell.margin_top = Inches(0.05)
                cell.margin_bottom = Inches(0.05)
            
            # بيانات الجدول (ألوان متناوبة)
            for row_idx, note in enumerate(notes, start=1):
                # لون الصف (متناوب)
                row_color = RGBColor(217, 225, 242) if row_idx % 2 == 1 else RGBColor(255, 255, 255)
                
                # العمود 1: الحالة (فارغ)
                cell1 = table.rows[row_idx].cells[0]
                cell1.text = ""
                cell1.fill.solid()
                cell1.fill.fore_color.rgb = row_color
                cell1.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                
                # العمود 2: الخطة التصحيحية (فارغ)
                cell2 = table.rows[row_idx].cells[1]
                cell2.text = ""
                cell2.fill.solid()
                cell2.fill.fore_color.rgb = row_color
                cell2.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                
                # العمود 3: الملاحظة
                category = note['category']
                subcategory = note['subCategory']
                observation = note['observation']
                
                # النص بدون عدد التكرار
                full_text = f"في {category} ( {subcategory} ) {observation}"
                
                cell = table.rows[row_idx].cells[2]
                cell.text = full_text
                
                cell.fill.solid()
                cell.fill.fore_color.rgb = row_color
                
                cell.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
                
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Calibri'
                        run.font.size = Pt(14)
                        run.font.color.rgb = RGBColor(0, 0, 0)
                
                cell.text_frame.word_wrap = True
                cell.vertical_anchor = 1
                cell.margin_left = Inches(0.1)
                cell.margin_right = Inches(0.1)
                cell.margin_top = Inches(0.05)
                cell.margin_bottom = Inches(0.05)
        
        # حذف الشرائح الفاضية
        total_slides_used = slide_index
        total_slides = len(prs.slides)
        
        if total_slides_used < total_slides:
            slides_to_delete = total_slides - total_slides_used
            
            for _ in range(slides_to_delete):
                rId = prs.slides._sldIdLst[-1].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[-1]
        
        # حفظ الملف
        output = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
        prs.save(output.name)
        output.close()
        
        return send_file(
            output.name,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name=f'تحليل_الزائر_السري_{data.get("ticketId", "")}.pptx'
        )
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port, debug=False)

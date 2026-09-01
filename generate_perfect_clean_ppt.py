import os
import zipfile
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree

def create_replaced_slide10_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Theme Colors
    COLOR_BG = RGBColor(15, 23, 42)          # Dark Slate
    COLOR_CARD = RGBColor(30, 41, 59)        # Card Slate
    COLOR_CARD_BORDER = RGBColor(51, 65, 85)
    COLOR_PRIMARY = RGBColor(129, 140, 248)  # Indigo
    COLOR_AMBER = RGBColor(251, 191, 36)     # Amber
    COLOR_TEXT_MAIN = RGBColor(248, 250, 252)# White
    COLOR_TEXT_MUTED = RGBColor(203, 213, 225) # Soft Grey
    COLOR_CODE_BG = RGBColor(10, 15, 29)     # Code Box
    COLOR_CODE_TEXT = RGBColor(56, 189, 248) # Cyan

    slides_content = [
        # SLIDE 1: Direct Comparison on Slide 1
        {
            "tag": "SEMINAR INTRODUCTION & COMPARISON",
            "title": "JavaScript Array Methods vs 'for' Loops",
            "desc": "Why Modern JavaScript Replaces 6-Line Loops with 1-Line Safe Methods.",
            "bullets": [
                "✦ ❌ Traditional 'for' Loop: Requires empty arrays, index counters (i=0), and manual .push().",
                "✦ ✅ Modern Array Method: 1-line clean, readable, and declarative code.",
                "✦ 🛡️ ORIGINAL ARRAY STAYS 100% SAME: Modern methods NEVER modify or destroy original data!",
                "✦ Today's Agenda: map, filter, find, reduce, some, every, and method chaining pipelines."
            ],
            "code_title": "Slide 1: Direct Code Comparison",
            "code": "let nums = [1, 2, 3];\n\n// ❌ OLD WAY (for loop - 5 lines):\nlet doubled1 = [];\nfor (let i = 0; i < nums.length; i++) {\n    doubled1.push(nums[i] * 2);\n}\nconsole.log(doubled1); // 👉 Output: [ 2, 4, 6 ]\n\n// ✅ MODERN WAY (map() - 1 line):\nlet doubled2 = nums.map(n => n * 2);\nconsole.log(doubled2); // 👉 Output: [ 2, 4, 6 ]\n\nconsole.log(nums);     // 👉 Output: [ 1, 2, 3 ] (ORIGINAL STAYS SAME!)",
            "notes": "Good morning/afternoon everyone! Look at Slide 1: On the right, you can immediately see the difference between 5 lines of traditional for-loop and 1 line of modern map(). Notice that the original array [1, 2, 3] stays completely untouched!"
        },
        # Slide 2: Why Immutability Matters
        {
            "tag": "CORE ADVANTAGE",
            "title": "Why is 'Original Array Stays Same' So Important?",
            "desc": "How modern methods prevent accidental data destruction (Immutability).",
            "bullets": [
                "✦ 1. Bug Prevention: Accidental changes to original data cause hidden bugs across your app.",
                "✦ 2. Pure & Safe: Modern methods return a FRESH NEW array and NEVER touch the original.",
                "✦ 3. Predictable: You can reuse the original array multiple times without fear.",
                "✦ 4. Essential for React / Vue: Modern UI frameworks strictly require original data to remain unchanged!"
            ],
            "code_title": "Original Protection",
            "code": "let originalData = [10, 20, 30];\n\n// Filter creates new copy\nlet filtered = originalData.filter(x => x > 15);\n\nconsole.log(filtered);     // 👉 [ 20, 30 ] (New)\nconsole.log(originalData); // 👉 [ 10, 20, 30 ] (Untouched!)",
            "notes": "Explain: In modern apps, multiple parts of the app read the same array. If one part modifies it, everything breaks. Methods like map and filter prevent this by keeping original data safe."
        },
        # Slide 3: MAP
        {
            "tag": "1. MAP METHOD",
            "title": "map() vs 'for' Loop (Original Stays Same)",
            "desc": "Goal: Add 5 grace marks to all students in the list.",
            "bullets": [
                "✦ Input Data: [50, 60, 70]",
                "✦ ❌ for Loop: Requires creating empty array and manual pushes.",
                "✦ ✅ map(): Returns a brand new array with transformed items.",
                "✦ 🛡️ ORIGINAL STAYS SAME: Original 'marks' is still [50, 60, 70]!"
            ],
            "code_title": "map_comparison.js",
            "code": "let marks = [50, 60, 70];\n\n// ❌ OLD WAY (for loop):\nlet res1 = [];\nfor (let i = 0; i < marks.length; i++) {\n    res1.push(marks[i] + 5);\n}\nconsole.log(res1); // 👉 [ 55, 65, 75 ]\n\n// ✅ MODERN WAY (map):\nlet res2 = marks.map(m => m + 5);\nconsole.log(res2); // 👉 [ 55, 65, 75 ] (New Array)\nconsole.log(marks);// 👉 [ 50, 60, 70 ] (ORIGINAL UNTOUCHED!)",
            "notes": "Highlight line 'console.log(marks)': Notice that marks is still [50, 60, 70]. The original array did not change at all."
        },
        # Slide 4: FILTER
        {
            "tag": "2. FILTER METHOD",
            "title": "filter() vs 'for' Loop (Original Stays Same)",
            "desc": "Goal: Keep only passing marks (35 or above).",
            "bullets": [
                "✦ Input Data: [85, 30, 92, 25, 70]",
                "✦ ❌ for Loop: Needs empty array, loop, and if condition.",
                "✦ ✅ filter(): Collects passing marks into a fresh new array.",
                "✦ 🛡️ ORIGINAL STAYS SAME: Original 'marks' still holds all 5 marks!"
            ],
            "code_title": "filter_comparison.js",
            "code": "let marks = [85, 30, 92, 25, 70];\n\n// ❌ OLD WAY (for loop):\nlet pass1 = [];\nfor (let i = 0; i < marks.length; i++) {\n    if (marks[i] >= 35) pass1.push(marks[i]);\n}\nconsole.log(pass1); // 👉 [ 85, 92, 70 ]\n\n// ✅ MODERN WAY (filter):\nlet pass2 = marks.filter(m => m >= 35);\nconsole.log(pass2); // 👉 [ 85, 92, 70 ] (New Array)\nconsole.log(marks); // 👉 [ 85, 30, 92, 25, 70 ] (ORIGINAL UNTOUCHED!)",
            "notes": "Even though pass2 only has 3 items, the original marks array still safely contains all 5 marks."
        },
        # Slide 5: FIND
        {
            "tag": "3. FIND METHOD",
            "title": "find() vs 'for' Loop (Original Stays Same)",
            "desc": "Goal: Find the user with role = 'Admin'.",
            "bullets": [
                "✦ Input Data: List of 3 user objects",
                "✦ ❌ for Loop: Needs 'let found = null' and manual 'break'.",
                "✦ ✅ find(): Stops instantly at 1st match and returns the item.",
                "✦ 🛡️ ORIGINAL STAYS SAME: Original 'users' array still has all 3 users!"
            ],
            "code_title": "find_comparison.js",
            "code": "let users = [\n  { id: 1, role: 'Admin', name: 'Arun' },\n  { id: 2, role: 'User',  name: 'Priya' },\n  { id: 3, role: 'Admin', name: 'Divakar' }\n];\n\n// ❌ OLD WAY (for loop + break):\nlet found1 = null;\nfor (let i = 0; i < users.length; i++) {\n    if (users[i].role === 'Admin') { found1 = users[i]; break; }\n}\nconsole.log(found1.name); // 👉 \"Arun\"\n\n// ✅ MODERN WAY (find):\nlet found2 = users.find(u => u.role === 'Admin');\nconsole.log(found2.name); // 👉 \"Arun\"\nconsole.log(users.length);// 👉 3 (ORIGINAL ARRAY UNTOUCHED!)",
            "notes": "find() simply inspects the items and returns the reference—it never removes or modifies items in the original array."
        },
        # Slide 6: REDUCE
        {
            "tag": "4. REDUCE METHOD",
            "title": "reduce() vs 'for' Loop (Original Stays Same)",
            "desc": "Goal: Calculate total shopping bill.",
            "bullets": [
                "✦ Input Data: [100, 250, 50, 200]",
                "✦ ❌ for Loop: Modifies an external 'let total = 0' variable.",
                "✦ ✅ reduce(): Encapsulates the sum inside an accumulator.",
                "✦ 🛡️ ORIGINAL STAYS SAME: Original 'prices' array is completely unchanged!"
            ],
            "code_title": "reduce_comparison.js",
            "code": "let prices = [100, 250, 50, 200];\n\n// ❌ OLD WAY (for loop):\nlet total1 = 0;\nfor (let i = 0; i < prices.length; i++) total1 += prices[i];\nconsole.log('Total: ₹' + total1); // 👉 Total: ₹600\n\n// ✅ MODERN WAY (reduce):\nlet total2 = prices.reduce((sum, p) => sum + p, 0);\nconsole.log('Total: ₹' + total2); // 👉 Total: ₹600\nconsole.log(prices); // 👉 [ 100, 250, 50, 200 ] (ORIGINAL UNTOUCHED!)",
            "notes": "reduce() takes the numbers, computes a total, and leaves the prices array completely safe."
        },
        # Slide 7: SOME & EVERY
        {
            "tag": "5 & 6. SOME & EVERY",
            "title": "some() & every() vs 'for' Loop (Original Stays Same)",
            "desc": "Goal: Quick True / False integrity checks without modifying data.",
            "bullets": [
                "✦ Input Data: [80, 90, 40, 95]",
                "✦ .some(): 'Did anyone fail (<35)?' ➔ false",
                "✦ .every(): 'Did all pass (>=35)?' ➔ true",
                "✦ 🛡️ ORIGINAL STAYS SAME: Returns a simple boolean; original marks array stays 100% unchanged!"
            ],
            "code_title": "some_every_comparison.js",
            "code": "let marks = [80, 90, 40, 95];\n\n// ✅ Check with some() & every():\nlet anyFail = marks.some(m => m < 35);   // 👉 false\nlet allPass = marks.every(m => m >= 35); // 👉 true\n\nconsole.log('Any Failed?:', anyFail);\nconsole.log('All Passed?:', allPass);\nconsole.log('Original marks:', marks); \n// 👉 [ 80, 90, 40, 95 ] (ORIGINAL UNTOUCHED!)",
            "notes": "These inspection methods are pure read-only operations."
        },
        # Slide 8: SAFE vs MODIFYING (was slide 9)
        {
            "tag": "EASY UNDERSTANDING",
            "title": "Safe Methods vs Modifying Methods",
            "desc": "The 'Photocopy' Analogy vs 'Writing in Pen on the Original'.",
            "bullets": [
                "✦ 📄 SAFE METHODS (map, filter, reduce, find):",
                "   👉 Like making a Photocopy: Modifies the copy, leaves original clean!",
                "   👉 Returns a brand new array without touching the original data.",
                "✦ ✍️ MODIFYING METHODS (push, pop, sort, splice):",
                "   👉 Like writing in ink directly on the original document!",
                "   👉 Permanently changes the original array in memory."
            ],
            "code_title": "Photocopy vs Overwrite",
            "code": "// 1. ✅ SAFE (.map) -> Like a Photocopy:\nlet originalA = [10, 20];\nlet copyA = originalA.map(x => x * 2);\nconsole.log(copyA);     // 👉 [ 20, 40 ] (New Copy)\nconsole.log(originalA); // 👉 [ 10, 20 ] (ORIGINAL SAFE!)\n\n// 2. ⚠️ MODIFYING (.push) -> Writes on Original:\nlet originalB = [10, 20];\noriginalB.push(30);\nconsole.log(originalB); // 👉 [ 10, 20, 30 ] (ORIGINAL CHANGED!)",
            "notes": "Explain with the photocopy analogy: map() makes a photocopy, colors it, and gives it to you while your original stays safe in your drawer. push() writes directly on your original paper."
        },
        # Slide 9: METHOD CHAINING STEP BY STEP
        {
            "tag": "METHOD CHAINING EXPLAINED STEP BY STEP",
            "title": "How .filter() → .map() → .reduce() Works Together",
            "desc": "Real-Life Example: Calculate Total Grocery Bill with 5% Tax — Step by Step Pipeline.",
            "bullets": [
                "✦ INPUT: 3 items — Rice ₹100 (grocery), Shirt ₹500 (cloth), Dal ₹200 (grocery)",
                "✦ STEP A - .filter(): Pick ONLY grocery items → [Rice ₹100, Dal ₹200]",
                "✦ STEP B - .map(): Add 5% Tax to each → [₹105, ₹210]",
                "✦ STEP C - .reduce(): Add all together → ₹315",
                "✦ 🛡️ ORIGINAL STAYS SAME: items still has all 3 products (Rice, Shirt, Dal)!"
            ],
            "code_title": "chaining_pipeline.js — Output: ₹315",
            "code": "let items = [\n  { name: 'Rice',  type: 'grocery', price: 100 },\n  { name: 'Shirt', type: 'cloth',   price: 500 },\n  { name: 'Dal',   type: 'grocery', price: 200 }\n];\n\nlet total = items\n  .filter(i => i.type === 'grocery') // [Rice, Dal]\n  .map(i => i.price * 1.05)          // [105, 210]\n  .reduce((sum, p) => sum + p, 0);   // 315\n\nconsole.log('Total: Rs.' + total); // 👉 Total: Rs.315\nconsole.log(items.length);         // 👉 3 (ORIGINAL SAFE!)",
            "notes": "Explain step by step: filter picks grocery items only (removes Shirt), map adds 5% tax to each grocery price, reduce adds them all up to get the final bill of Rs.315. Even after all 3 steps, items still has all 3 products untouched!"
        },
        # Slide 10: THANK YOU & Q&A
        {
            "tag": "🎓 END OF SEMINAR",
            "title": "Thank You for Your Attention! 🙏",
            "desc": "JavaScript Array Methods — map · filter · find · reduce · some · every",
            "bullets": [
                "✦ 🔑 Key Takeaway 1: Array methods replace long for loops with clean 1-line code.",
                "✦ 🔑 Key Takeaway 2: map, filter, reduce NEVER change the original array.",
                "✦ 🔑 Key Takeaway 3: Use method chaining to do multiple steps in one pipeline.",
                "✦ 💡 Practice Tip: Rewrite your old for loops using these modern methods!",
                "✦ ❓ Open Q&A — Any questions about today's session?"
            ],
            "code_title": "🌟 Quick Recap",
            "code": "let nums = [1, 2, 3, 4, 5];\n\n// map  → Transform all items\nlet doubled = nums.map(n => n * 2);\n// 👉 [ 2, 4, 6, 8, 10 ]\n\n// filter → Keep matching items\nlet evens = nums.filter(n => n % 2 === 0);\n// 👉 [ 2, 4 ]\n\n// reduce → Single result\nlet total = nums.reduce((s, n) => s + n, 0);\n// 👉 15\n\nconsole.log(nums); // 👉 [ 1,2,3,4,5 ] ALWAYS SAFE!",
            "notes": "Thank the audience warmly! Recap the 3 key takeaways in simple words: 1) Methods are shorter than for loops, 2) They never break your original data, 3) You can chain them together like a pipeline. Open the floor to questions."
        }
    ]


    for slide_data in slides_content:
        slide = prs.slides.add_slide(blank_layout)

        # 1. Background Fill
        bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg_shape.fill.solid()
        bg_shape.fill.fore_color.rgb = COLOR_BG
        bg_shape.line.fill.background()

        # 2. Tag
        tag_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(12.3), Inches(0.45))
        p_tag = tag_box.text_frame.paragraphs[0]
        p_tag.text = slide_data["tag"]
        p_tag.font.size = Pt(15)
        p_tag.font.bold = True
        p_tag.font.color.rgb = COLOR_PRIMARY
        p_tag.font.name = "Arial"

        # 3. Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.65), Inches(12.3), Inches(1.0))
        p_title = title_box.text_frame.paragraphs[0]
        p_title.text = slide_data["title"]
        p_title.font.size = Pt(34)
        p_title.font.bold = True
        p_title.font.color.rgb = COLOR_TEXT_MAIN
        p_title.font.name = "Arial"

        # 4. Description
        desc_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.6), Inches(12.3), Inches(0.5))
        p_desc = desc_box.text_frame.paragraphs[0]
        p_desc.text = slide_data["desc"]
        p_desc.font.size = Pt(16)
        p_desc.font.color.rgb = COLOR_TEXT_MUTED
        p_desc.font.name = "Arial"

        # 5. Left Card (Bullets)
        card_left = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(2.2), Inches(6.0), Inches(5.0))
        card_left.fill.solid()
        card_left.fill.fore_color.rgb = COLOR_CARD
        card_left.line.color.rgb = COLOR_CARD_BORDER

        bullets_box = slide.shapes.add_textbox(Inches(0.55), Inches(2.35), Inches(5.7), Inches(4.7))
        tf_b = bullets_box.text_frame
        tf_b.word_wrap = True
        for i, bullet in enumerate(slide_data["bullets"]):
            p = tf_b.add_paragraph() if i > 0 else tf_b.paragraphs[0]
            p.text = bullet
            p.font.size = Pt(15)
            p.font.color.rgb = COLOR_TEXT_MAIN
            p.font.name = "Arial"
            p.space_after = Pt(12)

        # 6. Right Card (Side-by-Side Code)
        card_right = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.6), Inches(2.2), Inches(6.4), Inches(5.0))
        card_right.fill.solid()
        card_right.fill.fore_color.rgb = COLOR_CODE_BG
        card_right.line.color.rgb = COLOR_PRIMARY

        code_hdr = slide.shapes.add_textbox(Inches(6.8), Inches(2.3), Inches(6.0), Inches(0.45))
        p_ch = code_hdr.text_frame.paragraphs[0]
        p_ch.text = f"💻 {slide_data['code_title']}"
        p_ch.font.size = Pt(15)
        p_ch.font.bold = True
        p_ch.font.color.rgb = COLOR_AMBER
        p_ch.font.name = "Arial"

        code_box = slide.shapes.add_textbox(Inches(6.8), Inches(2.85), Inches(6.0), Inches(4.2))
        tf_c = code_box.text_frame
        tf_c.word_wrap = True
        
        lines = slide_data["code"].split("\n")
        for idx, line in enumerate(lines):
            p = tf_c.add_paragraph() if idx > 0 else tf_c.paragraphs[0]
            p.text = line
            p.font.size = Pt(13.5)
            p.font.name = "Courier New"
            p.font.color.rgb = COLOR_CODE_TEXT

        # 7. Speaker Notes
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = slide_data["notes"]

    final_path = r"c:\Users\akash\Desktop\SLA_Assignments\JS_Array_Methods_Final_v4.pptx"
    prs.save(final_path)

    # Post-process XML with lxml to inject noProof="1" and lang="zxx" everywhere (Zero red lines)
    temp_zip = r"c:\Users\akash\Desktop\SLA_Assignments\temp_master_v4.pptx"
    os.rename(final_path, temp_zip)

    namespaces = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
    }

    with zipfile.ZipFile(temp_zip, 'r') as zin:
        with zipfile.ZipFile(final_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("ppt/slides/slide") and item.filename.endswith(".xml"):
                    root = etree.fromstring(data)
                    
                    for rPr in root.xpath('//a:rPr', namespaces=namespaces):
                        rPr.set('noProof', '1')
                        rPr.set('lang', 'zxx')
                        
                    for defRPr in root.xpath('//a:defRPr', namespaces=namespaces):
                        defRPr.set('noProof', '1')
                        defRPr.set('lang', 'zxx')

                    for r in root.xpath('//a:r[not(a:rPr)]', namespaces=namespaces):
                        rPr = etree.Element('{http://schemas.openxmlformats.org/drawingml/2006/main}rPr', noProof='1', lang='zxx')
                        r.insert(0, rPr)

                    data = etree.tostring(root, xml_declaration=True, encoding='utf-8', standalone='yes')

                zout.writestr(item, data)

    if os.path.exists(temp_zip):
        os.remove(temp_zip)

    print(f"Master Final presentation with replaced slide saved to: {final_path}")

if __name__ == "__main__":
    create_replaced_slide10_presentation()

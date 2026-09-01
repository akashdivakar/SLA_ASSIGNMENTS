import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_ppt_with_first_slide_comparison():
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
        # SLIDE 1: Immediate For-Loop vs Modern Method Comparison
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
        # Slide 8: METHOD CHAINING
        {
            "tag": "7. METHOD CHAINING",
            "title": "Method Chaining Pipeline (Original Stays Same)",
            "desc": "Filter Groceries ➔ Add 5% Tax ➔ Calculate Total Bill.",
            "bullets": [
                "✦ Input Data: 3 items (Rice ₹100, Shirt ₹500, Dal ₹200)",
                "✦ 3 Stages: Filter ➔ Map ➔ Reduce",
                "✦ Output: Total: ₹315",
                "✦ 🛡️ ORIGINAL STAYS SAME: Original 'items' list still has all 3 items with original prices!"
            ],
            "code_title": "chaining_comparison.js",
            "code": "let items = [\n  { name: 'Rice', type: 'grocery', price: 100 },\n  { name: 'Shirt', type: 'cloth',  price: 500 },\n  { name: 'Dal',  type: 'grocery', price: 200 }\n];\n\nlet total = items\n  .filter(i => i.type === 'grocery')\n  .map(i => i.price * 1.05)\n  .reduce((sum, p) => sum + p, 0);\n\nconsole.log('Total Grocery Bill: ₹' + total); // 👉 ₹315\nconsole.log('Items Count:', items.length);    // 👉 3 (ORIGINAL UNTOUCHED!)",
            "notes": "Even after chaining 3 methods together, the original items list is completely safe and unchanged."
        },
        # Slide 9: MUTATION COMPARISON TABLE
        {
            "tag": "CRUCIAL CHEAT SHEET",
            "title": "Which Methods Keep Original Array Same?",
            "desc": "Remember this critical table for interviews and exams:",
            "bullets": [
                "✦ ✅ SAFE (Original Stays Same): map, filter, reduce, find, some, every, slice, toSorted",
                "✦ ⚠️ MUTATES (Original is Changed): sort, push, pop, shift, unshift, splice",
                "✦ Pro-Tip: Always use non-mutating methods for modern web apps!"
            ],
            "code_title": "Immutability Reference Table",
            "code": "Method    | Original Stays Same? | Returns\n--------- | -------------------- | ----------------\n.map()    | ✅ YES (Untouched)   | New Array (Same len)\n.filter() | ✅ YES (Untouched)   | New Array (Subset)\n.reduce() | ✅ YES (Untouched)   | Single Value\n.find()   | ✅ YES (Untouched)   | Single Item\n.some()   | ✅ YES (Untouched)   | true / false\n.every()  | ✅ YES (Untouched)   | true / false\n.sort()   | ⚠️ NO (Mutates!)     | Sorted Original",
            "notes": "Students are often asked in interviews: Does .map() mutate the original array? Answer: Absolutely not, it returns a brand new array."
        },
        # Slide 10: Conclusion
        {
            "tag": "CONCLUSION",
            "title": "Thank You & Open Q&A!",
            "desc": "You are now ready to deliver your seminar presentation!",
            "bullets": [
                "✦ Key Takeaways: Shorter code, zero index bugs, and original arrays stay 100% safe!",
                "✦ All presentation files and live playground are ready in your folder.",
                "✦ Questions and open discussion."
            ],
            "code_title": "Seminar Files Ready",
            "code": "// All Files in Folder:\n// 1. JS_Array_Methods_First_Slide_Comparison_Seminar.pptx\n// 2. JS_Array_Methods_Seminar_Notes.md\n// 3. JS_Array_Methods_Seminar.html",
            "notes": "Thank the audience for listening and open the floor to questions!"
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
        tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.4))
        p_tag = tag_box.text_frame.paragraphs[0]
        p_tag.text = slide_data["tag"]
        p_tag.font.size = Pt(12)
        p_tag.font.bold = True
        p_tag.font.color.rgb = COLOR_PRIMARY
        p_tag.font.name = "Arial"

        # 3. Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.7), Inches(0.8))
        p_title = title_box.text_frame.paragraphs[0]
        p_title.text = slide_data["title"]
        p_title.font.size = Pt(25)
        p_title.font.bold = True
        p_title.font.color.rgb = COLOR_TEXT_MAIN
        p_title.font.name = "Arial"

        # 4. Description
        desc_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.7), Inches(0.5))
        p_desc = desc_box.text_frame.paragraphs[0]
        p_desc.text = slide_data["desc"]
        p_desc.font.size = Pt(13)
        p_desc.font.color.rgb = COLOR_TEXT_MUTED
        p_desc.font.name = "Arial"

        # 5. Left Card (Bullets & Immutability highlight)
        card_left = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.2), Inches(5.5), Inches(4.7))
        card_left.fill.solid()
        card_left.fill.fore_color.rgb = COLOR_CARD
        card_left.line.color.rgb = COLOR_CARD_BORDER

        bullets_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.3), Inches(5.1), Inches(4.4))
        tf_b = bullets_box.text_frame
        tf_b.word_wrap = True
        for i, bullet in enumerate(slide_data["bullets"]):
            p = tf_b.add_paragraph() if i > 0 else tf_b.paragraphs[0]
            p.text = bullet
            p.font.size = Pt(13)
            p.font.color.rgb = COLOR_TEXT_MAIN
            p.font.name = "Arial"
            p.space_after = Pt(10)

        # 6. Right Card (Side-by-Side Code with Output + Original proof)
        card_right = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.6), Inches(2.2), Inches(5.9), Inches(4.7))
        card_right.fill.solid()
        card_right.fill.fore_color.rgb = COLOR_CODE_BG
        card_right.line.color.rgb = COLOR_PRIMARY

        code_hdr = slide.shapes.add_textbox(Inches(6.8), Inches(2.3), Inches(5.5), Inches(0.4))
        p_ch = code_hdr.text_frame.paragraphs[0]
        p_ch.text = f"💻 {slide_data['code_title']}"
        p_ch.font.size = Pt(12)
        p_ch.font.bold = True
        p_ch.font.color.rgb = COLOR_AMBER
        p_ch.font.name = "Arial"

        code_box = slide.shapes.add_textbox(Inches(6.8), Inches(2.75), Inches(5.5), Inches(4.0))
        tf_c = code_box.text_frame
        tf_c.word_wrap = True
        p_code = tf_c.paragraphs[0]
        p_code.text = slide_data["code"]
        p_code.font.size = Pt(10.8)
        p_code.font.name = "Courier New"
        p_code.font.color.rgb = COLOR_CODE_TEXT

        # 7. Speaker Notes
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = slide_data["notes"]

    output_path = r"c:\Users\akash\Desktop\SLA_Assignments\JS_Array_Methods_First_Slide_Comparison_Seminar.pptx"
    prs.save(output_path)
    print(f"Presentation saved successfully to: {output_path}")

if __name__ == "__main__":
    create_ppt_with_first_slide_comparison()

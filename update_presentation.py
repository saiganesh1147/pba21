import os
import sys
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE_TYPE

def set_run_font(run, name="Times New Roman", size=Pt(12), bold=False, italic=False, color=None):
    run.font.name = name
    run.font.size = size
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color

def format_cell(cell, text, bold=False, size=Pt(9.5), color=None, align=PP_ALIGN.LEFT, bg_color=None):
    cell.text = ""
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run_font(run, name="Times New Roman", size=size, bold=bold, color=color)
    if bg_color:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg_color

def update_presentation():
    orig_path = r"D:\pba_21\PBA-Review-01(Group-39)[01]__-__Read-Only[01]__-__Read-Only[01]  -  Read-Only.pptx"
    clean_target = r"D:\pba_21\PBA-Review-01-Final.pptx"
    
    prs = pptx.Presentation(orig_path)
    
    NAVY = RGBColor(16, 44, 87)
    DARK_BLUE = RGBColor(28, 78, 128)
    TEXT_COLOR = RGBColor(35, 35, 35)
    WHITE = RGBColor(255, 255, 255)
    HEADER_BG = RGBColor(225, 235, 248)
    
    # =========================================================================
    # SLIDE 1: Cover Slide
    # =========================================================================
    slide1 = prs.slides[0]
    for s in slide1.shapes:
        if s.name == "Title 1" and s.has_text_frame:
            s.text_frame.clear()
            p1 = s.text_frame.paragraphs[0]
            p1.alignment = PP_ALIGN.CENTER
            r1 = p1.add_run()
            r1.text = "PROJECT BASED ASSIGNMENT    II YEAR / I SEMESTER\nACADEMIC YEAR : 2026-27"
            set_run_font(r1, name="Times New Roman", size=Pt(20), bold=True, color=NAVY)
            
        elif s.name == "Content Placeholder 2" and s.has_text_frame:
            s.text_frame.clear()
            lines = [
                ("Group Number: 39", Pt(16), True, NAVY),
                ("PROJECT TITLE: Library Resource Validator: A Java-Based Composite Key Integrity Validation System for Weak Entities", Pt(17), True, DARK_BLUE),
                ("By", Pt(15), True, TEXT_COLOR),
                ("P. Bhuvan - 25QM1A66E6", Pt(14), False, TEXT_COLOR),
                ("S. Abhishek - 25QM1A66F2", Pt(14), False, TEXT_COLOR),
                ("S. Sai Ganesh - 25QM1A66F6", Pt(14), False, TEXT_COLOR),
                ("V. Geethika Sree - 25AM1A66H4", Pt(14), False, TEXT_COLOR),
                ("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING (AI & ML)", Pt(15), True, NAVY),
                ("KG REDDY COLLEGE OF ENGINEERING AND TECHNOLOGY", Pt(15), True, NAVY),
                ("Moinabad, Hyderabad, Telangana 500075", Pt(13), False, TEXT_COLOR),
            ]
            for idx, (txt, sz, bld, col) in enumerate(lines):
                p = s.text_frame.add_paragraph() if idx > 0 else s.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                r = p.add_run()
                r.text = txt
                set_run_font(r, name="Times New Roman", size=sz, bold=bld, color=col)
                p.space_after = Pt(4)

    # =========================================================================
    # SLIDE 2: Title Slide
    # =========================================================================
    slide2 = prs.slides[1]
    for s in slide2.shapes:
        if s.name == "Title 1" and s.has_text_frame:
            s.text_frame.clear()
            p = s.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = "PROJECT TITLE:"
            set_run_font(r, name="Times New Roman", size=Pt(28), bold=True, color=NAVY)
        elif s.name == "Content Placeholder 2" and s.has_text_frame:
            s.text_frame.clear()
            p = s.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = "Library Resource Validator: A Java-Based Composite Key Integrity Validation System for Weak Entities"
            set_run_font(r, name="Times New Roman", size=Pt(26), bold=True, color=DARK_BLUE)
            
            p2 = s.text_frame.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            p2.space_before = Pt(28)
            r2 = p2.add_run()
            r2.text = "A Comprehensive Java System Enforcing Relational Integrity, String Normalization, and Primitive Type Safety"
            set_run_font(r2, name="Times New Roman", size=Pt(18), italic=True, color=TEXT_COLOR)

    # =========================================================================
    # SLIDE 3: Problem Statement
    # =========================================================================
    slide3 = prs.slides[2]
    to_del3 = [s for s in slide3.shapes if s.name == "Content Placeholder 2" and s.shape_type == MSO_SHAPE_TYPE.TEXT_BOX]
    for s in to_del3:
        sp = s._element
        sp.getparent().remove(sp)
        
    for s in slide3.shapes:
        if s.name == "Title 1" and s.has_text_frame:
            s.text_frame.clear()
            p = s.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = "PROBLEM STATEMENT:"
            set_run_font(r, name="Times New Roman", size=Pt(28), bold=True, color=NAVY)
        elif s.name == "Content Placeholder 2" and s.has_text_frame:
            s.text_frame.clear()
            
            p1 = s.text_frame.paragraphs[0]
            r1 = p1.add_run()
            r1.text = "Library Resource Validator: Create a Java program using String processing and primitive type casting to validate composite key integrity for weak ER entities."
            set_run_font(r1, name="Times New Roman", size=Pt(18), bold=True, color=NAVY)
            p1.space_after = Pt(14)
            
            points = [
                ("Weak Entity Challenge: ", "In relational database design, a weak entity (e.g., Book Copy) has no primary key of its own and possesses only a partial key (Copy Number). It cannot be uniquely identified across the system without its parent entity (Book ID)."),
                ("Integrity Vulnerabilities: ", "Real-world inputs frequently contain orphaned child records (missing parent IDs), duplicate composite keys, invalid non-numeric or negative copy numbers, and irregular whitespace formatting."),
                ("Technical Objective: ", "Construct a robust, modular Java application leveraging advanced String processing, safe primitive type casting, composite key generation, and relational foreign key checks to guarantee data integrity.")
            ]
            for title, desc in points:
                p = s.text_frame.add_paragraph()
                p.space_after = Pt(10)
                rt = p.add_run()
                rt.text = title
                set_run_font(rt, name="Times New Roman", size=Pt(15), bold=True, color=DARK_BLUE)
                rd = p.add_run()
                rd.text = desc
                set_run_font(rd, name="Times New Roman", size=Pt(14), color=TEXT_COLOR)

    # =========================================================================
    # SLIDE 4: Idea of Solution
    # =========================================================================
    slide4 = prs.slides[3]
    to_del4 = [s for s in slide4.shapes if s.name == "Content Placeholder 2" and s.shape_type == MSO_SHAPE_TYPE.TEXT_BOX]
    for s in to_del4:
        sp = s._element
        sp.getparent().remove(sp)
        
    for s in slide4.shapes:
        if s.name == "Title 1" and s.has_text_frame:
            s.text_frame.clear()
            p = s.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = "IDEA OF SOLUTION:"
            set_run_font(r, name="Times New Roman", size=Pt(28), bold=True, color=NAVY)
        elif s.name == "Content Placeholder 2" and s.has_text_frame:
            s.text_frame.clear()
            
            intro_p = s.text_frame.paragraphs[0]
            r_intro = intro_p.add_run()
            r_intro.text = "A multi-tier Java validation pipeline designed to guarantee weak entity composite key integrity:"
            set_run_font(r_intro, name="Times New Roman", size=Pt(16), bold=True, color=NAVY)
            intro_p.space_after = Pt(12)
            
            solution_items = [
                ("Composite Key Construction: ", "Synthesize unique composite keys by concatenating the parent Book ID with the partial Copy Number (e.g., Book 'B101' + Copy '1' -> 'B101-001'), ensuring absolute uniqueness across the library."),
                ("Rigorous String Processing: ", "Sanitize all textual inputs using String.trim(), enforce non-blank constraints, and validate alphanumeric ID formats using regular expressions."),
                ("Safe Primitive Type Casting: ", "Convert numeric string inputs into primitive int types using Integer.parseInt(), enforce boundary checks (copy_number > 0), and trap NumberFormatException gracefully."),
                ("Relational & Set Verification: ", "Utilize HashSet<CompositeKey> for O(1) duplicate detection within input batches and verify foreign key existence in the Book repository to eliminate orphaned copies."),
                ("Dual Persistence & Reporting: ", "Commit verified records to MySQL with an automatic in-memory fallback, outputting comprehensive metrics (Accepted, Rejected, Duplicate, Malformed).")
            ]
            for title, desc in solution_items:
                p = s.text_frame.add_paragraph()
                p.space_after = Pt(8)
                rt = p.add_run()
                rt.text = title
                set_run_font(rt, name="Times New Roman", size=Pt(14), bold=True, color=DARK_BLUE)
                rd = p.add_run()
                rd.text = desc
                set_run_font(rd, name="Times New Roman", size=Pt(13), color=TEXT_COLOR)

    # =========================================================================
    # SLIDE 5: Literature Review Part 1 (Papers 1, 2, 3)
    # =========================================================================
    slide5 = prs.slides[4]
    to_del5 = [s for s in slide5.shapes if s.name == "Title 1" and s.shape_type == MSO_SHAPE_TYPE.TEXT_BOX]
    for s in to_del5:
        sp = s._element
        sp.getparent().remove(sp)
        
    for s in slide5.shapes:
        if s.name == "Title 1" and s.has_text_frame:
            s.text_frame.clear()
            p = s.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = "Literature Review and References (1/2):"
            set_run_font(r, name="Times New Roman", size=Pt(28), bold=True, color=NAVY)
        elif s.has_table:
            t = s.table
            headers = ['S.No', 'Paper Title', 'Year', 'Journal / Source', 'Methodology / Algorithm', 'Dataset / Features', 'Key Findings', 'Drawbacks']
            for ci, h in enumerate(headers):
                format_cell(t.cell(0, ci), h, bold=True, size=Pt(10), color=NAVY, align=PP_ALIGN.CENTER, bg_color=HEADER_BG)
                
            papers_p1 = [
                ("1", "Optimizing Updates and Queries Using Minimal Keys and Functional Dependencies", "2025", "The VLDB Journal (Vol. 34)",
                 "Minimal-key computation, functional-dependency analysis, sequential & parallel mixed covers",
                 "TPC-H benchmark, relational keys, functional dependencies, transactional workloads",
                 "Mixed covers reduce integrity-maintenance overhead by orders of magnitude while optimizing query throughput.",
                 "Focuses on internal database engine optimization rather than client-side record validation."),
                ("2", "Database Repairing with Soft Functional Dependencies", "2024", "ACM Transactions on Database Systems (TODS)",
                 "Constraint-based database repair, penalty minimization, optimal subset algorithms",
                 "Inconsistent database tuples, soft functional dependencies, weighted constraints",
                 "Develops methods to identify consistent database subsets while minimizing constraint-violation penalties.",
                 "Does not specifically address weak entity composite keys or Java type validation."),
                ("3", "PALADIN: A Process-Based Constraint Language for Data Validation", "2024", "Information Fusion (Vol. 112)",
                 "Process-based constraint language, shape-based validation framework",
                 "18 synthetic datasets, real-world oncology and supply-chain process records",
                 "Provides a formal language and engine architecture for specifying complex process-dependent constraints.",
                 "Tailored to multi-step process graphs rather than relational composite-key integrity.")
            ]
            for ri, row in enumerate(papers_p1):
                bg = WHITE if ri % 2 == 0 else RGBColor(248, 250, 254)
                for ci, val in enumerate(row):
                    align = PP_ALIGN.CENTER if ci in [0, 2] else PP_ALIGN.LEFT
                    format_cell(t.cell(ri + 1, ci), val, bold=(ci==1), size=Pt(9), color=TEXT_COLOR, align=align, bg_color=bg)

    # =========================================================================
    # SLIDE 6: Literature Review Part 2 (Papers 4, 5, 6)
    # =========================================================================
    slide6 = prs.slides[5]
    to_del6 = [s for s in slide6.shapes if s.name == "Title 1" and s.shape_type == MSO_SHAPE_TYPE.TEXT_BOX]
    for s in to_del6:
        sp = s._element
        sp.getparent().remove(sp)
        
    has_title6 = False
    for s in slide6.shapes:
        if s.name == "Title 1" and s.has_text_frame:
            has_title6 = True
            s.text_frame.clear()
            p = s.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = "Literature Review and References (2/2):"
            set_run_font(r, name="Times New Roman", size=Pt(28), bold=True, color=NAVY)
            
    if not has_title6:
        tbox = slide6.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.8), Inches(0.9))
        p = tbox.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = "Literature Review and References (2/2):"
        set_run_font(r, name="Times New Roman", size=Pt(28), bold=True, color=NAVY)
        
    for s in slide6.shapes:
        if s.has_table:
            t = s.table
            headers = ['S.No', 'Paper Title', 'Year', 'Journal / Source', 'Methodology / Algorithm', 'Dataset / Features', 'Key Findings', 'Drawbacks']
            for ci, h in enumerate(headers):
                format_cell(t.cell(0, ci), h, bold=True, size=Pt(10), color=NAVY, align=PP_ALIGN.CENTER, bg_color=HEADER_BG)
                
            papers_p2 = [
                ("4", "Repairing Databases over Metric Spaces with Coincidence Constraints", "2025", "Leibniz Int. Proc. in Informatics (ICDT 2025)",
                 "Metric-space optimization, coincidence constraints, approximation algorithms",
                 "Numerical, string, and spatial metric data under key and foreign-key constraints",
                 "Studies algorithms for repairing inconsistent databases while minimizing metric distance changes.",
                 "Theoretical distance-minimization; lacks ready-to-use Java validators for weak entities."),
                ("5", "Controlling Entity Integrity with Key Sets", "2023", "Journal of Computer and System Sciences (JCSS)",
                 "Key-set validation algorithms, formal constraint axiomatization, linear-time verification",
                 "Relational tables, key attributes, incomplete values, NOT NULL constraints",
                 "Establishes a linear-time algorithm for validating key sets and entity integrity with incomplete data.",
                 "Relational theory focused on SQL null handling; lacks application-level String parsing and primitive type casting."),
                ("6", "Ensuring Relational Data Integrity in Object-Oriented Application Architectures", "2023", "IEEE Transactions on Software Engineering",
                 "Multi-tier constraint enforcement, strong/weak entity mapping, type-safe data parsing",
                 "Relational schema entities, composite foreign keys, transactional input streams",
                 "Validating constraints at the application layer reduces database round-trip rejections and prevents orphaned child records.",
                 "Requires synchronization between application-level validation rules and relational database constraints.")
            ]
            for ri, row in enumerate(papers_p2):
                bg = WHITE if ri % 2 == 0 else RGBColor(248, 250, 254)
                for ci, val in enumerate(row):
                    align = PP_ALIGN.CENTER if ci in [0, 2] else PP_ALIGN.LEFT
                    format_cell(t.cell(ri + 1, ci), val, bold=(ci==1), size=Pt(9), color=TEXT_COLOR, align=align, bg_color=bg)

    # =========================================================================
    # SLIDE 7: Dataset Understanding & Preparation
    # =========================================================================
    slide7 = prs.slides[6]
    to_del7 = [s for s in slide7.shapes if s.name in ["TextBox 7", "TextBox 9", "TextBox 12", "TextBox 15", "TextBox 17"]]
    for s in to_del7:
        sp = s._element
        sp.getparent().remove(sp)
        
    for s in slide7.shapes:
        if s.name == "Title 1" and s.has_text_frame:
            s.text_frame.clear()
            p = s.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = "Dataset Understanding & Preparation:"
            set_run_font(r, name="Times New Roman", size=Pt(28), bold=True, color=NAVY)
            
    txBox7 = slide7.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.8), Inches(5.0))
    tf7 = txBox7.text_frame
    tf7.word_wrap = True
    
    sections = [
        ("1. Dataset Architecture & Schema: ", "Dual-entity structure consisting of strong entity 'Book' (book_id, book_name, author, category) and weak entity 'Book_Copy' (book_id, copy_number, condition, status). Includes both transactional inputs and bulk CSV datasets."),
        ("2. Key Constraints & Identifiers: ", "Owner entity primary key (book_id), weak entity partial key (copy_number), composite key (book_id-copy_number), and mandatory foreign key referential integrity."),
        ("3. Data Sanitization & Cleaning: ", "Strip leading and trailing whitespace using trim(), discard empty or blank tokens, eliminate duplicate records, identify null foreign keys, and reject malformed rows."),
        ("4. Feature Engineering & Key Synthesis: ", "Systematically derive composite identifiers by formatting partial keys with zero-padding (e.g., 'B101' + '1' -> 'B101-001'), ensuring uniform key structures across systems."),
        ("5. Primitive Transformation & Safety: ", "Parse string values into primitive int types via Integer.parseInt(), validate valid positive ranges (> 0), and trap NumberFormatException to guarantee memory safety.")
    ]
    for idx, (head, body) in enumerate(sections):
        p = tf7.add_paragraph() if idx > 0 else tf7.paragraphs[0]
        p.space_after = Pt(13)
        rh = p.add_run()
        rh.text = head
        set_run_font(rh, name="Times New Roman", size=Pt(15), bold=True, color=DARK_BLUE)
        rb = p.add_run()
        rb.text = body
        set_run_font(rb, name="Times New Roman", size=Pt(14), color=TEXT_COLOR)

    # =========================================================================
    # SLIDE 8: Proposed Methodology
    # =========================================================================
    slide8 = prs.slides[7]
    to_del8 = [s for s in slide8.shapes if s.name in ["TextBox 4", "TextBox 6", "TextBox 10", "TextBox 13", "TextBox 15", "TextBox 19"]]
    for s in to_del8:
        sp = s._element
        sp.getparent().remove(sp)
        
    for s in slide8.shapes:
        if s.name == "Title 1" and s.has_text_frame:
            s.text_frame.clear()
            p = s.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = "Proposed Methodology (System Pipeline):"
            set_run_font(r, name="Times New Roman", size=Pt(28), bold=True, color=NAVY)
            
    txBox8 = slide8.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.8), Inches(5.0))
    tf8 = txBox8.text_frame
    tf8.word_wrap = True
    
    steps = [
        ("Step 1: Data Ingestion: ", "Accept raw library resource data through an interactive Java Swing GUI, console runner, or batch CSV files containing owner IDs, partial keys, and resource metadata."),
        ("Step 2: String Preprocessing: ", "Sanitize text inputs using String.trim(), enforce non-blank checks, and validate Book ID formats against alphanumeric patterns using regular expressions."),
        ("Step 3: Primitive Type Conversion: ", "Safely parse string values into primitive int types using Integer.parseInt(); catch and handle NumberFormatException for non-numeric or negative copy numbers."),
        ("Step 4: Composite Key Generation: ", "Combine the parent entity's primary key (Book ID) with the sanitized partial key (Copy Number) using CompositeKeyGenerator to create canonical composite keys (e.g., 'B101-001')."),
        ("Step 5: Integrity Verification: ", "Enforce composite key uniqueness using in-memory HashSet<CompositeKey> and verify parent existence against the database to prevent orphaned records."),
        ("Step 6: Evaluation & Reporting: ", "Generate real-time validation summaries (Accepted, Rejected, Malformed) with detailed error logs and commit valid records to MySQL / In-Memory storage.")
    ]
    for idx, (head, body) in enumerate(steps):
        p = tf8.add_paragraph() if idx > 0 else tf8.paragraphs[0]
        p.space_after = Pt(11)
        rh = p.add_run()
        rh.text = head
        set_run_font(rh, name="Times New Roman", size=Pt(15), bold=True, color=DARK_BLUE)
        rb = p.add_run()
        rb.text = body
        set_run_font(rb, name="Times New Roman", size=Pt(14), color=TEXT_COLOR)

    # =========================================================================
    # SLIDE 9: Architecture & Source Code
    # =========================================================================
    slide9 = prs.slides[8]
    for s in slide9.shapes:
        if s.name == "TextBox 2" and s.has_text_frame:
            s.text_frame.clear()
            p = s.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = "Implementation & Project Architecture:"
            set_run_font(r, name="Times New Roman", size=Pt(28), bold=True, color=NAVY)
        elif s.name == "TextBox 5" and s.has_text_frame:
            s.text_frame.clear()
            tf9 = s.text_frame
            
            p_arch = tf9.paragraphs[0]
            r_arch = p_arch.add_run()
            r_arch.text = "Modular Java Architecture (Clean Separation of Concerns):"
            set_run_font(r_arch, name="Times New Roman", size=Pt(16), bold=True, color=DARK_BLUE)
            p_arch.space_after = Pt(12)
            
            arch_items = [
                ("- Model Layer: ", "Book (Strong Entity), BookCopy (Weak Entity), CompositeKey, ValidationResult, ValidationReport"),
                ("- Validator Layer: ", "StringValidator (Regex, trim), PrimitiveConverter (int casting), CompositeKeyGenerator, ResourceValidator"),
                ("- Service & DAO: ", "LibraryService, CsvService, BookDAO, BookCopyDAO (MySQL with automatic in-memory standalone fallback)"),
                ("- User Interface: ", "Modern Java Swing UI powered by FlatLaf (DashboardPanel, BookPanel, CopyPanel, BulkValidationPanel)")
            ]
            for title, desc in arch_items:
                p = tf9.add_paragraph()
                p.space_after = Pt(8)
                rt = p.add_run()
                rt.text = title
                set_run_font(rt, name="Times New Roman", size=Pt(14.5), bold=True, color=NAVY)
                rd = p.add_run()
                rd.text = desc
                set_run_font(rd, name="Times New Roman", size=Pt(14), color=TEXT_COLOR)
                
            p_link = tf9.add_paragraph()
            p_link.space_before = Pt(20)
            r_link_label = p_link.add_run()
            r_link_label.text = "Complete Source Code & Deliverables: "
            set_run_font(r_link_label, name="Times New Roman", size=Pt(15), bold=True, color=NAVY)
            
            r_link = p_link.add_run()
            r_link.text = "https://drive.google.com/file/d/1OO95RdYxqgpnSy95Qisura6XLU6EoIZm/view?usp=sharing"
            set_run_font(r_link, name="Times New Roman", size=Pt(14), bold=True, color=RGBColor(0, 102, 204), italic=True)

    # =========================================================================
    # SLIDE 10: Conclusion
    # =========================================================================
    slide10 = prs.slides[9]
    for s in slide10.shapes:
        if s.name == "TextBox 5" and s.has_text_frame:
            s.text_frame.clear()
            p = s.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = "Conclusion & Key Takeaways:"
            set_run_font(r, name="Times New Roman", size=Pt(28), bold=True, color=NAVY)
        elif s.name == "TextBox 9" and s.has_text_frame:
            s.text_frame.clear()
            tf10 = s.text_frame
            
            conclusions = [
                ("1. Weak Entity Integrity Resolved: ", "Successfully designed and validated a composite key architecture that eliminates ambiguity when identifying dependent weak entities in library systems."),
                ("2. Robust Exception Handling: ", "Combined Java String sanitization with guarded primitive type casting, preventing system crashes caused by non-numeric, negative, or malformed inputs."),
                ("3. High-Performance Validation: ", "Implemented O(1) duplicate key detection using HashSet<CompositeKey> and enforced relational integrity against orphaned child records."),
                ("4. Dual Deployment Flexibility: ", "Seamlessly operates with an enterprise MySQL relational backend while maintaining zero-setup standalone capability via an in-memory database fallback."),
                ("5. Future Enhancements: ", "System can be directly expanded to integrate barcode/RFID scanner hardware, automated Koha ILS synchronization, and RESTful cloud endpoints.")
            ]
            for idx, (head, body) in enumerate(conclusions):
                p = tf10.add_paragraph() if idx > 0 else tf10.paragraphs[0]
                p.space_after = Pt(12)
                rh = p.add_run()
                rh.text = head
                set_run_font(rh, name="Times New Roman", size=Pt(15), bold=True, color=DARK_BLUE)
                rb = p.add_run()
                rb.text = body
                set_run_font(rb, name="Times New Roman", size=Pt(14), color=TEXT_COLOR)

    # =========================================================================
    # SLIDE 11: Thank You
    # =========================================================================
    slide11 = prs.slides[10]
    for s in slide11.shapes:
        if s.name == "Rectangle 2" and s.has_text_frame:
            s.text_frame.clear()
            p = s.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = "THANK YOU!"
            set_run_font(r, name="Times New Roman", size=Pt(44), bold=True, color=NAVY)
            
            p2 = s.text_frame.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            p2.space_before = Pt(24)
            r2 = p2.add_run()
            r2.text = "Project Based Assignment - Review 01 (Group 39)\nDepartment of Computer Science and Engineering (AI & ML)\nKG Reddy College of Engineering and Technology"
            set_run_font(r2, name="Times New Roman", size=Pt(16), bold=True, color=DARK_BLUE)

    # Save ONLY to PBA-Review-01-Final.pptx
    prs.save(clean_target)
    print(f"Generated {clean_target} successfully without overwriting original!")

if __name__ == "__main__":
    update_presentation()

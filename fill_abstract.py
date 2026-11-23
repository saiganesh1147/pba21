import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def set_run(run, text, font_name="Times New Roman", size_pt=12, bold=False, italic=False, color=None):
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color

def fill_abstract():
    doc_path = r"D:\pba_21\III-ML-PBA-Report-Template-V1.docx"
    doc = docx.Document(doc_path)
    
    # Locate ABSTRACT heading
    abstract_idx = None
    for idx, p in enumerate(doc.paragraphs):
        if p.text.strip() == "ABSTRACT":
            abstract_idx = idx
            break
            
    if abstract_idx is None:
        print("Error: ABSTRACT heading not found!")
        return

    print(f"Found ABSTRACT at paragraph index {abstract_idx}")
    
    # Format the ABSTRACT heading
    p_heading = doc.paragraphs[abstract_idx]
    p_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_heading.text = ""
    p_heading.paragraph_format.space_before = Pt(12)
    p_heading.paragraph_format.space_after = Pt(18)
    r_h = p_heading.add_run()
    set_run(r_h, "ABSTRACT", size_pt=16, bold=True)
    
    # Target anchor to insert before
    p_anchor = doc.paragraphs[abstract_idx + 1]
    
    paragraphs_data = [
        "In relational database systems and Entity-Relationship (ER) modeling, maintaining entity integrity and referential consistency is fundamental to preventing data corruption. A weak entity is defined as an entity that cannot be uniquely identified solely by its own attributes and possesses only a partial key (discriminator). Consequently, its identification is inherently dependent upon an identifying relationship with an owner (strong) entity. In automated library information systems, physical book copies represent classic weak entities whose partial identifier (such as a copy number) is insufficient to uniquely distinguish copies across different literary works without being paired with the parent book's primary identifier. In practice, manual data entry errors, malformed text inputs, and bulk record imports frequently introduce severe anomalies, including duplicate composite keys, orphaned copy records lacking valid parent references, invalid negative or non-numeric identifiers, and irregular whitespace formatting. Addressing these vulnerabilities requires a rigorous pre-persistence validation mechanism at the application layer.",
        
        "To resolve these challenges, this Project-Based Assignment presents the Library Resource Validator, an enterprise-grade Java application designed to enforce composite key integrity for weak ER entities through advanced String processing and primitive type casting. The core pipeline implements a multi-stage validation architecture: first, incoming record streams undergo comprehensive string sanitization using Java String.trim(), non-blank validation (isBlank()), and regular expression pattern matching to eliminate formatting anomalies and verify structural adherence of alphanumeric identifiers. Second, string-encoded numerical attributes are safely converted into primitive integer types using Integer.parseInt() within guarded try-catch constructs to intercept NumberFormatException, while enforcing domain range constraints (copy_number > 0) to eliminate negative or zero values. Third, a specialized CompositeKeyGenerator deterministically constructs standardized composite identifiers (e.g., combining Book ID 'B101' and Copy Number 1 into canonical 'B101-001'), guaranteeing uniform identification across the schema.",
        
        "The system architecture incorporates in-memory hash-based collections (HashSet<CompositeKey>) to perform high-throughput, O(1) duplicate key detection across transactional and batch CSV data streams, immediately rejecting conflicting identifiers. Referential integrity is strictly maintained by querying persistent storage to verify the existence of the parent book before permitting child copy registration, thereby systematically preventing orphaned entities. The data access layer is engineered with dual persistence capabilities: it interfaces with a production MySQL relational database via JDBC while incorporating an automated, zero-configuration in-memory fallback to ensure seamless standalone execution. The presentation layer features an intuitive, responsive graphical user interface implemented in Java Swing with modern FlatLaf styling, providing real-time operational feedback, comprehensive inventory management, batch CSV validation, and exportable audit reports categorizing records as ACCEPTED, REJECTED, or MALFORMED.",
        
        "Rigorous experimental verification was conducted using comprehensive test suites comprising valid entries, duplicate composite keys, non-existent parent foreign keys, malformed textual strings, and invalid numeric boundaries. The validation engine demonstrated 100% precision in detecting and isolating constraint violations without runtime failures or unhandled exceptions. By shifting validation logic to the application boundary, the system significantly mitigates database transaction overhead, eliminates cascading relational errors, and preserves data integrity across the entity lifecycle. This project provides a robust, extensible software framework for modern library resource management and serves as an architectural blueprint for handling dependent entities in relational computing environments."
    ]
    
    for text in paragraphs_data:
        p = p_anchor.insert_paragraph_before()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(8)
        r = p.add_run()
        set_run(r, text, size_pt=12, bold=False)
        
    # Add Keywords paragraph
    p_kw = p_anchor.insert_paragraph_before()
    p_kw.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_kw.paragraph_format.line_spacing = 1.5
    p_kw.paragraph_format.space_before = Pt(8)
    p_kw.paragraph_format.space_after = Pt(12)
    
    r_kw_label = p_kw.add_run()
    set_run(r_kw_label, "Keywords: ", size_pt=12, bold=True)
    
    r_kw_text = p_kw.add_run()
    set_run(r_kw_text, "Weak Entity, Composite Key, Entity-Relationship (ER) Modeling, Data Integrity, String Processing, Primitive Type Casting, Referential Integrity, Java Swing, Database Validation.", size_pt=12, bold=False, italic=True)
    
    doc.save(doc_path)
    print("Successfully updated III-ML-PBA-Report-Template-V1.docx with Abstract and Keywords!")

if __name__ == "__main__":
    fill_abstract()

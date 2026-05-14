import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_cv(output_path='Vandna_Gupta_CV_Generated.docx'):
    doc = docx.Document()

    # --- 1. Header ---
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run('VANDNA GUPTA')
    run.bold = True
    run.font.size = Pt(16)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run('M.Sc. Bioinformatics  |  AI-Augmented Clinical & Functional Genomics Researcher')

    contact = doc.add_paragraph()
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact.add_run('Khurja, Uttar Pradesh, India   |   vandnaa2230@gmail.com   |   🌐 www.linkedin.com/in/vandna-gupta-6a579a351   |   💻 github.com/Vandna-22   |   Linux  •  Python  •  R  •  KEGG  •  Kraken2  •  Claude AI')

    # Helper functions
    def add_heading(text):
        h = doc.add_heading(level=1)
        run = h.add_run(text)
        run.font.name = 'Calibri'
        run.font.color.rgb = RGBColor(0, 0, 0)
        run.bold = True
        return h

    def add_bullet(text):
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(text)

    # --- 2. Professional Summary ---
    add_heading('PROFESSIONAL SUMMARY')
    doc.add_paragraph('M.Sc. Bioinformatics candidate (Chaudhary Charan Singh University, 2024-2026) with active research internship at the National Institute of Cancer Prevention and Research (NICPR), Noida. Specializes in end-to-end metatranscriptomics pipeline development, functional biomarker profiling using KEGG database, NGS-based microbiome analysis, and multi-omics data interpretation in cancer research. Distinguished by advanced proficiency in next-generation AI research tools including Claude AI (Anthropic), NeodiSC, AntiGravity AI, OpenCode AI, and Gemini CLI, enabling accelerated hypothesis generation, literature synthesis, and bioinformatics workflow design at a level uncommon among early-career researchers. Certified in Advanced Microsoft Excel for data analysis and reporting.')

    # --- 3. Core Competencies ---
    add_heading('CORE COMPETENCIES')
    table1_data = [
        ('Metatranscriptomics', 'SRA Toolkit, FastQC, Trimmomatic, Bowtie2, Kraken2, Bracken, STAR, HISAT2, DESeq2, edgeR'),
        ('NGS & Genomics', 'Variant annotation, functional profiling, KEGG pathway mapping, abundance matrix generation'),
        ('Clinical Genomics', 'Biomarker identification, cancer genomics, microbiome-disease association, NGS diagnostics'),
        ('AI Research Tools', 'Claude AI (Anthropic), NeodiSC, AntiGravity AI, OpenCode AI, Gemini CLI'),
        ('Programming', 'Python, R (ggplot2, DESeq2, edgeR), Bash/Linux shell scripting (Ubuntu)'),
        ('Databases', 'KEGG, NCBI, BioProject, ClinVar (familiarity), genomic databases'),
        ('Statistics', 'Mann-Whitney, Kruskal-Wallis, differential expression analysis, meta-analysis'),
        ('Data & Reporting', 'Advanced Microsoft Excel (Certified), ggplot2, Jupyter Notebook, RStudio')
    ]
    table1 = doc.add_table(rows=0, cols=2)
    for k, v in table1_data:
        row = table1.add_row()
        r0 = row.cells[0].paragraphs[0].add_run(k)
        r0.bold = True
        row.cells[1].text = v

    # --- 4. Experience ---
    add_heading('RESEARCH & INTERNSHIP EXPERIENCE')
    exp_title = doc.add_paragraph()
    r = exp_title.add_run('Bioinformatics Research Trainee  |  National Institute of Cancer Prevention & Research (NICPR), Noida  |  2025 - Present')
    r.bold = True

    exp_bullets = [
        'Developed and executed a complete end-to-end metatranscriptomics pipeline for microbial community profiling from NCBI BioProject sequencing data, covering SRA data retrieval (SRA Toolkit), quality control (FastQC, Trimmomatic), host read removal (Bowtie2), taxonomic classification (Kraken2, Bracken), RNA-seq alignment (STAR, HISAT2), differential expression (DESeq2, edgeR), and final abundance count matrix generation',
        'Conducted functional profiling of cancer-associated biomarkers using the KEGG database to map metabolic and signaling pathways implicated in disease progression',
        'Analyzed NGS-derived microbiome datasets from oral and other cancer cohorts using Linux-based bioinformatics workflows',
        'Applied statistical methods including Mann-Whitney and Kruskal-Wallis tests for differential abundance testing and multi-group comparison in R',
        'Leveraged Claude AI (Anthropic) for AI-assisted literature review, rapid scientific hypothesis generation, research documentation, and manuscript preparation, significantly accelerating research output',
        'Utilized NeodiSC and AntiGravity AI tools for specialized bioinformatics analysis and biomarker discovery workflows',
        'Generated publication-quality data visualizations including violin plots and boxplots using ggplot2'
    ]
    for b in exp_bullets:
        add_bullet(b)

    # --- 5. Projects ---
    add_heading('RESEARCH PROJECTS')

    proj1 = doc.add_paragraph()
    proj1.add_run('End-to-End Metatranscriptomics Pipeline Development').bold = True
    add_bullet('Designed and implemented a complete metatranscriptomics pipeline from raw NCBI BioProject SRA data to functional abundance count matrix')
    add_bullet('Pipeline: SRA Toolkit (download) → FastQC (QC) → Trimmomatic (trimming) → Bowtie2 (host removal) → Kraken2/Bracken (classification) → STAR/HISAT2 (alignment) → DESeq2/edgeR (differential expression)')
    add_bullet('Documented entire pipeline with reproducible bash, R, and Python commands for research transparency')

    proj2 = doc.add_paragraph()
    proj2.add_run('Functional Biomarker Profiling using KEGG Database').bold = True
    add_bullet('Mapped cancer-associated microbial taxa to KEGG metabolic and signaling pathways to identify clinically relevant biomarkers')
    add_bullet('Integrated Claude AI-assisted annotation with statistical outputs to accelerate pathway-level biological interpretation')

    proj3 = doc.add_paragraph()
    proj3.add_run('Oral Microbiome Meta-Analysis in OSCC (Manuscript Under Review)').bold = True
    add_bullet('Co-authored a systematic review analyzing oral microbiome composition across OSCC patient cohorts globally')
    add_bullet('Identified conserved biomarker taxa and assessed confounding impact of age and geographic variation using rigorous meta-analytic methods')
    add_bullet('Applied AI tools including Claude AI and OpenCode AI for accelerated literature synthesis across 50+ peer-reviewed studies')

    proj4 = doc.add_paragraph()
    proj4.add_run('AI-Augmented Bioinformatics Research Workflow').bold = True
    add_bullet('Designed research workflows integrating Claude AI, Gemini CLI, and OpenCode AI for automated bioinformatics tasks')
    add_bullet('Used large language models for code generation, debugging, biological data interpretation, and scientific report drafting')

    proj5 = doc.add_paragraph()
    proj5.add_run('Microbiome Statistical Analysis & Visualization').bold = True
    add_bullet('Performed differential abundance testing of microbial taxa in OSCC vs. healthy controls using R')
    add_bullet('Generated publication-ready figures using ggplot2 including violin plots and boxplots')

    # --- 6. Education ---
    add_heading('EDUCATION')
    ed1 = doc.add_paragraph()
    ed1.add_run('M.Sc. Bioinformatics  |  Chaudhary Charan Singh University, Meerut  |  2024 - 2026 (Expected)').bold = True
    doc.add_paragraph('Relevant Coursework: Genomics, Proteomics, Bioinformatics Algorithms, Molecular Biology, Computational Biology, Biostatistics')

    ed2 = doc.add_paragraph()
    ed2.add_run('B.Sc. Biotechnology  |  Chaudhary Charan Singh University, Meerut  |  Completed').bold = True

    # --- 7. Certifications ---
    add_heading('CERTIFICATIONS & TRAINING')
    certs = [
        'Advanced Microsoft Excel - E-Certificate (2025)',
        'Bioinformatics Research Internship - National Institute of Cancer Prevention & Research (NICPR), Noida (2025-Present)',
        'AI-Augmented Research Methodology - Self-directed training in Claude AI, Gemini CLI, OpenCode AI for scientific research applications',
        'M.Sc. Advanced Training in Genomics, Computational Biology & Molecular Biology - CCSU, Meerut (2024-2026)'
    ]
    for c in certs:
        add_bullet(c)

    # --- 8. Technical Skills ---
    add_heading('TECHNICAL SKILLS')
    table2_data = [
        ('OS & Shell', 'Linux (Ubuntu) - command line, shell scripting, pipeline execution'),
        ('Languages', 'Python, R, Bash'),
        ('Metatranscriptomics', 'SRA Toolkit, FastQC, Trimmomatic, Bowtie2, Kraken2, Bracken, STAR, HISAT2, DESeq2, edgeR'),
        ('AI Tools', 'Claude AI (Anthropic), NeodiSC, AntiGravity AI, OpenCode AI, Gemini CLI'),
        ('Bioinformatics', 'KEGG, NCBI, BioProject, microbiome analysis, NGS workflows, meta-analysis'),
        ('Data Tools', 'Advanced Excel (Certified), Jupyter Notebook, RStudio, ggplot2'),
        ('Communication', 'English (Professional), Hindi (Native)')
    ]
    table2 = doc.add_table(rows=0, cols=2)
    for k, v in table2_data:
        row = table2.add_row()
        r0 = row.cells[0].paragraphs[0].add_run(k)
        r0.bold = True
        row.cells[1].text = v

    # --- 9. Why Vandna Gupta ---
    add_heading('WHY VANDNA GUPTA')
    why_bullets = [
        'One of very few early-career researchers actively building complete metatranscriptomics pipelines AND integrating frontier AI tools (Claude AI, NeodiSC, Gemini CLI) into live cancer genomics research',
        'Hands-on experience with full NGS workflow: from raw SRA data download to differential expression and functional interpretation using KEGG',
        'Co-author of a Springer Nature manuscript at MSc level - demonstrating exceptional research output for career stage',
        'Unique combination: wet-lab biology foundation + metatranscriptomics pipelines + AI tool proficiency + Advanced Excel certification',
        'Fast learner who independently adopts and applies cutting-edge technologies in real research environments'
    ]
    for w in why_bullets:
        add_bullet(w)

    doc.save(output_path)
    print(f"CV successfully generated at {output_path}")

if __name__ == '__main__':
    create_cv()

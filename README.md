# **Quinns Training Generator – AI-Powered Training Material Builder**

##  **Problem Statement**

Training organisations like **Quinns Training Services** face a major challenge:
creating high-quality training materials (PowerPoints, manuals, assessments, outlines) takes **a lot of manual effort**, requires **experienced trainers**, and often leads to:

* inconsistent training modules
* long preparation time
* difficulty updating materials
* duplicate work across teams
* lack of standardisation in training content
* poor scalability as the number of courses grows

Traditional methods also require trainers to manually:

* analyse documents
* extract topics
* create outlines
* design slides
* write assessments
* format training guides

This slows down delivery and reduces consistency.

# **Project Overview**

The **Quinns Training Generator** is a modern AI-powered training-material generation system built using Streamlit and OpenAI.

It allows a user to upload **PDF, PPTX, DOCX, TXT, CSV, Excel, or Markdown files**, and automatically transforms them into:

✔ Structured topics
✔ A full multi-module training outline
✔ PowerPoint presentation
✔ AI-generated trainer guide
✔ AI-generated assessments
✔ A ZIP package for easy distribution

#  **Key Features**

### **1️⃣ Multi-format Document Processing**

Supports PDF, DOCX, PPTX, TXT, CSV, XLSX, and MD files. Extracts text, bullet points, tables, and slide content.
✔ OCR fallback for scanned PDFs
✔ Page/slide-level extraction
✔ Bullet detection

### **2️⃣ AI Topic Identification**

The system uses LLMs to analyse uploaded content and extract:

* Key topics
* Duration estimates
* Importance levels
* Key concepts
* Topic descriptions

Handled using the **TopicAnalyzer** class. 

### **3️⃣ Automatic Training Outline Generator**

Generates a complete learning outline with:

* Modules
* Module objectives
* Key points
* Slide estimates
* Overall program duration

Driven by the **OutlineGenerator** class. 

### **4️⃣ Slide Generation Engine**

AI creates detailed slide content:

* Title slides
* Content slides
* Summary slides
* Teaching notes (200+ words each)

Powered by the **SlideGenerator** class. 

### **5️⃣ Trainer Guide Builder**

Creates a well-formatted Word document containing:

* Program overview
* Module breakdown
* Slide-by-slide trainer instructions
* Teaching notes

Provided through **DocumentBuilder**. 

### **6️⃣ AI-Powered Assessments**

Includes multiple-choice and short-answer questions extracted directly from source documents.
✔ With correct answers
✔ Explanations
✔ Marking guide
✔ Sample answers

### **7️⃣ Beautiful UI With Light & Dark Modes**

The Dynamic Theme System includes:

* Animated gradient UI
* Styled input components
* Colour-coded importance levels
* Phase badges
* Responsive layout

Defined in `get_theme_css()` and UI sections. 

### **8️⃣ 4-Phase Guided Workflow**

#### **Phase 1 – Upload**

Users upload any training document.
Documents are cleaned, extracted, analysed.

#### **Phase 2 – Analyze**

AI detects topics, concepts and creates early structure.

#### **Phase 3 – Edit**

User-friendly editor allows manual refinement:

* reorder modules
* rename topics
* adjust key points
* remove or add modules

#### **Phase 4 – Generate**

Creates:

* PowerPoint
* Trainer Guide
* Assessments
* ZIP export

# 🧠 **Technology Stack**

### **Core Technologies**

* Python
* Streamlit
* OpenAI GPT-4o
* Pandas
* Pytesseract OCR
* pdfplumber
* python-pptx
* python-docx

### **ML Components**

* Topic extraction
* Latent Dirichlet Allocation (LDA) for topic modelling
* NLP preprocessing

### **File Generation**

* PowerPoint (.pptx)
* Word documents (.docx)
* ZIP packaging

# 💡 **How It Works**

### **Step 1 – Upload Files**

You upload your training materials (PDF, PPTX, DOCX…).

### **Step 2 – AI Extracts Information**

The system reads and breaks down content into:

* topics
* concepts
* learning duration
* key sentences
* bullet points

### **Step 3 – You Edit (Optional)**

Modify titles, durations, key points, modules.

### **Step 4 – Generate All Materials**

In one click, the system outputs:

* Full PowerPoint
* Trainer guide
* Assessment pack
* ZIP file


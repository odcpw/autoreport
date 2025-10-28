Below is a technical summary outlining the architecture and modules for the AutoBericht System. This document serves as a reference for developers to understand the vision and framework:

---

# AutoBericht System Architecture

**Overview:**  
A VBA-based solution contained in a single Excel workbook, designed to automate safety culture reporting by integrating Excel, Word, and PowerPoint. It streamlines photo sorting, input processing (self-assessment and moderator data), report generation, and presentation export.

**Architecture Layers:**

- **UI Layer:**  
  - **modUI:** Contains Excel event handlers (button click events) to trigger each major workflow: Photo Sorting, Report Generation, and PPT Export.

- **Controller/Coordinator:**  
  - **modCoordinator:** Acts as the central orchestrator. It sequences tasks, validates inputs, and delegates responsibilities to the domain service modules.

- **Domain Services:**
  - **Photo Processing:**  
    - **clsPhotoSorter (IPhotoProcessor):** Sorts and categorizes photos from a shared directory, tags them, and extracts metadata into a structured Excel format.
  - **Input Processing:**  
    - **clsInputProcessor:** Reads structured customer self-assessment data and moderator findings, mapping them to criteria defined in the master document.
  - **Report Generation:**  
    - **clsReportMaker (IReportGenerator):** Combines photos, moderator comments, and self-assessment scores to generate a draft Word report by dynamically inserting text blocks from a master document.
  - **Presentation Generation:**  
    - **clsPresentationMaker (IPresentationGenerator):** Creates a formatted PowerPoint presentation from key report highlights and visuals.

- **Automation Services:**
  - **clsWordService:** Encapsulates all interactions with Word’s object model, handling document opening, editing, and saving.
  - **clsPPTService:** Manages PowerPoint automation, including slide creation and layout consistency.

- **Utilities:**  
  - **modUtilities:** Provides shared functions for file handling, configuration management, error logging, and error handling routines.

**Coding Standards & Practices:**

- Use modern VBA practices adhering to the Rubberduck style guide.
- Ensure modular, well-commented code with clear separation of concerns.
- Avoid global variables; use structured data passing (custom types or objects) where applicable.
- Implement robust error handling and logging in a centralized manner.
- Isolate Office automation specifics to dedicated classes to ease future modifications.

---

This concise technical overview can be placed in the repository to align developers with the intended design and facilitate further development without significant refactoring.

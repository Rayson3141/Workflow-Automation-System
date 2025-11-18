# Workflow Automation System

This repository contains a workflow automation system built using **Google Apps Script**. The system streamlines the creation, review, and finalization of documents for activities and events.

## Overview
The workflow consists of four main components:

1. **ActivityDoc_Builder**
2. **ActivityForm_Submission**
3. **EventScheduling**
4. **EventDoc_Builder**

Each module automates part of the process, reducing manual work and ensuring consistency.

---

## 1. ActivityDoc_Builder
This script automatically generates the initial activity document based on user inputs.

- Creates a structured document according to predefined templates.
- Saves the generated document into the database for later editing and review.

**Output:** A draft activity document stored in the system.

---

## 2. ActivityForm_Submission
Once the user edits the drafted document, they submit it through this module.

- Provides an interface for users to submit their updated activity documents.
- Forwards the edited document to the admin for verification and approval.

**Purpose:** Ensures the admin can review all changes before the workflow proceeds.

---

## 3. EventScheduling
After the activity document is approved, the workflow moves to the scheduling phase.

- Handles event-related logistics and timeline organization.
- Ensures all activity details are properly scheduled before generating the remaining documents.

**Result:** A finalized event schedule ready for documentation.

---

## 4. EventDoc_Builder
With the schedule confirmed, the remaining required documents can be automatically generated.

- Produces all additional event-related documents using the provided details.
- Reduces repetitive manual work by leveraging templates and automated generation.

**Final Output:** A complete set of event documents.

---

## Workflow Summary
```text
ActivityDoc_Builder → ActivityForm_Submission → EventScheduling → EventDoc_Builder
```

This streamlined system ensures:
- Faster document preparation  
- Lower risk of manual errors  
- Centralized storage and management  

---

## Technologies Used
- **Google Apps Script** for automation  
- **Google Drive / Sheets / Docs** for storage and document manipulation  

---

## Future Improvements
- Add notifications for submission status  
- Integrate version tracking for submitted documents  
- Provide customizable templates for more event types  

---
For any questions or contributions, feel free to open an issue or submit a pull request.



Addtional Reference:
Workflow Automation System Flowchart
![alt text]

Public Database (Experiment Database):
[Activity Database (Google Sheet)](https://docs.google.com/spreadsheets/d/1e4GD2dp9pZwx0wMeYgwavWVZuX0ebjLvoOQp8YaVVR4/edit?usp=sharing)

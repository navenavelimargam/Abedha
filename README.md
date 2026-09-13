# Orion — Sovereign Workbench

> **Thinking locally. Working globally.**

Orion — Sovereign Workbench is a **local, air-gapped enterprise AI assistant** designed to process confidential organizational documents and provide intelligent answers without depending on external cloud AI services.

The system is designed for organizations that require **privacy, local processing, controlled knowledge access, document analysis, team collaboration, and administrative governance**.

---

## 🚀 Key Features

### 🔐 Local / Air-Gapped AI

- Runs AI models locally using Ollama.
- Designed for environments where confidential data should remain inside the organization.
- No requirement to send enterprise documents to external AI APIs.
- Supports locally installed language and vision models.

### 🤖 AI Chatbot

Orion can answer normal user questions and assist with:

- General questions
- Document-based questions
- Summarization
- Data analysis
- Financial analysis
- Technical questions
- Knowledge-base queries
- File-based question answering

Users can ask normal questions even without uploading a file.

---

## 📄 Document Analysis

Orion supports analysis of multiple document formats, including:

- PDF
- DOCX
- PPTX
- CSV
- Excel/XLSX
- TXT
- Markdown
- Images and scanned documents

The system extracts document content locally and uses the appropriate processing pipeline.

For scanned/image-based documents, OCR can be used for local text extraction.

---

## 📊 Data Analysis & Visualization

Orion can analyze structured datasets and generate charts from CSV and Excel files.

Examples include:

- Revenue trends
- Net income trends
- Company comparisons
- Financial performance
- Year-wise analysis
- Employee/data statistics
- Other numerical relationships

Charts can be generated as downloadable PNG files.

### Example

User query:

> Can you create a chart comparing the financial performance of two companies?

Orion can identify comparable financial fields from the uploaded dataset and generate a visualization such as:

**AAPL vs MSFT — Revenue**

---

## 📈 Generated Charts

Charts are displayed directly inside the Orion chat interface.

Users can:

- View the generated chart
- Download the chart as PNG
- Continue asking questions about the uploaded data

The visualization pipeline uses deterministic local data processing for structured CSV/Excel analysis wherever possible.

---

## 🧠 Knowledge Base

Orion includes a local knowledge-base system for organizational documents.

Knowledge-base functionality includes:

- Local document indexing
- Retrieval of relevant information
- Controlled organizational knowledge
- Query-based retrieval
- Offline/local operation

Knowledge-base files are stored locally rather than being sent to an external cloud service.

---

## 👥 Employee Registration & Approval

Orion provides administrator-controlled employee onboarding.

### Employee workflow

1. Employee opens the registration form.
2. Employee provides required details.
3. The account is created with a **Pending** status.
4. A registration request is generated.
5. Administrators receive a notification.
6. Administrator reviews the employee.
7. Administrator approves or rejects the registration.
8. Only an approved employee can sign in and use the workbench.

### Security rule

A pending employee **cannot access the chatbot until administrator approval is completed**.

---

## 👨‍💼 Administrator Center

Administrators have additional privileges for managing the system.

Administrators can:

- Review employee registration requests
- Approve employees
- Reject employees
- Create employees directly
- Manage teams
- Add approved employees to teams
- Upload controlled SOP / knowledge documents
- Manage organizational access

Employees do not receive administrator privileges.

---

## 🏢 Team Chat

Orion supports controlled team-based collaboration.

### Team creation

Only administrators can create team chats.

When creating a team, the administrator can select approved employees from the organization.

The team contains:

- Team owner
- Team members
- Team chat
- Shared conversation history

### Team query attribution

When a member sends a query in a team chat, the query is associated with the **actual employee who submitted it**.

For example:

```text
Raj:
Can you summarize the financial report?

Orion:
The report contains...

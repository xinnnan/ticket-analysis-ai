# Support Ticket Analysis Tool

## Overview

Support Ticket Analysis Tool is a Python-based application designed to analyze support tickets. It extracts ticket data from Excel files, stores the re-organized data in a local SQLite database, and performs both local and OpenAI-driven analysis to generate insights and visual dashboards.

## Features

- **Data Import & Storage:**
  - **Input Format:** Excel files with multiple sheets.
  - **Sheets:** The file must contain the following sheets:
    - 硬件工单
    - 系统工单
    - 服务工单
    - 网络工单
  - **Required Columns:**  
    - 工单编号 (Ticket NO)
    - 工单状态 (Ticket Status)
    - 工单类型 (Ticket Type)
    - ISU项目名称 (Project Name)
    - 标题 (Title)
    - 工单描述 (Description)
    - 处理方法 (Resolve Method)
    - 事件等级 (Level)
    - 响应时长 (Response SLA)
    - 处理时长 (Process Duration)
    - 完成时长 (Complete duration)
  - **Processing:**  
    - Only the required columns are extracted.
    - Chinese column names are mapped to English-compatible names (e.g., 工单编号(Ticket NO) → `ticket_no`).
    - An additional `category` column is added based on the source sheet.
  - **Storage:** Data is stored in a local SQLite database.

- **Data Analysis & Visualization:**
  - **Local Analysis:**
    - Counts tickets per category.
    - Performs text clustering on ticket titles and descriptions using TF-IDF and KMeans.
    - Displays results in a Tkinter-based dashboard with Matplotlib charts.
  - **OpenAI Analysis:**
    - Aggregates a sample of tickets (including unique ticket number, title, and description) and sends them to the OpenAI API.
    - Receives a JSON response containing correlation insights and recommendations.
    - Visualizes the OpenAI analysis results in a separate dashboard.

- **User Interface:**
  - Built with Tkinter.
  - Provides buttons for:
    - Uploading Excel files.
    - Running local analysis.
    - Triggering OpenAI analysis.

## Installation

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/yourusername/ticket-analysis-ai.git
   cd ticket-analysis-ai
   pip3 install -r requirements.txt
   ```
2. **Create or update the config.json file in the project root with your configuration:**

{
  "openai_api_key": "your_openai_api_key_here",
  "db_name": "tickets.db"
}

3. **Run the Application:**
     ```bash
    python3 main.py
     ```
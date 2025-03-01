import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import sqlite3
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans
import openai

# Load configuration from config.json
CONFIG_FILE = "config.json"
if os.path.exists(CONFIG_FILE):
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = json.load(f)
else:
    messagebox.showerror("Configuration Error", f"Missing {CONFIG_FILE} file.")
    exit(1)

# Set OpenAI API key from the configuration
openai.api_key = config.get("openai_api_key", "")
if not openai.api_key:
    messagebox.showerror("Configuration Error", "OpenAI API key not found in config.json")
    exit(1)

# Set the SQLite database name from configuration
DB_NAME = config.get("db_name", "tickets.db")

def create_database():
    """Creates the SQLite database and the tickets table with correct column names."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS tickets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ticket_no TEXT,
        ticket_status TEXT,
        ticket_type TEXT,
        project_name TEXT,
        title TEXT,
        description TEXT,
        resolve_method TEXT,
        level TEXT,
        response_sla TEXT,
        process_duration TEXT,
        complete_duration TEXT,
        category TEXT
    )
    """)
    conn.commit()
    conn.close()

def load_and_store_data(file_path):
    """
    Reads an Excel file with multiple sheets, extracts only the needed columns,
    renames them to English-compatible names, and stores them in the SQLite database.
    """
    sheets = ["硬件工单", "系统工单", "服务工单", "网络工单"]
    
    columns_map = {
        "工单编号(Ticket NO)": "ticket_no",
        "工单状态(Ticket Status)": "ticket_status",
        "工单类型(Ticket Type)": "ticket_type",
        "ISU项目名称(Project Name)": "project_name",
        "标题(Title)": "title",
        "工单描述(Description)": "description",
        "处理方法(Resolve Method)": "resolve_method",
        "事件等级(Level)": "level",
        "响应时长(Response SLA)": "response_sla",
        "处理时长(Process Duration)": "process_duration",
        "完成时长(Complete duration)": "complete_duration"
    }
    data_list = []
    for sheet_name in sheets:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            available_cols = [col for col in columns_map.keys() if col in df.columns]
            df = df[available_cols]
            df = df.rename(columns=columns_map)  # Rename columns to English names
            df["category"] = sheet_name  # Add category column based on the sheet
            data_list.append(df)
        except Exception as e:
            print(f"Error processing sheet {sheet_name}: {e}")
    
    if data_list:
        all_data = pd.concat(data_list, ignore_index=True)
        conn = sqlite3.connect(DB_NAME)
        all_data.to_sql("tickets", conn, if_exists="append", index=False)
        conn.close()
        messagebox.showinfo("Success", "Data loaded and stored successfully!")
    else:
        messagebox.showerror("Error", "No data loaded from file.")

def analyze_data():
    """
    Performs local analysis on the ticket data:
    - Ticket count per category.
    - Text clustering analysis using TF-IDF and KMeans on Title and Description.
    Displays results in a new dashboard window.
    """
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT * FROM tickets", conn)
    conn.close()
    
    if df.empty:
        messagebox.showerror("Error", "No data available for analysis.")
        return
    
    # Use the lowercase "category" column
    category_counts = df["category"].value_counts()
    
    dashboard = tk.Toplevel()
    dashboard.title("Dashboard - Local Analysis")
    
    fig, axs = plt.subplots(2, 1, figsize=(8, 10))
    axs[0].bar(category_counts.index, category_counts.values, color='skyblue')
    axs[0].set_title("Ticket Count per Category")
    axs[0].set_xlabel("Category")
    axs[0].set_ylabel("Count")
    
    # Analysis 2: Text clustering using renamed columns (title and description)
    df['combined_text'] = df['title'].fillna('') + " " + df['description'].fillna('')
    vectorizer = TfidfVectorizer(stop_words='english')
    try:
        X = vectorizer.fit_transform(df['combined_text'])
    except Exception as e:
        X = None
    
    if X is not None and X.shape[0] > 0:
        kmeans = KMeans(n_clusters=3, random_state=42)
        clusters = kmeans.fit_predict(X)
        df['Cluster'] = clusters
        cluster_counts = df['Cluster'].value_counts()
        axs[1].bar(cluster_counts.index.astype(str), cluster_counts.values, color='salmon')
        axs[1].set_title("Ticket Clusters (Based on Title & Description)")
        axs[1].set_xlabel("Cluster")
        axs[1].set_ylabel("Count")
    else:
        axs[1].text(0.5, 0.5, "Not enough text data for clustering", 
                    horizontalalignment='center', verticalalignment='center')
    
    plt.tight_layout()
    canvas = FigureCanvasTkAgg(fig, master=dashboard)
    canvas.draw()
    canvas.get_tk_widget().pack()
    
    summary_text = tk.Text(dashboard, height=10, width=80)
    summary_text.pack(pady=10)
    summary_text.insert(tk.END, "Local Analysis Summary:\n")
    summary_text.insert(tk.END, f"Total Tickets: {len(df)}\n")
    summary_text.insert(tk.END, "Tickets per Category:\n")
    for category, count in category_counts.items():
        summary_text.insert(tk.END, f"  {category}: {count}\n")
    
    summary_text.insert(tk.END, "\nCluster Analysis:\n")
    if 'Cluster' in df.columns:
        for cluster, count in cluster_counts.items():
            summary_text.insert(tk.END, f"  Cluster {cluster}: {count}\n")
    else:
        summary_text.insert(tk.END, "  Clustering not performed.\n")
    
    summary_text.config(state=tk.DISABLED)

def analyze_with_openai():
    """
    Uses the OpenAI API to provide a deeper analysis on the support ticket data.
    This function aggregates ticket numbers, titles, and descriptions and sends them to OpenAI
    with instructions to correlate the tickets by issue type and by related workstation/robot.
    The API is requested to return a structured JSON containing category counts and a summary.
    """
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT * FROM tickets", conn)
    conn.close()
    
    if df.empty:
        messagebox.showerror("Error", "No data available for OpenAI analysis.")
        return

    # Use the renamed columns: ticket_no, title, and description
    sample_texts = df.apply(lambda row: (
        f"Ticket Number: {row['ticket_no']}\n"
        f"Title: {row['title']}\n"
        f"Description: {row['description']}"
    ), axis=1)
    text_sample = "\n\n".join(sample_texts.tolist()[:10])  # Adjust sample size as needed

    # Construct the prompt for correlation analysis
    prompt = (
        "You are an expert in support ticket analysis. Below is a sample of support ticket data. "
        "Each record includes a unique ticket number, title, and description. Your task is to analyze the text and correlate "
        "the tickets by identifying the main issue types and any mentions of specific workstations or robots. Use the ticket "
        "numbers to precisely count how many tickets are correlated with each other.\n\n"
        "Please group the tickets into distinct categories based on these criteria and count the number of tickets in each category. "
        "Also, provide a brief summary of the key findings and recommendations for reducing ticket numbers. \n\n"
        "Return your answer as a valid JSON object with the following structure:\n\n"
        "{\n"
        '  "categories": {\n'
        '      "<Category Name>": <Count>,\n'
        '      "...": ...\n'
        "  },\n"
        '  "summary": "<A short summary of your findings and recommendations>"\n'
        "}\n\n"
        "Here is the sample data:\n"
        f"{text_sample}"
    )
    
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.5,
            max_tokens=400,
        )
        analysis_response = response.choices[0].message['content']
    except Exception as e:
        analysis_response = f"Error in OpenAI API call: {e}"
    
    # Attempt to parse the API response as JSON
    try:
        analysis_data = json.loads(analysis_response)
        categories = analysis_data.get("categories", {})
        summary = analysis_data.get("summary", "No summary provided.")
    except Exception as e:
        categories = {}
        summary = analysis_response

    # Create a dashboard window for OpenAI analysis results
    analysis_window = tk.Toplevel()
    analysis_window.title("OpenAI Analysis Result")
    
    if categories:
        fig, ax = plt.subplots(figsize=(6, 4))
        names = list(categories.keys())
        counts = list(categories.values())
        ax.bar(names, counts, color='mediumseagreen')
        ax.set_title("Ticket Correlation by Issue / Workstation/Robot")
        ax.set_xlabel("Category")
        ax.set_ylabel("Count")
        plt.xticks(rotation=45, ha="right")
        
        canvas = FigureCanvasTkAgg(fig, master=analysis_window)
        canvas.draw()
        canvas.get_tk_widget().pack(pady=10)
    else:
        label = tk.Label(analysis_window, text="No valid category data returned from OpenAI.", fg="red")
        label.pack(pady=10)
    
    summary_label = tk.Label(analysis_window, text="OpenAI Analysis Summary:", font=("Arial", 12, "bold"))
    summary_label.pack(pady=(10, 0))
    result_text = tk.Text(analysis_window, wrap=tk.WORD, height=10, width=100)
    result_text.pack(padx=10, pady=10)
    result_text.insert(tk.END, summary)
    result_text.config(state=tk.DISABLED)

def main():
    """Main function to set up the UI and initialize the database."""
    create_database()
    
    root = tk.Tk()
    root.title("Support Ticket Analysis Tool")
    root.geometry("400x300")
    
    def upload_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file_path:
            load_and_store_data(file_path)
    
    # Button to upload the ticket data file
    upload_button = tk.Button(root, text="Upload Ticket Data File", command=upload_file, width=30)
    upload_button.pack(pady=10)
    
    # Button to perform local analysis and display a dashboard
    analyze_button = tk.Button(root, text="Analyze Data & Show Dashboard", command=analyze_data, width=30)
    analyze_button.pack(pady=10)
    
    # Button to perform deeper analysis with OpenAI and display correlation insights
    openai_button = tk.Button(root, text="Analyze with OpenAI", command=analyze_with_openai, width=30)
    openai_button.pack(pady=10)
    
    root.mainloop()

if __name__ == "__main__":
    main()

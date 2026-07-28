import pandas as pd
from core.etl_builder import build_engagement_batch_data_replacements
from core.pptx_engine import run_batch_presentations

def main():
    # 1. Paths
    TEMPLATE_PATH = "templates/template.pptx"
    
    # 2. Load Raw Data
    print("Loading source data...")
    # Mock dataframes to stand in for your actual CSV/SQL loads
    data_sources = {
        "clients_df": pd.DataFrame({
            "ClientID": ["CLIENT_001", "CLIENT_002"], 
            "Client_Name": ["Acme Corp", "Globex"], 
            "Report_Date_Formatted": ["Q3 2026", "Q3 2026"],
            "Summary_Text": ["Good quarter.", "Needs improvement."]
        }),
        "revenue_df": pd.DataFrame({
            "ClientID": ["CLIENT_001", "CLIENT_001"], 
            "Category": ["SaaS", "Services"], 
            "Amount": ["$1M", "$500k"]
        }),
    }
    
    # List of IDs you want to run right now
    trial_list = ["CLIENT_001", "CLIENT_002"]

    # 3. ETL Pipeline: Build the dictionaries
    batch_configs = build_engagement_batch_data_replacements(
        trial_ids=trial_list,
        data_sources=data_sources,
        stop_on_error=False
    )

    # 4. PPTX Pipeline: Generate the files
    if batch_configs:
        run_batch_presentations(
            template_path=TEMPLATE_PATH, 
            trials=batch_configs
        )
    else:
        print("No successful trials to process. Exiting.")

if __name__ == "__main__":
    main()

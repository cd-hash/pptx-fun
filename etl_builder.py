import pandas as pd
from core.exceptions import TrialBuildError

# --- Extractor Functions ---

def extract_text_placeholders(client_row: pd.Series) -> dict:
    return {
        "{CLIENT_NAME}": client_row.get("Client_Name", "N/A"),
        "{DDMmmYYYY}": client_row.get("Report_Date_Formatted", ""),
        "{EXEC_SUMMARY}": client_row.get("Summary_Text", "No summary provided.")
    }

def extract_revenue_table_config(raw_revenue_df: pd.DataFrame, client_id: str) -> dict:
    client_rev = raw_revenue_df[raw_revenue_df["ClientID"] == client_id].copy()
    if client_rev.empty:
        raise ValueError(f"No revenue data found for ClientID: {client_id}")
    formatted_df = client_rev[["Category", "Amount"]].rename(columns={"Category": "Metric", "Amount": "Total"})
    
    return {
        "slide_title": "Financial Overview",
        "table_name": "Revenue Breakdown",
        "dataframe": formatted_df,
        "include_headers": False,
        "start_row": 1,
        "start_col": 0
    }

def extract_loss_reasons(client_id: str) -> dict:
    """Mocks the extraction of complex hierarchical data for the new feature."""
    # In reality, you'd filter a DataFrame or database query here.
    complex_data = [
        {
            "reason": "Other",
            "sub_reasons": ["Pricing was too high", "Missing critical integration"],
            "count": 150,
            "percent": "25%",
            "delta": "+5%"
        }
    ]
    
    return {
        "slide_title": "Churn Analysis",
        "table_name": "Top Loss Reasons",
        "complex_data": complex_data,
        "start_row": 1
    }

# --- Single Trial Builder ---

def build_single_trial(trial_id: str, client_row: pd.Series, data_sources: dict) -> dict:
    current_stage = "Initialization"
    try:
        current_stage = "Extracting Text"
        text_replace_dict = extract_text_placeholders(client_row)

        current_stage = "Extracting Revenue Table"
        revenue_config = extract_revenue_table_config(data_sources["revenue_df"], trial_id)

        current_stage = "Extracting Complex Loss Reasons"
        loss_reasons_config = extract_loss_reasons(trial_id)

        current_stage = "Formatting Output Path"
        safe_client_name = client_row["Client_Name"].replace(" ", "_")
        output_filename = f"output/Report_{trial_id}_{safe_client_name}.pptx"

        return {
            "trial_id": trial_id,
            "output_path": output_filename,
            "text_replace_dict": text_replace_dict,
            "table_data_configs": [revenue_config],
            "complex_table_configs": [loss_reasons_config] # Added your new feature here
        }

    except Exception as e:
        raise TrialBuildError(f"\n[ERROR] Trial '{trial_id}' failed during stage: '{current_stage}'\nReason: {type(e).__name__} - {str(e)}") from e

# --- Batch Orchestrator ---

def build_engagement_batch_data_replacements(trial_ids: list, data_sources: dict, stop_on_error: bool = False) -> list:
    batch_trials = []
    failed_trials = []

    print(f"Building data configurations for {len(trial_ids)} trials...")
    for trial_id in trial_ids:
        try:
            clients_df = data_sources["clients_df"]
            client_row = clients_df[clients_df["ClientID"] == trial_id].iloc[0]
            trial_config = build_single_trial(trial_id, client_row, data_sources)
            batch_trials.append(trial_config)
        except TrialBuildError as err:
            print(err)
            failed_trials.append(trial_id)
            if stop_on_error: raise err
        except IndexError:
            print(f"\n[ERROR] Trial '{trial_id}' failed: Client ID not found.")
            failed_trials.append(trial_id)
            if stop_on_error: raise

    print(f"\n--- Data Build Summary ---\nSuccessfully built: {len(batch_trials)} / {len(trial_ids)}")
    if failed_trials: print(f"Failed Trials: {failed_trials}")
    return batch_trials

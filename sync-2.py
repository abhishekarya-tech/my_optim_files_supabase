import os
import requests
from supabase import create_client
 
PROJECT_IDS = [
    "5136939842535424",  # RAC
    "5058562934702080",  # FX RAC
    "19926114527",       # ACIMA AMC Merchant Portal
    "5790171803680768",  # ACIMA Core Site
]
 
OPTIMIZELY_TOKEN = os.environ["OPTIMIZELY_TOKEN"]
SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_KEY = os.environ["SUPABASE_KEY"]
 
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
 
 
def get_all_experiments(project_id):
    url = "https://api.optimizely.com/v2/experiments"
 
    headers = {
        "Authorization": f"Bearer {OPTIMIZELY_TOKEN}"
    }
 
    all_experiments = []
    page = 1
    per_page = 100
 
    while True:
        params = {
            "project_id": project_id,
            "page": page,
            "per_page": per_page,
        }
 
        response = requests.get(
            url,
            headers=headers,
            params=params,
            timeout=30,
        )
        response.raise_for_status()
 
        experiments = response.json()
        all_experiments.extend(experiments)
 
        print(
            f"Project {project_id}, page {page}: "
            f"retrieved {len(experiments)} experiments"
        )
 
        if len(experiments) < per_page:
            break
 
        page += 1
 
    return all_experiments
 
 
def map_experiment(experiment):
    return {
        "project_id": str(experiment["project_id"]),
        "experiment_id": str(experiment["id"]),
        "campaign_id": str(experiment["campaign_id"]),
        "campaign_name": experiment["name"],
        "experience_name": experiment["name"],
        "type": experiment["type"],
        "variations": experiment["variations"],
        "traffic_split": experiment["traffic_allocation"],
        "holdback": experiment["holdback"],
        "status": experiment["status"],
        "start_date": experiment.get("created"),
        "end_date": experiment.get("last_modified"),
    }
 
 
if __name__ == "__main__":
    all_experiments = []
 
    for project_id in PROJECT_IDS:
        print(f"Syncing project {project_id}...")
        experiments = get_all_experiments(project_id)
        all_experiments.extend(experiments)
 
    rows = [map_experiment(exp) for exp in all_experiments]
 
    batch_size = 100
 
    for start in range(0, len(rows), batch_size):
        batch = rows[start:start + batch_size]
 
        (
            supabase
            .table("optimizely_experiments")
            .upsert(batch, on_conflict="experiment_id")
            .execute()
        )
 
        print(f"Synced records {start + 1}–{start + len(batch)}")
 
    print(
        f"Successfully synced {len(rows)} experiments "
        f"across {len(PROJECT_IDS)} projects"
    )

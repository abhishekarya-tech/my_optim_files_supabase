import os
import json
from datetime import datetime
from typing import Any

import xlsxwriter
from supabase import create_client

# ---------------------------------------------------------------------------
# EDIT THESE BEFORE RUNNING
# ---------------------------------------------------------------------------
PROJECT_ID = "5790171803680768"       # e.g: ACIMA Core
START_DATE = "2026-06-01T00:00:00Z"   # inclusive
END_DATE = "2026-07-01T00:00:00Z"     # exclusive (so this covers all of June)
# ---------------------------------------------------------------------------

SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_KEY = os.environ["SUPABASE_KEY"]

TABLE_NAME = "optimizely_experiments"
OUTPUT_FILE = f"{PROJECT_ID}_{START_DATE[:7]}.xlsx"

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)


def get_experiments() -> list[dict[str, Any]]:
    """
    Returns experiments for PROJECT_ID that started within
    [START_DATE, END_DATE). Edit the constants above to change scope.
    """
    response = (
        supabase
        .table(TABLE_NAME)
        .select("*")
        .eq("project_id", PROJECT_ID)
        .gte("start_date", START_DATE)
        .lt("start_date", END_DATE)
        .order("start_date")
        .execute()
    )

    return response.data or []


def make_excel_safe(value: Any) -> Any:
    """Convert complex values into readable Excel text."""
    if value is None:
        return ""

    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)

    return value


def export_to_excel(rows: list[dict[str, Any]]) -> None:
    workbook = xlsxwriter.Workbook(OUTPUT_FILE)
    worksheet = workbook.add_worksheet("Experiments")

    header_format = workbook.add_format(
        {
            "bold": True,
            "border": 1,
            "text_wrap": True,
            "valign": "top",
        }
    )

    text_format = workbook.add_format({"num_format": "@", "valign": "top"})
    date_format = workbook.add_format(
        {"num_format": "yyyy-mm-dd hh:mm:ss", "valign": "top"}
    )
    wrap_format = workbook.add_format({"text_wrap": True, "valign": "top"})

    if not rows:
        worksheet.write(0, 0, "No matching records found.")
        workbook.close()
        return

    columns = list(rows[0].keys())

    for col_index, column_name in enumerate(columns):
        worksheet.write(0, col_index, column_name, header_format)

    text_columns = {"experiment_id", "campaign_id", "project_id"}
    date_columns = {"start_date", "end_date"}

    for row_index, row in enumerate(rows, start=1):
        for col_index, column_name in enumerate(columns):
            value = make_excel_safe(row.get(column_name))

            if column_name in text_columns:
                worksheet.write_string(row_index, col_index, str(value), text_format)

            elif column_name in date_columns and value:
                try:
                    parsed_date = datetime.fromisoformat(
                        str(value).replace("Z", "+00:00")
                    )
                    parsed_date = parsed_date.replace(tzinfo=None)
                    worksheet.write_datetime(row_index, col_index, parsed_date, date_format)
                except ValueError:
                    worksheet.write(row_index, col_index, str(value), text_format)

            elif isinstance(row.get(column_name), (dict, list)):
                worksheet.write(row_index, col_index, value, wrap_format)

            else:
                worksheet.write(row_index, col_index, value)

    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, len(rows), len(columns) - 1)

    for col_index, column_name in enumerate(columns):
        if column_name in {"variations", "traffic_split"}:
            worksheet.set_column(col_index, col_index, 45)
        elif column_name in text_columns:
            worksheet.set_column(col_index, col_index, 22)
        elif column_name in date_columns:
            worksheet.set_column(col_index, col_index, 21)
        else:
            worksheet.set_column(col_index, col_index, 18)

    workbook.close()


if __name__ == "__main__":
    experiments = get_experiments()
    export_to_excel(experiments)

    print(
        f"Exported {len(experiments)} rows for project {PROJECT_ID} "
        f"({START_DATE[:10]} to {END_DATE[:10]}) to {OUTPUT_FILE}"
    )

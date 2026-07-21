import os

import json

from datetime import datetime
 
import xlsxwriter

from supabase import create_client
 
 
SUPABASE_URL = os.environ["SUPABASE_URL"]

SUPABASE_KEY = os.environ["SUPABASE_KEY"]
 
TABLE_NAME = "optimizely_experiments"

OUTPUT_FILE = "experiments_2026.xlsx"
 
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
 
 
def get_2026_experiments():

    """

    Returns experiments that:

    1. Started during 2026, or

    2. Started before 2026 and continued into 2026, or

    3. Are still running and started before January 1, 2027.

    """

    response = (

        supabase

        .table(TABLE_NAME)

        .select("*")

        .or_(

            "and(start_date.gte.2026-01-01T00:00:00Z,"

            "start_date.lt.2027-01-01T00:00:00Z),"

            "and(start_date.lt.2026-01-01T00:00:00Z,"

            "end_date.gte.2026-01-01T00:00:00Z),"

            "and(status.eq.running,"

            "start_date.lt.2027-01-01T00:00:00Z)"

        )

        .order("start_date")

        .execute()

    )
 
    return response.data or []
 
 
def make_excel_safe(value):

    if value is None:

        return ""
 
    if isinstance(value, (dict, list)):

        return json.dumps(value, ensure_ascii=False)
 
    return value
 
 
def export_to_excel(rows):

    workbook = xlsxwriter.Workbook(OUTPUT_FILE)

    worksheet = workbook.add_worksheet("2026 Experiments")
 
    header_format = workbook.add_format(

        {

            "bold": True,

            "border": 1,

            "text_wrap": True,

            "valign": "top",

        }

    )
 
    text_format = workbook.add_format(

        {

            "num_format": "@",

            "valign": "top",

        }

    )
 
    date_format = workbook.add_format(

        {

            "num_format": "yyyy-mm-dd hh:mm:ss",

            "valign": "top",

        }

    )
 
    wrap_format = workbook.add_format(

        {

            "text_wrap": True,

            "valign": "top",

        }

    )
 
    if not rows:

        worksheet.write(0, 0, "No matching records found.")

        workbook.close()

        return
 
    preferred_columns = [

        "project_id",

        "experiment_id",

        "campaign_id",

        "campaign_name",

        "experience_name",

        "type",

        "status",

        "start_date",

        "end_date",

        "variations",

        "traffic_split",

        "holdback",

        "web_order_impacted",

        "month",

        "last_synced_at",

    ]
 
    columns = [

        column

        for column in preferred_columns

        if column in rows[0]

    ]
 
    additional_columns = [

        column

        for column in rows[0].keys()

        if column not in columns

    ]
 
    columns.extend(additional_columns)
 
    for col_index, column_name in enumerate(columns):

        worksheet.write(

            0,

            col_index,

            column_name,

            header_format,

        )
 
    text_columns = {

        "project_id",

        "experiment_id",

        "campaign_id",

    }
 
    date_columns = {

        "month",

        "start_date",

        "end_date",

        "last_synced_at",

    }
 
    for row_index, row in enumerate(rows, start=1):

        for col_index, column_name in enumerate(columns):

            original_value = row.get(column_name)

            value = make_excel_safe(original_value)
 
            if column_name in text_columns:

                worksheet.write_string(

                    row_index,

                    col_index,

                    str(value),

                    text_format,

                )
 
            elif column_name in date_columns and value:

                try:

                    parsed_date = datetime.fromisoformat(

                        str(value).replace("Z", "+00:00")

                    )

                    parsed_date = parsed_date.replace(tzinfo=None)
 
                    worksheet.write_datetime(

                        row_index,

                        col_index,

                        parsed_date,

                        date_format,

                    )
 
                except ValueError:

                    worksheet.write_string(

                        row_index,

                        col_index,

                        str(value),

                        text_format,

                    )
 
            elif isinstance(original_value, (dict, list)):

                worksheet.write(

                    row_index,

                    col_index,

                    value,

                    wrap_format,

                )
 
            else:

                worksheet.write(

                    row_index,

                    col_index,

                    value,

                )
 
    worksheet.freeze_panes(1, 0)

    worksheet.autofilter(

        0,

        0,

        len(rows),

        len(columns) - 1,

    )
 
    for col_index, column_name in enumerate(columns):

        if column_name == "project_id":

            worksheet.set_column(col_index, col_index, 20)
 
        elif column_name in {"experiment_id", "campaign_id"}:

            worksheet.set_column(col_index, col_index, 22)
 
        elif column_name in {"campaign_name", "experience_name"}:

            worksheet.set_column(col_index, col_index, 35)
 
        elif column_name == "variations":

            worksheet.set_column(col_index, col_index, 50)
 
        elif column_name in date_columns:

            worksheet.set_column(col_index, col_index, 21)
 
        else:

            worksheet.set_column(col_index, col_index, 18)
 
    workbook.close()
 
 
if __name__ == "__main__":

    experiments = get_2026_experiments()

    export_to_excel(experiments)
 
    print(

        f"Exported {len(experiments)} rows "

        f"to {OUTPUT_FILE}"

    )
 

from datetime import date, datetime
import os
import gspread
import pandas as pd
from app.helpers import get_date_format, normalize_date, normalize_flight_date, normalize_flight_time
from googleapiclient.discovery import build


LOGBOOK_FIXED_ROWS = os.getenv("LOGBOOK_FIXED_ROWS", 2)

class PilotLogBook:
    def __init__(self, credentials, spreadsheet_key: str):
        aircraft_model_sheet_name = "Aircraft model"
        flight_log_glider_sheet_name = "FlightLogGlider"
        summary_glider_sheet_name = "Summary Glider"
        self.credentials = credentials
        self.gc = gspread.authorize(credentials)
        self.spreadsheet_key = spreadsheet_key
        self.document = self.gc.open_by_key(spreadsheet_key)

        # Summary Glider
        self.worksheet_summary_glider = self.document.worksheet(
            summary_glider_sheet_name
        )
        self.pilot_name = self.worksheet_summary_glider.acell("B1").value

        # Instructor
        self.is_instructor = (
            True if self.worksheet_summary_glider.acell("G1").value == "Yes" else False
        )
        _instructor_from_date = self.worksheet_summary_glider.acell("G2").value
        if _instructor_from_date:
            self.instructor_from_date = datetime.strptime(
                _instructor_from_date, "%Y-%m-%d"
            ).date()
        else:
            self.instructor_from_date = None
        
        # Aircraft models
        self.worksheet_aircraft_model = self.document.worksheet(
            aircraft_model_sheet_name
        )
        self.aircraft_models = self.worksheet_aircraft_model.get_all_values()
        self.aircraft_models_to_add = []
        self.aircraft_models_to_add_row_index = None

        # Flight log glider
        self.worksheet_flight_log_glider =self.document.worksheet(flight_log_glider_sheet_name)
        self.flight_log_glider = [
            i for i in self.worksheet_flight_log_glider.get_all_values() if i[0]
        ]
        self.date_format = get_date_format([i[0] for i in self.flight_log_glider[:10]])
        self.flight_log_id_list = [
            self._make_flight_log_id(i) for i in self.flight_log_glider
        ]
        self.flight_log_glider_to_add = []
        self.flight_log_glider_to_add_row_index = max(
            len(self.worksheet_flight_log_glider.col_values(1)), LOGBOOK_FIXED_ROWS
        ) + 1

    def _get_formula(self, key: str, row_index: int):
        formula_dict = {
            "glider_model": f"""=IF(G{row_index}="";"";XLOOKUP(G{row_index};'Aircraft model'!$B$1:$B$1000;'Aircraft model'!$A$1:$A$1000;""))""",
            "total_time_flights": f'=IF(G{row_index}>0;I{row_index}-G{row_index};"")',
            # "pic_time": f"""=IF(K{row_index}='Summary Glider'!$B$1;E{row_index}-C{row_index};"")""",
            # "dual_time": f"""=IF(K{row_index}='Summary Glider'!$B$1;"";IF(L{row_index}='Summary Glider'!$B$1;E{row_index}-C{row_index};""))""",
            # "instructor_time": f"""=IF(M{row_index}=TRUE;E{row_index}-C{row_index};"")""",
            "pic_time": f"""=IF(B{row_index}='Summary Glider'!$B$1,I{row_index}-G{row_index},"")""",
            "dual_time": f"""=IF(B{row_index}='Summary Glider'!$B$1,"",IF(C{row_index}='Summary Glider'!$B$1,I{row_index}-G{row_index},""))""",
            "instructor_time": f"""=IF(M{row_index}=TRUE,I{row_index}-G{row_index},"")""",
        }
        return formula_dict.get(key, "")

    def _make_flight_log_id(self, data: list) -> str:
        d = normalize_flight_date(data[0])
        start_time = normalize_flight_time(data[6])
        lend_time = normalize_flight_time(data[8])
        result = f"{d}{data[5]}{start_time}{data[7]}{lend_time}"
        return result

    def add_aircraft_model(self, model: str, registration: str) -> bool:
        exist_registration_list = [i[1].lower() for i in self.aircraft_models] + [
            i[1].lower() for i in self.aircraft_models_to_add
        ]
        if registration.lower() not in exist_registration_list:
            row_index = len(self.aircraft_models) + 1
            self.aircraft_models_to_add_row_index = (
                self.aircraft_models_to_add_row_index or row_index
            )
            # self.worksheet_aircraft_model.update(f"A{row_index}", [[model, registration]])
            self.aircraft_models_to_add.append([model, registration])
            return True
        return False

    def save_aircraft_model(self):
        if self.aircraft_models_to_add:
            row_index = self.aircraft_models_to_add_row_index
            self.worksheet_aircraft_model.update(
                f"A{row_index}:P{row_index + len(self.aircraft_models_to_add) - 1}",
                self.aircraft_models_to_add,
                value_input_option="USER_ENTERED",
            )
            self.aircraft_models_to_add = []

    def add_flight_log_glider(
        self,
        d: pd.Timestamp,
        departure_place: str,
        departure_time: str,
        arrival_place: str,
        arrival_time: str,
        glider_model: str,
        glider_registration: str,
        type_of_launch: str,
        landings: int,
        name_p1: str,
        name_p2: str,
    ) -> bool:
        row_index = self.flight_log_glider_to_add_row_index + len(
            self.flight_log_glider_to_add
        )
        is_instructor = False
        if (
            self.instructor_from_date is not None
            and d >= pd.Timestamp(self.instructor_from_date)
        ):
            is_instructor = self.is_instructor
        if is_instructor is True and not name_p2:
            is_instructor = False
        if is_instructor is True and name_p2 == self.pilot_name:
            is_instructor = False

        # Date (yyyy-mm-dd)	
        # Name PIC	
        # Name P2	Glider		
        # Departure		
        # Arrival		
        # Total time of flight	
        # Type of launch	
        # Landings	
        # Instructor
        data = [
            normalize_date(d, self.date_format),
            name_p1,
            name_p2,
            # self._get_formula("glider_model", row_index),
            glider_model,
            glider_registration,
            departure_place,
            departure_time,
            arrival_place,
            arrival_time,
            self._get_formula("total_time_flights", row_index),
            type_of_launch,
            landings,
            is_instructor,
            self._get_formula("pic_time", row_index),
            self._get_formula("dual_time", row_index),
            self._get_formula("instructor_time", row_index),
        ]
        flight_log_id = self._make_flight_log_id(data)
        if flight_log_id not in self.flight_log_id_list:
            self.flight_log_glider_to_add.append(data)
            self.flight_log_id_list.append(flight_log_id)
            return True
        return False

    def save_flight_log_glider(self):
        if len(self.flight_log_glider_to_add) > 0:
            current_rows = self.worksheet_flight_log_glider.row_count
            if current_rows < self.flight_log_glider_to_add_row_index:
                self.worksheet_flight_log_glider.add_rows(len(self.flight_log_glider_to_add))
            row_index = self.flight_log_glider_to_add_row_index
            self.worksheet_flight_log_glider.update(
                f"A{row_index}:P{row_index + len(self.flight_log_glider_to_add) - 1}",
                self.flight_log_glider_to_add,
                value_input_option="USER_ENTERED",
            )
            self.flight_log_glider_to_add = []

    def update_filters(self):
        rows_count = self.worksheet_flight_log_glider.row_count + len(self.flight_log_glider_to_add)
        requests = [
            {
                "setBasicFilter": {
                    "filter": {
                        "range": {
                            "sheetId": self.worksheet_flight_log_glider.id,
                            "startRowIndex": 0,
                            "endRowIndex": rows_count,
                            "startColumnIndex": 0,
                            "endColumnIndex": 18
                        }
                    }
                }
            }
        ]

        body = {"requests": requests}
        service = build('sheets', 'v4', credentials=self.credentials)
        response = service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_key,
            body=body
        ).execute()

    def update_tick_boxes(self):
        rows_count = self.worksheet_flight_log_glider.row_count + len(self.flight_log_glider_to_add)
        requests = [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": self.worksheet_flight_log_glider.id,
                        "startRowIndex": LOGBOOK_FIXED_ROWS,
                        "endRowIndex": rows_count,
                        "startColumnIndex": 12,
                        "endColumnIndex": 13
                    },
                    "cell": {
                        "dataValidation": {
                            "condition": {
                                "type": "BOOLEAN"
                            },
                            "showCustomUi": True
                        },
                        "userEnteredValue": {"boolValue": False}
                    },
                    "fields": "dataValidation,userEnteredValue"
                }
            }
        ]
        body = {
            'requests': requests
        }
        service = build('sheets', 'v4', credentials=self.credentials)
        response = service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_key,
            body=body
        ).execute()

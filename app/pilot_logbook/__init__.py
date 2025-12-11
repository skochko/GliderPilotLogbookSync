from datetime import date, datetime

from app.helpers import normalize_flight_date, normalize_flight_time


class PilotLogBook:
    def __init__(self, gc, spreadsheet_key: str):
        aircraft_model_sheet_name = "Aircraft model"
        flight_log_glider_sheet_name = "FlightLogGlider"
        summary_glider_sheet_name = "Summary Glider"
        self.gc = gc
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
        self.worksheet_flight_gog_glider =self.document.worksheet(flight_log_glider_sheet_name)
        self.flight_log_glider = [
            i for i in self.worksheet_flight_gog_glider.get_all_values() if i[0]
        ]
        self.flight_log_id_list = [
            self._make_flight_log_id(i) for i in self.flight_log_glider
        ]
        self.flight_log_glider_to_add = []
        self.flight_log_glider_to_add_row_index = (
            len(self.worksheet_flight_gog_glider.col_values(1)) + 1
        )

    def _get_formula(self, key: str, row_index: int):
        formula_dict = {
            "glider_model": f"""=IF(G{row_index}="";"";XLOOKUP(G{row_index};'Aircraft model'!$B$1:$B$1000;'Aircraft model'!$A$1:$A$1000;""))""",
            "total_time_flights": f'=IF(E{row_index}>0;E{row_index}-C{row_index};"")',
            "pic_time": f"""=IF(K{row_index}='Summary Glider'!$B$1;E{row_index}-C{row_index};"")""",
            "dual_time": f"""=IF(K{row_index}='Summary Glider'!$B$1;"";IF(L{row_index}='Summary Glider'!$B$1;E{row_index}-C{row_index};""))""",
            "instructor_time": f"""=IF(M{row_index}=TRUE;E{row_index}-C{row_index};"")""",
        }
        return formula_dict.get(key, "")

    def _make_flight_log_id(self, data: list) -> str:
        d = normalize_flight_date(data[0])

        start_time = normalize_flight_time(data[2])
        lend_time = normalize_flight_time(data[4])
        result = f"{d}{data[1]}{start_time}{data[3]}{lend_time}"
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
        d: date,
        departure_place: str,
        departure_time: str,
        arrival_place: str,
        arrival_time: str,
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
            and datetime.strptime(d, "%Y-%m-%d").date() >= self.instructor_from_date
        ):
            is_instructor = self.is_instructor
        if is_instructor is True and not name_p2:
            is_instructor = False
        if is_instructor is True and name_p2 == self.pilot_name:
            is_instructor = False
        data = [
            d,
            departure_place,
            departure_time,
            arrival_place,
            arrival_time,
            self._get_formula("glider_model", row_index),
            glider_registration,
            type_of_launch,
            landings,
            self._get_formula("total_time_flights", row_index),
            name_p1,
            name_p2,
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
            current_rows = self.worksheet_flight_gog_glider.row_count
            if current_rows < self.flight_log_glider_to_add_row_index:
                self.worksheet_flight_gog_glider.add_rows(len(self.flight_log_glider_to_add))
            row_index = self.flight_log_glider_to_add_row_index
            self.worksheet_flight_gog_glider.update(
                f"A{row_index}:P{row_index + len(self.flight_log_glider_to_add) - 1}",
                self.flight_log_glider_to_add,
                value_input_option="USER_ENTERED",
            )
            self.flight_log_glider_to_add = []
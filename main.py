from tqdm import tqdm

from access_parser import AccessParser
import pandas as pd

from app.club_members import ClubMembers
from app.db import read_table
from app.pilot_logbook import PilotLogBook
import gspread
from google.oauth2.service_account import Credentials
from app.helpers import (
    normalize_date,
    normalize_time,
)
from dotenv import load_dotenv
import os


load_dotenv()

PLACE_NAME = os.getenv("PLACE_NAME")
DATABASE_PATH = os.getenv("DATABASE_PATH")
DEFAULT_LAUNCH_TYPE = os.getenv("DEFAULT_LAUNCH_TYPE")

SERVICE_ACCOUNT_FILE = "keys.json"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

gc = gspread.authorize(credentials)

print("Load members from Members.xlsx")
club_members = ClubMembers("Members.xlsx")
print("Loaded members", len(club_members.members))
print(f"Load database from {DATABASE_PATH}...")

# print("Reading tables")
# table_flight_time_dict = db.parse_table("tblFlightTime")
# table_glider_daetails_dict = db.parse_table("tblGliderDetails")
# table_glider_type_dict = db.parse_table("TblGliderType")
# table_member_dict = db.parse_table("tblMember")

# glider_daetails_dict = dict(
#     zip(table_glider_daetails_dict["AutoID"], table_glider_daetails_dict["GliderID"])
# )
# glider_dict = dict(
#     zip(table_glider_type_dict["TypeId"], table_glider_type_dict["GliderType"])
# )
# member_dict = dict(zip(table_member_dict["MemberID"], table_member_dict["Name"]))

# df_flight_time = pd.DataFrame(table_flight_time_dict)

print("Reading tables")
table_flight_time_dict = read_table("tblFlightTime")
table_glider_daetails_dict = read_table("tblGliderDetails")
table_glider_type_dict = read_table("TblGliderType")
table_member_dict = read_table("tblMember")

glider_daetails_dict = dict(
    zip(table_glider_daetails_dict["AutoID"], table_glider_daetails_dict["GliderID"])
)
glider_dict = dict(
    zip(table_glider_type_dict["TypeId"], table_glider_type_dict["GliderType"])
)
member_dict = dict(zip(table_member_dict["MemberID"], table_member_dict["Name"]))
df_flight_time = pd.DataFrame(table_flight_time_dict)


df_flight_time["DateFlown"] = pd.to_datetime(
    df_flight_time["DateFlown"], errors="coerce"
)

pbar = tqdm(club_members.members, total=len(club_members.members))
for member in pbar:
    pbar.set_description(f"Sync Logbook for {member.name}")
    if member.sync_date:
        pilot_flights = df_flight_time[
            (
                (df_flight_time["P1"] == member.club_id)
                | (df_flight_time["P2"] == member.club_id)
            )
            & (df_flight_time["DateFlown"] >= pd.Timestamp(member.sync_date))
        ]
    else:
        pilot_flights = df_flight_time[
            (df_flight_time["P1"] == member.club_id)
            | (df_flight_time["P2"] == member.club_id)
        ]
    # Sorting by DateFlown, LaunchTime, LandTime
    pilot_flights = pilot_flights.sort_values(
        by=["DateFlown", "LaunchTime", "LandTime"], ascending=[True, True, True]
    )
    pilog_log_book = PilotLogBook(gc, member.spreadsheet_key)
    count = 0
    _sync_date = member.sync_date
    for _, row in pilot_flights.iterrows():
    # for _, row in tqdm(pilot_flights.iterrows(), total=len(pilot_flights), desc=f"Sync {member.name}"):
        glider_registration = glider_daetails_dict.get(row["GliderID"])
        glider_model = glider_dict.get(row["GliderType"])
        p1_name = member_dict.get(row["P1"])
        if pd.isna(row["P2"]):
            p2_name = ""
        else:
            p2_name = member_dict.get(row["P2"], "<hidden>")

        pilog_log_book.add_aircraft_model(glider_model, glider_registration)
        added = pilog_log_book.add_flight_log_glider(
            normalize_date(row["DateFlown"]),
            PLACE_NAME,
            normalize_time(row["LaunchTime"]),
            PLACE_NAME,
            normalize_time(row["LandTime"]),
            glider_model,
            glider_registration,
            DEFAULT_LAUNCH_TYPE,
            1,
            p1_name,
            p2_name,
        )
        _sync_date = row["DateFlown"].date()
        if added is True:
            count += 1
    tqdm.write(f"Added {count} rows for {member.name}")
    if count > 0:
        tqdm.write(
            f"Save {count} flight log and aircraft models for {member.name} - processing..."
        )
        try:
            pilog_log_book.save_aircraft_model()
            pilog_log_book.save_flight_log_glider()
            member.sync_date = _sync_date
            tqdm.write(
                f"Save {count} flight log and aircraft models for {member.name} - saved"
            )
        except Exception as e:
            tqdm.write(
                f"Save {count} flight log and aircraft models for {member.name} - error ({e})"
            )
        tqdm.write(f"Sync date for {member.name} has been updated to {_sync_date}")
        club_members.save()

print("All steps done.")

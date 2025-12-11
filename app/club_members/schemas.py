from datetime import date
from typing import Optional
from pydantic import BaseModel


class ClubMemberSchema(BaseModel):
    club_id: int
    name: str
    spreadsheet_key: str
    sync_date: Optional[date] = None